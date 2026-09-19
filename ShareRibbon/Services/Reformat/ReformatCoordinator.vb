' ShareRibbon\Services\Reformat\ReformatCoordinator.vb
' 排版协调器 - 编排"临时文档 → OpenXML排版 → 预览 → 合并"完整流程
' 核心设计：所有排版操作在临时文档上进行，原文档不受影响

Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq

''' <summary>
''' 排版操作类型
''' </summary>
Public Enum ReformatOperationType
    Template
    Semantic
    Smart
End Enum

''' <summary>
''' 用户决策
''' </summary>
Public Enum ReformatUserDecision
    Accept
    Reject
    SaveAs
End Enum

''' <summary>
''' 排版结果
''' </summary>
Public Class ReformatResult
    Public Property Success As Boolean = False
    Public Property ModifiedCount As Integer = 0
    Public Property ErrorMessage As String = ""
    Public Property TempDocPath As String = ""
    Public Property GeneratedInstructions As List(Of Instruction)
    Public Property ExecutionResult As ExecutionResult
End Class

''' <summary>
''' 排版协调器 - 编排整个排版流程
'''
''' 流程：
''' 1. 创建临时文档副本
''' 2. AI语义标注
''' 3. 构建DSL指令
''' 4. OpenXML执行排版（在临时文档上）
''' 5. 启动Word预览临时文档
''' 6. 用户确认后，在原文档上重放修改（支持UndoRecord撤销）
''' </summary>
Public Class ReformatCoordinator

    ''' <summary>
    ''' 执行完整排版流程
    ''' </summary>
    Public Async Function ExecuteReformatPipelineAsync(
        sourceDoc As Object,
        operationType As ReformatOperationType,
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping,
        Optional templateName As String = "",
        Optional sourceParagraphs As List(Of Object) = Nothing,
        Optional sourceParagraphTypes As List(Of String) = Nothing) As Task(Of ReformatResult)

        Dim result As New ReformatResult()
        Dim tempDocPath As String = Nothing

        Try
            ' === Step 1: 创建临时文档 ===
            tempDocPath = TempDocumentService.CreateTempDocument(sourceDoc)
            result.TempDocPath = tempDocPath

            ' === Step 2: 构建DSL指令集 ===
            Dim instructions = BuildInstructions(taggedParagraphs, mapping, sourceDoc, sourceParagraphs)
            result.GeneratedInstructions = instructions

            If instructions.Count = 0 Then
                result.ErrorMessage = "Не сформировано ни одной инструкции форматирования"
                TempDocumentService.Cleanup(tempDocPath)
                Return result
            End If

            ' === Step 3: OpenXML在临时文档上执行排版 ===
            Dim execResult = OpenXmlReformatEngine.ExecuteInstructions(tempDocPath, instructions)
            result.ExecutionResult = execResult
            result.ModifiedCount = execResult.SuccessCount

            If execResult.FailureCount > 0 Then
                Debug.WriteLine($"[ReformatCoordinator] {execResult.FailureCount} 条指令执行失败")
            End If

            ' === Step 4: 用户预览确认 ===
            Dim userDecision = Await ShowPreviewAsync(sourceDoc, tempDocPath, templateName, execResult)

            Select Case userDecision
                Case ReformatUserDecision.Accept
                    ' 在原文档上重放修改（使用UndoRecord包装，支持Ctrl+Z撤销）
                    Dim mergeCount = MergeToSourceDocument(sourceDoc, taggedParagraphs, mapping, sourceParagraphs, sourceParagraphTypes)
                    result.Success = True
                    result.ModifiedCount = mergeCount
                    GlobalStatusStrip.ShowInfo($"Форматирование завершено: изменено абзацев — {mergeCount}")

                Case ReformatUserDecision.SaveAs
                    ' 另存为新文档
                    Dim savePath = PromptSaveAs(tempDocPath)
                    If Not String.IsNullOrEmpty(savePath) Then
                        result.Success = True
                        result.TempDocPath = savePath
                        GlobalStatusStrip.ShowInfo($"Сохранено как: {Path.GetFileName(savePath)}")
                    Else
                        result.Success = False
                        result.ErrorMessage = "Пользователь отменил сохранение"
                    End If

                Case ReformatUserDecision.Reject
                    result.Success = False
                    result.ErrorMessage = "Пользователь отменил форматирование"
                    GlobalStatusStrip.ShowInfo("Форматирование отменено")
            End Select

        Catch ex As Exception
            result.Success = False
            result.ErrorMessage = $"Ошибка процесса форматирования: {ex.Message}"
            Debug.WriteLine($"[ReformatCoordinator] 异常: {ex}")
        Finally
            ' 清理临时文件（如果用户不是"另存为"）
            If tempDocPath IsNot Nothing AndAlso
               result.TempDocPath = tempDocPath AndAlso
               File.Exists(tempDocPath) Then
                TempDocumentService.Cleanup(tempDocPath)
            End If
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 构建完整的DSL指令集
    ''' </summary>
    Private Function BuildInstructions(
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping,
        sourceDoc As Object,
        Optional sourceParagraphs As List(Of Object) = Nothing) As List(Of Instruction)

        Dim instructions As New List(Of Instruction)()

        ' 构建 tagId → SemanticTag 查找字典
        Dim tagDict As New Dictionary(Of String, SemanticTag)()
        For Each tag In mapping.SemanticTags
            If Not tagDict.ContainsKey(tag.TagId) Then
                tagDict(tag.TagId) = tag
            End If
        Next

        ' 收集段落用于构建指令时的rollback信息（优先使用传入的段落列表）
        Dim wordParagraphs As List(Of Object) = sourceParagraphs
        If wordParagraphs Is Nothing Then
            Try
                wordParagraphs = CollectParagraphs(sourceDoc)
            Catch
            End Try
        End If

        For Each tagged In taggedParagraphs
            Dim semanticTag = FindTagWithFallback(tagged.TagId, tagDict, mapping)
            If semanticTag Is Nothing Then Continue For

            Dim para As Object = Nothing
            If wordParagraphs IsNot Nothing AndAlso
               tagged.ParaIndex >= 0 AndAlso tagged.ParaIndex < wordParagraphs.Count Then
                para = wordParagraphs(tagged.ParaIndex)
            End If

            ' 段落样式指令
            If semanticTag.Paragraph IsNot Nothing Then
                Dim pInstr = BuildParagraphStyleInstruction(tagged.ParaIndex, semanticTag.Paragraph, para)
                If pInstr IsNot Nothing Then instructions.Add(pInstr)
            End If

            ' 字符格式指令
            If semanticTag.Font IsNot Nothing OrElse semanticTag.Color IsNot Nothing Then
                Dim cInstr = BuildCharacterFormatInstruction(tagged.ParaIndex, semanticTag, para)
                If cInstr IsNot Nothing Then instructions.Add(cInstr)
            End If
        Next

        Return instructions
    End Function

    ''' <summary>
    ''' 构建段落样式指令（带rollback信息）
    ''' </summary>
    Private Function BuildParagraphStyleInstruction(
        paraIndex As Integer,
        paraCfg As ParagraphConfig,
        para As Object) As Instruction

        Dim target As New JObject()
        target("type") = "paraIndex"
        target("index") = paraIndex

        Dim params As New JObject()
        Dim rollback As New JObject()

        Try
            If para IsNot Nothing Then
                Dim rng = para.Range
                ' 对齐
                If Not String.IsNullOrEmpty(paraCfg.Alignment) Then
                    params("alignment") = paraCfg.Alignment.ToLower()
                    Try : rollback("alignment") = CInt(rng.ParagraphFormat.Alignment) : Catch : End Try
                End If
                ' 首行缩进
                If paraCfg.FirstLineIndent > 0 Then
                    params("firstLineIndent") = paraCfg.FirstLineIndent
                    Try : rollback("firstLineIndent") = CSng(rng.ParagraphFormat.FirstLineIndent) : Catch : End Try
                End If
                ' 行距
                If paraCfg.LineSpacing > 0 Then
                    params("lineSpacing") = paraCfg.LineSpacing
                    Try : rollback("lineSpacingRule") = CInt(rng.ParagraphFormat.LineSpacingRule) : Catch : End Try
                    Try : rollback("lineSpacing") = CSng(rng.ParagraphFormat.LineSpacing) : Catch : End Try
                End If
                ' 段前
                If paraCfg.SpaceBefore > 0 Then
                    params("spaceBefore") = paraCfg.SpaceBefore
                    Try : rollback("spaceBefore") = CSng(rng.ParagraphFormat.SpaceBefore) : Catch : End Try
                End If
                ' 段后
                If paraCfg.SpaceAfter > 0 Then
                    params("spaceAfter") = paraCfg.SpaceAfter
                    Try : rollback("spaceAfter") = CSng(rng.ParagraphFormat.SpaceAfter) : Catch : End Try
                End If
                ' 孤行控制
                If paraCfg.KeepWithNext Then
                    params("keepWithNext") = True
                    Try : rollback("keepWithNext") = CBool(rng.ParagraphFormat.KeepWithNext) : Catch : End Try
                End If
                ' 段前分页
                If paraCfg.PageBreakBefore Then
                    params("pageBreakBefore") = True
                    Try : rollback("pageBreakBefore") = CBool(rng.ParagraphFormat.PageBreakBefore) : Catch : End Try
                End If
                ' 段中不分页
                If paraCfg.KeepLinesTogether Then
                    params("keepLinesTogether") = True
                    Try : rollback("keepLinesTogether") = CBool(rng.ParagraphFormat.KeepLinesTogether) : Catch : End Try
                End If
            Else
                ' 无para对象时，仅使用配置值（无rollback）
                If Not String.IsNullOrEmpty(paraCfg.Alignment) Then params("alignment") = paraCfg.Alignment.ToLower()
                If paraCfg.FirstLineIndent > 0 Then params("firstLineIndent") = paraCfg.FirstLineIndent
                If paraCfg.LineSpacing > 0 Then params("lineSpacing") = paraCfg.LineSpacing
                If paraCfg.SpaceBefore > 0 Then params("spaceBefore") = paraCfg.SpaceBefore
                If paraCfg.SpaceAfter > 0 Then params("spaceAfter") = paraCfg.SpaceAfter
                If paraCfg.KeepWithNext Then params("keepWithNext") = True
                If paraCfg.PageBreakBefore Then params("pageBreakBefore") = True
                If paraCfg.KeepLinesTogether Then params("keepLinesTogether") = True
            End If
        Catch ex As Exception
            Debug.WriteLine($"BuildParagraphStyleInstruction 异常: {ex.Message}")
        End Try

        If params.Count = 0 Then Return Nothing

        Dim instr As New Instruction("setParagraphStyle", params, Nothing)
        instr.Target = target
        instr.Rollback = rollback
        Return instr
    End Function

    ''' <summary>
    ''' 构建字符格式指令（带rollback信息）
    ''' </summary>
    Private Function BuildCharacterFormatInstruction(
        paraIndex As Integer,
        tag As SemanticTag,
        para As Object) As Instruction

        Dim target As New JObject()
        target("type") = "paraIndex"
        target("index") = paraIndex

        Dim params As New JObject()
        Dim rollback As New JObject()

        Try
            If para IsNot Nothing Then
                Dim rng = para.Range
                If tag.Font IsNot Nothing Then
                    If Not String.IsNullOrEmpty(tag.Font.FontNameCN) Then
                        params("fontNameCN") = tag.Font.FontNameCN
                        Try : rollback("fontNameFarEast") = If(rng.Font.NameFarEast?.ToString(), "") : Catch : End Try
                    End If
                    If Not String.IsNullOrEmpty(tag.Font.FontNameEN) Then
                        params("fontNameEN") = tag.Font.FontNameEN
                        Try : rollback("fontName") = If(rng.Font.Name?.ToString(), "") : Catch : End Try
                    End If
                    If tag.Font.FontSize > 0 Then
                        params("fontSize") = tag.Font.FontSize
                        Try : rollback("fontSize") = CSng(rng.Font.Size) : Catch : End Try
                    End If
                    params("bold") = tag.Font.Bold
                    Try : rollback("bold") = CInt(rng.Font.Bold) : Catch : End Try
                    params("italic") = tag.Font.Italic
                    Try : rollback("italic") = CInt(rng.Font.Italic) : Catch : End Try
                    params("underline") = tag.Font.Underline
                    Try : rollback("underline") = CInt(rng.Font.Underline) : Catch : End Try
                End If
                If tag.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(tag.Color.FontColor) Then
                    params("fontColor") = tag.Color.FontColor.TrimStart("#"c)
                    Try : rollback("fontColor") = CInt(rng.Font.Color) : Catch : End Try
                End If
            Else
                ' 无para对象时，仅使用配置值
                If tag.Font IsNot Nothing Then
                    If Not String.IsNullOrEmpty(tag.Font.FontNameCN) Then params("fontNameCN") = tag.Font.FontNameCN
                    If Not String.IsNullOrEmpty(tag.Font.FontNameEN) Then params("fontNameEN") = tag.Font.FontNameEN
                    If tag.Font.FontSize > 0 Then params("fontSize") = tag.Font.FontSize
                    params("bold") = tag.Font.Bold
                    params("italic") = tag.Font.Italic
                    params("underline") = tag.Font.Underline
                End If
                If tag.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(tag.Color.FontColor) Then
                    params("fontColor") = tag.Color.FontColor.TrimStart("#"c)
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"BuildCharacterFormatInstruction 异常: {ex.Message}")
        End Try

        If params.Count = 0 Then Return Nothing

        Dim instr As New Instruction("setCharacterFormat", params, Nothing)
        instr.Target = target
        instr.Rollback = rollback
        Return instr
    End Function

    ''' <summary>
    ''' 查找标签（精确匹配 → 父级回退）
    ''' </summary>
    Private Shared Function FindTagWithFallback(
        tagId As String,
        tagDict As Dictionary(Of String, SemanticTag),
        mapping As SemanticStyleMapping) As SemanticTag

        If tagDict.ContainsKey(tagId) Then Return tagDict(tagId)

        Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
        If Not String.IsNullOrEmpty(parentId) AndAlso tagDict.ContainsKey(parentId) Then
            Return tagDict(parentId)
        End If

        Return mapping.FindTag(tagId)
    End Function

    ''' <summary>
    ''' 收集原文档段落列表（用于构建rollback快照）
    ''' </summary>
    Private Shared Function CollectParagraphs(sourceDoc As Object) As List(Of Object)
        Dim list As New List(Of Object)()
        Try
            For Each para As Object In sourceDoc.Paragraphs
                list.Add(para)
            Next
        Catch
        End Try
        Return list
    End Function

    ''' <summary>
    ''' 显示预览对话框，返回用户决策
    ''' </summary>
    Private Async Function ShowPreviewAsync(
        sourceDoc As Object,
        tempDocPath As String,
        templateName As String,
        execResult As ExecutionResult) As Task(Of ReformatUserDecision)

        ' 在UI线程上显示对话框
        Dim tcs As New TaskCompletionSource(Of ReformatUserDecision)()

        Dim showAction As New Action(
            Sub()
                Dim docName As String = ""
                Try
                    docName = Path.GetFileNameWithoutExtension(sourceDoc.FullName)
                Catch
                End Try

                Dim msg = $"Предпросмотр форматирования{vbCrLf}{vbCrLf}" &
                          $"Документ: {docName}{vbCrLf}" &
                          $"Шаблон: {templateName}{vbCrLf}" &
                          $"Успешно: {execResult.SuccessCount} инструкций{vbCrLf}" &
                          $"Ошибок: {execResult.FailureCount} инструкций{vbCrLf}{vbCrLf}" &
                          $"Будет запущен Word и открыт временный документ для предпросмотра.{vbCrLf}" &
                          $"После подтверждения форматирование будет применено к исходному документу (отмена по Ctrl+Z)."

                If execResult.Errors.Count > 0 Then
                    msg &= $"{vbCrLf}{vbCrLf}Предупреждение:{vbCrLf}{String.Join(vbCrLf, execResult.Errors.Take(3))}"
                End If

                ' 在当前 Word 进程中打开临时文档供预览（避免新进程的文件占用冲突）
                Dim tempDoc As Object = Nothing
                Try
                    Dim wordApp = sourceDoc.Application
                    tempDoc = wordApp.Documents.Open(
                        tempDocPath,
                        [ReadOnly]:=False,
                        Visible:=True)
                    If tempDoc IsNot Nothing Then
                        tempDoc.Activate()
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ReformatCoordinator] 在当前Word中打开临时文档失败: {ex.Message}")
                End Try

                Dim result = MessageBox.Show(
                    Nothing,
                    msg,
                    "Предпросмотр форматирования ИИ",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1)

                Select Case result
                    Case DialogResult.Yes
                        tcs.SetResult(ReformatUserDecision.Accept)
                    Case DialogResult.No
                        tcs.SetResult(ReformatUserDecision.Reject)
                    Case DialogResult.Cancel
                        tcs.SetResult(ReformatUserDecision.SaveAs)
                    Case Else
                        tcs.SetResult(ReformatUserDecision.Reject)
                End Select

                ' 根据用户决策处理临时文档：接受/拒绝时关闭；另存为时保留
                ' decision already known from MessageBox above — do not block on tcs.Task.Result
                Try
                    If tempDoc IsNot Nothing Then
                        Dim decision As ReformatUserDecision
                        Select Case result
                            Case DialogResult.Yes
                                decision = ReformatUserDecision.Accept
                            Case DialogResult.No
                                decision = ReformatUserDecision.Reject
                            Case DialogResult.Cancel
                                decision = ReformatUserDecision.SaveAs
                            Case Else
                                decision = ReformatUserDecision.Reject
                        End Select
                        If decision = ReformatUserDecision.Accept OrElse decision = ReformatUserDecision.Reject Then
                            tempDoc.Close(SaveChanges:=False)
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ReformatCoordinator] 关闭临时文档失败: {ex.Message}")
                End Try
            End Sub)

        If System.Windows.Forms.Form.ActiveForm IsNot Nothing AndAlso
           System.Windows.Forms.Form.ActiveForm.InvokeRequired Then
            System.Windows.Forms.Form.ActiveForm.BeginInvoke(showAction)
        Else
            showAction()
        End If

        Return Await tcs.Task
    End Function

    ''' <summary>
    ''' 提示用户另存为
    ''' </summary>
    Private Function PromptSaveAs(tempDocPath As String) As String
        Using dlg As New SaveFileDialog()
            dlg.Filter = "Документ Word (*.docx)|*.docx"
            dlg.FileName = Path.GetFileNameWithoutExtension(tempDocPath) & "_отформатировано"
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    File.Copy(tempDocPath, dlg.FileName, overwrite:=True)
                    Return dlg.FileName
                Catch ex As Exception
                    MessageBox.Show($"Не удалось сохранить: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
        Return ""
    End Function

    ''' <summary>
    ''' 将排版修改合并回原文档（使用UndoRecord包装，支持Ctrl+Z撤销）
    ''' </summary>
    Private Function MergeToSourceDocument(
        sourceDoc As Object,
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping,
        Optional sourceParagraphs As List(Of Object) = Nothing,
        Optional sourceParagraphTypes As List(Of String) = Nothing) As Integer

        Dim appliedCount As Integer = 0

        Dim wordApp As Object = Nothing
        Dim screenUpdatingChanged As Boolean = False
        Dim undoRecordStarted As Boolean = False

        Try
            wordApp = sourceDoc.Application
            wordApp.ScreenUpdating = False
            screenUpdatingChanged = True
            Try
                wordApp.UndoRecord.StartCustomRecord("Форматирование ИИ")
                undoRecordStarted = True
            Catch ex As Exception
                Debug.WriteLine($"[ReformatCoordinator] StartCustomRecord failed: {ex.Message}")
            End Try

            ' 使用传入的段落列表（保证索引一致），否则回退到全文收集
            Dim wordParagraphs As List(Of Object) = sourceParagraphs
            Dim paragraphTypes As List(Of String) = sourceParagraphTypes

            If wordParagraphs Is Nothing Then
                wordParagraphs = New List(Of Object)()
                paragraphTypes = New List(Of String)()
                For Each para As Object In sourceDoc.Paragraphs
                    wordParagraphs.Add(para)
                    paragraphTypes.Add("text")
                Next
            End If

            If paragraphTypes Is Nothing Then
                paragraphTypes = Enumerable.Repeat("text", wordParagraphs.Count).ToList()
            End If

            Dim result = SemanticRenderingEngine.ApplySemanticFormatting(
                taggedParagraphs, mapping, wordParagraphs, paragraphTypes, wordApp)

            appliedCount = result.AppliedCount

        Catch ex As Exception
            Debug.WriteLine($"[ReformatCoordinator] 合并到原文档失败: {ex}")
            Throw
        Finally
            If undoRecordStarted AndAlso wordApp IsNot Nothing Then
                Try
                    wordApp.UndoRecord.EndCustomRecord()
                Catch ex As Exception
                    Debug.WriteLine($"[ReformatCoordinator] EndCustomRecord failed: {ex.Message}")
                End Try
            End If

            If screenUpdatingChanged AndAlso wordApp IsNot Nothing Then
                Try
                    wordApp.ScreenUpdating = True
                Catch
                End Try
            End If
        End Try

        Return appliedCount
    End Function

End Class
