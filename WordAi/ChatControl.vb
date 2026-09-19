Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Math
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Markdig
Imports ShareRibbon.Extensions
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Imports DocumentFormat.OpenXml.Packaging

Public Class ChatControl
    Inherits BaseChatControl



    ' 排版上下文：存储待格式化的段落和样式信息
    Private _reformatParagraphs As List(Of Object) = Nothing
    Private _reformatStyles As List(Of String) = Nothing
    Private _reformatTypes As List(Of String) = Nothing ' text/image/table/formula
    Private _reformatMapping As SemanticStyleMapping = Nothing ' 语义映射上下文
    Private _reformatRetryCount As Integer = 0 ' 重试计数器，防止死循环
    Private Const MAX_REFORMAT_RETRIES As Integer = 2 ' 最大重试次数
    Private _mirrorFormatDocName As String = "" ' 格式克隆时记录源文档名

    ' 智能排版 v2 字段
    Private _formatterAgent As ChatFormatterAgent = Nothing
    Private _activeReformatPlan As ReformatPreviewPlan = Nothing
    Private _activeReformatJob As ReformatJob = Nothing
    Private _activeWordFormattingTaskPlan As Services.WordFormattingTaskPlan = Nothing
    Private _reformatVariantIndex As Integer = 0

    ' 排版撤销快照
    Private _reformatSnapshot As SemanticRenderingEngine.ReformatSnapshot = Nothing
    Private _reformatSnapshotParagraphs As List(Of Object) = Nothing
    Private _reformatSnapshotTypes As List(Of String) = Nothing

    ''' <summary>
    ''' 设置排版上下文，用于规则匹配后应用格式
    ''' </summary>
    Public Sub SetReformatContext(paragraphs As List(Of Object), styles As List(Of String), Optional types As List(Of String) = Nothing, Optional mapping As SemanticStyleMapping = Nothing)
        _reformatParagraphs = paragraphs
        _reformatStyles = styles
        _reformatTypes = types
        _reformatMapping = mapping
        _reformatRetryCount = 0 ' 重置重试计数器
    End Sub

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overrides Function GetOfficeApplicationName() As String
        Return "Word"
    End Function

    ''' <summary>
    ''' 处理模板预览请求（Word 特定实现）
    ''' </summary>
    Protected Overrides Sub HandlePreviewTemplateInWord(jsonDoc As Newtonsoft.Json.Linq.JObject)
        Dim templateId As String = jsonDoc("templateId")?.ToString()
        If String.IsNullOrEmpty(templateId) Then
            GlobalStatusStrip.ShowWarning("ID шаблона не может быть пустым")
            Return
        End If

        ' docx映射卡片：直接打开关联的.docx文件预览
        If templateId.StartsWith("docx_") Then
            Dim mappingId = templateId.Substring(5)
            Dim mapping = SemanticMappingManager.Instance.GetMappingById(mappingId)
            If mapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("Семантическое сопоставление не существует")
                Return
            End If
            If String.IsNullOrEmpty(mapping.SourceFilePath) OrElse Not IO.File.Exists(mapping.SourceFilePath) Then
                GlobalStatusStrip.ShowWarning("Исходный файл шаблона утерян, загрузите его заново")
                Return
            End If

            ' 直接用系统默认方式打开文档（新Word实例）
            Try
                Process.Start(mapping.SourceFilePath)
                GlobalStatusStrip.ShowInfo($"Открыт предпросмотр документа шаблона: {mapping.Name}")
            Catch ex As Exception
                GlobalStatusStrip.ShowWarning($"Не удалось открыть документ: {ex.Message}")
            End Try
            Return
        End If

        ' 常规模板预览：先保存为临时文件，然后用系统默认方式打开
        Dim template As ReformatTemplate = ReformatTemplateManager.Instance.GetTemplateById(templateId)
        If template Is Nothing Then
            GlobalStatusStrip.ShowWarning($"Шаблон с ID {templateId} не найден")
            Return
        End If

        Try
            ' 创建临时文件
            Dim tempPath = IO.Path.Combine(IO.Path.GetTempPath(), $"模板预览_{template.Name}_{DateTime.Now:yyyyMMddHHmmss}.docx")

            ' 使用当前Word实例创建临时文档
            Dim currentApp = Globals.ThisAddIn.Application
            Dim tempDoc = currentApp.Documents.Add()

            ApplyTemplateToDocument(tempDoc, template)

            ' 保存并关闭临时文档
            tempDoc.SaveAs2(tempPath)
            tempDoc.Close(SaveChanges:=False)

            ' 用系统默认方式打开（新Word实例）
            Process.Start(tempPath)

            GlobalStatusStrip.ShowInfo($"Открыт предпросмотр шаблона: {template.Name}")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"Не удалось показать предпросмотр шаблона: {ex.Message}")
            Debug.WriteLine($"HandlePreviewTemplateInWord 错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 将模板应用到Word文档（预览用）
    ''' </summary>
    Private Sub ApplyTemplateToDocument(doc As Microsoft.Office.Interop.Word.Document, template As ReformatTemplate)
        Try
            ' 添加预览标记标题
            Dim para = doc.Paragraphs.Add()
            para.Range.Text = $"【Предпросмотр шаблона】{template.Name}"
            para.Range.Font.Size = 16
            para.Range.Font.Bold = 1
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            para.Range.InsertParagraphAfter()

            ' 添加分隔线
            para = doc.Paragraphs.Add()
            para.Range.Text = "─────────────────────────────"
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            para.Range.InsertParagraphAfter()

            ' 如果有布局配置，添加相应的示例内容
            If template.Layout IsNot Nothing AndAlso template.Layout.Elements IsNot Nothing Then
                For i As Integer = 0 To Math.Min(template.Layout.Elements.Count - 1, 5)
                    Dim element = template.Layout.Elements(i)

                    ' 添加新段落
                    para = doc.Paragraphs.Add()
                    para.Range.Text = If(String.IsNullOrEmpty(element.DefaultValue), $"【{element.Name}】Пример содержимого", element.DefaultValue)

                    ' 安全地应用字体设置
                    If element.Font IsNot Nothing Then
                        If Not String.IsNullOrEmpty(element.Font.FontNameCN) Then
                            para.Range.Font.Name = element.Font.FontNameCN
                            para.Range.Font.NameFarEast = element.Font.FontNameCN
                        End If
                        If element.Font.FontSize > 0 Then
                            para.Range.Font.Size = CSng(element.Font.FontSize)
                        End If
                        para.Range.Font.Bold = If(element.Font.Bold, 1, 0)
                    End If

                    ' 安全地应用段落设置
                    If element.Paragraph IsNot Nothing Then
                        Select Case element.Paragraph.Alignment?.ToLower()
                            Case "center"
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            Case "right"
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            Case "justify"
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Case Else
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                        End Select
                    End If

                    ' 应用颜色设置
                    If element.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(element.Color.FontColor) Then
                        Try
                            Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(element.Color.FontColor)
                            para.Range.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                        Catch ex As Exception
                            Debug.WriteLine($"应用版式元素颜色失败: {ex.Message}")
                        End Try
                    End If

                    para.Range.InsertParagraphAfter()
                Next
            End If

            ' 添加正文样式预览
            If template.BodyStyles IsNot Nothing AndAlso template.BodyStyles.Count > 0 Then
                para = doc.Paragraphs.Add()
                para.Range.Text = "─────Предпросмотр стиля основного текста─────"
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                para.Range.InsertParagraphAfter()

                For Each style In template.BodyStyles
                    para = doc.Paragraphs.Add()
                    para.Range.Text = $"【{style.RuleName}】Это пример текста основного стиля для демонстрации форматирования."

                    If style.Font IsNot Nothing Then
                        If Not String.IsNullOrEmpty(style.Font.FontNameCN) Then
                            para.Range.Font.Name = style.Font.FontNameCN
                            para.Range.Font.NameFarEast = style.Font.FontNameCN
                        End If
                        If style.Font.FontSize > 0 Then
                            para.Range.Font.Size = CSng(style.Font.FontSize)
                        End If
                        para.Range.Font.Bold = If(style.Font.Bold, 1, 0)
                    End If

                    If style.Paragraph IsNot Nothing Then
                        If style.Paragraph.FirstLineIndent > 0 AndAlso para.Range.Font.Size > 0 Then
                            para.Range.ParagraphFormat.FirstLineIndent = CSng(style.Paragraph.FirstLineIndent * para.Range.Font.Size)
                        End If
                    End If

                    ' 应用颜色设置
                    If style.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(style.Color.FontColor) Then
                        Try
                            Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(style.Color.FontColor)
                            para.Range.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                        Catch ex As Exception
                            Debug.WriteLine($"应用正文样式颜色失败: {ex.Message}")
                        End Try
                    End If

                    para.Range.InsertParagraphAfter()
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyTemplateToDocument 错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用模板进行排版（覆盖基类方法）- 语义标注流水线
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithTemplate(template As ReformatTemplate)
        Try
            ' 转换旧模板为SemanticStyleMapping
            Dim mapping = LegacyTemplateConverter.Convert(template)
            If mapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("Не удалось преобразовать шаблон")
                Return
            End If

            Await StartSemanticReformatPipeline(mapping, template.Name)
        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка форматирования: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用docx解析的SemanticStyleMapping直接排版（覆盖基类方法）
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithMapping(mapping As SemanticStyleMapping)
        Try
            Await StartSemanticReformatPipeline(mapping, mapping.Name)
        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithMapping 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка форматирования: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 从当前 Word 选区收集段落信息，供语义排版使用（消除重复代码）
    ''' </summary>
    Private Function CollectParagraphsFromSelection(
        selRange As Microsoft.Office.Interop.Word.Range,
        ByRef paragraphs As List(Of Microsoft.Office.Interop.Word.Paragraph),
        ByRef styles As List(Of String),
        ByRef types As List(Of String),
        ByRef texts As List(Of String)) As Boolean

        paragraphs = New List(Of Microsoft.Office.Interop.Word.Paragraph)()
        styles = New List(Of String)()
        types = New List(Of String)()
        texts = New List(Of String)()

        For Each p As Microsoft.Office.Interop.Word.Paragraph In selRange.Paragraphs
            Dim paraText As String = If(p.Range.Text IsNot Nothing, p.Range.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)
            Dim paraType As String = "text"
            Try
                If p.Range.InlineShapes.Count > 0 Then
                    paraType = "image"
                ElseIf p.Range.Tables.Count > 0 Then
                    paraType = "table"
                ElseIf p.Range.OMaths.Count > 0 Then
                    paraType = "formula"
                End If
            Catch
            End Try
            If Not String.IsNullOrWhiteSpace(paraText) OrElse paraType <> "text" Then
                paragraphs.Add(p)
                Dim styleName As String = ""
                Try
                    styleName = p.Style.NameLocal
                Catch
                    styleName = "Основной текст"
                End Try
                styles.Add(styleName)
                types.Add(paraType)
                texts.Add(paraText)
            End If
        Next
        Return paragraphs.Count > 0
    End Function

    Private Function CreateReformatJobFromRange(selRange As Microsoft.Office.Interop.Word.Range,
                                                Optional scopeKind As ReformatScopeKind = ReformatScopeKind.Selection) As ReformatJob
        If selRange Is Nothing Then Return Nothing

        Dim allParagraphs As List(Of Microsoft.Office.Interop.Word.Paragraph) = Nothing
        Dim paragraphStyles As List(Of String) = Nothing
        Dim paragraphTypes As List(Of String) = Nothing
        Dim paragraphTexts As List(Of String) = Nothing

        If Not CollectParagraphsFromSelection(selRange, allParagraphs, paragraphStyles, paragraphTypes, paragraphTexts) Then
            Return Nothing
        End If

        Dim job As New ReformatJob()
        job.ScopeKind = scopeKind
        Try
            job.SourceDocumentName = Globals.ThisAddIn.Application.ActiveDocument.Name
        Catch
            job.SourceDocumentName = ""
        End Try
        Try
            job.ScopeStart = selRange.Start
            job.ScopeEnd = selRange.End
        Catch
            job.ScopeStart = -1
            job.ScopeEnd = -1
        End Try

        job.WordParagraphs = allParagraphs.Cast(Of Object).ToList()
        job.ParagraphTexts = paragraphTexts
        job.ParagraphStyles = paragraphStyles
        job.ParagraphTypes = paragraphTypes

        For Each p As Microsoft.Office.Interop.Word.Paragraph In allParagraphs
            Try
                job.ParagraphFontSizes.Add(CSng(p.Range.Font.Size))
            Catch
                job.ParagraphFontSizes.Add(12.0F)
            End Try

            Try
                Dim boldRaw As Object = p.Range.Font.Bold
                job.ParagraphIsBold.Add(boldRaw IsNot Nothing AndAlso CInt(boldRaw) <> 0)
            Catch
                job.ParagraphIsBold.Add(False)
            End Try
        Next

        Return job
    End Function

    Private Function ResolveReformatTargetRange(wordApp As Microsoft.Office.Interop.Word.Application,
                                                userMessage As String,
                                                ByRef scopeKind As ReformatScopeKind) As Microsoft.Office.Interop.Word.Range
        scopeKind = ReformatScopeKind.Selection
        If wordApp Is Nothing Then Return Nothing

        Dim hasSelection = HasUsableProofreadSelection(wordApp)
        Dim explicitWhole = HasExplicitWholeDocumentScope(userMessage)
        Dim explicitSelection = HasExplicitSelectionScope(userMessage)

        Try
            If explicitSelection AndAlso hasSelection Then
                scopeKind = ReformatScopeKind.Selection
                Return wordApp.Selection.Range
            End If

            If explicitWhole AndAlso wordApp.ActiveDocument IsNot Nothing Then
                scopeKind = ReformatScopeKind.WholeDocument
                Return wordApp.ActiveDocument.Content
            End If

            If hasSelection Then
                scopeKind = ReformatScopeKind.Selection
                Return wordApp.Selection.Range
            End If

            If wordApp.ActiveDocument IsNot Nothing Then
                scopeKind = ReformatScopeKind.WholeDocument
                Return wordApp.ActiveDocument.Content
            End If
        Catch ex As Exception
            Debug.WriteLine($"ResolveReformatTargetRange error: {ex.Message}")
        End Try

        Return Nothing
    End Function

    Private Shared Function HasExplicitWholeDocumentScope(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False
        Return ContainsAnyText(message, {"全文", "整篇", "整个文档", "全部", "所有", "通篇", "统一"})
    End Function

    Private Shared Function HasExplicitSelectionScope(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False
        Return ContainsAnyText(message, {"选中", "选区", "所选", "当前选择", "选择的内容"})
    End Function

    Private Shared Function ContainsAnyText(text As String, keywords As IEnumerable(Of String)) As Boolean
        If String.IsNullOrWhiteSpace(text) Then Return False
        For Each keyword In keywords
            If text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Next
        Return False
    End Function

    Private Async Function StartSemanticReformatPipeline(mapping As SemanticStyleMapping, displayName As String) As Task
        Dim wordApp = Globals.ThisAddIn.Application
        If wordApp Is Nothing OrElse wordApp.Selection Is Nothing OrElse wordApp.Selection.Range Is Nothing Then
            GlobalStatusStrip.ShowWarning("Сначала выберите текст для форматирования.")
            Return
        End If

        Dim job = CreateReformatJobFromRange(wordApp.Selection.Range)
        If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
            GlobalStatusStrip.ShowWarning("В выбранном содержимом нет допустимых абзацев.")
            Return
        End If

        Await StartSemanticReformatPipeline(job, mapping, displayName)
    End Function

    Private Async Function StartSemanticReformatPipeline(job As ReformatJob,
                                                        mapping As SemanticStyleMapping,
                                                        displayName As String) As Task
        If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
            GlobalStatusStrip.ShowWarning("Контекст задачи форматирования потерян. Сформируйте план форматирования заново.")
            Return
        End If

        If mapping Is Nothing Then
            GlobalStatusStrip.ShowWarning("Контекст сопоставления форматирования потерян.")
            Return
        End If

        Await ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")
        ExitReformatTemplateMode()

        ' --- 智能排版 v2 优化：过滤非文本段落，传入样式上下文、字号、加粗和检测到的标题 ---
        Dim textOnlyParagraphs As New List(Of String)()
        Dim textOnlyStyles As New List(Of String)()
        Dim textOnlyOrigIndices As New List(Of Integer)()
        Dim textOnlyFontSizes As New List(Of Single)()
        Dim textOnlyIsBold As New List(Of Boolean)()

        For i = 0 To job.ParagraphTexts.Count - 1
            If job.ParagraphTypes(i) = "text" Then
                textOnlyParagraphs.Add(job.ParagraphTexts(i))
                textOnlyStyles.Add(job.ParagraphStyles(i))
                textOnlyOrigIndices.Add(i)
                If i < job.ParagraphFontSizes.Count Then textOnlyFontSizes.Add(job.ParagraphFontSizes(i)) Else textOnlyFontSizes.Add(12.0F)
                If i < job.ParagraphIsBold.Count Then textOnlyIsBold.Add(job.ParagraphIsBold(i)) Else textOnlyIsBold.Add(False)
            End If
        Next

        ' 使用 DocumentAnalyzer 检测标题结构，辅助 AI 判断
        Dim detectedHeadingInfo As String = Nothing
        Try
            Dim analyzer As New DocumentAnalyzer()
            Dim analysis = analyzer.Analyze(job.ParagraphTexts)
            If analysis.DocStructure IsNot Nothing AndAlso analysis.DocStructure.Headings.Count > 0 Then
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("Следующие абзацы предварительно определены системой как заголовки (для справки):")
                For Each h In analysis.DocStructure.Headings
                    sb.AppendLine($"  Абзац[{h.ParagraphIndex}] уровень {h.Level}: {h.Text.Substring(0, Math.Min(h.Text.Length, 60))}")
                Next
                detectedHeadingInfo = sb.ToString()
            End If
        Catch ex As Exception
            Debug.WriteLine($"DocumentAnalyzer heading detection failed: {ex.Message}")
        End Try

        ' 从映射的标签MatchHint构建文档类型上下文，帮助AI理解识别规则
        Dim docTypeCtx As New System.Text.StringBuilder()
        docTypeCtx.AppendLine(displayName)
        If mapping.SemanticTags.Count > 0 Then
            docTypeCtx.AppendLine()
            docTypeCtx.AppendLine("Правила распознавания семантических тегов:")
            Dim tags As List(Of SemanticTag) = mapping.SemanticTags
            For i As Integer = 0 To tags.Count - 1
                Dim tag As SemanticTag = tags(i)
                If Not String.IsNullOrEmpty(tag.MatchHint) Then
                    docTypeCtx.AppendLine($"- {tag.TagId}({tag.DisplayName}): {tag.MatchHint}")
                End If
            Next
        End If

        Dim systemPrompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(
            mapping, textOnlyParagraphs, textOnlyStyles, textOnlyOrigIndices, detectedHeadingInfo,
            documentTypeContext:=docTypeCtx.ToString(),
            paragraphFontSizes:=textOnlyFontSizes,
            paragraphIsBold:=textOnlyIsBold)

        job.PreviewPlan = If(job.PreviewPlan, _activeReformatPlan)
        _activeReformatJob = job
        SetReformatContext(job.WordParagraphs, job.ParagraphStyles, job.ParagraphTypes, mapping)
        Await Send("Выполни семантическую разметку выбранного содержимого с помощью «" & displayName & "».", systemPrompt, False, "semantic_reformat")
        GlobalStatusStrip.ShowInfo("Форматирование по «" & displayName & "»... диапазон: " & job.GetScopeSummary())
    End Function

    ''' <summary>
    ''' 使用规范进行排版（覆盖基类方法）- 语义标注流水线
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithStyleGuide(guide As StyleGuideResource)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim selText As String = String.Empty
            Try
                If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                    selText = If(wordApp.Selection.Range IsNot Nothing, wordApp.Selection.Range.Text, String.Empty)
                End If
            Catch
                selText = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(selText) Then
                GlobalStatusStrip.ShowWarning("Сначала выберите текст для форматирования.")
                Return
            End If

            Dim allParagraphs As List(Of Microsoft.Office.Interop.Word.Paragraph) = Nothing
            Dim paragraphStyles As List(Of String) = Nothing
            Dim paragraphTypes As List(Of String) = Nothing
            Dim paragraphTexts As List(Of String) = Nothing
            If Not CollectParagraphsFromSelection(wordApp.Selection.Range, allParagraphs, paragraphStyles, paragraphTypes, paragraphTexts) Then
                GlobalStatusStrip.ShowWarning("В выбранном содержимом нет допустимых абзацев.")
                Return
            End If

            Await ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")
            ExitReformatTemplateMode()

            Dim mapping = SemanticMappingManager.Instance.GetMappingBySourceId(guide.Id)
            If mapping IsNot Nothing Then
                Dim systemPrompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(mapping, paragraphTexts)
                SetReformatContext(allParagraphs.Cast(Of Object).ToList(), paragraphStyles, paragraphTypes, mapping)
                Await Send("Выполни семантическую разметку выбранного содержимого по стандарту форматирования «" & guide.Name & "».", systemPrompt, False, "semantic_reformat")
            Else
                Dim conversionPrompt = StyleGuideConverter.BuildConversionPrompt(guide.GuideContent)
                SetReformatContext(allParagraphs.Cast(Of Object).ToList(), paragraphStyles, paragraphTypes, Nothing)
                Await Send("Разбери стандарт форматирования «" & guide.Name & "» и извлеки параметры формата.", conversionPrompt, False, "styleguide_convert")
            End If

            GlobalStatusStrip.ShowInfo("Нормативное форматирование по «" & guide.Name & "»...")
        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithStyleGuide 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка форматирования: {ex.Message}")
        End Try
    End Sub

    ' ================== Smart Reformat v2 Methods ==================

    ''' <summary>获取或创建 ChatFormatterAgent 实例</summary>
    Private Function GetFormatterAgent() As ChatFormatterAgent
        If _formatterAgent Is Nothing Then
            _formatterAgent = ReformatSvc.ChatFormatterAgent
        End If
        Return _formatterAgent
    End Function

    ''' <summary>
    ''' 打开对话排版入口：保持普通聊天页面，不进入模板/规范资源页。
    ''' 用户在聊天框输入明确排版需求后，由 HandleSendMessage 路由到排版卡片流程。
    ''' </summary>
    Public Async Function OpenReformatPageAsync() As Task
        Dim js As String =
            "if (typeof exitReformatTemplateMode === 'function') { exitReformatTemplateMode(true); }" &
            "var smartInput = document.getElementById('smart-input');" &
            "if (smartInput) { smartInput.focus(); }" &
            "var chatInput = document.getElementById('chat-input');" &
            "if (chatInput) { chatInput.placeholder = 'Опишите требования к форматированию, например: оформить по стандарту документа'; }"

        Await ExecuteJavaScriptAsyncJS(js)
        GlobalStatusStrip.ShowInfo("Форматирование в диалоге открыто. Введите требования к форматированию в чат.")
    End Function

    Public Async Function TriggerSmartReformat() As Task
        Try
            Dim wordApp = Globals.ThisAddIn.Application

            Dim targetRange As Microsoft.Office.Interop.Word.Range = Nothing
            Dim scopeKind As ReformatScopeKind = ReformatScopeKind.Selection
            targetRange = ResolveReformatTargetRange(wordApp, "", scopeKind)

            Dim job = CreateReformatJobFromRange(targetRange, scopeKind)
            If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
                GlobalStatusStrip.ShowWarning("Допустимые абзацы не найдены.")
                Return
            End If

            ' 执行编排器分析（增强版，传入Word富文本信息）
            Dim agent = GetFormatterAgent()
            Dim plan = agent.Orchestrator.AnalyzeAndRecommend(
                job.ParagraphTexts,
                job.ParagraphStyles,
                job.ParagraphFontSizes,
                job.ParagraphIsBold)

            ' 存储段落类型
            plan.ParagraphTypes = job.ParagraphTypes
            AttachReformatJobMetadata(plan, job)
            _reformatVariantIndex = 0
            _activeWordFormattingTaskPlan = CreateSemanticWordFormattingTaskPlan("Умное форматирование в один клик", job, plan)

            ' 保存方案供后续应用
            _activeReformatPlan = plan
            job.PreviewPlan = plan
            _activeReformatJob = job
            SetReformatContext(job.WordParagraphs, job.ParagraphStyles, job.ParagraphTypes, plan.SemanticMapping)

            Await ShowReformatPlanCard(plan)
            GlobalStatusStrip.ShowInfo($"Анализ завершён, рекомендуемый стандарт: {plan.StandardName}. Подтвердите применение форматирования. Диапазон: {job.GetScopeSummary()}")
        Catch ex As Exception
            Debug.WriteLine($"TriggerSmartReformat error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка умного форматирования: {ex.Message}")
        End Try
    End Function

    Private Async Function ShowReformatPlanCard(plan As ReformatPreviewPlan) As Task
        Dim agent = GetFormatterAgent()
        Dim html = agent.GenerateFormattingCardHtml(plan)
        Dim responseUuid As String = Guid.NewGuid().ToString()

        Dim jsCreate As String = $"createChatSection('AI-помощник по форматированию', formatDateTime(new Date()), '{responseUuid}');"
        Await ExecuteJavaScriptAsyncJS(jsCreate)

        Dim jsonPayload As New JObject()
        jsonPayload("uuid") = responseUuid
        jsonPayload("html") = html
        Await ExecuteJavaScriptAsyncJS($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")
    End Function

    Private Sub AttachReformatJobMetadata(plan As ReformatPreviewPlan, job As ReformatJob)
        If plan Is Nothing OrElse job Is Nothing Then Return

        plan.ScopeSummary = job.GetScopeSummary()
        plan.TextParagraphCount = job.TextParagraphCount
        plan.TotalParagraphs = Math.Max(plan.TotalParagraphs, job.ParagraphCount)

        Select Case job.ScopeKind
            Case ReformatScopeKind.WholeDocument
                plan.ScopeKindName = "wholeDocument"
            Case ReformatScopeKind.Section
                plan.ScopeKindName = "section"
            Case Else
                plan.ScopeKindName = "selection"
        End Select
    End Sub

    Private Function CreateSemanticWordFormattingTaskPlan(userRequest As String,
                                                          job As ReformatJob,
                                                          plan As ReformatPreviewPlan) As Services.WordFormattingTaskPlan
        If job Is Nothing OrElse plan Is Nothing Then Return Nothing

        Dim targetSummary As String = ""
        If plan.TextParagraphCount > 0 AndAlso plan.TextParagraphCount <> plan.TotalParagraphs Then
            targetSummary = $"Текстовых абзацев {plan.TextParagraphCount} / всего абзацев {plan.TotalParagraphs}"
        Else
            targetSummary = $"Абзацев {plan.TotalParagraphs}"
        End If

        Return Services.WordFormattingTaskPlan.FromSemanticReformat(
            userRequest,
            job.GetScopeSummary(),
            plan.StandardName,
            targetSummary,
            plan.TotalChanges)
    End Function

    Private Function FormatWordFormattingTaskSummary(taskPlan As Services.WordFormattingTaskPlan,
                                                     executionSummary As String) As String
        If taskPlan Is Nothing Then Return executionSummary
        If String.IsNullOrWhiteSpace(executionSummary) Then Return taskPlan.ToHumanReadableSummary()
        Return taskPlan.ToHumanReadableSummary() & "; выполнение: " & executionSummary
    End Function

    ''' <summary>应用当前预览的排版方案到 Word 文档</summary>
    Private Async Function ApplyReformatPlan() As Task
        Try
            If _activeReformatPlan Is Nothing OrElse _activeReformatPlan.SemanticMapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("Нет плана форматирования для применения.")
                Return
            End If

            Dim mapping = _activeReformatPlan.SemanticMapping
            If _activeReformatJob IsNot Nothing AndAlso _activeReformatJob.HasUsableParagraphs() Then
                _activeReformatJob.PreviewPlan = _activeReformatPlan
                Await StartSemanticReformatPipeline(_activeReformatJob, mapping, _activeReformatPlan.StandardName)
            Else
                Await StartSemanticReformatPipeline(mapping, _activeReformatPlan.StandardName)
            End If
        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatPlan error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка применения форматирования: {ex.Message}")
        End Try
    End Function

    ''' <summary>智能排版 v2：应用排版方案</summary>
    Protected Overrides Async Sub HandleApplySmartReformat(jsonDoc As JObject)
        Await ApplyReformatPlan()
    End Sub

    ''' <summary>智能排版 v2：撤销排版（优先使用Word UndoRecord，失败时使用格式快照兜底）</summary>
    Protected Overrides Sub HandleUndoReformat(jsonDoc As JObject)
        Dim wordApp = Globals.ThisAddIn.Application
        If wordApp Is Nothing Then
            GlobalStatusStrip.ShowWarning("Не удалось получить объект приложения Word.")
            Return
        End If

        Dim doc = wordApp.ActiveDocument
        If doc Is Nothing Then
            GlobalStatusStrip.ShowWarning("Нет активного документа.")
            Return
        End If

        Dim screenUpdated As Boolean = False
        Dim undoSucceeded As Boolean = False
        Try
            wordApp.ScreenUpdating = False
            screenUpdated = True

            Try
                undoSucceeded = doc.Undo(1)
                If undoSucceeded Then
                    GlobalStatusStrip.ShowInfo("Форматирование отменено.")
                    Debug.WriteLine("使用Word Document.Undo撤销排版成功")
                Else
                    Debug.WriteLine("Word Document.Undo 返回False，将尝试快照恢复")
                End If
            Catch undoEx As Exception
                Debug.WriteLine($"Word Document.Undo 失败，将尝试快照恢复: {undoEx.Message}")
            End Try

            If Not undoSucceeded AndAlso _reformatSnapshot IsNot Nothing AndAlso _reformatSnapshotParagraphs IsNot Nothing Then
                Dim restoredCount = SemanticRenderingEngine.RestoreFormatSnapshot(
                    _reformatSnapshot,
                    _reformatSnapshotParagraphs,
                    _reformatSnapshotTypes)

                If restoredCount > 0 Then
                    undoSucceeded = True
                    GlobalStatusStrip.ShowInfo($"Форматирование отменено по снимку, восстановлено абзацев: {restoredCount}.")
                    Debug.WriteLine($"使用快照恢复排版成功: {restoredCount}")
                End If
            End If

            If Not undoSucceeded Then
                GlobalStatusStrip.ShowWarning("Нет операций форматирования для отмены.")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleUndoReformat 失败: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка отмены форматирования: {ex.Message}")
        Finally
            If screenUpdated Then
                Try : wordApp.ScreenUpdating = True : Catch : End Try
            End If
            If undoSucceeded Then
                _reformatSnapshot = Nothing
                _reformatSnapshotParagraphs = Nothing
                _reformatSnapshotTypes = Nothing
            End If
        End Try
    End Sub

    ''' <summary>智能排版 v2：微调排版方案</summary>
    Protected Overrides Async Sub HandleRefineSmartReformat(jsonDoc As JObject)
        Try
            Dim instruction As String = If(jsonDoc("instruction")?.ToString(), jsonDoc("command")?.ToString())
            If String.IsNullOrEmpty(instruction) Then
                GlobalStatusStrip.ShowWarning("Не указана команда тонкой настройки.")
                Return
            End If

            Dim job = _activeReformatJob
            If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
                Dim wordApp = Globals.ThisAddIn.Application
                job = CreateReformatJobFromRange(wordApp.Selection.Range)
            End If

            If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
                GlobalStatusStrip.ShowWarning("Допустимые абзацы не найдены.")
                Return
            End If

            ' 通过编排器应用微调
            Dim agent = GetFormatterAgent()
            Dim refinedPlan = agent.Orchestrator.ApplyRefinement(instruction)
            If refinedPlan IsNot Nothing Then
                AttachReformatJobMetadata(refinedPlan, job)
                _activeWordFormattingTaskPlan = CreateSemanticWordFormattingTaskPlan(instruction, job, refinedPlan)
                _activeReformatPlan = refinedPlan
                job.PreviewPlan = refinedPlan
                _activeReformatJob = job

                ' 同步 V1 上下文，确保"应用排版"使用固定任务范围
                SetReformatContext(job.WordParagraphs, job.ParagraphStyles, job.ParagraphTypes, refinedPlan.SemanticMapping)

                ' 重新生成卡片
                Dim html = agent.GenerateFormattingCardHtml(refinedPlan)
                Dim responseUuid As String = Guid.NewGuid().ToString()
                Dim jsonPayload As New JObject()
                jsonPayload("uuid") = responseUuid
                jsonPayload("html") = html
                Await ExecuteJavaScriptAsyncJS($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")
                GlobalStatusStrip.ShowInfo("План форматирования скорректирован.")
            Else
                GlobalStatusStrip.ShowWarning("Тонкая настройка не удалась. Попробуйте более конкретную команду.")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleRefineSmartReformat error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка тонкой настройки форматирования: {ex.Message}")
        End Try
    End Sub

    ''' <summary>智能排版 v2：切换排版标准/模板</summary>
    Protected Overrides Async Sub HandleSwitchReformatTemplate(jsonDoc As JObject)
        Try
            Dim templateName As String = If(jsonDoc("templateName")?.ToString(), "")
            If String.IsNullOrEmpty(templateName) Then
                Dim job = _activeReformatJob
                If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
                    Dim wordApp = Globals.ThisAddIn.Application
                    Dim scopeKind As ReformatScopeKind = ReformatScopeKind.Selection
                    Dim targetRange = ResolveReformatTargetRange(wordApp, "", scopeKind)
                    job = CreateReformatJobFromRange(targetRange, scopeKind)
                End If

                If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
                    GlobalStatusStrip.ShowWarning("Допустимые абзацы не найдены.")
                    Return
                End If

                Dim agent = GetFormatterAgent()
                _reformatVariantIndex += 1
                Dim currentStandardName = If(_activeReformatPlan Is Nothing, "", _activeReformatPlan.StandardName)
                Dim plan = agent.Orchestrator.GenerateAlternativePlan(
                    job.ParagraphTexts,
                    job.ParagraphStyles,
                    job.ParagraphFontSizes,
                    job.ParagraphIsBold,
                    currentStandardName,
                    _reformatVariantIndex)
                If plan IsNot Nothing Then
                    plan.ParagraphTypes = job.ParagraphTypes
                    AttachReformatJobMetadata(plan, job)
                    _activeWordFormattingTaskPlan = CreateSemanticWordFormattingTaskPlan("Другой вариант форматирования", job, plan)
                    _activeReformatPlan = plan
                    job.PreviewPlan = plan
                    _activeReformatJob = job
                    SetReformatContext(job.WordParagraphs, job.ParagraphStyles, job.ParagraphTypes, plan.SemanticMapping)

                    Dim html = agent.GenerateFormattingCardHtml(plan)
                    Dim responseUuid As String = Guid.NewGuid().ToString()
                    Dim jsonPayload As New JObject()
                    jsonPayload("uuid") = responseUuid
                    jsonPayload("html") = html
                    Await ExecuteJavaScriptAsyncJS($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")
                    GlobalStatusStrip.ShowInfo($"Переключено на альтернативный план форматирования: {plan.StandardName}")
                End If
            Else
                GlobalStatusStrip.ShowInfo($"Переключено на стандарт: {templateName}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleSwitchReformatTemplate error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка переключения стандарта форматирования: {ex.Message}")
        End Try
    End Sub

    ''' <summary>智能排版 v2：显示排版前后对比</summary>
    Protected Overrides Async Sub HandlePreviewReformatCompare(jsonDoc As JObject)
        Try
            If _activeReformatPlan Is Nothing Then
                GlobalStatusStrip.ShowWarning("Нет плана форматирования для предпросмотра.")
                Return
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine("【Сравнение до и после форматирования】")
            If Not String.IsNullOrWhiteSpace(_activeReformatPlan.ScopeSummary) Then
                sb.AppendLine($"Область применения: {_activeReformatPlan.ScopeSummary}")
            End If
            sb.AppendLine($"Рекомендуемый стандарт: {_activeReformatPlan.StandardName}")
            sb.AppendLine($"Тип документа: {_activeReformatPlan.DetectedType.ToString()}")
            sb.AppendLine($"Всего изменений формата: {_activeReformatPlan.TotalStyleChanges}:")
            If _activeReformatPlan.Changes Is Nothing OrElse _activeReformatPlan.Changes.Count = 0 Then
                sb.AppendLine("  - Структурный анализ выполнен; области, требующие немедленной настройки, не обнаружены.")
            Else
                For Each change In _activeReformatPlan.Changes
                    Dim tagName = GetReformatTagDisplayName(change.NewTag, _activeReformatPlan.SemanticMapping)
                    sb.AppendLine($"  - [{tagName}] {change.ChangeDescription}")
                Next
            End If

            Dim responseUuid As String = Guid.NewGuid().ToString()
            Dim jsCreate As String = $"createChatSection('Сравнение форматирования', formatDateTime(new Date()), '{responseUuid}');"
            Await ExecuteJavaScriptAsyncJS(jsCreate)
            Dim escapedText = sb.ToString().Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "\n").Replace(vbLf, "")
            Await ExecuteJavaScriptAsyncJS($"appendRenderer('{responseUuid}','{escapedText}');")
        Catch ex As Exception
            Debug.WriteLine($"HandlePreviewReformatCompare error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка сравнения предпросмотра: {ex.Message}")
        End Try
    End Sub

    Private Function GetReformatTagDisplayName(tagId As String, mapping As SemanticStyleMapping) As String
        If String.IsNullOrWhiteSpace(tagId) Then Return "Стиль для распознавания"

        Try
            If mapping IsNot Nothing Then
                Dim tag = mapping.FindTag(tagId)
                If tag IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tag.DisplayName) Then
                    Return tag.DisplayName
                End If
            End If
        Catch
        End Try

        Select Case tagId.ToLowerInvariant()
            Case "body.normal"
                Return "Основной текст"
            Case "title.1", "heading.1"
                Return "Заголовок 1 уровня"
            Case "title.2", "heading.2"
                Return "Заголовок 2 уровня"
            Case "title.3", "heading.3"
                Return "Заголовок 3 уровня"
            Case "list.ordered"
                Return "Нумерованный список"
            Case "list.unordered"
                Return "Маркированный список"
            Case Else
                Return tagId
        End Select
    End Function

    ''' <summary>
    ''' 重写消息发送：拦截排版相关消息，走专门的排版流程
    ''' </summary>
    Protected Overrides Sub HandleSendMessage(jsonDoc As JObject)
        Debug.WriteLine($"[DEBUG WordAi.HandleSendMessage] ENTER")
        Dim messageValue As JToken = jsonDoc("value")
        Dim question As String = ""

        If messageValue IsNot Nothing AndAlso messageValue.Type = JTokenType.Object Then
            question = messageValue("text")?.ToString()
        End If
        Debug.WriteLine($"[DEBUG WordAi.HandleSendMessage] question='{question}', IsExplicit={IsExplicitFormattingRequest(question)}")

        If IsMixedProofreadFormattingRequest(question) Then
            Debug.WriteLine("[DEBUG WordAi.HandleSendMessage] routing to HandleMixedProofreadFormattingCommand")
            HandleMixedProofreadFormattingCommand(question, jsonDoc)
            Return
        End If

        If TryRouteWordActionHarness(question, jsonDoc) Then Return

        ' 非排版消息走正常流程
        Debug.WriteLine($"[DEBUG WordAi.HandleSendMessage] routing to MyBase.HandleSendMessage")
        MyBase.HandleSendMessage(jsonDoc)
    End Sub

    ''' <summary>
    ''' Host capability fast path (P0-2): high-confidence Word capabilities run deterministically.
    ''' This is not a parallel NL product router; unresolved messages continue to ChatRouting → AgentKernel.
    ''' Future: fold observe/explain results back into Agent loop memory.
    ''' </summary>
    Private Function TryRouteWordActionHarness(userMessage As String, originalJsonDoc As JObject) As Boolean
        If String.IsNullOrWhiteSpace(userMessage) Then Return False
        If Not Agent.ExecutionPathPolicy.AllowHostCapabilityFastPath Then Return False

        Try
            Dim intent As IntentResult = Nothing
            Try
                intent = IntentService.IdentifyIntent(userMessage)
            Catch ex As Exception
                Debug.WriteLine($"[WordActionHarness] IntentService failed: {ex.Message}")
            End Try

            Dim harness As New Services.WordActionHarness(Globals.ThisAddIn.Application)
            Dim plan = harness.Plan(userMessage, intent)
            If plan Is Nothing OrElse Not plan.ShouldHandle Then Return False

            Dim recorder As New Agent.Harness.HostCapabilityRunRecorder(
                "Word",
                If(plan.Capability?.Id, plan.Kind.ToString()),
                userMessage,
                If(plan.Capability Is Nothing, "medium", plan.Capability.RiskLevel.ToString()))
            plan.RunRecorder = recorder
            Dim safetyDecision = recorder.EvaluateSafety()
            If safetyDecision IsNot Nothing AndAlso safetyDecision.Action <> Agent.Execution.SafetyAction.Allow Then
                recorder.CompleteSafetyDecision(safetyDecision)
                GlobalStatusStrip.ShowWarning(If(safetyDecision.UserMessage, "Операция требует подтверждения; выполняется через безопасный путь Agent"))
                Return False
            End If

            Debug.WriteLine($"[WordActionHarness] capability fast-path kind={plan.Kind}, confidence={plan.Confidence:0.00}, reason={plan.Reason}, capability={plan.CapabilitySummary}")

            Select Case plan.Kind
                Case Services.WordActionKind.Numbering
                    HandleNumberingCommand(userMessage, originalJsonDoc, plan)
                    Return True

                Case Services.WordActionKind.Proofread
                    HandleProofreadCommand(userMessage, originalJsonDoc, plan)
                    Return True

                Case Services.WordActionKind.DirectFormatting
                    HandleDirectFormattingCommand(userMessage, originalJsonDoc, plan)
                    Return True

                Case Services.WordActionKind.SemanticReformat
                    HandleChatDrivenReformat(userMessage, originalJsonDoc, plan)
                    Return True
            End Select
        Catch ex As Exception
            Debug.WriteLine($"[WordActionHarness] route failed: {ex.Message}")
            AppLogger.Error("WordActionHarness", "Capability fast-path route failed", ex)
        End Try

        Return False
    End Function

    Private Sub ReportWordCapabilityResult(plan As Services.WordActionPlan, result As Services.WordCapabilityExecutionResult)
        If result Is Nothing Then Return

        Dim capabilityId = If(String.IsNullOrWhiteSpace(result.CapabilityId), If(plan?.Capability?.Id, "word.unknown"), result.CapabilityId)
        Dim message = result.ToObserveSummary()
        If result.Success Then
            AppLogger.Info("WordCapability", message)
        Else
            AppLogger.Warn("WordCapability", message)
        End If
        Debug.WriteLine($"[WordCapability] {message}")

        Try
            plan?.RunRecorder?.Complete(result.ToToolResult())
        Catch ex As Exception
            AppLogger.Warn("WordCapability", $"fast-path trace failed: {AppLogger.Redact(ex.Message)}")
        End Try

        Try
            If result.Status = Services.WordCapabilityExecutionStatus.Failed Then
                GlobalStatusStrip.ShowWarning(If(String.IsNullOrWhiteSpace(result.UserMessage), $"{capabilityId}: ошибка выполнения", result.UserMessage))
            End If
        Catch
        End Try
    End Sub

    Private Async Sub HandleDirectFormattingCommand(userMessage As String, originalJsonDoc As JObject, plan As Services.WordActionPlan)
        Try
            Dim applied = Await ExecuteDirectFormattingCommandAsync(userMessage, True)
            ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                plan,
                If(applied, Services.WordCapabilityExecutionStatus.Succeeded, Services.WordCapabilityExecutionStatus.Fallback),
                applied,
                If(applied, "Настройка формата применена.", "Не удалось применить настройку формата; возврат к основному пути Agent."),
                plan?.Reason,
                New With {.applied = applied}))
            If Not applied Then
                If originalJsonDoc IsNot Nothing Then
                    MyBase.HandleSendMessage(originalJsonDoc)
                Else
                    SendChatMessage(userMessage)
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleDirectFormattingCommand error: {ex.Message}")
            ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                plan,
                Services.WordCapabilityExecutionStatus.Failed,
                False,
                $"Не удалось изменить формат: {ex.Message}",
                ex.ToString(),
                recoverable:=True))
            GlobalStatusStrip.ShowWarning($"Ошибка настройки формата: {ex.Message}")
            Try
                If originalJsonDoc IsNot Nothing Then
                    MyBase.HandleSendMessage(originalJsonDoc)
                Else
                    SendChatMessage(userMessage)
                End If
            Catch fallbackEx As Exception
                Debug.WriteLine($"HandleDirectFormattingCommand fallback error: {fallbackEx.Message}")
            End Try
        End Try
    End Sub

    Private Async Sub HandleNumberingCommand(userMessage As String, originalJsonDoc As JObject, plan As Services.WordActionPlan)
        Try
            Dim userMsgUuid = Guid.NewGuid().ToString()
            Await ExecuteJavaScriptAsyncJS($"createChatSection('Вы', formatDateTime(new Date()), '{userMsgUuid}');")
            Await ExecuteJavaScriptAsyncJS($"document.getElementById('content-{userMsgUuid}').innerHTML = '<p>{EscapeHtmlForInline(userMessage)}</p>';")

            Dim agent As New Services.WordNumberingAgent(Globals.ThisAddIn.Application)
            Dim result = agent.RebuildSequentialNumbering(userMessage)

            If result Is Nothing OrElse Not result.Success Then
                Dim reason = If(result Is Nothing, "Результат выполнения нумерации не сформирован", result.ToHumanReadableSummary())
                Debug.WriteLine($"[Word] 自动编号重排未执行: {reason}")
                ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                    plan,
                    Services.WordCapabilityExecutionStatus.Failed,
                    False,
                    reason,
                    reason,
                    result,
                    recoverable:=True))
                Dim failureResponseUuid = Guid.NewGuid().ToString()
                Await ExecuteJavaScriptAsyncJS($"createChatSection('AI-помощник по нумерации', formatDateTime(new Date()), '{failureResponseUuid}');")
                Dim escapedReason = reason.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "").Replace(vbLf, "\n")
                Await ExecuteJavaScriptAsyncJS($"appendRenderer('{failureResponseUuid}','Автонумерация не перестроена: {escapedReason}');")
                GlobalStatusStrip.ShowWarning(reason)
                Return
            End If

            Dim responseUuid = Guid.NewGuid().ToString()
            Await ExecuteJavaScriptAsyncJS($"createChatSection('AI-помощник по нумерации', formatDateTime(new Date()), '{responseUuid}');")
            Dim summary = result.ToHumanReadableSummary()
            ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                plan,
                Services.WordCapabilityExecutionStatus.Succeeded,
                True,
                summary,
                summary,
                result))
            Dim escapedSummary = summary.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "").Replace(vbLf, "\n")
            Await ExecuteJavaScriptAsyncJS($"appendRenderer('{responseUuid}','Автонумерация успешно перестроена: {escapedSummary}');")
            GlobalStatusStrip.ShowSuccess("Автонумерация перестроена в непрерывную")
        Catch ex As Exception
            Debug.WriteLine($"HandleNumberingCommand error: {ex.Message}")
            ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                plan,
                Services.WordCapabilityExecutionStatus.Failed,
                False,
                $"Не удалось перестроить автонумерацию: {ex.Message}",
                ex.ToString(),
                recoverable:=True))
            GlobalStatusStrip.ShowWarning($"Ошибка автоматической перенумерации: {ex.Message}")
            Try
                If originalJsonDoc IsNot Nothing Then
                    MyBase.HandleSendMessage(originalJsonDoc)
                Else
                    SendChatMessage(userMessage)
                End If
            Catch fallbackEx As Exception
                Debug.WriteLine($"HandleNumberingCommand fallback error: {fallbackEx.Message}")
            End Try
        End Try
    End Sub

    Private Async Function ExecuteDirectFormattingCommandAsync(userMessage As String,
                                                               showUserMessage As Boolean) As Task(Of Boolean)
        Try
            If showUserMessage Then
                Dim userMsgUuid = Guid.NewGuid().ToString()
                Await ExecuteJavaScriptAsyncJS($"createChatSection('Вы', formatDateTime(new Date()), '{userMsgUuid}');")
                Await ExecuteJavaScriptAsyncJS($"document.getElementById('content-{userMsgUuid}').innerHTML = '<p>{EscapeHtmlForInline(userMessage)}</p>';")
            End If

            Dim agent As New Services.WordFormattingAgent(Globals.ThisAddIn.Application)
            Dim result = agent.RunDirectFormatting(userMessage)
            _activeWordFormattingTaskPlan = result?.TaskPlan

            If result Is Nothing OrElse Not result.Success Then
                Dim reason = If(result Is Nothing, "Результат выполнения не сформирован", result.ToHumanReadableSummary())
                Debug.WriteLine($"[Word] 直接格式调整未执行: {reason}")
                Return False
            End If

            Dim responseUuid = Guid.NewGuid().ToString()
            Await ExecuteJavaScriptAsyncJS($"createChatSection('AI-помощник по форматированию', formatDateTime(new Date()), '{responseUuid}');")
            Dim unifiedSummary = FormatWordFormattingTaskSummary(result.TaskPlan, result.ToHumanReadableSummary())
            Dim escapedSummary = unifiedSummary.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "").Replace(vbLf, "\n")
            Await ExecuteJavaScriptAsyncJS($"appendRenderer('{responseUuid}','Настройка формата применена: {escapedSummary}');")
            GlobalStatusStrip.ShowSuccess("Настройка формата применена")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteDirectFormattingCommandAsync error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка настройки формата: {ex.Message}")
            Return False
        End Try
    End Function

    Private Shared Function IsMixedProofreadFormattingRequest(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False
        If Not Services.ProofreadIntentCompiler.LooksLikeProofreadCommand(message) Then Return False
        Return Services.SmartFormatter.LooksLikeDirectFormattingCommand(message) OrElse IsExplicitFormattingRequest(message)
    End Function

    Private Async Sub HandleMixedProofreadFormattingCommand(userMessage As String, originalJsonDoc As JObject)
        Try
            Dim userMsgUuid = Guid.NewGuid().ToString()
            Await ExecuteJavaScriptAsyncJS($"createChatSection('Вы', formatDateTime(new Date()), '{userMsgUuid}');")
            Await ExecuteJavaScriptAsyncJS($"document.getElementById('content-{userMsgUuid}').innerHTML = '<p>{EscapeHtmlForInline(userMessage)}</p>';")

            _pendingPostProofreadFormattingRequest = ExtractPostProofreadFormattingRequest(userMessage)
            GlobalStatusStrip.ShowInfo("Распознана составная задача: сначала проверка, затем форматирование.")

            Dim wordApp = Globals.ThisAddIn.Application
            Dim compiler As New Services.ProofreadIntentCompiler()
            Dim plan = compiler.Compile(userMessage, HasUsableProofreadSelection(wordApp))
            Await ExecuteProofreadAsync(plan, userMessage)
        Catch ex As Exception
            Debug.WriteLine($"HandleMixedProofreadFormattingCommand error: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка запуска составной задачи: {ex.Message}")
            Try
                If originalJsonDoc IsNot Nothing Then
                    MyBase.HandleSendMessage(originalJsonDoc)
                Else
                    SendChatMessage(userMessage)
                End If
            Catch fallbackEx As Exception
                Debug.WriteLine($"HandleMixedProofreadFormattingCommand fallback error: {fallbackEx.Message}")
            End Try
        End Try
    End Sub

    Private Shared Function ExtractPostProofreadFormattingRequest(message As String) As String
        If String.IsNullOrWhiteSpace(message) Then Return ""

        Dim separators As String() = {"再", "然后", "之后", "接着", "并且", "并"}
        For Each separator In separators
            Dim index = message.IndexOf(separator, StringComparison.Ordinal)
            If index >= 0 AndAlso index + separator.Length < message.Length Then
                Dim tail = message.Substring(index + separator.Length).Trim()
                If Not String.IsNullOrWhiteSpace(tail) Then Return tail
            End If
        Next

        Return message
    End Function

    Private Async Function SendWithIntentPromptAsync(message As String, intent As IntentResult) As Task
        Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
        Await Send(message, optimizedPrompt, True, "")
    End Function

    Private Async Function RouteProofreadIntentAsync(message As String, intent As IntentResult) As Task
        Dim fallbackNeeded As Boolean = False
        Try
            Dim compiler As New Services.ProofreadIntentCompiler()
            Dim plan = compiler.Compile(If(String.IsNullOrWhiteSpace(message), "校对", message),
                                        HasUsableProofreadSelection(Globals.ThisAddIn.Application))
            Await ExecuteProofreadAsync(plan, message)
        Catch ex As Exception
            Debug.WriteLine($"[Word] 校对路由失败: {ex.Message}")
            fallbackNeeded = True
        End Try

        If fallbackNeeded Then
            Await SendWithIntentPromptAsync(message, intent)
        End If
    End Function

    Private Async Function RouteFormattingIntentAsync(message As String, intent As IntentResult) As Task
        Dim fallbackNeeded As Boolean = False
        Dim directAttempted As Boolean = False
        Try
            If Services.SmartFormatter.LooksLikeDirectFormattingCommand(message) Then
                directAttempted = True
                Dim applied = Await ExecuteDirectFormattingCommandAsync(message, True)
                If applied Then Return
            End If

            Dim wordApp = Globals.ThisAddIn.Application
            Dim hasSelection As Boolean = False
            Try
                If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing AndAlso wordApp.Selection.Range IsNot Nothing Then
                    hasSelection = Not String.IsNullOrWhiteSpace(wordApp.Selection.Range.Text)
                End If
            Catch
            End Try

            If hasSelection Then
                Await TriggerSmartReformat()
            Else
                Dim applied = Await ExecuteDirectFormattingCommandAsync(message, Not directAttempted)
                If Not applied Then
                    GlobalStatusStrip.ShowWarning("Не удалось автоматически сформировать исполняемый план форматирования; переход в обычный диалог.")
                    fallbackNeeded = True
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[Word] 排版路由失败: {ex.Message}")
            fallbackNeeded = True
        End Try

        If fallbackNeeded Then
            Await SendWithIntentPromptAsync(message, intent)
        End If
    End Function

    Private Async Sub HandleProofreadCommand(userMessage As String, originalJsonDoc As JObject, plan As Services.WordActionPlan)
        Try
            Dim userMsgUuid = Guid.NewGuid().ToString()
            Await ExecuteJavaScriptAsyncJS($"createChatSection('Вы', formatDateTime(new Date()), '{userMsgUuid}');")
            Await ExecuteJavaScriptAsyncJS($"document.getElementById('content-{userMsgUuid}').innerHTML = '<p>{EscapeHtmlForInline(userMessage)}</p>';")

            Dim wordApp = Globals.ThisAddIn.Application
            Dim compiler As New Services.ProofreadIntentCompiler()
            Dim proofreadPlan = compiler.Compile(userMessage, HasUsableProofreadSelection(wordApp))
            Dim started = Await ExecuteProofreadAsync(proofreadPlan, userMessage)
            ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                plan,
                If(started, Services.WordCapabilityExecutionStatus.Succeeded, Services.WordCapabilityExecutionStatus.Fallback),
                started,
                If(started, "Процесс проверки запущен.", "Процесс проверки не запущен; возврат к основному пути Agent."),
                If(proofreadPlan Is Nothing, "", proofreadPlan.ToHumanReadableSummary()),
                proofreadPlan))
            If Not started AndAlso originalJsonDoc IsNot Nothing Then
                MyBase.HandleSendMessage(originalJsonDoc)
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleProofreadCommand error: {ex.Message}")
            ReportWordCapabilityResult(plan, Services.WordCapabilityExecutionResult.FromPlan(
                plan,
                Services.WordCapabilityExecutionStatus.Failed,
                False,
                $"Не удалось запустить проверку: {ex.Message}",
                ex.ToString(),
                recoverable:=True))
            GlobalStatusStrip.ShowWarning($"Ошибка запуска проверки: {ex.Message}")
            Try
                MyBase.HandleSendMessage(originalJsonDoc)
            Catch fallbackEx As Exception
                Debug.WriteLine($"HandleProofreadCommand fallback error: {fallbackEx.Message}")
            End Try
        End Try
    End Sub

    Private Function EscapeHtmlForInline(value As String) As String
        If value Is Nothing Then Return ""
        Return WebUtility.HtmlEncode(value).Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "").Replace(vbLf, "<br>")
    End Function

    ''' <summary>
    ''' 判断是否为明确的排版指令（比IsFormattingRelated更严格）
    ''' 必须包含排版核心关键词+动作词，避免误截日常对话
    ''' </summary>
    Private Shared Function IsExplicitFormattingRequest(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        ' 核心排版动作词（必须命中至少一个）
        Dim actionKeywords As String() = {
            "排版", "重新排版", "帮我排版", "按.*排版", "按照.*排版",
            "格式化", "套用格式", "应用格式", "换一种排版"
        }
        ' 补充场景词（配合动作词使用）
        Dim topicKeywords As String() = {
            "公文", "国标", "GB/T", "gbt", "标准格式",
            "仿宋", "宋体", "黑体", "楷体", "微软雅黑", "小标宋",
            "序号", "编号", "标题", "标题层级", "标题编号",
            "стандарт", "шрифт", "заголовок", "нумерац", "стиль"
        }

        Dim msg = message.Trim()

        ' 直接包含"排版"动作
        If msg.Contains("排版") Then Return True
        ' Русские действия форматирования
        If msg.Contains("отформатир") OrElse msg.Contains("форматировани") OrElse msg.Contains("переформат") OrElse msg.Contains("оформ") Then Return True
        ' "按XX格式/标准" 句式
        If (msg.Contains("按") OrElse msg.Contains("按照") OrElse msg.Contains("参照")) AndAlso
           (msg.Contains("格式") OrElse msg.Contains("标准") OrElse msg.Contains("模板")) Then Return True
        If msg.Contains("по стандарту") OrElse msg.Contains("по образцу") OrElse msg.Contains("по шаблону") Then Return True
        ' "格式化" 动作
        If msg.Contains("格式化") Then Return True
        ' "套用/应用格式" 动作
        If (msg.Contains("套用") OrElse msg.Contains("应用")) AndAlso msg.Contains("格式") Then Return True
        If msg.Contains("примен") AndAlso (msg.Contains("формат") OrElse msg.Contains("стил")) Then Return True
        ' 标题/编号结构重构
        If (msg.Contains("重构") OrElse msg.Contains("整理") OrElse msg.Contains("规范") OrElse msg.Contains("优化") OrElse msg.Contains("调整")) AndAlso
           (msg.Contains("序号") OrElse msg.Contains("编号") OrElse msg.Contains("标题") OrElse msg.Contains("层级")) Then Return True
        If (msg.Contains("нумерац") OrElse msg.Contains("перестро") OrElse msg.Contains("упорядоч") OrElse msg.Contains("нормализ")) AndAlso
           (msg.Contains("заголов") OrElse msg.Contains("нумерац") OrElse msg.Contains("уровн")) Then Return True
        If msg.Contains("标题") AndAlso (msg.Contains("序号") OrElse msg.Contains("编号") OrElse msg.Contains("层级")) Then Return True
        If msg.Contains("заголов") AndAlso (msg.Contains("нумерац") OrElse msg.Contains("уровн")) Then Return True
        ' 主题词+动作组合（如"公文标准"、"宋体样式"）
        For Each topic In topicKeywords
            If msg.Contains(topic) AndAlso
               (msg.Contains("排") OrElse msg.Contains("格式") OrElse msg.Contains("样式") OrElse msg.Contains("规范") OrElse
                msg.Contains("форм") OrElse msg.Contains("стил") OrElse msg.Contains("оформ")) Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Chat驱动的排版流程：收集段落→意图解析→选择标准→预览卡片
    ''' 用户点击"应用排版"后走V1的AI语义标注管道（StartSemanticReformatPipeline）
    ''' </summary>
    Private Async Sub HandleChatDrivenReformat(userMessage As String, originalJsonDoc As JObject, planContext As Services.WordActionPlan)
        Try
            Dim wordApp = Globals.ThisAddIn.Application

            ' 1. 从用户表达和当前 Word 状态推断选区/全文，并创建固定排版任务
            Dim targetRange As Microsoft.Office.Interop.Word.Range = Nothing
            Dim scopeKind As ReformatScopeKind = ReformatScopeKind.Selection
            targetRange = ResolveReformatTargetRange(wordApp, userMessage, scopeKind)

            Dim job = CreateReformatJobFromRange(targetRange, scopeKind)
            If job Is Nothing OrElse Not job.HasUsableParagraphs() Then
                ReportWordCapabilityResult(planContext, Services.WordCapabilityExecutionResult.FromPlan(
                    planContext,
                    Services.WordCapabilityExecutionStatus.Fallback,
                    False,
                    "В текущем диапазоне нет абзацев, пригодных для семантического форматирования; возврат к основному пути Agent.",
                    "CreateReformatJobFromRange returned no usable paragraphs",
                    New With {.scope = scopeKind.ToString()}))
                ' 没有选区时回退到正常AI对话流程
                If originalJsonDoc IsNot Nothing Then
                    MyBase.HandleSendMessage(originalJsonDoc)
                Else
                    Await Send(userMessage, "", True, "")
                End If
                Return
            End If

            ' 3. 显示用户消息到Chat（普通气泡）
            Dim userMsgUuid = Guid.NewGuid().ToString()
            Dim jsUserMsg = $"createChatSection('Вы', formatDateTime(new Date()), '{userMsgUuid}');"
            Await ExecuteJavaScriptAsyncJS(jsUserMsg)
            Dim escapedMsg = userMessage.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "\n").Replace(vbLf, "")
            Await ExecuteJavaScriptAsyncJS($"document.getElementById('content-{userMsgUuid}').innerHTML = '<p>{escapedMsg}</p>';")

            ' 4. 通过编排器解析意图并推荐标准（增强版，传入Word富文本信息）
            GlobalStatusStrip.ShowInfo("Анализ документа...")
            Dim agent = GetFormatterAgent()
            Dim intentResult = Await agent.RecognizeReformatIntentAsync(
                userMessage,
                job.ParagraphTexts,
                job.ParagraphStyles,
                job.ParagraphFontSizes,
                job.ParagraphIsBold)
            If intentResult IsNot Nothing Then
                Debug.WriteLine($"[ReformatIntent] source={intentResult.Source}, confidence={intentResult.Confidence:0.00}, type={intentResult.Intent.IntentType}, standard={intentResult.Intent.TargetStandardName}")
                GlobalStatusStrip.ShowInfo($"Распознано намерение форматирования: {intentResult.Intent.IntentType} ({intentResult.Source})")

                If intentResult.ScopeHint = ReformatScopeKind.WholeDocument AndAlso
                   scopeKind <> ReformatScopeKind.WholeDocument AndAlso
                   Not HasExplicitSelectionScope(userMessage) AndAlso
                   wordApp IsNot Nothing AndAlso wordApp.ActiveDocument IsNot Nothing Then

                    Dim wholeDocJob = CreateReformatJobFromRange(wordApp.ActiveDocument.Content, ReformatScopeKind.WholeDocument)
                    If wholeDocJob IsNot Nothing AndAlso wholeDocJob.HasUsableParagraphs() Then
                        job = wholeDocJob
                        scopeKind = ReformatScopeKind.WholeDocument
                        Debug.WriteLine("[ReformatIntent] scope hint switched target to whole document")
                    End If
                End If
            End If

            Dim plan = Await agent.Orchestrator.ChatReformatAsync(
                userMessage,
                job.ParagraphTexts,
                job.WordParagraphs,
                job.ParagraphStyles,
                job.ParagraphFontSizes,
                job.ParagraphIsBold,
                If(intentResult IsNot Nothing, intentResult.Intent, Nothing))

            ' 5. 存储段落类型到方案中
            plan.ParagraphTypes = job.ParagraphTypes
            AttachReformatJobMetadata(plan, job)
            _reformatVariantIndex = 0
            _activeWordFormattingTaskPlan = CreateSemanticWordFormattingTaskPlan(userMessage, job, plan)

            ' 6. 存储上下文供后续应用
            _activeReformatPlan = plan
            job.PreviewPlan = plan
            _activeReformatJob = job
            SetReformatContext(job.WordParagraphs, job.ParagraphStyles, job.ParagraphTypes, plan.SemanticMapping)

            ' 7. 生成排版卡片并推送到Chat
            Dim html = agent.GenerateFormattingCardHtml(plan)
            Dim responseUuid = Guid.NewGuid().ToString()
            Dim jsCreate = $"createChatSection('AI-помощник по форматированию', formatDateTime(new Date()), '{responseUuid}');"
            Await ExecuteJavaScriptAsyncJS(jsCreate)

            Dim jsonPayload As New JObject()
            jsonPayload("uuid") = responseUuid
            jsonPayload("html") = html
            Await ExecuteJavaScriptAsyncJS($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")

            ReportWordCapabilityResult(planContext, Services.WordCapabilityExecutionResult.FromPlan(
                planContext,
                Services.WordCapabilityExecutionStatus.Succeeded,
                True,
                $"Предпросмотр семантического форматирования сформирован, рекомендуемый стандарт: {plan.StandardName}.",
                $"scope={job.GetScopeSummary()}, paragraphs={job.ParagraphTexts.Count}, standard={plan.StandardName}",
                New With {.scope = job.GetScopeSummary(), .paragraphCount = job.ParagraphTexts.Count, .standardName = plan.StandardName}))
            GlobalStatusStrip.ShowInfo($"Анализ завершён, рекомендуемый стандарт: {plan.StandardName}. Диапазон: {job.GetScopeSummary()}")
        Catch ex As Exception
            Debug.WriteLine($"HandleChatDrivenReformat error: {ex.Message}")
            ReportWordCapabilityResult(planContext, Services.WordCapabilityExecutionResult.FromPlan(
                planContext,
                Services.WordCapabilityExecutionStatus.Failed,
                False,
                $"Сбой умного форматирования: {ex.Message}",
                ex.ToString(),
                recoverable:=True))
            GlobalStatusStrip.ShowWarning($"Ошибка умного форматирования: {ex.Message}")
            ' 排版失败时回退到正常AI对话流程
            Try
                MyBase.HandleSendMessage(originalJsonDoc)
            Catch fallbackEx As Exception
                Debug.WriteLine($"HandleChatDrivenReformat fallback error: {fallbackEx.Message}")
            End Try
        End Try
    End Sub

    ' ================== 校对专注模式 ==================

    ''' <summary>
    ''' 校对专注模式实例
    ''' </summary>
    Private _proofreadFocusMode As SmartProofreadFocusMode = Nothing

    ''' <summary>
    ''' 当前校对的段落列表（用于结果处理）
    ''' </summary>
    Private _proofreadParagraphs As List(Of String) = Nothing

    ''' <summary>
    ''' 校对时用户选中的 Range（用于替换操作限定范围）
    ''' </summary>
    Private _proofreadSelectionRange As Object = Nothing
    Private _pendingPostProofreadFormattingRequest As String = ""

    ''' <summary>
    ''' 获取或创建校对专注模式实例
    ''' </summary>
    Private Function GetProofreadFocusMode() As SmartProofreadFocusMode
        If _proofreadFocusMode Is Nothing Then
            _proofreadFocusMode = New SmartProofreadFocusMode(
                AddressOf ExecuteJavaScriptAsyncJS,
                AddressOf ApplyCorrectionToWord)
        End If
        Return _proofreadFocusMode
    End Function

    ''' <summary>
    ''' 处理校对专注模式的消息
    ''' </summary>
    Protected Overrides Sub HandleProofreadFocusMode(jsonDoc As JObject)
        Try
            Dim action As String = If(jsonDoc("action")?.ToString(), "")
            Dim proofreadMode = GetProofreadFocusMode()

            Select Case action
                Case "accept"
                    Dim issueId As String = If(jsonDoc("issueId")?.ToString(), "")
                    Task.Run(Async Function()
                                 Await proofreadMode.AcceptCorrectionAsync(issueId)
                             End Function)

                Case "ignore"
                    Dim issueId As String = If(jsonDoc("issueId")?.ToString(), "")
                    Task.Run(Async Function()
                                 Await proofreadMode.IgnoreIssueAsync(issueId)
                             End Function)

                Case "acceptAll"
                    Task.Run(Async Function()
                                 Await proofreadMode.AcceptAllAsync()
                             End Function)

                Case "exit"
                    Task.Run(Async Function()
                                 Await proofreadMode.ExitAsync()
                             End Function)
            End Select
        Catch ex As Exception
            Debug.WriteLine($"HandleProofreadFocusMode 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 将修正应用到Word文档（限定在选区内搜索替换）
    ''' </summary>
    Private Async Function ApplyCorrectionToWord(original As String, corrected As String) As Task(Of Boolean)
        Return Await Task.Run(Function()
                                  Try
                                      Dim wordApp = Globals.ThisAddIn.Application
                                      If wordApp Is Nothing Then Return False

                                      ' 使用存储的选区范围进行替换，避免替换文档中其他位置
                                      Dim searchRange = TryCast(_proofreadSelectionRange, Microsoft.Office.Interop.Word.Range)
                                      If searchRange Is Nothing Then
                                          ' 回退到当前选中范围
                                          searchRange = TryCast(wordApp.Selection.Range, Microsoft.Office.Interop.Word.Range)
                                      End If
                                      If searchRange Is Nothing Then Return False

                                      ' 在选区内查找并替换
                                      With searchRange.Find
                                          .Text = original
                                          .Replacement.Text = corrected
                                          .Forward = True
                                          .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                                          .Format = False
                                          .MatchCase = True
                                          .MatchWholeWord = False

                                          Dim replaced As Boolean = .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne)
                                          Return replaced
                                      End With
                                  Catch ex As Exception
                                      Debug.WriteLine($"ApplyCorrectionToWord 出错: {ex.Message}")
                                      Return False
                                  End Try
                              End Function)
    End Function

    ''' <summary>
    ''' 执行校对分析（由Ribbon按钮调用）
    ''' </summary>
    Public Async Function ExecuteProofreadAsync(Optional plan As Services.ProofreadIntentPlan = Nothing,
                                                 Optional userRequest As String = "") As Task(Of Boolean)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            If wordApp Is Nothing Then Return False

            If plan Is Nothing Then
                Dim compiler As New Services.ProofreadIntentCompiler()
                plan = compiler.Compile(If(String.IsNullOrWhiteSpace(userRequest), "校对", userRequest), HasUsableProofreadSelection(wordApp))
            End If

            Dim selRange = ResolveProofreadRange(wordApp, plan)
            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(If(selRange.Text, "")) Then
                GlobalStatusStrip.ShowWarning("В текущем документе нет текста для проверки.")
                Return False
            End If

            ' 按段落收集文本
            Dim paragraphs As New List(Of String)()
            Dim sb As New StringBuilder()
            sb.AppendLine("Ниже приведено содержимое для проверки (с номерами абзацев):")
            sb.AppendLine("План проверки: " & plan.ToHumanReadableSummary())
            sb.AppendLine()

            For Each p In selRange.Paragraphs
                Dim paraText As String = If(p.Range.Text IsNot Nothing, p.Range.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)
                If Not String.IsNullOrWhiteSpace(paraText) Then
                    paragraphs.Add(paraText)
                    sb.AppendLine($"[Абзац {paragraphs.Count - 1}] {paraText}")
                End If
            Next

            If paragraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("В выбранном содержимом нет допустимых абзацев.")
                Return False
            End If

            ' 存储段落列表供结果处理使用
            _proofreadParagraphs = paragraphs
            ' 存储当前选区范围，供替换操作限定范围使用
            _proofreadSelectionRange = selRange

            ' 将选中文本传递到JS端，供校对模式下聊天时注入上下文
            Dim selectedTextPreview = selRange.Text
            If selectedTextPreview IsNot Nothing AndAlso selectedTextPreview.Length > 200 Then
                selectedTextPreview = selectedTextPreview.Substring(0, 200) & "..."
            End If
            Dim escapedText = selectedTextPreview.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, " ").Replace(vbLf, " ").Replace("""", "\""")
            Await ExecuteJavaScriptAsyncJS($"window.proofreadSelectedText = '{escapedText}';")

            ' 进入校对专注模式
            Dim proofreadMode = GetProofreadFocusMode()
            Await proofreadMode.EnterAsync()
            Await ExecuteJavaScriptAsyncJS($"updateProofreadPlanSummary({Newtonsoft.Json.JsonConvert.SerializeObject(plan.ToHumanReadableSummary())});")

            ' 显示加载中提示
            Await ExecuteJavaScriptAsyncJS("showProofreadModeIndicator(); showProofreadLoading();")
            GlobalStatusStrip.ShowInfo("Проверка выполняется, подождите... " & plan.ToHumanReadableSummary())

            ' 构建校对提示词
            Dim systemPrompt = ProofreadPromptBuilder.BuildFullDocumentPrompt(paragraphs) &
                vbCrLf & BuildProofreadPlanInstruction(plan)

            ' 发送校对请求
            Await Send(sb.ToString(), systemPrompt, False, "proofread")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteProofreadAsync 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка проверки: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function HasUsableProofreadSelection(wordApp As Object) As Boolean
        Try
            If wordApp Is Nothing OrElse wordApp.Selection Is Nothing OrElse wordApp.Selection.Range Is Nothing Then Return False
            Dim text = If(wordApp.Selection.Range.Text, "").Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
            Return text.Length > 0
        Catch
            Return False
        End Try
    End Function

    Private Function ResolveProofreadRange(wordApp As Object, plan As Services.ProofreadIntentPlan) As Microsoft.Office.Interop.Word.Range
        Try
            If wordApp Is Nothing OrElse wordApp.ActiveDocument Is Nothing Then Return Nothing

            Select Case plan.Scope
                Case Services.ProofreadTargetScope.Selection
                    If HasUsableProofreadSelection(wordApp) Then Return wordApp.Selection.Range
                    Return wordApp.ActiveDocument.Content

                Case Services.ProofreadTargetScope.CurrentParagraph
                    If wordApp.Selection IsNot Nothing AndAlso wordApp.Selection.Paragraphs.Count > 0 Then
                        Return wordApp.Selection.Paragraphs(1).Range
                    End If
                    Return wordApp.ActiveDocument.Content

                Case Services.ProofreadTargetScope.Document
                    Return wordApp.ActiveDocument.Content

                Case Else
                    If HasUsableProofreadSelection(wordApp) Then Return wordApp.Selection.Range
                    Return wordApp.ActiveDocument.Content
            End Select
        Catch ex As Exception
            Debug.WriteLine($"ResolveProofreadRange error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Function BuildProofreadPlanInstruction(plan As Services.ProofreadIntentPlan) As String
        Dim sb As New StringBuilder()
        sb.AppendLine()
        sb.AppendLine("【План текущей проверки】")
        sb.AppendLine(plan.ToHumanReadableSummary())
        sb.AppendLine("Проверяй строго только типы, перечисленные в плане; если в плане указаны все проблемы, выполни полную проверку.")
        If plan.ApplyMode = Services.ProofreadApplyMode.AutoApplyHighConfidence Then
            sb.AppendLine("Для высоконадёжных и низкорисковых опечаток/пунктуационных ошибок укажи в JSON severity=high; остальные предложения оставь со severity medium/low для подтверждения пользователем.")
        End If
        sb.AppendLine("Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть.")
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 处理校对结果（在校对完成后调用）
    ''' 需要在UI线程执行以安全操作Word COM对象
    ''' </summary>
    Public Sub ProcessProofreadResult(aiResponse As String, paragraphs As List(Of String))
        Dim ignored = ProcessProofreadResultAsync(aiResponse, paragraphs)
    End Sub

    Private Async Function ProcessProofreadResultAsync(aiResponse As String, paragraphs As List(Of String)) As Task
        Try
            Dim proofreadMode = GetProofreadFocusMode()
            If proofreadMode Is Nothing Then Return

            ' 获取选区起始偏移量（用于精确定位波浪线）
            Dim selStartOffset As Integer = 0
            Try
                Dim selRange = TryCast(_proofreadSelectionRange, Microsoft.Office.Interop.Word.Range)
                If selRange IsNot Nothing Then
                    selStartOffset = selRange.Start
                End If
            Catch
                selStartOffset = 0
            End Try

            ' 在UI线程异步执行校对分析，避免阻塞消息循环或跨线程访问COM对象。
            Await UiDispatcher.InvokeAsync(Me,
                Async Function()
                    Await proofreadMode.AnalyzeAsync(aiResponse, paragraphs, Globals.ThisAddIn.Application, selStartOffset)
                End Function)

            StartPendingPostProofreadFormatting()
        Catch ex As Exception
            Debug.WriteLine($"ProcessProofreadResult 出错: {ex.Message}")
        End Try
    End Function

    Private Sub StartPendingPostProofreadFormatting()
        Dim request = _pendingPostProofreadFormattingRequest
        If String.IsNullOrWhiteSpace(request) Then Return

        _pendingPostProofreadFormattingRequest = ""

        If Me.InvokeRequired Then
            Me.BeginInvoke(New System.Action(Sub() ContinuePostProofreadFormattingAsync(request)))
        Else
            ContinuePostProofreadFormattingAsync(request)
        End If
    End Sub

    Private Async Sub ContinuePostProofreadFormattingAsync(formattingRequest As String)
        Try
            Dim responseUuid = Guid.NewGuid().ToString()
            Await ExecuteJavaScriptAsyncJS($"createChatSection('AI-помощник по форматированию', formatDateTime(new Date()), '{responseUuid}');")
            Await ExecuteJavaScriptAsyncJS($"appendRenderer('{responseUuid}','Проверка завершена, продолжаем задачи форматирования.');")

            If Services.SmartFormatter.LooksLikeDirectFormattingCommand(formattingRequest) Then
                Dim applied = Await ExecuteDirectFormattingCommandAsync(formattingRequest, False)
                If applied Then Return
            End If

            If IsExplicitFormattingRequest(formattingRequest) Then
                HandleChatDrivenReformat(formattingRequest, Nothing, Nothing)
            Else
                Await Send(formattingRequest, "", True, "")
            End If
        Catch ex As Exception
            Debug.WriteLine($"ContinuePostProofreadFormattingAsync 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка выполнения последующей задачи форматирования: {ex.Message}")
        End Try
    End Sub

    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' 订阅Word的SelectionChange 事件
        ' 帮我补全word选择的内容事件
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent

        ' 初始化智能格式化服务
        Try
            _smartFormatter = New Services.SmartFormatter(Globals.ThisAddIn.Application)
            _paragraphService = New Services.ParagraphService(Globals.ThisAddIn.Application)
        Catch ex As Exception
            Debug.WriteLine($"初始化服务失败: {ex.Message}")
        End Try
    End Sub

    ' 智能格式化服务实例
    Private _smartFormatter As Services.SmartFormatter
    ' 段落服务实例
    Private _paragraphService As Services.ParagraphService

    '获取选中的内容
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' 转换为 Word.Selection 对象
            Dim selection = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Word.Selection)
            If selection Is Nothing Then
                Return
            End If

            ' 检查是否有实际选中内容（通过比较Start和End位置）
            If selection.Start = selection.End Then
                ' 光标在单一位置，没有选中内容，清除之前的选中显示
                ClearSelectedContentBySheetName("Word文档")
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 检查是否选中了表格
            If selection.Tables.Count > 0 Then
                ' 如果选中的是表格
                Dim table = selection.Tables(1)
                Dim sb As New StringBuilder()

                ' 遍历表格内容
                For row As Integer = 1 To table.Rows.Count
                    For col As Integer = 1 To table.Columns.Count
                        sb.Append(table.Cell(row, col).Range.Text.TrimEnd(ChrW(13), ChrW(7)))
                        If col < table.Columns.Count Then sb.Append(vbTab)
                    Next
                    sb.AppendLine()
                Next
                content = sb.ToString()

            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' 如果选中的是图片或形状
                content = "[Изображение или фигура]"
            Else
                ' 普通文本选择
                content = selection.Text
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "Word文档",  ' 使用文档名称作为标识
                If(selection.Tables.Count > 0,
                   "[Содержимое таблицы]",
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
            )
            Else
                ClearSelectedContentBySheetName("Word文档")
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取Word选中内容时出错: {ex.Message}")
        End Try
    End Sub


    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    ' 返回应用信息
    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    ' 返回Office应用类型
    Protected Overrides Function GetOfficeAppType() As String
        Return "Word"
    End Function

    ''' <summary>
    ''' 获取样式预览回调（Word实现：将样式应用到选中的文本）
    ''' </summary>
    Protected Overrides Function GetStylePreviewCallback() As PreviewStyleCallback
        Return AddressOf ApplyStylePreviewToSelection
    End Function

    ''' <summary>
    ''' 显示模板编辑器面板（使用 CustomTaskPane）
    ''' </summary>
    Protected Overrides Function ShowTemplateEditorPane(template As ReformatTemplate) As Boolean
        Try
            Globals.ThisAddIn.ShowTemplateEditorTaskPane(template)
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ShowTemplateEditorPane 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 将样式预览应用到Word中选中的文本（Public供外部调用）
    ''' </summary>
    Public Sub ApplyStylePreviewToSelection(fontName As String, fontSize As Double, bold As Boolean, alignment As String, firstLineIndent As Double, lineSpacing As Double)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim doc = wordApp.ActiveDocument
            If doc Is Nothing Then Return

            Dim selection = wordApp.Selection
            Dim previewRange As Word.Range = Nothing

            ' 检查是否有有效的选中内容
            Dim hasValidSelection = selection IsNot Nothing AndAlso
                                   selection.Type <> Word.WdSelectionType.wdNoSelection AndAlso
                                   Not String.IsNullOrWhiteSpace(selection.Text?.Replace(vbCr, "").Replace(vbLf, ""))

            If hasValidSelection Then
                ' 使用选中的文本
                previewRange = selection.Range
            Else
                ' 没有选中文本，查找或创建预览段落
                Dim previewMarker = "【Предпросмотр стиля】"
                Dim found = False

                ' 查找已有的预览段落
                For Each para As Word.Paragraph In doc.Paragraphs
                    If para.Range.Text.StartsWith(previewMarker) Then
                        previewRange = para.Range
                        found = True
                        Exit For
                    End If
                Next

                ' 如果没有找到，在文档末尾创建预览段落
                If Not found Then
                    Dim endRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                    endRange.InsertParagraphAfter()
                    endRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                    endRange.Text = previewMarker & "Это пример текста для предпросмотра стиля; здесь видны шрифт, размер, выравнивание и другие эффекты."
                    previewRange = doc.Paragraphs.Last.Range
                End If
            End If

            If previewRange Is Nothing Then Return

            ' 应用字体
            If Not String.IsNullOrEmpty(fontName) Then
                previewRange.Font.Name = fontName
                previewRange.Font.NameFarEast = fontName
            End If

            ' 应用字号（磅值）
            If fontSize > 0 Then
                previewRange.Font.Size = CSng(fontSize)
            End If

            ' 应用加粗
            previewRange.Font.Bold = If(bold, 1, 0)

            ' 应用对齐方式
            Select Case alignment?.ToLower()
                Case "left"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                Case "center"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Case "right"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                Case "justify"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
            End Select

            ' 应用首行缩进（字符数转换为磅值）
            If firstLineIndent > 0 AndAlso fontSize > 0 Then
                previewRange.ParagraphFormat.FirstLineIndent = CSng(firstLineIndent * fontSize)
            Else
                previewRange.ParagraphFormat.FirstLineIndent = 0
            End If

            ' 应用行距
            If lineSpacing > 0 Then
                previewRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
                previewRange.ParagraphFormat.LineSpacing = CSng(lineSpacing * 12)
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyStylePreviewToSelection 出错: {ex.Message}")
        End Try
    End Sub

    ' 返回 Word Application 对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function

    ' 返回当前文档的 VBProject（可能为 Nothing）
    Protected Overrides Function GetVBProject() As VBProject
        Try
            Return Globals.ThisAddIn.Application.ActiveDocument.VBProject
        Catch
            Return Nothing
        End Try
    End Function

    ' 预览运行：展示代码并询问是否继续（返回 True 执行）
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
        If Not preview Then Return True
        Dim prompt As String = "Предпросмотр выполняемого кода VBA. Продолжить?" & vbCrLf & "----" & vbCrLf & vbaCode
        Return (MessageBox.Show(prompt, "Предпросмотр VBA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)
    End Function

    ' 真正执行宏（通过 Application.Run 调用模块.过程）
    Protected Overrides Function RunCode(vbaCode As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(vbaCode)
        Catch ex As Exception
            MessageBox.Show("Ошибка выполнения макроса: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    ' 将要发送到 LLM 的消息委托到底层 Send 方法（异步）
    Protected Overrides Sub SendChatMessage(message As String)
        Task.Run(Async Function()
                     Await Send(message, "", True, "")
                 End Function)
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息（重写基类方法）
    ''' Word子类拦截校对/排版意图，路由到专业流程
    ''' </summary>
    Protected Overrides Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        If intent IsNot Nothing AndAlso intent.Confidence > 0.4 Then
            ' 拦截校对意图 → 路由到校对专注模式
            If intent.OfficeIntent = OfficeIntentType.PROOFREAD Then
                Debug.WriteLine("[Word] 检测到校对意图，路由到ExecuteProofreadAsync")
                Task.Run(Async Function()
                             Await RouteProofreadIntentAsync(message, intent)
                         End Function)
                Return
            End If

            ' 拦截排版意图 → 路由到智能排版流程
            If intent.OfficeIntent = OfficeIntentType.FORMAT_STYLE OrElse
               intent.OfficeIntent = OfficeIntentType.TEXT_FORMAT Then
                Debug.WriteLine("[Word] 检测到排版意图，路由到计划化排版流程")
                Task.Run(Async Function()
                             Await RouteFormattingIntentAsync(message, intent)
                         End Function)
                Return
            End If
        End If

        ' 其他意图走默认流程
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"Word使用意图优化提示词: {intent.IntentType}, 置信度: {intent.Confidence:F2}")

            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

    ' 解析 Word 文件为文本（用于 file 引用）
    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' 创建一个新的Word应用程序实例（避免影响当前文档）
            Dim wordApp As New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
            Try
                doc = wordApp.Documents.Open(FileName:=filePath, ReadOnly:=True, Visible:=False)
                Dim contentBuilder As New StringBuilder()

                contentBuilder.AppendLine($"Файл: {Path.GetFileName(filePath)}")
                contentBuilder.AppendLine($"Всего абзацев: {doc.Paragraphs.Count}")
                contentBuilder.AppendLine()

                ' 限制处理的段落数量
                Dim maxParagraphs As Integer = Math.Min(doc.Paragraphs.Count, 50)
                Dim paraIndex As Integer = 0

                For Each para As Microsoft.Office.Interop.Word.Paragraph In doc.Paragraphs
                    paraIndex += 1
                    If paraIndex > maxParagraphs Then Exit For

                    Dim text As String = para.Range.Text.Trim()
                    If Not String.IsNullOrEmpty(text) AndAlso text <> vbCr Then
                        ' 获取段落样式
                        Dim styleName As String = ""
                        Try
                            styleName = para.Style.NameLocal
                        Catch
                        End Try

                        ' 判断是否是标题
                        Dim prefix As String = $"Абзац {paraIndex}"
                        If styleName.Contains("标题") OrElse styleName.ToLower().Contains("heading") OrElse styleName.ToLower().Contains("заголов") Then
                            prefix = $"[{styleName}]"
                        End If

                        contentBuilder.AppendLine($"{prefix}: {text}")
                    End If
                Next

                ' 处理表格
                If doc.Tables.Count > 0 Then
                    contentBuilder.AppendLine()
                    contentBuilder.AppendLine($"=== Документ содержит таблиц: {doc.Tables.Count} ===")

                    Dim tableIndex As Integer = 0
                    For Each tbl As Microsoft.Office.Interop.Word.Table In doc.Tables
                        tableIndex += 1
                        If tableIndex > 5 Then Exit For ' 限制表格数量

                        contentBuilder.AppendLine($"Таблица {tableIndex}: {tbl.Rows.Count} строк × {tbl.Columns.Count} столбцов")

                        ' 读取表格前几行
                        Dim maxRows = Math.Min(tbl.Rows.Count, 5)
                        For rowIdx = 1 To maxRows
                            Dim rowContent As New StringBuilder("  ")
                            For colIdx = 1 To tbl.Columns.Count
                                Try
                                    Dim cellText = tbl.Cell(rowIdx, colIdx).Range.Text.Trim()
                                    cellText = cellText.Replace(vbCr, "").Replace(Chr(7), "")
                                    If cellText.Length > 20 Then cellText = cellText.Substring(0, 17) & "..."
                                    rowContent.Append(cellText & " | ")
                                Catch
                                End Try
                            Next
                            contentBuilder.AppendLine(rowContent.ToString().TrimEnd(" |".ToCharArray()))
                        Next
                        contentBuilder.AppendLine()
                    Next
                End If

                If doc.Paragraphs.Count > maxParagraphs Then
                    contentBuilder.AppendLine()
                    contentBuilder.AppendLine($"... всего абзацев: {doc.Paragraphs.Count}, показаны только первые {maxParagraphs}")
                End If

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "Word",
                    .ParsedContent = contentBuilder.ToString(),
                    .RawData = Nothing
                }

            Finally
                If doc IsNot Nothing Then
                    doc.Close(False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc)
                End If
                wordApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            Debug.WriteLine($"解析Word文件时出错: {ex.Message}")
            Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Word",
                .ParsedContent = $"[Ошибка при разборе файла Word: {ex.Message}]"
            }
        End Try
    End Function

    ' 返回当前文档所在目录（未保存返回空字符串）
    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            Dim p = Globals.ThisAddIn.Application.ActiveDocument.Path
            If String.IsNullOrEmpty(p) Then Return String.Empty
            Return p
        Catch
            Return String.Empty
        End Try
    End Function

    ' 将当前选区内容附加到提示，并记录 PendingSelectionInfo 供写回使用
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)

            Dim info As New SelectionInfo()
            info.DocumentPath = If(Globals.ThisAddIn.Application.ActiveDocument.Path, "")
            info.SelectedText = txt
            Try
                info.StartPos = sel.Range.Start
                info.EndPos = sel.Range.End
            Catch
                info.StartPos = 0
                info.EndPos = 0
            End Try

            PendingSelectionInfo = info

            If String.IsNullOrWhiteSpace(txt) Then
                Return message
            Else
                Return message & vbCrLf & vbCrLf & txt
            End If
        Catch
            Return message
        End Try
    End Function


    ' 修订、审阅功能（简化版：使用段落索引定位）
    Protected Overrides Sub HandleApplyRevisionSegment(jsonDoc As JObject)
        Try
            ' 期望收到字段： uuid, paraIndex, original, corrected
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim paraIndex As Integer = If(jsonDoc("paraIndex") IsNot Nothing, CInt(jsonDoc("paraIndex")), -1)
            Dim original As String = If(jsonDoc("original") IsNot Nothing, jsonDoc("original").ToString(), String.Empty)
            Dim corrected As String = If(jsonDoc("corrected") IsNot Nothing, jsonDoc("corrected").ToString(), String.Empty)

            If paraIndex < 0 Then
                GlobalStatusStrip.ShowWarning("Отсутствует параметр paraIndex")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("Проверка поддерживается только в Word")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("Не удалось получить объект приложения Word")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range

            ' 使用选中范围内的段落索引定位
            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
                GlobalStatusStrip.ShowWarning("Сначала выберите содержимое для проверки")
                Return
            End If

            ' 获取选中范围内的段落
            Dim paragraphs = selRange.Paragraphs
            If paraIndex >= paragraphs.Count Then
                GlobalStatusStrip.ShowWarning($"Индекс абзаца {paraIndex} вне диапазона")
                Return
            End If

            ' 定位目标段落（段落索引从1开始）
            Dim targetPara = paragraphs(paraIndex + 1)
            Dim targetRange = targetPara.Range

            ' 在目标段落中查找并替换原文
            If Not String.IsNullOrEmpty(original) Then
                Dim paraText As String = targetRange.Text
                Dim startPos As Integer = paraText.IndexOf(original, StringComparison.Ordinal)
                If startPos >= 0 Then
                    ' 创建精确的替换范围
                    Dim replaceRange = doc.Range(targetRange.Start + startPos, targetRange.Start + startPos + original.Length)

                    ' 开启审阅模式
                    Try
                        doc.TrackRevisions = True
                    Catch
                    End Try

                    ' 执行替换
                    replaceRange.Text = corrected
                    GlobalStatusStrip.ShowInfo($"Содержимое абзаца {paraIndex} заменено (режим рецензирования)")
                Else
                    GlobalStatusStrip.ShowWarning($"В абзаце {paraIndex} не найден исходный текст: {original}")
                End If
            Else
                GlobalStatusStrip.ShowWarning("Отсутствует исходный текст")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionSegment 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("Ошибка записи результатов проверки: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub HandleApplyRevisionAccept(jsonDoc As JObject)
        Try
            ' 期望 { type:'applyRevisionAccept', responseUuid:..., globalIndex: n }
            Dim responseUuid As String = If(jsonDoc("responseUuid") IsNot Nothing, jsonDoc("responseUuid").ToString(), If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty))
            Dim globalIndex As Integer = If(jsonDoc("globalIndex") IsNot Nothing, CInt(jsonDoc("globalIndex")), -1)

            If globalIndex < 0 Then
                GlobalStatusStrip.ShowWarning("applyRevisionAccept: отсутствует globalIndex")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("Принятие отдельной правки поддерживается только в Word (реализация по умолчанию)")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("Не удалось получить объект приложения Word; принятие правки не выполнено")
                Return
            End If

            Try
                Dim doc = officeApp.ActiveDocument
                ' Word Revisions 集合是 1 基的；尝试保护性调用
                If globalIndex >= 1 And globalIndex <= doc.Revisions.Count Then
                    doc.Revisions(globalIndex).Accept()
                    GlobalStatusStrip.ShowInfo($"Правка #{globalIndex} принята")
                Else
                    GlobalStatusStrip.ShowWarning("Указанный индекс правки вне диапазона или не существует")
                End If
            Catch ex As Exception
                Debug.WriteLine("接受修订失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("Не удалось принять правку: " & ex.Message)
            End Try
        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionAccept 出错: {ex.Message}")
        End Try
    End Sub

    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)

        ' 如果此次会话绑定了选区信息，则发送对比预览（原文 vs AI 输出）
        Try
            ' 使用 response->request 的映射查找对应的选区信息（修正原有逻辑中使用 _finalUuid 直接查找的错误）
            Dim requestId As String = Nothing
            If _responseToRequestMap.ContainsKey(_finalUuid) Then
                requestId = _responseToRequestMap(_finalUuid)
            End If

            Dim mode As String = ""
            If _responseModeMap.ContainsKey(_finalUuid) Then
                mode = _responseModeMap(_finalUuid)
            End If

            ' 语义标注排版模式：AI返回标注JSON，由渲染引擎确定性应用格式
            If String.Equals(mode, "semantic_reformat", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    ApplySemanticTaggingResult(aiText)
                Catch ex As Exception
                    Debug.WriteLine("语义标注排版处理失败: " & ex.Message)
                End Try
                MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
                Return
            End If

            ' 规范转换模式：AI返回规范的结构化JSON，解析后进入语义标注阶段
            If String.Equals(mode, "styleguide_convert", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    HandleStyleGuideConversionResult(aiText)
                Catch ex As Exception
                    Debug.WriteLine("规范转换处理失败: " & ex.Message)
                End Try
                MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
                Return
            End If

            ' 格式克隆模式：AI返回 SemanticStyleMapping JSON，直接保存并推送前端预览
            If String.Equals(mode, "mirror_format", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    HandleMirrorFormatResult(aiText)
                Catch ex As Exception
                    Debug.WriteLine("格式克隆处理失败: " & ex.Message)
                End Try
                MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
                Return
            End If

            ' 如果是排版重构动作，则触发 showComparison
            If _responseSelectionMap.ContainsKey(_finalUuid) AndAlso String.Equals(mode, "reformat", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim selInfo = _responseSelectionMap(_finalUuid)
                    Dim originalText As String = If(selInfo?.SelectedText, "")
                    Dim aiFinal As String = allPlainMarkdownBuffer.ToString()

                    Dim js As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(aiFinal)});"
                    ExecuteJavaScriptAsyncJS(js)
                Catch ex As Exception
                    Debug.WriteLine("尝试解析并发送 comparison 时出错: " & ex.Message)
                End Try
            End If

            ' 如果是校对动作，调用校对专注模式处理结果
            If String.Equals(mode, "proofread", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    ' 调用校对专注模式处理结果（内联标注 + 侧边面板）
                    If _proofreadParagraphs IsNot Nothing AndAlso _proofreadParagraphs.Count > 0 Then
                        ProcessProofreadResult(aiText, _proofreadParagraphs)
                    Else
                        ' 如果没有段落列表，只显示前端 revisions
                        Dim revisions As JArray = TryExtractJsonArrayFromText(aiText)
                        If revisions IsNot Nothing AndAlso revisions.Count > 0 Then
                            _revisionsMap(_finalUuid) = revisions
                            Dim jsRev As String = $"showRevisions('{_finalUuid}', {revisions.ToString(Formatting.None)});"
                            ExecuteJavaScriptAsyncJS(jsRev)
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine("校对结果处理出错: " & ex.Message)
                End Try
            End If

            ' 解析并发送文档计划或预览 HTML 给前端，作为唯一内容
            If String.Equals(mode, "documentPlan", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(mode, "previewHtml", StringComparison.OrdinalIgnoreCase) Then
                Try
                    ' 尝试直接解析 JSON 对象（可能是 documentPlan 数组 / previewHtml / previewHtmlMap / 单个 planItem）
                    Dim rawText As String = allPlainMarkdownBuffer.ToString()
                    Dim jsonObj As JObject = TryExtractJsonObjectFromText(rawText)

                    If jsonObj IsNot Nothing Then
                        ' 如果后端/模型仅返回单个 planItem（键为 planItem），将其包装为 documentPlan 数组以便前端统一处理
                        Dim sendObj As JObject = Nothing
                        If jsonObj("planItem") IsNot Nothing Then
                            Dim arr As New JArray()
                            arr.Add(jsonObj("planItem"))
                            sendObj = New JObject()
                            sendObj("documentPlan") = arr
                            ' 若同时包含 previewHtmlMap，保留之
                            If jsonObj("previewHtmlMap") IsNot Nothing Then
                                sendObj("previewHtmlMap") = jsonObj("previewHtmlMap")
                            End If
                            ' 若 planItem 自身已包含 previewHtmlMap（极少见），合并也可按需处理
                        Else
                            ' 直接使用解析到的对象：可能含 documentPlan、previewHtml、previewHtmlMap 等
                            sendObj = jsonObj
                        End If

                        ' 获取原始选区文本（若存在）
                        Dim originalText As String = ""
                        If _responseSelectionMap.ContainsKey(_finalUuid) Then
                            Dim selInfo = _responseSelectionMap(_finalUuid)
                            originalText = If(selInfo?.SelectedText, "")
                        End If

                        ' 将整个对象序列化为字符串后传给前端的 showComparison，前端会解析 previewHtmlMap 或 documentPlan
                        Dim payload As String = sendObj.ToString(Formatting.None)
                        Dim jsPlan As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(payload)});"
                        ExecuteJavaScriptAsyncJS(jsPlan)
                    Else
                        ' 退回尝试解析为 JSON 数组（旧版可能只返回数组）
                        Dim planArr As JArray = TryExtractJsonArrayFromText(rawText)
                        If planArr IsNot Nothing AndAlso planArr.Count > 0 Then
                            Dim wrapper As New JObject()
                            wrapper("documentPlan") = planArr

                            Dim originalText As String = ""
                            If _responseSelectionMap.ContainsKey(_finalUuid) Then
                                Dim selInfo = _responseSelectionMap(_finalUuid)
                                originalText = If(selInfo?.SelectedText, "")
                            End If

                            Dim payload As String = wrapper.ToString(Formatting.None)
                            Dim jsPlan As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(payload)});"
                            ExecuteJavaScriptAsyncJS(jsPlan)
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine("处理 documentPlan/previewHtml 失败: " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"发送对比预览失败: {ex.Message}")
        End Try

        ' 调用基类处理续写模式
        MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
    End Sub

    ' 提取文本中第一个 JSON 数组（严格数组格式），返回 JArray 或 Nothing
    Private Function TryExtractJsonArrayFromText(text As String) As JArray
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            ' 尝试用正则抽取第一个 [...] 数组块（Singleline 允许跨行）
            Dim m As Match = Regex.Match(text, "\[.*\]", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                ' 验证并解析为 JArray
                Try
                    Dim arr As JArray = JArray.Parse(jsonCandidate)
                    Return arr
                Catch ex As Exception
                    Debug.WriteLine("解析 JSON 数组失败: " & ex.Message)
                    Return Nothing
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("TryExtractJsonArrayFromText 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' 提取文本中第一个 JSON 对象（如 {"documentPlan":..., "previewHtml":...}），返回 JObject 或 Nothing
    Private Function TryExtractJsonObjectFromText(text As String) As JObject
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            ' 尝试用正则抽取第一个 { ... } 对象块（Singleline 允许跨行）
            Dim m As Match = Regex.Match(text, "\{[\s\S]*\}", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                ' 验证并解析为 JObject
                Try
                    Dim obj As JObject = JObject.Parse(jsonCandidate)
                    Return obj
                Catch ex As Exception
                    Debug.WriteLine("解析 JSON 对象失败: " & ex.Message)
                    Return Nothing
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("TryExtractJsonObjectFromText 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' 排版功能（支持语义标注模式、规则模式和旧的逐段落模式）
    Protected Overrides Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
        Try
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)

            ' 语义标注模式（新流水线）：有tags字段
            If jsonDoc("tags") IsNot Nothing AndAlso jsonDoc("tags").Type = JTokenType.Array Then
                ApplySemanticTaggingResult(jsonDoc("tags").ToString(Newtonsoft.Json.Formatting.None))
                Return
            End If

            ' 检测是否为新的规则模式（有rules字段）
            If jsonDoc("rules") IsNot Nothing AndAlso jsonDoc("rules").Type = JTokenType.Array Then
                ApplyReformatRules(jsonDoc)
                Return
            End If

            ' 旧模式：逐段落格式化（保留向后兼容）
            Dim paraIndex As Integer = If(jsonDoc("paraIndex") IsNot Nothing, CInt(jsonDoc("paraIndex")), -1)
            Dim formatting As JObject = Nothing
            If jsonDoc("formatting") IsNot Nothing Then
                formatting = DirectCast(jsonDoc("formatting"), JObject)
            End If

            If paraIndex < 0 Then
                GlobalStatusStrip.ShowWarning("Отсутствует параметр paraIndex")
                Return
            End If

            If formatting Is Nothing Then
                GlobalStatusStrip.ShowWarning("Отсутствует параметр formatting")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("Форматирование поддерживается только в Word")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("Не удалось получить объект приложения Word")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range

            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
                GlobalStatusStrip.ShowWarning("Сначала выберите содержимое для форматирования")
                Return
            End If

            ' 获取选中范围内的段落
            Dim paragraphs = selRange.Paragraphs
            If paraIndex >= paragraphs.Count Then
                GlobalStatusStrip.ShowWarning($"Индекс абзаца {paraIndex} вне диапазона")
                Return
            End If

            ' 定位目标段落
            Dim targetPara = paragraphs(paraIndex + 1)
            Dim targetRange = targetPara.Range

            ' 使用Word对象模型应用格式化
            Try
                ApplyFormattingToRange(targetRange, formatting)
                GlobalStatusStrip.ShowInfo($"Форматирование абзаца {paraIndex} применено")
            Catch ex As Exception
                Debug.WriteLine("排版写回失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("Ошибка записи форматирования: " & ex.Message)
            End Try

        Catch ex As Exception
            Debug.WriteLine("HandleApplyDocumentPlanItem 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("Ошибка применения форматирования: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 应用规则模式的排版（优化版：减少token消耗）
    ''' </summary>
    Private Sub ApplyReformatRules(jsonDoc As JObject)
        Try
            Dim rules = jsonDoc("rules").ToObject(Of List(Of JObject))()
            Dim sampleClassification = jsonDoc("sampleClassification")?.ToObject(Of List(Of JObject))()

            If rules Is Nothing OrElse rules.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Не получены допустимые правила форматирования")
                Return
            End If

            ' 构建规则字典
            Dim ruleDict As New Dictionary(Of String, JObject)()
            For Each rule In rules
                Dim ruleType = rule("type")?.ToString()
                If Not String.IsNullOrEmpty(ruleType) AndAlso rule("formatting") IsNot Nothing Then
                    ruleDict(ruleType) = DirectCast(rule("formatting"), JObject)
                End If
            Next

            ' 如果没有保存的段落上下文，使用当前选中内容
            If _reformatParagraphs Is Nothing OrElse _reformatParagraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Контекст форматирования потерян. Выберите содержимое и запустите форматирование заново")
                Return
            End If

            ' 基于样本分类推断所有段落的规则
            Dim sampleRuleMap As New Dictionary(Of Integer, String)()
            If sampleClassification IsNot Nothing Then
                For Each sc In sampleClassification
                    Dim idx = sc("sampleIndex")?.ToObject(Of Integer)()
                    Dim appliedRule = sc("appliedRule")?.ToString()
                    If idx IsNot Nothing AndAlso Not String.IsNullOrEmpty(appliedRule) Then
                        sampleRuleMap(idx) = appliedRule
                    End If
                Next
            End If

            ' 应用格式到所有段落
            Dim appliedCount As Integer = 0
            Dim skippedCount As Integer = 0
            Dim defaultRule As String = If(ruleDict.ContainsKey("body"), "body", ruleDict.Keys.FirstOrDefault())

            For i As Integer = 0 To _reformatParagraphs.Count - 1
                Try
                    ' 检查段落类型，跳过非文本元素
                    Dim paraType As String = "text"
                    If _reformatTypes IsNot Nothing AndAlso i < _reformatTypes.Count Then
                        paraType = _reformatTypes(i)
                    End If

                    If paraType <> "text" Then
                        ' 跳过图片、表格、公式等非文本元素
                        skippedCount += 1
                        Continue For
                    End If

                    Dim para = _reformatParagraphs(i)
                    Dim styleName = If(i < _reformatStyles.Count, _reformatStyles(i), "")

                    ' 确定使用哪个规则
                    Dim ruleToApply As String = defaultRule

                    ' 先检查是否有样本分类指定
                    If sampleRuleMap.ContainsKey(i) Then
                        ruleToApply = sampleRuleMap(i)
                    Else
                        ' 基于样式名推断规则
                        If styleName.Contains("标题") OrElse styleName.ToLower().Contains("heading") OrElse styleName.ToLower().Contains("заголов") Then
                            ' 找到第一个标题类规则
                            For Each key In ruleDict.Keys
                                If key.ToLower().Contains("title") OrElse key.ToLower().Contains("heading") Then
                                    ruleToApply = key
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    ' 应用规则
                    If ruleDict.ContainsKey(ruleToApply) Then
                        Dim formatting = ruleDict(ruleToApply)
                        ApplyFormattingToRange(para.Range, formatting)
                        appliedCount += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"应用段落 {i} 格式失败: " & ex.Message)
                End Try
            Next

            ' 清理上下文
            _reformatParagraphs = Nothing
            _reformatStyles = Nothing
            _reformatTypes = Nothing

            Dim resultMsg = $"Форматирование завершено, обработано текстовых абзацев: {appliedCount}"
            If skippedCount > 0 Then
                resultMsg &= $", пропущено специальных элементов: {skippedCount}"
            End If
            GlobalStatusStrip.ShowInfo(resultMsg)

        Catch ex As Exception
            Debug.WriteLine("ApplyReformatRules 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("Ошибка применения правил форматирования: " & ex.Message)
        End Try
    End Sub

    ' 使用Word对象模型应用格式化属性
    Private Sub ApplyFormattingToRange(targetRange As Object, formatting As JObject)
        Try
            ' 字体名称（中文）
            If formatting("fontNameCN") IsNot Nothing Then
                Dim fontNameCN As String = formatting("fontNameCN").ToString()
                If Not String.IsNullOrEmpty(fontNameCN) Then
                    Try
                        targetRange.Font.NameFarEast = fontNameCN
                    Catch
                        ' 某些 Word 版本可能不支持 NameFarEast
                    End Try
                End If
            End If

            ' 字体名称（英文/西文）
            If formatting("fontNameEN") IsNot Nothing Then
                Dim fontNameEN As String = formatting("fontNameEN").ToString()
                If Not String.IsNullOrEmpty(fontNameEN) Then
                    Try
                        targetRange.Font.Name = fontNameEN
                    Catch
                    End Try
                End If
            End If

            ' 字号
            If formatting("fontSize") IsNot Nothing Then
                Dim fontSize As Single = 0
                Single.TryParse(formatting("fontSize").ToString(), fontSize)
                If fontSize > 0 Then
                    Try
                        targetRange.Font.Size = fontSize
                    Catch
                    End Try
                End If
            End If

            ' 加粗
            If formatting("bold") IsNot Nothing Then
                Try
                    Dim bold As Boolean = formatting("bold").ToObject(Of Boolean)()
                    targetRange.Font.Bold = If(bold, -1, 0) ' Word: -1 = True, 0 = False
                Catch
                End Try
            End If

            ' 对齐方式
            If formatting("alignment") IsNot Nothing Then
                Dim alignment As String = formatting("alignment").ToString().ToLower()
                Try
                    Select Case alignment
                        Case "left"
                            targetRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                        Case "center"
                            targetRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                        Case "right"
                            targetRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                        Case "justify", "justified"
                            targetRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
                    End Select
                Catch
                End Try
            End If

            ' 首行缩进（字符数）
            If formatting("firstLineIndent") IsNot Nothing Then
                Dim indent As Single = 0
                Single.TryParse(formatting("firstLineIndent").ToString(), indent)
                If indent > 0 Then
                    Try
                        ' CharacterUnitFirstLineIndent 以字符为单位
                        targetRange.ParagraphFormat.CharacterUnitFirstLineIndent = indent
                    Catch
                        ' 回退：使用磅值（1字符约=10.5磅 for 五号字）
                        Try
                            targetRange.ParagraphFormat.FirstLineIndent = indent * 10.5
                        Catch
                        End Try
                    End Try
                End If
            End If

            ' 行距
            If formatting("lineSpacing") IsNot Nothing Then
                Dim lineSpacing As Single = 0
                Single.TryParse(formatting("lineSpacing").ToString(), lineSpacing)
                If lineSpacing > 0 Then
                    Try
                        ' LineSpacingRule: 0=wdLineSpaceSingle, 1=wdLineSpace1pt5, 2=wdLineSpaceDouble, 5=wdLineSpaceMultiple
                        If lineSpacing = 1.0 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 0 ' wdLineSpaceSingle
                        ElseIf lineSpacing = 1.5 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 1 ' wdLineSpace1pt5
                        ElseIf lineSpacing = 2.0 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 2 ' wdLineSpaceDouble
                        Else
                            ' 使用多倍行距
                            targetRange.ParagraphFormat.LineSpacingRule = 5 ' wdLineSpaceMultiple
                            targetRange.ParagraphFormat.LineSpacing = 12 * lineSpacing ' 12磅 * 倍数
                        End If
                    Catch
                    End Try
                End If
            End If

            ' 字体颜色
            If formatting("color") IsNot Nothing Then
                Dim colorStr As String = formatting("color").ToString()
                If Not String.IsNullOrEmpty(colorStr) AndAlso colorStr <> "auto" Then
                    Try
                        Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                        targetRange.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                    Catch ex As Exception
                        Debug.WriteLine("设置字体颜色失败: " & ex.Message)
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine("ApplyFormattingToRange 出错: " & ex.Message)
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' AI 语义标注过于保守时，用预览方案中的确定性结构标签覆盖对应段落。
    ''' 例如用户要求“重构序号和标题”，但模型全部返回 body.normal，仍应执行 plan 中的标题/编号候选。
    ''' </summary>
    Private Function MergePreviewPlanTags(aiTags As List(Of TaggedParagraph)) As List(Of TaggedParagraph)
        Dim merged As New Dictionary(Of Integer, TaggedParagraph)()

        If aiTags IsNot Nothing Then
            For Each tagged As TaggedParagraph In aiTags
                If tagged Is Nothing Then Continue For
                merged(tagged.ParaIndex) = New TaggedParagraph(tagged.ParaIndex, tagged.TagId, tagged.Reason)
            Next
        End If

        Dim plan = If(_activeReformatJob?.PreviewPlan, _activeReformatPlan)
        If plan Is Nothing OrElse plan.Changes Is Nothing Then
            Return merged.Values.OrderBy(Function(t) t.ParaIndex).ToList()
        End If

        Dim overrideCount As Integer = 0
        For Each change In plan.Changes
            If change Is Nothing OrElse change.ParagraphIndex < 0 Then Continue For
            If String.IsNullOrWhiteSpace(change.NewTag) Then Continue For
            If String.Equals(change.NewTag, SemanticTagRegistry.TAG_BODY_NORMAL, StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim paraType As String = "text"
            If _reformatTypes IsNot Nothing AndAlso change.ParagraphIndex < _reformatTypes.Count Then
                paraType = If(_reformatTypes(change.ParagraphIndex), "text")
            End If
            If Not String.Equals(paraType, "text", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim reason = If(String.IsNullOrWhiteSpace(change.ChangeDescription),
                            "Структурный абзац, определённый планом предпросмотра",
                            "План предпросмотра: " & change.ChangeDescription)
            merged(change.ParagraphIndex) = New TaggedParagraph(change.ParagraphIndex, change.NewTag, reason)
            overrideCount += 1
        Next

        If overrideCount > 0 Then
            Debug.WriteLine($"[Reformat] 使用预览方案覆盖 AI 标注: {overrideCount} 个结构段落")
        End If

        Return merged.Values.OrderBy(Function(t) t.ParaIndex).ToList()
    End Function

    ''' <summary>
    ''' 处理AI语义标注结果 - 渲染引擎确定性应用格式
    ''' </summary>
    Private Async Sub ApplySemanticTaggingResult(taggingJson As String)
        Try
            If _reformatMapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("Контекст сопоставления форматирования потерян")
                Return
            End If

            If _reformatParagraphs Is Nothing OrElse _reformatParagraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Контекст абзацев форматирования потерян")
                Return
            End If

            ' 校验AI标注结果
            Dim validation = TaggingValidator.Validate(taggingJson, _reformatMapping, _reformatParagraphs.Count)

            If Not validation.IsValid Then
                ' 严重错误：尝试重试
                Debug.WriteLine("语义标注校验失败，错误数: " & validation.Errors.Count)

                ' 检查重试次数限制
                _reformatRetryCount += 1
                If _reformatRetryCount > MAX_REFORMAT_RETRIES Then
                    ' 超过重试限制，显示错误给用户
                    Debug.WriteLine($"重试次数超过限制({MAX_REFORMAT_RETRIES})，停止重试")
                    GlobalStatusStrip.ShowWarning($"Не удалось разобрать разметку ИИ; повторных попыток: {MAX_REFORMAT_RETRIES}")
                    _reformatRetryCount = 0 ' 重置计数器

                    ' 显示错误详情
                    Dim errorMsg = String.Join(vbCrLf, validation.Errors.Take(5))
                    If validation.Errors.Count > 5 Then
                        errorMsg &= vbCrLf & $"...и ещё ошибок: {validation.Errors.Count - 5}"
                    End If
                    Await ShowReformatError($"Не удалось разобрать разметку ИИ:{vbCrLf}{errorMsg}")
                    Return
                End If

                ' 收集段落文本用于重试
                Dim paragraphTexts As New List(Of String)()
                For Each para In _reformatParagraphs
                    Try
                        paragraphTexts.Add(If(para.Range.Text?.ToString().TrimEnd(vbCr, vbLf), ""))
                    Catch
                        paragraphTexts.Add("")
                    End Try
                Next

                Debug.WriteLine($"第{_reformatRetryCount}次重试...")
                Dim retryPrompt = SemanticPromptBuilder.BuildRetryPrompt(_reformatMapping, paragraphTexts, validation.Errors)
                Await Send("В результатах разметки есть ошибки, исправь их.", retryPrompt, False, "semantic_reformat")
                Return
            End If

            ' 校验通过，重置重试计数器
            _reformatRetryCount = 0

            ' 使用ReformatCoordinator执行"临时文档 → OpenXML排版 → 预览 → 合并"流程
            Dim wordApp = Globals.ThisAddIn.Application
            Dim renderResult As SemanticRenderingEngine.RenderResult = Nothing
            Dim effectiveTags = MergePreviewPlanTags(validation.ValidatedTags)
            If _reformatMapping IsNot Nothing Then
                _reformatMapping.EnsureBaselineTags()
            End If

            ' 在应用前捕获格式快照（独立于Word UndoRecord的撤销机制）
            Try
                _reformatSnapshot = SemanticRenderingEngine.CaptureFormatSnapshot(_reformatParagraphs, _reformatTypes)
                _reformatSnapshotParagraphs = _reformatParagraphs
                _reformatSnapshotTypes = _reformatTypes
                Debug.WriteLine($"排版快照已捕获: {_reformatSnapshot.ParagraphCount} 个段落")
            Catch snapEx As Exception
                _reformatSnapshot = Nothing
                Debug.WriteLine($"捕获排版快照失败: {snapEx.Message}")
            End Try

            ' 使用新的协调器：临时文档 → OpenXML执行 → 预览 → 合并回原文档
            Dim coordinator As New ReformatCoordinator()
            Dim reformatResult = Await coordinator.ExecuteReformatPipelineAsync(
                wordApp.ActiveDocument,
                ReformatOperationType.Semantic,
                effectiveTags,
                _reformatMapping,
                _reformatMapping?.Name,
                _reformatParagraphs,
                _reformatTypes)

            Dim repairCount As Integer = 0
            If reformatResult.Success Then
                repairCount = RepairStructuralSemanticFormattingIfNeeded(effectiveTags, _reformatMapping)
                If repairCount > 0 Then
                    reformatResult.ModifiedCount += repairCount
                    Debug.WriteLine($"[ReformatObserve] 自动修复结构段落: {repairCount}")
                End If

                ' 构建兼容的RenderResult用于前端展示
                renderResult = New SemanticRenderingEngine.RenderResult() With {
                    .AppliedCount = reformatResult.ModifiedCount,
                    .SkippedCount = Math.Max(0, effectiveTags.Count - reformatResult.ModifiedCount)
                }
                If reformatResult.GeneratedInstructions IsNot Nothing Then
                    renderResult.GeneratedInstructions = reformatResult.GeneratedInstructions
                End If
            Else
                ' 用户取消或失败
                If String.IsNullOrEmpty(reformatResult.ErrorMessage) OrElse
                   reformatResult.ErrorMessage.Contains("取消") Then
                    GlobalStatusStrip.ShowInfo("Форматирование отменено")
                Else
                    Await ShowReformatError($"Сбой форматирования: {reformatResult.ErrorMessage}")
                End If
                _reformatParagraphs = Nothing
                _reformatStyles = Nothing
                _reformatTypes = Nothing
                Return
            End If

            ' 推送排版结果到前端（含内置撤销按钮）
            If renderResult IsNot Nothing Then
                Dim resultJson = renderResult.ToJson()
                Await ExecuteJavaScriptAsyncJS($"showReformatResult({resultJson.ToString(Formatting.None)});")
                Dim observationSummary = BuildSemanticReformatObservationSummary(effectiveTags, _reformatMapping, reformatResult, validation.AutoFixedCount, repairCount)
                If Not String.IsNullOrWhiteSpace(observationSummary) Then
                    Await ShowSemanticReformatObservation(observationSummary)
                End If
            Else
                ' ApplySemanticFormatting 返回了 Nothing（理论上不应该）
                Await ShowReformatError("Движок отрисовки форматирования вернул пустой результат")
                ' 清理但不抛异常
                _reformatParagraphs = Nothing
                _reformatStyles = Nothing
                _reformatTypes = Nothing
                Return
            End If

            ' 显示状态
            Dim resultMsg = $"Форматирование завершено, обработано абзацев: {renderResult.AppliedCount}"
            If renderResult.SkippedCount > 0 Then
                resultMsg &= $", пропущено специальных элементов: {renderResult.SkippedCount}"
            End If
            If validation.AutoFixedCount > 0 Then
                resultMsg &= $", автоматически исправлено тегов: {validation.AutoFixedCount}"
            End If
            GlobalStatusStrip.ShowInfo(resultMsg)

            ' 清理段落引用（段落对象会失效），但保留映射上下文以支持后续排版
            _reformatParagraphs = Nothing
            _reformatStyles = Nothing
            _reformatTypes = Nothing
            _activeReformatJob = Nothing
            ' _reformatMapping 保留，用户可继续使用同一映射排版其他内容

        Catch ex As Exception
            Debug.WriteLine($"ApplySemanticTaggingResult 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка применения семантического форматирования: {ex.Message}")
        End Try
    End Sub

    Private Function RepairStructuralSemanticFormattingIfNeeded(tags As List(Of TaggedParagraph),
                                                               mapping As SemanticStyleMapping) As Integer
        If tags Is Nothing OrElse mapping Is Nothing OrElse _reformatParagraphs Is Nothing Then Return 0

        Dim failedTags As New List(Of TaggedParagraph)()
        For Each tagged In tags
            If tagged Is Nothing Then Continue For
            If tagged.ParaIndex < 0 OrElse tagged.ParaIndex >= _reformatParagraphs.Count Then Continue For

            Dim semanticTag = mapping.FindTag(tagged.TagId)
            If Not IsStructuralSemanticTag(tagged.TagId, semanticTag) Then Continue For
            If Not IsSemanticTagObserved(_reformatParagraphs(tagged.ParaIndex), semanticTag) Then
                failedTags.Add(tagged)
            End If
        Next

        If failedTags.Count = 0 Then Return 0

        Try
            Dim result = SemanticRenderingEngine.ApplySemanticFormatting(
                failedTags,
                mapping,
                _reformatParagraphs,
                _reformatTypes,
                Globals.ThisAddIn.Application)
            Return If(result Is Nothing, 0, result.AppliedCount)
        Catch ex As Exception
            Debug.WriteLine($"[ReformatObserve] 结构段落修复失败: {ex.Message}")
            Return 0
        End Try
    End Function

    Private Function IsStructuralSemanticTag(tagId As String, tag As SemanticTag) As Boolean
        Dim parentId = If(tag?.ParentTagId, "")
        Dim value = If(tagId, "")
        Return parentId.Equals("title", StringComparison.OrdinalIgnoreCase) OrElse
               parentId.Equals("heading", StringComparison.OrdinalIgnoreCase) OrElse
               parentId.Equals("list", StringComparison.OrdinalIgnoreCase) OrElse
               value.StartsWith("title.", StringComparison.OrdinalIgnoreCase) OrElse
               value.StartsWith("heading.", StringComparison.OrdinalIgnoreCase) OrElse
               value.StartsWith("list.", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function IsSemanticTagObserved(paragraphObj As Object, tag As SemanticTag) As Boolean
        If paragraphObj Is Nothing OrElse tag Is Nothing OrElse tag.Font Is Nothing Then Return True

        Try
            Dim rng = paragraphObj.Range
            If rng Is Nothing OrElse rng.Font Is Nothing Then Return True

            If tag.Font.FontSize > 0 Then
                Dim currentSize As Double = CDbl(rng.Font.Size)
                If currentSize > 0 AndAlso currentSize < 200 AndAlso
                   Math.Abs(currentSize - tag.Font.FontSize) > 0.75 Then
                    Return False
                End If
            End If

            If tag.Font.Bold Then
                Dim boldValue As Integer = CInt(rng.Font.Bold)
                If boldValue = 0 Then Return False
            End If
        Catch ex As Exception
            Debug.WriteLine($"[ReformatObserve] 样式观察跳过: {ex.Message}")
        End Try

        Return True
    End Function

    Private Function BuildSemanticReformatObservationSummary(tags As List(Of TaggedParagraph),
                                                            mapping As SemanticStyleMapping,
                                                            result As ReformatResult,
                                                            autoFixedCount As Integer,
                                                            repairCount As Integer) As String
        If tags Is Nothing OrElse result Is Nothing Then Return ""

        Dim expectedTextCount As Integer = 0
        For Each tagged In tags
            If tagged Is Nothing Then Continue For
            Dim paraType = "text"
            If _reformatTypes IsNot Nothing AndAlso
               tagged.ParaIndex >= 0 AndAlso tagged.ParaIndex < _reformatTypes.Count Then
                paraType = If(_reformatTypes(tagged.ParaIndex), "text")
            End If
            If String.Equals(paraType, "text", StringComparison.OrdinalIgnoreCase) Then
                expectedTextCount += 1
            End If
        Next

        Dim groups = tags.
            Where(Function(t) t IsNot Nothing).
            GroupBy(Function(t) GetReformatTagDisplayName(t.TagId, mapping)).
            Select(Function(g) $"{g.Key}: {g.Count()} абз.").
            Take(5).
            ToList()

        Dim details As New List(Of String)()
        If groups.Count > 0 Then
            details.Add("Основные стили: " & String.Join(", ", groups))
        End If
        If autoFixedCount > 0 Then
            details.Add($"Автоматически исправлено тегов: {autoFixedCount}")
        End If

        If expectedTextCount > 0 AndAlso result.ModifiedCount < expectedTextCount Then
            details.Add("Часть абзацев не применена; рекомендуется сравнить с предпросмотром или применить после тонкой настройки.")
        Else
            details.Add("Результат добавлен в стек отмены Word; можно восстановить отменой.")
        End If

        Dim agentResult = Services.WordFormattingAgentResult.FromSemanticReformat(
            _activeWordFormattingTaskPlan,
            result.ModifiedCount,
            expectedTextCount,
            repairCount,
            String.Join("; ", details))

        Return agentResult.ToHumanReadableSummary()
    End Function

    Private Async Function ShowSemanticReformatObservation(summary As String) As Task
        Dim responseUuid As String = Guid.NewGuid().ToString()
        Await ExecuteJavaScriptAsyncJS($"createChatSection('AI-наблюдение за форматированием', formatDateTime(new Date()), '{responseUuid}');")
        Await ExecuteJavaScriptAsyncJS($"appendRenderer('{responseUuid}', {JsonConvert.SerializeObject(summary)});")
    End Function

    ''' <summary>
    ''' 显示排版错误到前端
    ''' </summary>
    Private Async Function ShowReformatError(errorMsg As String) As Task
        Try
            Dim errorJson = New JObject From {
                {"success", False},
                {"error", errorMsg}
            }
            Await ExecuteJavaScriptAsyncJS($"showReformatResult({errorJson.ToString(Formatting.None)});")
        Catch ex As Exception
            Debug.WriteLine($"ShowReformatError 出错: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理规范转换结果 - AI将规范文本转为结构化映射后，继续发起语义标注
    ''' </summary>
    Private Async Sub HandleStyleGuideConversionResult(aiResponseJson As String)
        Try
            ' 解析AI返回的规范映射
            Dim guideName As String = "Стандарт форматирования"
            Dim guideId As String = ""

            ' 尝试从响应映射获取规范信息
            Dim mapping = StyleGuideConverter.ParseAiResponse(aiResponseJson, guideName, guideId)

            If mapping Is Nothing OrElse mapping.SemanticTags.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Не удалось разобрать результат нормативного преобразования")
                Return
            End If

            ' 缓存映射
            SemanticMappingManager.Instance.AddMapping(mapping)

            ' 更新上下文中的mapping
            _reformatMapping = mapping

            ' 收集段落文本
            Dim paragraphTexts As New List(Of String)()
            If _reformatParagraphs IsNot Nothing Then
                For Each para In _reformatParagraphs
                    Try
                        paragraphTexts.Add(If(para.Range.Text?.ToString().TrimEnd(vbCr, vbLf), ""))
                    Catch
                        paragraphTexts.Add("")
                    End Try
                Next
            End If

            ' 构建语义标注提示词
            Dim systemPrompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(mapping, paragraphTexts)

            ' 发送语义标注请求
            Await Send("Стандарт разобран, теперь выполни семантическую разметку.", systemPrompt, False, "semantic_reformat")

        Catch ex As Exception
            Debug.WriteLine($"HandleStyleGuideConversionResult 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка разметки после нормативного преобразования: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理格式克隆（mirror_format）的 AI 响应：解析 SemanticStyleMapping 并保存
    ''' </summary>
    Private Sub HandleMirrorFormatResult(aiResponseJson As String)
        Try
            Dim mapping = StyleGuideConverter.ParseAiResponse(aiResponseJson, _mirrorFormatDocName, "")
            If mapping Is Nothing OrElse mapping.SemanticTags.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Не удалось разобрать результат клонирования формата. Повторите попытку")
                Return
            End If

            mapping.Name = If(String.IsNullOrEmpty(_mirrorFormatDocName), "Клонированный формат", _mirrorFormatDocName)
            mapping.SourceType = SemanticMappingSourceType.FromDocxTemplate

            SemanticMappingManager.Instance.AddMapping(mapping)

            Dim json = JsonConvert.SerializeObject(mapping, Formatting.None)
            ExecuteJavaScriptAsyncJS($"showMappingPreview({json});")

            GlobalStatusStrip.ShowInfo($"Клонирование формата завершено, извлечено семантических тегов: {mapping.SemanticTags.Count}")
        Catch ex As Exception
            Debug.WriteLine($"HandleMirrorFormatResult 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось сохранить клонирование формата: {ex.Message}")
        End Try
    End Sub

    Protected Overrides Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)
            If String.IsNullOrEmpty(txt) Then
                Return Nothing
            End If

            Dim info As New SelectionInfo()
            info.SelectedText = txt
            info.DocumentPath = Globals.ThisAddIn.Application.ActiveDocument.FullName

            Try
                info.StartPos = sel.Range.Start
                info.EndPos = sel.Range.End
            Catch
                info.StartPos = 0
                info.EndPos = 0
            End Try

            Return info
        Catch
            Return Nothing
        End Try
    End Function

    ' ========== 续写功能 ==========

    Private _continuationService As WordContinuationService
    Private _cachedContinuationContext As ContinuationContext ' 缓存续写上下文，用于多轮续写

    ''' <summary>
    ''' 触发续写 - 获取光标上下文并发送AI请求
    ''' </summary>
    Protected Overrides Sub HandleTriggerContinuation(jsonDoc As JObject)
        Try
            ' 提取参数
            Dim style As String = ""
            Dim isContinuationMode As Boolean = False

            If jsonDoc IsNot Nothing Then
                If jsonDoc("style") IsNot Nothing Then
                    style = jsonDoc("style").ToString()
                End If
                If jsonDoc("isContinuationMode") IsNot Nothing Then
                    isContinuationMode = jsonDoc("isContinuationMode").ToObject(Of Boolean)()
                End If
            End If

            ' 初始化续写服务
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 检查是否可以续写
            If Not _continuationService.CanContinue() Then
                GlobalStatusStrip.ShowWarning("Не удалось получить информацию о документе. Убедитесь, что документ открыт")
                Return
            End If

            Dim context As ContinuationContext

            ' 如果是续写模式的后续请求，并且有缓存的上下文，则复用
            If isContinuationMode AndAlso _cachedContinuationContext IsNot Nothing Then
                ' 多轮续写：使用缓存的上下文，但style作为新的调整要求
                context = _cachedContinuationContext
                GlobalStatusStrip.ShowInfo("Продолжаем генерацию...")
            Else
                ' 首次续写或非续写模式：重新获取上下文
                context = _continuationService.GetCursorContext(3, 3)
                If context Is Nothing Then
                    GlobalStatusStrip.ShowWarning("Не удалось получить контекст документа")
                    Return
                End If
                ' 缓存上下文
                _cachedContinuationContext = context
                GlobalStatusStrip.ShowInfo("Анализ контекста и генерация продолжения...")
            End If

            ' 发送续写请求（带上风格参数）
            SendContinuationRequest(context, style)

        Catch ex As Exception
            Debug.WriteLine($"HandleTriggerContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка при запуске продолжения: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用续写结果到Word文档
    ''' </summary>
    Protected Overrides Sub HandleApplyContinuation(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("Содержимое продолжения пусто")
                Return
            End If

            ' 确保续写服务已初始化
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor
            End Select

            ' 插入续写内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("Продолжение вставлено в документ")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeContinuationActions('{uuid}');")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка при вставке продолжения: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用模板渲染结果到Word文档
    ''' </summary>
    Protected Overrides Sub HandleApplyTemplateContent(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("Содержимое шаблона пусто")
                Return
            End If

            ' 确保续写服务已初始化（复用其插入逻辑）
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor
            End Select

            ' 插入模板内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("Содержимое шаблона вставлено в документ")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeTemplateActions('{uuid}');")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyTemplateContent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка при вставке содержимого шаблона: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取当前Word上下文快照（用于自动补全）
    ''' </summary>
    Protected Overrides Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = "Word"

        Try
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection IsNot Nothing AndAlso selection.Start <> selection.End Then
                ' 有选中内容
                Dim selText = selection.Text
                If Not String.IsNullOrEmpty(selText) AndAlso selText.Length > 500 Then
                    selText = selText.Substring(0, 500) & "..."
                End If
                snapshot("selection") = If(selText, "")
            Else
                snapshot("selection") = ""
            End If

            ' 获取文档标题
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc IsNot Nothing Then
                snapshot("documentName") = If(doc.Name, "")
            End If

        Catch ex As Exception
            Debug.WriteLine($"GetContextSnapshot 出错: {ex.Message}")
        End Try

        Return snapshot
    End Function

    ''' <summary>
    ''' 重写保存设置方法，同步更新Word补全管理器状态
    ''' </summary>
    Protected Overrides Sub HandleSaveSettings(jsonDoc As JObject)
        MyBase.HandleSaveSettings(jsonDoc)

        ' 同步更新Word补全管理器的启用状态
        Try
            Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
            WordCompletionManager.Instance.Enabled = enableAutocomplete
            Debug.WriteLine($"[WordChatControl] 补全设置已同步: Enabled={enableAutocomplete}")
        Catch ex As Exception
            Debug.WriteLine($"[WordChatControl] 同步补全设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 执行JSON命令（重写基类方法）- 带严格验证，支持DSL和旧版格式
    ''' </summary>
    Private Function ExecuteJsonCommandCore(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 预览模式下跳过自动执行（排版/校对模式的JSON用于预览，由用户手动点击应用）
            If IsInPreviewMode() Then
                Debug.WriteLine($"[WordChatControl] 预览模式({GetCurrentResponseMode()})下跳过JSON命令自动执行")
                Return True ' 返回True表示"成功处理"，避免显示错误
            End If

            ' === DSL格式检测与转换 ===
            Dim detectedFormat = DslProtocolDetector.DetectFormat(jsonCode)
            If detectedFormat = InstructionFormat.DslJson OrElse detectedFormat = InstructionFormat.ProofreadJson Then
                Return ExecuteDslCommand(jsonCode, preview)
            End If

            ' === 旧版JSON命令格式（兼容） ===
            ' 使用严格的结构验证
            Dim errorMessage As String = ""
            Dim normalizedJson As JToken = Nothing

            If Not WordJsonCommandSchema.ValidateJsonStructure(jsonCode, errorMessage, normalizedJson) Then
                ' 格式验证失败
                Debug.WriteLine($"Word JSON格式验证失败: {errorMessage}")
                Debug.WriteLine($"原始JSON: {jsonCode.Substring(0, Math.Min(200, jsonCode.Length))}...")

                ShareRibbon.GlobalStatusStrip.ShowWarning($"Формат JSON не соответствует спецификации: {errorMessage}")
                Return False
            End If

            ' 验证通过，根据类型执行
            If normalizedJson.Type = JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, JObject)

                ' 命令数组格式
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecuteWordCommandsArray(jsonObj("commands"), jsonCode, preview)
                End If

                ' 单命令格式
                Return ExecuteWordSingleCommand(jsonObj, jsonCode, preview)
            End If

            ShareRibbon.GlobalStatusStrip.ShowWarning("Недопустимый формат JSON")
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Недопустимый формат JSON: {ex.Message}")
            Return False
        Catch ex As Exception
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка выполнения: {ex.Message}")
            Return False
        End Try
    End Function

    Protected Overrides Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
        Try
            Dim errorMessage As String = ""
            Dim normalizedJson As JToken = Nothing

            If Not WordJsonCommandSchema.ValidateJsonStructure(jsonCode, errorMessage, normalizedJson) Then
                Return Agent.ToolResult.Failed("", $"Формат JSON не соответствует спецификации: {errorMessage}",
                                               errorCode:=ExceptionClassifier.CodeJson,
                                               userMessage:=$"Формат JSON не соответствует спецификации: {errorMessage}",
                                               recoverable:=True)
            End If

            If normalizedJson IsNot Nothing AndAlso normalizedJson.Type = JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, JObject)
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecuteWordCommandsArrayWithToolResult(jsonObj("commands"), preview)
                Else
                    Dim command = jsonObj("command")?.ToString()
                    If Not String.IsNullOrWhiteSpace(command) Then
                        Return ExecuteWordCommandWithToolResult(jsonObj)
                    End If
                End If
            End If

            Dim ok = ExecuteJsonCommandCore(jsonCode, preview)
            If ok Then Return Agent.ToolResult.Succeed("", "Выполнено успешно")
            Return Agent.ToolResult.Failed("", "Сбой выполнения команды JSON")

        Catch ex As Newtonsoft.Json.JsonReaderException
            Return Agent.ToolResult.Failed("", $"Недопустимый формат JSON: {ex.Message}",
                                           errorCode:=ExceptionClassifier.CodeJson,
                                           userMessage:=$"Недопустимый формат JSON: {ex.Message}",
                                           debugDetail:=ex.Message,
                                           recoverable:=True)
        Catch ex As Exception
            Return Agent.ToolResult.FromException("", ex)
        End Try
    End Function

    Private Function IsReadToolCommand(command As String) As Boolean
        If String.IsNullOrWhiteSpace(command) Then Return False
        Select Case command.Trim().ToLowerInvariant()
            Case "listparagraphs", "getparagraphinfo"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function ExecuteWordCommandsArrayWithToolResult(commandsArray As JToken, preview As Boolean) As Agent.ToolResult
        Try
            If preview Then
                Dim previewObj As New JObject From {{"commands", commandsArray}}
                If Not ShareRibbon.CommandPreviewForm.ShowPreview("Предпросмотр команд Word", previewObj) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return Agent.ToolResult.Succeed("WordCommands", "Пользователь отменил выполнение")
                End If
            End If

            If commandsArray Is Nothing OrElse commandsArray.Type <> JTokenType.Array Then
                Return Agent.ToolResult.Failed("WordCommands",
                                               "commands должен быть массивом",
                                               errorCode:=ExceptionClassifier.CodeJson,
                                               userMessage:="commands должен быть массивом",
                                               recoverable:=True)
            End If

            Dim observations As New JArray()
            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim firstFailure As Agent.ToolResult = Nothing

            For Each cmd In CType(commandsArray, JArray)
                If cmd.Type <> JTokenType.Object Then
                    failCount += 1
                    Continue For
                End If

                Dim result = ExecuteWordCommandWithToolResult(CType(cmd, JObject))
                If result.Success Then
                    successCount += 1
                Else
                    failCount += 1
                    If firstFailure Is Nothing Then firstFailure = result
                End If

                If result.Observation IsNot Nothing Then
                    observations.Add(JToken.FromObject(result.Observation))
                End If
            Next

            Dim aggregateObservation As New JObject From {
                {"kind", "batch"},
                {"summary", $"Пакетное выполнение команд Word: успешно {successCount}, ошибок {failCount}"},
                {"changed", successCount > 0},
                {"targetRefs", New JArray("Word:Document")},
                {"warnings", New JArray()},
                {"items", observations}
            }

            If failCount = 0 Then
                Return Agent.ToolResult.Succeed("WordCommands",
                                                $"Все команды выполнены успешно: {successCount}",
                                                observation:=aggregateObservation)
            End If

            Return Agent.ToolResult.Failed("WordCommands",
                                           $"Пакетное выполнение завершено: успешно {successCount}, ошибок {failCount}",
                                           errorCode:=If(firstFailure?.ErrorCode, ExceptionClassifier.CodeUnknown),
                                           userMessage:=If(firstFailure?.UserMessage, $"Пакетное выполнение завершено: успешно {successCount}, ошибок {failCount}"),
                                           recoverable:=True,
                                           observation:=aggregateObservation)
        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommandsArrayWithToolResult 出错: {ex.Message}")
            Return Agent.ToolResult.FromException("WordCommands", ex)
        End Try
    End Function

    Private Function ExecuteWordCommandWithToolResult(commandJson As JObject) As Agent.ToolResult
        Dim command = commandJson("command")?.ToString()

        If IsReadToolCommand(command) Then
            Return ExecuteWordReadCommandWithToolResult(commandJson)
        End If

        Try
            Dim params = commandJson("params")
            Dim normalizedToolId = NormalizeWordToolId(command)
            If String.IsNullOrWhiteSpace(normalizedToolId) Then normalizedToolId = If(command, "WordCommand")
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = Globals.ThisAddIn.Application.Selection
            Dim beforeSnapshot = CaptureWordWriteSnapshot(normalizedToolId, params, doc, selection)
            Dim success As Boolean = False

            Select Case command.Trim().ToLowerInvariant()
                Case "inserttext"
                    success = ExecuteInsertText(params, selection)
                Case "formattext"
                    success = ExecuteFormatText(params, selection)
                Case "replacetext"
                    success = ExecuteReplaceText(params, doc)
                Case "deletetext"
                    success = ExecuteDeleteText(params, doc, selection)
                Case Else
                    ' 其余 Word 原生命令继续复用已有执行器，但不再退化为空 ToolId 的 Boolean 假成功。
                    success = ExecuteWordSingleCommand(commandJson, commandJson.ToString(Formatting.None), False)
            End Select

            Dim afterSnapshot = CaptureWordWriteSnapshot(normalizedToolId, params, doc, selection)
            Dim observation = BuildWordWriteObservation(normalizedToolId, params, success, beforeSnapshot, afterSnapshot)
            Dim summary = observation("summary")?.ToString()

            If success Then
                Return Agent.ToolResult.Succeed(normalizedToolId,
                                                If(String.IsNullOrWhiteSpace(summary), "Выполнено успешно", summary),
                                                observation:=observation)
            End If

            Return Agent.ToolResult.Failed(normalizedToolId,
                                           $"Сбой выполнения {normalizedToolId}",
                                           errorCode:=ExceptionClassifier.CodeUnknown,
                                           userMessage:=$"Сбой выполнения {normalizedToolId}",
                                           recoverable:=True,
                                           observation:=observation)
        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommandWithToolResult 出错: {ex.Message}")
            Return Agent.ToolResult.FromException(If(command, ""), ex)
        End Try
    End Function

    Private Function ExecuteWordReadCommandWithToolResult(commandJson As JObject) As Agent.ToolResult
        Dim command = commandJson("command")?.ToString()
        Dim params = commandJson("params")

        Select Case If(command, "").Trim().ToLowerInvariant()
            Case "listparagraphs"
                Return ExecuteListParagraphsWithToolResult(params)
            Case "getparagraphinfo"
                Return ExecuteGetParagraphInfoWithToolResult(params)
            Case Else
                Return Agent.ToolResult.Failed(If(command, ""), $"Неподдерживаемая команда чтения Word: {command}",
                                               errorCode:=ExceptionClassifier.CodeNotFound,
                                               userMessage:=$"Неподдерживаемая команда чтения Word: {command}",
                                               recoverable:=False)
        End Select
    End Function

    Private Function NormalizeWordToolId(command As String) As String
        Select Case If(command, "").Trim().ToLowerInvariant()
            Case "inserttext"
                Return "InsertText"
            Case "formattext"
                Return "FormatText"
            Case "replacetext"
                Return "ReplaceText"
            Case "deletetext"
                Return "DeleteText"
            Case Else
                Return If(command, "")
        End Select
    End Function

    Private Function BuildWordWriteObservation(toolId As String,
                                               params As JToken,
                                               commandSucceeded As Boolean,
                                               beforeSnapshot As JObject,
                                               afterSnapshot As JObject) As JObject
        Dim targetRefs As JArray = GetWordWriteTargetRefs(toolId, params)
        Dim warnings As New JArray()
        Dim diff = BuildWordSnapshotDiff(beforeSnapshot, afterSnapshot)
        Dim changed = SnapshotDiffChanged(diff)

        If Not commandSucceeded Then
            warnings.Add("Исполнитель узла вернул ошибку; изменение документа не подтверждено")
        ElseIf Not changed Then
            warnings.Add("Исполнитель узла вернул успех, но изменения содержимого документа или формата выделения не обнаружены")
        End If

        Return New JObject From {
            {"kind", "write"},
            {"summary", BuildWordWriteSummary(toolId, params, commandSucceeded, changed)},
            {"targetRefs", targetRefs},
            {"changed", changed},
            {"before", If(beforeSnapshot, New JObject())},
            {"after", If(afterSnapshot, New JObject())},
            {"diff", diff},
            {"warnings", warnings}
        }
    End Function

    Private Function BuildWordWriteSummary(toolId As String, params As JToken, commandSucceeded As Boolean, changed As Boolean) As String
        Dim prefix = ""
        If Not commandSucceeded Then
            prefix = "Не завершено: "
        ElseIf Not changed Then
            prefix = "Выполнено, но изменения не обнаружены: "
        End If

        Select Case toolId
            Case "InsertText"
                Dim content = If(params?("content")?.ToString(), "")
                Dim position = If(params?("position")?.ToString(), "cursor")
                Return $"{prefix}в позицию «{DescribeWordPosition(position)}» вставлено символов: {content.Length}"
            Case "FormatText"
                Return $"{prefix}форматирование текста текущего выделения"
            Case "ReplaceText"
                Dim find = If(params?("find")?.ToString(), "")
                Return $"{prefix}замена в документе совпадающего текста '{TruncateObservationText(find, 40)}'"
            Case "DeleteText"
                Dim rangeName = If(params?("range")?.ToString(), "selection")
                Return $"{prefix}удаление текста ({DescribeWordRange(rangeName)})"
            Case Else
                Return $"{prefix}{toolId}: выполнение завершено"
        End Select
    End Function

    Private Function CaptureWordWriteSnapshot(toolId As String, params As JToken, doc As Object, selection As Object) As JObject
        Dim snapshot As New JObject()

        Try
            Dim documentText = SafeGetDocumentText(doc)
            Dim selectionText = SafeGetSelectionText(selection)
            Dim targetText = SafeGetWordTargetText(toolId, params, doc, selection, documentText, selectionText)

            snapshot("paragraphCount") = SafeGetParagraphCount(doc)
            snapshot("documentCharCount") = documentText.Length
            snapshot("documentTextHash") = ComputeTextHash(documentText)
            snapshot("selectionStart") = SafeGetSelectionBoundary(selection, "Start")
            snapshot("selectionEnd") = SafeGetSelectionBoundary(selection, "End")
            snapshot("selectionTextPreview") = TruncateObservationText(NormalizeObservationText(selectionText), 120)
            snapshot("selectionTextHash") = ComputeTextHash(selectionText)
            snapshot("selectionFormat") = CaptureSelectionFormat(selection)
            snapshot("selectionFormatHash") = ComputeTextHash(snapshot("selectionFormat").ToString(Formatting.None))
            snapshot("targetTextPreview") = TruncateObservationText(NormalizeObservationText(targetText), 160)
            snapshot("targetTextHash") = ComputeTextHash(targetText)

            Dim findText = If(params?("find")?.ToString(), "")
            If Not String.IsNullOrEmpty(findText) Then
                snapshot("findMatchCount") = CountTextOccurrences(documentText, findText, If(params?("matchCase")?.Value(Of Boolean)(), False))
            End If
        Catch ex As Exception
            snapshot("snapshotError") = ex.Message
        End Try

        Return snapshot
    End Function

    Private Function BuildWordSnapshotDiff(beforeSnapshot As JObject, afterSnapshot As JObject) As JObject
        Dim beforeChars = GetIntegerValue(beforeSnapshot, "documentCharCount")
        Dim afterChars = GetIntegerValue(afterSnapshot, "documentCharCount")
        Dim beforeParagraphs = GetIntegerValue(beforeSnapshot, "paragraphCount")
        Dim afterParagraphs = GetIntegerValue(afterSnapshot, "paragraphCount")
        Dim beforeMatches = GetIntegerValue(beforeSnapshot, "findMatchCount")
        Dim afterMatches = GetIntegerValue(afterSnapshot, "findMatchCount")

        Return New JObject From {
            {"documentHashChanged", GetStringValue(beforeSnapshot, "documentTextHash") <> GetStringValue(afterSnapshot, "documentTextHash")},
            {"targetHashChanged", GetStringValue(beforeSnapshot, "targetTextHash") <> GetStringValue(afterSnapshot, "targetTextHash")},
            {"selectionTextChanged", GetStringValue(beforeSnapshot, "selectionTextHash") <> GetStringValue(afterSnapshot, "selectionTextHash")},
            {"selectionFormatChanged", GetStringValue(beforeSnapshot, "selectionFormatHash") <> GetStringValue(afterSnapshot, "selectionFormatHash")},
            {"charDelta", afterChars - beforeChars},
            {"paragraphDelta", afterParagraphs - beforeParagraphs},
            {"findMatchDelta", afterMatches - beforeMatches}
        }
    End Function

    Private Function SnapshotDiffChanged(diff As JObject) As Boolean
        If diff Is Nothing Then Return False
        Return GetBooleanValue(diff, "documentHashChanged") OrElse
               GetBooleanValue(diff, "targetHashChanged") OrElse
               GetBooleanValue(diff, "selectionTextChanged") OrElse
               GetBooleanValue(diff, "selectionFormatChanged") OrElse
               GetIntegerValue(diff, "charDelta") <> 0 OrElse
               GetIntegerValue(diff, "paragraphDelta") <> 0
    End Function

    Private Function SafeGetDocumentText(doc As Object) As String
        Try
            If doc Is Nothing OrElse doc.Content Is Nothing Then Return ""
            Return If(doc.Content.Text, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function SafeGetSelectionText(selection As Object) As String
        Try
            If selection Is Nothing OrElse selection.Range Is Nothing Then Return ""
            Return If(selection.Range.Text, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function SafeGetWordTargetText(toolId As String,
                                           params As JToken,
                                           doc As Object,
                                           selection As Object,
                                           documentText As String,
                                           selectionText As String) As String
        Select Case toolId
            Case "ReplaceText"
                Return documentText
            Case "DeleteText"
                Dim rangeName = If(params?("range")?.ToString(), "selection").Trim().ToLowerInvariant()
                If rangeName = "all" OrElse rangeName = "document" OrElse rangeName = "全文" Then Return documentText
                Return selectionText
            Case "InsertText"
                Dim position = If(params?("position")?.ToString(), "cursor").Trim().ToLowerInvariant()
                If position = "start" OrElse position = "end" Then Return documentText
                Return selectionText
            Case Else
                Return selectionText
        End Select
    End Function

    Private Function SafeGetParagraphCount(doc As Object) As Integer
        Try
            If doc Is Nothing OrElse doc.Paragraphs Is Nothing Then Return 0
            Return CInt(doc.Paragraphs.Count)
        Catch
            Return 0
        End Try
    End Function

    Private Function SafeGetSelectionBoundary(selection As Object, propertyName As String) As Integer
        Try
            If selection Is Nothing OrElse selection.Range Is Nothing Then Return -1
            If propertyName = "End" Then Return CInt(selection.Range.End)
            Return CInt(selection.Range.Start)
        Catch
            Return -1
        End Try
    End Function

    Private Function CaptureSelectionFormat(selection As Object) As JObject
        Dim result As New JObject()
        Try
            If selection Is Nothing OrElse selection.Range Is Nothing OrElse selection.Range.Font Is Nothing Then Return result
            Dim font = selection.Range.Font
            result("bold") = SafeObjectString(font.Bold)
            result("italic") = SafeObjectString(font.Italic)
            result("underline") = SafeObjectString(font.Underline)
            result("fontName") = SafeObjectString(font.Name)
            result("fontSize") = SafeObjectString(font.Size)
            result("color") = SafeObjectString(font.Color)
        Catch ex As Exception
            result("formatError") = ex.Message
        End Try
        Return result
    End Function

    Private Function SafeObjectString(value As Object) As String
        If value Is Nothing Then Return ""
        Try
            Return value.ToString()
        Catch
            Return ""
        End Try
    End Function

    Private Function CountTextOccurrences(text As String, findText As String, matchCase As Boolean) As Integer
        If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(findText) Then Return 0
        Dim comparison = If(matchCase, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)
        Dim count As Integer = 0
        Dim index As Integer = 0

        While index < text.Length
            Dim found = text.IndexOf(findText, index, comparison)
            If found < 0 Then Exit While
            count += 1
            index = found + Math.Max(findText.Length, 1)
        End While

        Return count
    End Function

    Private Function ComputeTextHash(text As String) As String
        Try
            Using sha = SHA256.Create()
                Dim bytes = Encoding.UTF8.GetBytes(If(text, ""))
                Dim hash = sha.ComputeHash(bytes)
                Return BitConverter.ToString(hash, 0, Math.Min(8, hash.Length)).Replace("-", "").ToLowerInvariant()
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Function NormalizeObservationText(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Return Regex.Replace(text, "\s+", " ").Trim()
    End Function

    Private Function GetIntegerValue(obj As JObject, propertyName As String) As Integer
        Try
            If obj Is Nothing OrElse obj(propertyName) Is Nothing Then Return 0
            Return obj(propertyName).Value(Of Integer)()
        Catch
            Return 0
        End Try
    End Function

    Private Function GetBooleanValue(obj As JObject, propertyName As String) As Boolean
        Try
            If obj Is Nothing OrElse obj(propertyName) Is Nothing Then Return False
            Return obj(propertyName).Value(Of Boolean)()
        Catch
            Return False
        End Try
    End Function

    Private Function GetStringValue(obj As JObject, propertyName As String) As String
        Try
            If obj Is Nothing OrElse obj(propertyName) Is Nothing Then Return ""
            Return obj(propertyName).ToString()
        Catch
            Return ""
        End Try
    End Function

    Private Function GetWordWriteTargetRefs(toolId As String, params As JToken) As JArray
        Select Case toolId
            Case "InsertText"
                Dim position = If(params?("position")?.ToString(), "cursor").Trim().ToLowerInvariant()
                Select Case position
                    Case "start"
                        Return New JArray("Word:DocumentStart")
                    Case "end"
                        Return New JArray("Word:DocumentEnd")
                    Case Else
                        Return New JArray("Word:Selection")
                End Select
            Case "ReplaceText"
                Return New JArray("Word:Document")
            Case "DeleteText"
                Dim rangeName = If(params?("range")?.ToString(), "selection").Trim().ToLowerInvariant()
                If rangeName = "all" OrElse rangeName = "document" OrElse rangeName = "全文" Then
                    Return New JArray("Word:Document")
                End If
                Return New JArray("Word:Selection")
            Case Else
                Return New JArray("Word:Selection")
        End Select
    End Function

    Private Function DescribeWordPosition(position As String) As String
        Select Case If(position, "").Trim().ToLowerInvariant()
            Case "start"
                Return "Начало документа"
            Case "end"
                Return "Конец документа"
            Case Else
                Return "Позиция курсора"
        End Select
    End Function

    Private Function DescribeWordRange(rangeName As String) As String
        Select Case If(rangeName, "").Trim().ToLowerInvariant()
            Case "all", "document", "全文"
                Return "Весь документ"
            Case Else
                Return "Текущее выделение"
        End Select
    End Function

    Private Function TruncateObservationText(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) OrElse text.Length <= maxLen Then Return If(text, "")
        Return text.Substring(0, maxLen) & "..."
    End Function

    ''' <summary>
    ''' 执行DSL指令 - 桥接到已有的Word排版渲染管道
    ''' </summary>
    Private Function ExecuteDslCommand(jsonCode As String, preview As Boolean) As Boolean
        Try
            Dim instructions = Instruction.ParseInstructions(jsonCode)

            If instructions Is Nothing OrElse instructions.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("Разобранный DSL-набор пуст")
                Return False
            End If

            ' 预览模式
            If preview Then
                Dim previewJson As JToken = Nothing
                Try
                    previewJson = JObject.Parse(jsonCode)
                Catch
                    previewJson = New JObject()
                End Try
                If Not ShareRibbon.CommandPreviewForm.ShowPreview($"Предпросмотр DSL-команд — всего {instructions.Count}", previewJson) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 桥接到已有排版管道：DSL指令 → 语义标签 → SemanticRenderingEngine
            If _reformatMapping Is Nothing OrElse _reformatParagraphs Is Nothing Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("Нет контекста форматирования. Сначала выберите содержимое и запустите форматирование")
                Return False
            End If

            Dim wordApp = Globals.ThisAddIn.Application
            Dim undoStarted As Boolean = False

            Try
                wordApp.UndoRecord.StartCustomRecord("AI-форматирование (DSL)")
                undoStarted = True
            Catch ex As Exception
                Debug.WriteLine($"DSL StartCustomRecord 失败: {ex.Message}")
            End Try

            Dim appliedCount As Integer = 0
            Try
                ' 遍历DSL指令，转换为Word DOM操作
                For Each inst In instructions
                    Try
                        Dim opResult = ApplySingleDslInstruction(inst, wordApp)
                        If opResult Then appliedCount += 1
                    Catch ex As Exception
                        Debug.WriteLine($"DSL指令 {inst.Id} 执行失败: {ex.Message}")
                    End Try
                Next
            Finally
                If undoStarted Then
                    Try
                        wordApp.UndoRecord.EndCustomRecord()
                    Catch ex As Exception
                        Debug.WriteLine($"DSL EndCustomRecord 失败: {ex.Message}")
                    End Try
                End If
            End Try

            If appliedCount > 0 Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"Форматирование DSL завершено, применено команд: {appliedCount}/{instructions.Count}")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"DSL-команды не применили ни одного формата (всего: {instructions.Count})")
            End If

            Return appliedCount > 0

        Catch ex As Exception
            Debug.WriteLine($"[ExecuteDslCommand] 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка выполнения DSL: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行单条DSL指令到Word DOM
    ''' </summary>
    Private Function ApplySingleDslInstruction(inst As Instruction, wordApp As Object) As Boolean
        Select Case inst.Operation
            Case "setParagraphStyle"
                Return ApplyDslParagraphStyle(inst, wordApp)
            Case "setCharacterFormat"
                Return ApplyDslCharacterFormat(inst, wordApp)
            Case "setPageSetup"
                Return ApplyDslPageSetup(inst, wordApp)
            Case "insertBreak"
                Return ApplyDslInsertBreak(inst, wordApp)
            Case "suggestCorrection", "suggestFormatFix", "suggestStyleUnify", "markForReview"
                ' 校对类指令不产生DOM修改，仅建议
                Debug.WriteLine($"DSL校对建议: {inst.GetDescription()}")
                Return True
            Case Else
                Debug.WriteLine($"DSL未知指令类型: {inst.Operation}")
                Return False
        End Select
    End Function

    ''' <summary>
    ''' DSL: 设置段落样式
    ''' </summary>
    Private Function ApplyDslParagraphStyle(inst As Instruction, wordApp As Object) As Boolean
        Dim targetIndex = CInt(inst.GetParam("target.index", -1))
        If targetIndex < 0 OrElse _reformatParagraphs Is Nothing OrElse targetIndex >= _reformatParagraphs.Count Then
            Return False
        End If

        Dim para = _reformatParagraphs(targetIndex)
        Dim range As Object = Nothing
        Try
            range = para.GetType().InvokeMember("Range", Reflection.BindingFlags.GetProperty, Nothing, para, Nothing)
        Catch
            Return False
        End Try

        ' 应用样式名
        Dim styleName = inst.GetParam("styleName", String.Empty)?.ToString()
        If Not String.IsNullOrEmpty(styleName) Then
            Try
                range.GetType().InvokeMember("Style", Reflection.BindingFlags.SetProperty, Nothing, range,
                    New Object() {wordApp.ActiveDocument.Styles(styleName)})
            Catch
            End Try
        End If

        ' 应用字体设置
        Dim font = inst.GetParam("font", Nothing)
        If font IsNot Nothing Then
            Try
                Dim fontObj = range.GetType().InvokeMember("Font", Reflection.BindingFlags.GetProperty, Nothing, range, Nothing)
                Dim fontName = font("name")?.ToString()
                If Not String.IsNullOrEmpty(fontName) Then
                    fontObj.GetType().InvokeMember("Name", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {fontName})
                End If
                Dim fontSize = font("size")
                If fontSize IsNot Nothing Then
                    fontObj.GetType().InvokeMember("Size", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {CSng(CType(fontSize, JToken).Value(Of Single)())})
                End If
            Catch
            End Try
        End If

        ' 应用对齐
        Dim alignment = inst.GetParam("alignment", String.Empty)?.ToString()
        If Not String.IsNullOrEmpty(alignment) Then
            Try
                Dim alignValue As Integer = GetWordAlignmentValue(alignment)
                range.GetType().InvokeMember("ParagraphFormat", Reflection.BindingFlags.GetProperty, Nothing, range, Nothing) _
                    .GetType().InvokeMember("Alignment", Reflection.BindingFlags.SetProperty, Nothing,
                    range.GetType().InvokeMember("ParagraphFormat", Reflection.BindingFlags.GetProperty, Nothing, range, Nothing),
                    New Object() {alignValue})
            Catch
            End Try
        End If

        Return True
    End Function

    ''' <summary>
    ''' DSL: 设置字符格式
    ''' </summary>
    Private Function ApplyDslCharacterFormat(inst As Instruction, wordApp As Object) As Boolean
        Dim targetIndex = CInt(inst.GetParam("target.index", -1))
        If targetIndex < 0 OrElse _reformatParagraphs Is Nothing OrElse targetIndex >= _reformatParagraphs.Count Then
            Return False
        End If

        Dim para = _reformatParagraphs(targetIndex)
        Dim range As Object = Nothing
        Try
            range = para.GetType().InvokeMember("Range", Reflection.BindingFlags.GetProperty, Nothing, para, Nothing)
        Catch
            Return False
        End Try

        Dim font = inst.GetParam("font", Nothing)
        If font IsNot Nothing Then
            Try
                Dim fontObj = range.GetType().InvokeMember("Font", Reflection.BindingFlags.GetProperty, Nothing, range, Nothing)

                Dim bold = font("bold")
                If bold IsNot Nothing Then
                    fontObj.GetType().InvokeMember("Bold", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {Convert.ToInt32(CType(bold, JToken).Value(Of Boolean)())})
                End If

                Dim italic = font("italic")
                If italic IsNot Nothing Then
                    fontObj.GetType().InvokeMember("Italic", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {Convert.ToInt32(CType(italic, JToken).Value(Of Boolean)())})
                End If

                Dim fontName = font("name")?.ToString()
                If Not String.IsNullOrEmpty(fontName) Then
                    fontObj.GetType().InvokeMember("Name", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {fontName})
                End If

                Dim fontSize = font("size")
                If fontSize IsNot Nothing Then
                    fontObj.GetType().InvokeMember("Size", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {CSng(CType(fontSize, JToken).Value(Of Single)())})
                End If

                Dim colorValue = font("color")
                If colorValue IsNot Nothing Then
                    fontObj.GetType().InvokeMember("Color", Reflection.BindingFlags.SetProperty, Nothing, fontObj, New Object() {CInt(CType(colorValue, JToken).Value(Of Integer)())})
                End If
            Catch ex As Exception
                Debug.WriteLine($"ApplyDslCharacterFormat: {ex.Message}")
            End Try
        End If

        Return True
    End Function

    ''' <summary>
    ''' DSL: 设置页面格式
    ''' </summary>
    Private Function ApplyDslPageSetup(inst As Instruction, wordApp As Object) As Boolean
        Try
            Dim section = wordApp.ActiveDocument.Sections(1)
            Dim pageSetup = section.GetType().InvokeMember("PageSetup", Reflection.BindingFlags.GetProperty, Nothing, section, Nothing)

            Dim margins = inst.GetParam("margins", Nothing)
            If margins IsNot Nothing Then
                Dim topVal = margins("top")
                If topVal IsNot Nothing Then
                    pageSetup.GetType().InvokeMember("TopMargin", Reflection.BindingFlags.SetProperty, Nothing, pageSetup, New Object() {CSng(CType(topVal, JToken).Value(Of Single)())})
                End If
                Dim bottomVal = margins("bottom")
                If bottomVal IsNot Nothing Then
                    pageSetup.GetType().InvokeMember("BottomMargin", Reflection.BindingFlags.SetProperty, Nothing, pageSetup, New Object() {CSng(CType(bottomVal, JToken).Value(Of Single)())})
                End If
                Dim leftVal = margins("left")
                If leftVal IsNot Nothing Then
                    pageSetup.GetType().InvokeMember("LeftMargin", Reflection.BindingFlags.SetProperty, Nothing, pageSetup, New Object() {CSng(CType(leftVal, JToken).Value(Of Single)())})
                End If
                Dim rightVal = margins("right")
                If rightVal IsNot Nothing Then
                    pageSetup.GetType().InvokeMember("RightMargin", Reflection.BindingFlags.SetProperty, Nothing, pageSetup, New Object() {CSng(CType(rightVal, JToken).Value(Of Single)())})
                End If
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ApplyDslPageSetup: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' DSL: 插入分隔符
    ''' </summary>
    Private Function ApplyDslInsertBreak(inst As Instruction, wordApp As Object) As Boolean
        Try
            Dim breakType = inst.GetParam("breakType", "page")?.ToString()
            Dim wdBreakType As Integer = If(breakType = "page", 7, If(breakType = "line", 6, 7))
            wordApp.Selection.GetType().InvokeMember("InsertBreak", Reflection.BindingFlags.InvokeMethod, Nothing, wordApp.Selection, New Object() {wdBreakType})
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ApplyDslInsertBreak: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取Word对齐常量值
    ''' </summary>
    Private Function GetWordAlignmentValue(alignment As String) As Integer
        Select Case alignment.ToLower()
            Case "left" : Return 0
            Case "center" : Return 1
            Case "right" : Return 2
            Case "justify" : Return 3
            Case Else : Return 0
        End Select
    End Function

    ''' <summary>
    ''' 执行Word命令数组
    ''' </summary>
    Private Function ExecuteWordCommandsArray(commandsArray As JToken, originalJson As String, preview As Boolean) As Boolean
        Try
            Dim commands = CType(commandsArray, JArray)
            If commands.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("Массив команд пуст")
                Return False
            End If

            ' 预览所有命令
            If preview Then
                ' 使用增强的预览表单
                If Not ShareRibbon.CommandPreviewForm.ShowPreview($"Предпросмотр команд Word — всего {commands.Count}", commandsArray) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行所有命令
            Dim successCount = 0
            Dim failCount = 0

            For Each cmd In commands
                If cmd.Type = JTokenType.Object Then
                    Dim cmdObj = CType(cmd, JObject)
                    If ExecuteWordCommand(cmdObj) Then
                        successCount += 1
                    Else
                        failCount += 1
                    End If
                End If
            Next

            If failCount = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"Все команды выполнены успешно: {successCount}")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"Выполнено: успешно {successCount}, с ошибкой {failCount}")
            End If

            Return failCount = 0

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommandsArray 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка пакетного выполнения: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行单个Word命令
    ''' </summary>
    Private Function ExecuteWordSingleCommand(commandJson As JObject, processedJson As String, preview As Boolean) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()

            ' 预览 - 使用增强的预览表单
            If preview Then
                If Not ShareRibbon.CommandPreviewForm.ShowPreview("Предпросмотр команд Word", commandJson) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行命令
            Dim success = ExecuteWordCommand(commandJson)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"Команда '{command}' выполнена успешно")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"Команда '{command}' выполнена с ошибкой")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordSingleCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行具体的Word命令
    ''' </summary>
    Private Function ExecuteWordCommand(commandJson As JObject) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")

            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = Globals.ThisAddIn.Application.Selection

            Select Case command.ToLower()
                Case "inserttext"
                    Return ExecuteInsertText(params, selection)
                Case "formattext"
                    Return ExecuteFormatText(params, selection)
                Case "replacetext"
                    Return ExecuteReplaceText(params, doc)
                Case "deletetext"
                    Return ExecuteDeleteText(params, doc, selection)
                Case "inserttable"
                    Return ExecuteInsertTable(params, selection)
                Case "applystyle"
                    Return ExecuteApplyStyle(params, selection)
                Case "generatetoc"
                    Return ExecuteGenerateTOC(params, doc)
                Case "beautifydocument"
                    Return ExecuteBeautifyDocument(params, doc)
                Case "findandformat"
                    Return ExecuteFindAndFormat(params)
                Case "listparagraphs"
                    Return ExecuteListParagraphs(params)
                Case "getparagraphinfo"
                    Return ExecuteGetParagraphInfo(params)
                Case "setparagraphformat"
                    Return ExecuteSetParagraphFormat(params)
                Case Else
                    Debug.WriteLine($"不支持的Word命令: {command}")
                    Return False
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertText(params As JToken, selection As Object) As Boolean
        Try
            Dim content = params("content")?.ToString()
            Dim position = If(params("position")?.ToString(), "cursor")

            Select Case position.ToLower()
                Case "start"
                    Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0).InsertBefore(content)
                Case "end"
                    Globals.ThisAddIn.Application.ActiveDocument.Content.InsertAfter(content)
                Case Else ' cursor
                    selection.TypeText(content)
            End Select

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteFormatText(params As JToken, selection As Object) As Boolean
        Try
            Dim range = selection.Range

            If params("bold") IsNot Nothing Then
                range.Font.Bold = If(params("bold").Value(Of Boolean)(), -1, 0)
            End If

            If params("italic") IsNot Nothing Then
                range.Font.Italic = If(params("italic").Value(Of Boolean)(), -1, 0)
            End If

            If params("underline") IsNot Nothing Then
                range.Font.Underline = If(params("underline").Value(Of Boolean)(), 1, 0)
            End If

            If params("fontSize") IsNot Nothing Then
                range.Font.Size = params("fontSize").Value(Of Single)()
            End If

            If params("fontName") IsNot Nothing Then
                range.Font.Name = params("fontName").ToString()
            End If

            If params("color") IsNot Nothing Then
                Dim colorStr As String = params("color").ToString()
                If Not String.IsNullOrEmpty(colorStr) AndAlso colorStr <> "auto" Then
                    Try
                        Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                    Catch ex As Exception
                        Debug.WriteLine("设置字体颜色失败: " & ex.Message)
                    End Try
                End If
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteReplaceText(params As JToken, doc As Object) As Boolean
        Try
            Dim find = params("find")?.ToString()
            Dim replace = If(params("replace")?.ToString(), "")
            Dim matchCase = If(params("matchCase")?.Value(Of Boolean)(), False)

            Dim findObj = doc.Content.Find
            findObj.ClearFormatting()
            findObj.Replacement.ClearFormatting()
            findObj.Text = find
            findObj.Replacement.Text = replace
            findObj.Forward = True
            findObj.Wrap = 1 ' wdFindContinue
            findObj.MatchCase = matchCase
            findObj.Execute(Replace:=2) ' wdReplaceAll

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteReplaceText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteDeleteText(params As JToken, doc As Object, selection As Object) As Boolean
        Try
            Dim rangeName = "selection"
            If params IsNot Nothing AndAlso params("range") IsNot Nothing Then
                rangeName = params("range").ToString()
            End If
            rangeName = rangeName.Trim().ToLowerInvariant()

            Select Case rangeName
                Case "all", "document", "全文"
                    Dim contentRange = doc.Content
                    contentRange.Delete()
                Case Else
                    If selection Is Nothing OrElse selection.Range Is Nothing Then Return False
                    selection.Range.Delete()
            End Select

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteDeleteText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertTable(params As JToken, selection As Object) As Boolean
        Try
            Dim rows = params("rows")?.Value(Of Integer)()
            Dim cols = params("cols")?.Value(Of Integer)()

            If rows <= 0 OrElse cols <= 0 Then Return False

            Dim table = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
                selection.Range, rows, cols)

            ' 如果有data，填充表格
            Dim data = params("data")
            If data IsNot Nothing AndAlso data.Type = JTokenType.Array Then
                Dim dataArr = CType(data, JArray)
                Dim x As Integer = dataArr.Count - 1
                Dim x2 As Integer = rows - 1
                For rowIdx = 0 To Math.Min(x, x2)
                    Dim rowData = dataArr(rowIdx)
                    If rowData.Type = JTokenType.Array Then
                        Dim rowArr = CType(rowData, JArray)
                        Dim y As Integer = rowArr.Count - 1
                        Dim y1 As Integer = cols - 1
                        For colIdx = 0 To Math.Min(y, y1)
                            table.Cell(rowIdx + 1, colIdx + 1).Range.Text = rowArr(colIdx).ToString()
                        Next
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertTable 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteApplyStyle(params As JToken, selection As Object) As Boolean
        Try
            Dim styleName = params("styleName")?.ToString()
            If String.IsNullOrEmpty(styleName) Then Return False

            ' 检查样式是否存在
            Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim styleExists As Boolean = False
            Try
                Dim testStyle = doc.Styles(styleName)
                styleExists = True
            Catch
                styleExists = False
            End Try

            If Not styleExists Then
                Debug.WriteLine($"ExecuteApplyStyle: 样式 '{styleName}' 不存在，跳过应用")
                ' 尝试使用内置样式名称映射
                Dim builtinStyleName = MapToBuiltinStyle(styleName)
                If Not String.IsNullOrEmpty(builtinStyleName) Then
                    Try
                        selection.Style = builtinStyleName
                        Return True
                    Catch
                        Debug.WriteLine($"ExecuteApplyStyle: 内置样式 '{builtinStyleName}' 也无法应用")
                    End Try
                End If
                Return True ' 返回True避免中断后续命令执行
            End If

            selection.Style = styleName
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyStyle 出错: {ex.Message}")
            Return True ' 返回True避免因样式问题中断整个流程
        End Try
    End Function

    ''' <summary>
    ''' 将常见样式名称映射到Word内置样式
    ''' </summary>
    Private Function MapToBuiltinStyle(styleName As String) As String
        Dim styleMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"标题", "标题 1"},
            {"Title", "Title"},
            {"标题1", "标题 1"},
            {"标题2", "标题 2"},
            {"标题3", "标题 3"},
            {"Heading1", "Heading 1"},
            {"Heading2", "Heading 2"},
            {"Heading3", "Heading 3"},
            {"正文", "正文"},
            {"Normal", "Normal"},
            {"副标题", "副标题"},
            {"Subtitle", "Subtitle"}
        }

        If styleMap.ContainsKey(styleName) Then
            Return styleMap(styleName)
        End If
        Return Nothing
    End Function

#Region "高级Word命令实现"

    ''' <summary>
    ''' 生成目录
    ''' </summary>
    Private Function ExecuteGenerateTOC(params As JToken, doc As Object) As Boolean
        Try
            Dim position = If(params("position")?.ToString(), "start")
            Dim levels = If(params("levels")?.Value(Of Integer)(), 3)
            Dim includePageNumbers = If(params("includePageNumbers")?.Value(Of Boolean)(), True)

            ' 确定插入位置
            Dim range As Object
            If position.ToLower() = "start" Then
                range = doc.Range(0, 0)
            Else
                range = Globals.ThisAddIn.Application.Selection.Range
            End If

            ' 删除已有目录
            For Each toc In doc.TablesOfContents
                toc.Delete()
            Next

            ' 插入新目录
            Dim newToc = doc.TablesOfContents.Add(
                Range:=range,
                UseHeadingStyles:=True,
                UpperHeadingLevel:=1,
                LowerHeadingLevel:=levels,
                IncludePageNumbers:=includePageNumbers
            )

            ' 更新目录
            newToc.Update()

            ShareRibbon.GlobalStatusStrip.ShowInfo($"Оглавление сформировано (уровней: {levels})")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteGenerateTOC 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 文档美化
    ''' </summary>
    Private Function ExecuteBeautifyDocument(params As JToken, doc As Object) As Boolean
        Try
            Dim theme = params("theme")
            Dim margins = params("margins")

            ' 应用页边距
            If margins IsNot Nothing Then
                ApplyMargins(doc, margins)
            End If

            ' 应用主题样式
            If theme IsNot Nothing Then
                ApplyThemeStyles(doc, theme)
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo("Оформление документа завершено")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteBeautifyDocument 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 查找并格式化 - 自然语言定位 + 自动操作
    ''' </summary>
    Private Function ExecuteFindAndFormat(params As JToken) As Boolean
        Try
            If _smartFormatter Is Nothing Then
                Debug.WriteLine("[ExecuteFindAndFormat] SmartFormatter 未初始化")
                Return False
            End If

            Dim query As String = params("query")?.ToString()
            Dim action As String = params("action")?.ToString()

            If String.IsNullOrEmpty(query) OrElse String.IsNullOrEmpty(action) Then
                Debug.WriteLine("[ExecuteFindAndFormat] query 或 action 参数为空")
                Return False
            End If

            Debug.WriteLine($"[ExecuteFindAndFormat] 查询: {query}, 操作: {action}")

            ' 调用智能格式化服务
            Dim success As Boolean = _smartFormatter.FormatByQuery(query, action)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"Выполнено: {query} → {action}")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка операции: {query} → {action}")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecuteFindAndFormat 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка выполнения: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 列出段落 - Harness 架构原子工具
    ''' </summary>
    Private Function ExecuteListParagraphs(params As JToken) As Boolean
        Try
            If _paragraphService Is Nothing Then
                Debug.WriteLine("[ExecuteListParagraphs] ParagraphService 未初始化")
                Return False
            End If

            Dim maxCount As Integer = If(params("maxCount")?.Value(Of Integer)(), 50)
            Dim result As JArray = _paragraphService.ListParagraphs(maxCount)

            ' 将结果显示给用户（通过状态栏或返回到 AI）
            ShareRibbon.GlobalStatusStrip.ShowInfo($"Найдено абзацев: {result.Count}")
            Debug.WriteLine($"[ExecuteListParagraphs] 返回: {result.ToString()}")

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteListParagraphs 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteListParagraphsWithToolResult(params As JToken) As Agent.ToolResult
        Try
            If _paragraphService Is Nothing Then
                Return Agent.ToolResult.Failed("ListParagraphs",
                                               "ParagraphService не инициализирован",
                                               errorCode:=ExceptionClassifier.CodeUnknown,
                                               userMessage:="Служба абзацев Word не инициализирована",
                                               recoverable:=True)
            End If

            Dim maxCount As Integer = If(params?("maxCount")?.Value(Of Integer)(), 50)
            Dim result As JArray = _paragraphService.ListParagraphs(maxCount)
            Dim total As Integer = result.Count
            Try
                total = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count
            Catch
            End Try

            Dim data As New JObject From {
                {"items", result},
                {"total", total},
                {"returned", result.Count},
                {"truncated", result.Count >= maxCount AndAlso total > result.Count}
            }
            Return Agent.ToolResult.Succeed("ListParagraphs",
                                            $"Прочитано абзацев: {result.Count}/{total}",
                                            data)
        Catch ex As Exception
            Debug.WriteLine($"ExecuteListParagraphsWithToolResult 出错: {ex.Message}")
            Return Agent.ToolResult.FromException("ListParagraphs", ex)
        End Try
    End Function

    ''' <summary>
    ''' 获取段落详情 - Harness 架构原子工具
    ''' </summary>
    Private Function ExecuteGetParagraphInfo(params As JToken) As Boolean
        Try
            If _paragraphService Is Nothing Then
                Debug.WriteLine("[ExecuteGetParagraphInfo] ParagraphService 未初始化")
                Return False
            End If

            Dim paragraphIndex As Integer = params("paragraphIndex")?.Value(Of Integer)()
            If paragraphIndex < 1 Then
                Debug.WriteLine("[ExecuteGetParagraphInfo] 无效的段落索引")
                Return False
            End If

            Dim result As JObject = _paragraphService.GetParagraphInfo(paragraphIndex)
            If result Is Nothing Then
                ShareRibbon.GlobalStatusStrip.ShowWarning($"Не удалось получить информацию об абзаце {paragraphIndex}")
                Return False
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"Абзац {paragraphIndex}: {result("style")} {result("fontSize")}pt")
            Debug.WriteLine($"[ExecuteGetParagraphInfo] 返回: {result.ToString()}")

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteGetParagraphInfo 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteGetParagraphInfoWithToolResult(params As JToken) As Agent.ToolResult
        Try
            If _paragraphService Is Nothing Then
                Return Agent.ToolResult.Failed("GetParagraphInfo",
                                               "ParagraphService не инициализирован",
                                               errorCode:=ExceptionClassifier.CodeUnknown,
                                               userMessage:="Служба абзацев Word не инициализирована",
                                               recoverable:=True)
            End If

            Dim paragraphIndex As Integer = If(params?("paragraphIndex")?.Value(Of Integer)(), 1)
            If paragraphIndex < 1 Then
                Return Agent.ToolResult.Failed("GetParagraphInfo",
                                               "Индекс абзаца должен быть не меньше 1",
                                               errorCode:=ExceptionClassifier.CodeArgument,
                                               userMessage:="Индекс абзаца должен быть не меньше 1",
                                               recoverable:=True)
            End If

            Dim result As JObject = _paragraphService.GetParagraphInfo(paragraphIndex)
            If result Is Nothing Then
                Return Agent.ToolResult.Failed("GetParagraphInfo",
                                               $"Абзац не найден: {paragraphIndex}",
                                               errorCode:=ExceptionClassifier.CodeNotFound,
                                               userMessage:=$"Абзац не найден: {paragraphIndex}",
                                               recoverable:=True)
            End If

            Return Agent.ToolResult.Succeed("GetParagraphInfo",
                                            $"Информация об абзаце {paragraphIndex}",
                                            result)
        Catch ex As Exception
            Debug.WriteLine($"ExecuteGetParagraphInfoWithToolResult 出错: {ex.Message}")
            Return Agent.ToolResult.FromException("GetParagraphInfo", ex)
        End Try
    End Function

    ''' <summary>
    ''' 设置段落格式 - Harness 架构原子工具
    ''' </summary>
    Private Function ExecuteSetParagraphFormat(params As JToken) As Boolean
        Try
            If _paragraphService Is Nothing Then
                Debug.WriteLine("[ExecuteSetParagraphFormat] ParagraphService 未初始化")
                Return False
            End If

            Dim paragraphIndex As Integer = params("paragraphIndex")?.Value(Of Integer)()
            If paragraphIndex < 1 Then
                Debug.WriteLine("[ExecuteSetParagraphFormat] 无效的段落索引")
                Return False
            End If

            Dim success As Boolean = _paragraphService.SetParagraphFormat(paragraphIndex, CType(params, JObject))

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"Формат абзаца {paragraphIndex} обновлён")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"Не удалось обновить формат абзаца {paragraphIndex}")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecuteSetParagraphFormat 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка выполнения: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 应用页边距
    ''' </summary>
    Private Sub ApplyMargins(doc As Object, margins As JToken)
        Try
            Dim pageSetup = doc.PageSetup

            ' 单位转换: 厘米 -> 磅 (1cm = 28.35磅)
            Const cmToPoints As Single = 28.35F

            If margins("top") IsNot Nothing Then
                pageSetup.TopMargin = margins("top").Value(Of Single)() * cmToPoints
            End If
            If margins("bottom") IsNot Nothing Then
                pageSetup.BottomMargin = margins("bottom").Value(Of Single)() * cmToPoints
            End If
            If margins("left") IsNot Nothing Then
                pageSetup.LeftMargin = margins("left").Value(Of Single)() * cmToPoints
            End If
            If margins("right") IsNot Nothing Then
                pageSetup.RightMargin = margins("right").Value(Of Single)() * cmToPoints
            End If
        Catch ex As Exception
            Debug.WriteLine($"ApplyMargins 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用主题样式到文档
    ''' </summary>
    Private Sub ApplyThemeStyles(doc As Object, theme As JToken)
        Try
            ' 应用标题1样式
            Dim h1Theme = theme("h1")
            If h1Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 1", h1Theme)
            End If

            ' 应用标题2样式
            Dim h2Theme = theme("h2")
            If h2Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 2", h2Theme)
            End If

            ' 应用标题3样式
            Dim h3Theme = theme("h3")
            If h3Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 3", h3Theme)
            End If

            ' 应用正文样式
            Dim bodyTheme = theme("body")
            If bodyTheme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "正文", bodyTheme)

                ' 应用行间距到所有段落
                If bodyTheme("lineSpacing") IsNot Nothing Then
                    Dim lineSpacing = bodyTheme("lineSpacing").Value(Of Single)()
                    For Each para In doc.Paragraphs
                        Try
                            para.LineSpacingRule = 5 ' wdLineSpaceMultiple
                            para.LineSpacing = 12 * lineSpacing ' 12磅 * 倍数
                        Catch
                        End Try
                    Next
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyThemeStyles 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用样式设置
    ''' </summary>
    Private Sub ApplyStyleFromTheme(doc As Object, styleName As String, themeSettings As JToken)
        Try
            Dim style = doc.Styles(styleName)

            If themeSettings("font") IsNot Nothing Then
                style.Font.Name = themeSettings("font").ToString()
            End If
            If themeSettings("size") IsNot Nothing Then
                style.Font.Size = themeSettings("size").Value(Of Single)()
            End If
            If themeSettings("color") IsNot Nothing Then
                Dim colorStr = themeSettings("color").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                style.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
            End If
            If themeSettings("bold") IsNot Nothing Then
                style.Font.Bold = If(themeSettings("bold").Value(Of Boolean)(), -1, 0)
            End If
            If themeSettings("italic") IsNot Nothing Then
                style.Font.Italic = If(themeSettings("italic").Value(Of Boolean)(), -1, 0)
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyStyleFromTheme ({styleName}) 出错: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "语义排版 - .docx模板解析"

    ''' <summary>
    ''' 覆盖基类方法：从.docx文件解析SemanticStyleMapping并推送前端预览
    ''' </summary>
    Protected Overrides Sub HandleUploadDocxTemplateFromPath(filePath As String)
        Try
            Dim mapping = WordTemplateParser.ExtractFromDocx(filePath)
            If mapping Is Nothing OrElse mapping.SemanticTags.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Не удалось разобрать шаблон: не извлечена корректная информация о стилях")
                Return
            End If

            ' 将原始docx拷贝到数据目录，关联到映射
            Try
                Dim templateDir = IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    ConfigSettings.OfficeAiAppDataFolder, "docx_templates")
                If Not IO.Directory.Exists(templateDir) Then
                    IO.Directory.CreateDirectory(templateDir)
                End If
                Dim destPath = IO.Path.Combine(templateDir, mapping.Id & IO.Path.GetExtension(filePath))
                IO.File.Copy(filePath, destPath, True)
                mapping.SourceFilePath = destPath
            Catch ex As Exception
                Debug.WriteLine($"拷贝模板文件到数据目录失败: {ex.Message}")
                ' 非致命错误，继续保存映射
            End Try

            ' 保存映射
            SemanticMappingManager.Instance.AddMapping(mapping)

            ' 序列化为JSON并推送前端预览
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(mapping, Newtonsoft.Json.Formatting.None)
            ExecuteJavaScriptAsyncJS($"showMappingPreview({json});")

            GlobalStatusStrip.ShowInfo($"Шаблон «{mapping.Name}» разобран, извлечено семантических тегов: {mapping.SemanticTags.Count}")
        Catch ex As Exception
            Debug.WriteLine($"HandleUploadDocxTemplateFromPath 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось разобрать шаблон: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 将当前打开文档的格式提取后发送给 AI，生成 SemanticStyleMapping（格式克隆）
    ''' </summary>
    Protected Overrides Async Sub HandleSaveCurrentDocumentAsTemplate()
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            If wordApp Is Nothing OrElse wordApp.Documents.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Нет открытого документа")
                Return
            End If

            Dim extracted = FormatMirrorService.ExtractFormattingFromDocument(wordApp, False)
            If extracted.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Не удалось извлечь информацию о формате из документа")
                Return
            End If

            ' 记录文档名，供 mirror_format 响应时命名映射
            _mirrorFormatDocName = Path.GetFileNameWithoutExtension(wordApp.ActiveDocument.Name)
            If String.IsNullOrEmpty(_mirrorFormatDocName) Then _mirrorFormatDocName = "Формат документа"

            Dim prompt = FormatMirrorService.BuildClonePrompt(extracted)
            GlobalStatusStrip.ShowInfo("Анализ формата документа, подождите…")
            Await Send("Сформируй SemanticStyleMapping на основе следующей информации о формате.", prompt, False, "mirror_format")
        Catch ex As Exception
            Debug.WriteLine($"HandleSaveCurrentDocumentAsTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось проанализировать формат документа: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 捕获 Agent 启动用 Word 上下文快照
    ''' </summary>
    Protected Overrides Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
        Try
            Return New Context.WordContextProvider(Globals.ThisAddIn.Application).GetContext()
        Catch ex As Exception
            Debug.WriteLine($"CaptureOfficeContext 出错: {ex.Message}")
            Return New Agent.Context.OfficeContext With {.AppType = appType}
        End Try
    End Function

#End Region

End Class
