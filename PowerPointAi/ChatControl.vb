Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports PowerPointAi.Extensions
Imports ShareRibbon.Extensions
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl



    ' 排版上下文：存储待格式化的形状和类型信息
    Private _reformatShapes As List(Of Object) = Nothing
    Private _reformatTypes As List(Of String) = Nothing

    ' Последняя подробная ошибка хостового PPT-инструмента для ToolResult/repair.
    Private _lastPptCommandError As String = ""

    ' Последнее некритичное предупреждение выполненной команды (например, обрезан список значений).
    Private _lastPptCommandWarning As String = ""

    ' Дополнительная информация об успешной команде (например, индекс созданной копии).
    Private _lastPptCommandInfo As String = ""

    ' Индекс слайда, на который реально повлияла последняя команда (копия после DuplicateSlide).
    Private _lastPptCommandTargetIndex As Integer = 0

    ' 排版撤销快照（PPT不支持UndoRecord，使用自定义快照）
    Private _reformatUndoSnapshots As List(Of ShapeFormatSnapshot) = Nothing

    ''' <summary>
    ''' 形状格式快照 - 用于PPT排版撤销
    ''' </summary>
    Private Class ShapeFormatSnapshot
        Public ShapeIndex As Integer
        Public FontNameFarEast As String = ""
        Public FontName As String = ""
        Public FontSize As Single = 0
        Public Bold As Microsoft.Office.Core.MsoTriState = Microsoft.Office.Core.MsoTriState.msoFalse
        Public Alignment As Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft

        Public Shared Function Capture(shp As Microsoft.Office.Interop.PowerPoint.Shape, index As Integer) As ShapeFormatSnapshot
            Dim snap As New ShapeFormatSnapshot()
            snap.ShapeIndex = index
            Try
                If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim tr = shp.TextFrame.TextRange
                    snap.FontNameFarEast = If(tr.Font.NameFarEast IsNot Nothing, tr.Font.NameFarEast, "")
                    snap.FontName = If(tr.Font.Name IsNot Nothing, tr.Font.Name, "")
                    snap.FontSize = tr.Font.Size
                    snap.Bold = tr.Font.Bold
                    snap.Alignment = tr.ParagraphFormat.Alignment
                End If
            Catch ex As Exception
                Debug.WriteLine($"捕获形状{index}快照失败: {ex.Message}")
            End Try
            Return snap
        End Function

        Public Sub Restore(shp As Microsoft.Office.Interop.PowerPoint.Shape)
            Try
                If shp.HasTextFrame <> Microsoft.Office.Core.MsoTriState.msoTrue Then Return
                Dim tr = shp.TextFrame.TextRange
                If Not String.IsNullOrEmpty(FontNameFarEast) Then tr.Font.NameFarEast = FontNameFarEast
                If Not String.IsNullOrEmpty(FontName) Then tr.Font.Name = FontName
                If FontSize > 0 Then tr.Font.Size = FontSize
                tr.Font.Bold = Bold
                tr.ParagraphFormat.Alignment = Alignment
            Catch ex As Exception
                Debug.WriteLine($"恢复形状{ShapeIndex}快照失败: {ex.Message}")
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' 设置排版上下文，用于规则匹配后应用格式
    ''' </summary>
    Public Sub SetReformatContext(shapes As List(Of Object), types As List(Of String))
        ' 新排版开始时，清空旧的撤销快照（防止不匹配）
        _reformatUndoSnapshots = Nothing
        _reformatShapes = shapes
        _reformatTypes = types
    End Sub

    ''' <summary>
    ''' 使用模板进行排版（覆盖基类方法）
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithTemplate(template As ReformatTemplate)
        Try
            Dim pptApp = Globals.ThisAddIn.Application
            Dim activeWindow = pptApp.ActiveWindow
            Dim selection = activeWindow.Selection

            ' 收集所有待排版的形状
            Dim selectedShapes As New List(Of Microsoft.Office.Interop.PowerPoint.Shape)()
            Dim shapeTypes As New List(Of String)()

            Try
                If selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                    ' 选中了形状
                    For Each shp As Microsoft.Office.Interop.PowerPoint.Shape In selection.ShapeRange
                        If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                           shp.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            selectedShapes.Add(shp)
                            shapeTypes.Add(GetShapeTypeName(shp))
                        End If
                    Next
                ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                    ' 选中了幻灯片，获取所有文本形状
                    For Each slide As Microsoft.Office.Interop.PowerPoint.Slide In selection.SlideRange
                        For Each shp As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                            If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                               shp.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                selectedShapes.Add(shp)
                                shapeTypes.Add(GetShapeTypeName(shp))
                            End If
                        Next
                    Next
                End If
            Catch ex As Exception
                Debug.WriteLine("获取选中内容失败: " & ex.Message)
            End Try

            If selectedShapes.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Сначала выберите слайды или текстовые поля для форматирования.")
                Return
            End If

            ' 统计形状类型
            Dim titleCount = shapeTypes.Where(Function(t) t.Contains("Заголовок")).Count()
            Dim bodyCount = shapeTypes.Where(Function(t) t = "Основной текст" OrElse t = "Текстовое поле").Count()

            ' 采样策略：只取代表性样本（最多5个）
            Dim sampleBlocks As New Newtonsoft.Json.Linq.JArray()
            Dim sampleIndices As New List(Of Integer)()
            Dim totalCount = selectedShapes.Count

            If totalCount <= 5 Then
                For i As Integer = 0 To totalCount - 1
                    sampleIndices.Add(i)
                Next
            Else
                ' 采样：首2、中1、尾2
                sampleIndices.Add(0)
                sampleIndices.Add(1)
                sampleIndices.Add(CInt(Math.Floor(totalCount / 2)))
                sampleIndices.Add(totalCount - 2)
                sampleIndices.Add(totalCount - 1)
            End If

            For Each idx In sampleIndices
                Dim shp = selectedShapes(idx)
                Dim textContent = ""
                Try
                    textContent = shp.TextFrame.TextRange.Text
                    ' 头尾采样：截断过长文本
                    If textContent.Length > 80 Then
                        textContent = textContent.Substring(0, 40) & "..." & textContent.Substring(textContent.Length - 30)
                    End If
                Catch
                End Try

                Dim sampleObj As New Newtonsoft.Json.Linq.JObject()
                sampleObj("sampleIndex") = idx
                sampleObj("text") = textContent
                sampleObj("currentType") = shapeTypes(idx)
                sampleBlocks.Add(sampleObj)
            Next

            ' 显示排版模式吸顶提示
            Await ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")

            ' 构建带模板的系统提示
            Dim systemPrompt As New System.Text.StringBuilder()
            systemPrompt.AppendLine("Ты — помощник по форматированию PowerPoint. Пользователь выбрал шаблон «" & template.Name & "» для форматирования.")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("【Конфигурация шаблона】")
            systemPrompt.AppendLine($"Название шаблона: {template.Name}")
            systemPrompt.AppendLine($"Категория шаблона: {template.Category}")
            systemPrompt.AppendLine($"Описание шаблона: {template.Description}")
            systemPrompt.AppendLine()

            ' 版式配置
            If template.Layout IsNot Nothing AndAlso template.Layout.Elements IsNot Nothing AndAlso template.Layout.Elements.Count > 0 Then
                systemPrompt.AppendLine("Элементы макета (скелет):")
                For Each el In template.Layout.Elements
                    systemPrompt.AppendLine($"  - {el.Name}: {el.Font?.FontNameCN} {el.Font?.FontSize}pt, {el.Paragraph?.Alignment}")
                Next
                systemPrompt.AppendLine()
            End If

            ' 正文样式
            If template.BodyStyles IsNot Nothing AndAlso template.BodyStyles.Count > 0 Then
                systemPrompt.AppendLine("Правила стиля основного текста:")
                For Each style In template.BodyStyles
                    systemPrompt.AppendLine($"  - {style.RuleName}: {style.Font?.FontNameCN} {style.Font?.FontSize}pt")
                Next
                systemPrompt.AppendLine()
            End If

            ' AI说明
            If Not String.IsNullOrEmpty(template.AiGuidance) Then
                systemPrompt.AppendLine("【Пояснения к шаблону】")
                systemPrompt.AppendLine(template.AiGuidance)
                systemPrompt.AppendLine()
            End If

            systemPrompt.AppendLine("【Информация о документе】")
            systemPrompt.AppendLine($"В презентации {totalCount} текстовых полей ({titleCount} заголовков, {bodyCount} основного текста/текстовых полей).")
            systemPrompt.AppendLine($"Тебе отправлено представительных образцов: {sampleIndices.Count}.")
            systemPrompt.AppendLine()

            systemPrompt.AppendLine("【Требования к задаче】")
            systemPrompt.AppendLine("На основе конфигурации шаблона и образцов текстовых полей верни JSON с конкретными правилами форматирования. Формат:")
            systemPrompt.AppendLine("```json")
            systemPrompt.AppendLine("{")
            systemPrompt.AppendLine("  ""rules"": [{""type"": ""title"", ""matchCondition"": ""..."", ""formatting"": {""fontNameCN"": ""黑体"", ""fontSize"": 36, ""bold"": true, ""alignment"": ""center""}}],")
            systemPrompt.AppendLine("  ""sampleClassification"": [{""sampleIndex"": 0, ""appliedRule"": ""title""}],")
            systemPrompt.AppendLine("  ""summary"": ""Пояснение стратегии форматирования""")
            systemPrompt.AppendLine("}")
            systemPrompt.AppendLine("```")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("Пояснение к полям formatting: fontNameCN(китайский шрифт), fontNameEN(латинский шрифт), fontSize(размер шрифта, pt), bold(полужирный), alignment(выравнивание left/center/right)")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть.")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("Ниже приведены образцы текстовых полей:")
            systemPrompt.AppendLine(sampleBlocks.ToString(Newtonsoft.Json.Formatting.Indented))

            ' 保存上下文用于后续应用
            SetReformatContext(selectedShapes.Cast(Of Object).ToList(), shapeTypes)

            ' 发送请求
            Await Send("Выполни форматирование выбранного содержимого по шаблону «" & template.Name & "».", systemPrompt.ToString(), False, "reformat")

            GlobalStatusStrip.ShowInfo("Форматирование по шаблону «" & template.Name & "»...")

        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Ошибка форматирования: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取形状类型名称（用于模板排版）
    ''' </summary>
    Private Function GetShapeTypeName(shp As Microsoft.Office.Interop.PowerPoint.Shape) As String
        Try
            If shp.PlaceholderFormat IsNot Nothing Then
                Select Case shp.PlaceholderFormat.Type
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                         Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        Return "Заголовок"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                        Return "Подзаголовок"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody
                        Return "Основной текст"
                End Select
            End If
        Catch
        End Try
        Return "Текстовое поле"
    End Function

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overrides Function GetOfficeApplicationName() As String
        Return "PowerPoint"
    End Function


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
    End Sub

    '获取选中的内容
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' 转换为 PowerPoint.Selection 对象
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 根据选择类型处理内容
            If selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                ' 处理形状选择
                Dim shapeRange = selection.ShapeRange
                If shapeRange.Count > 0 Then
                    ' 检查是否是表格
                    If shapeRange(1).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        ' 处理表格
                        Dim table = shapeRange(1).Table
                        Dim sb As New StringBuilder()
                        For row As Integer = 1 To table.Rows.Count
                            For col As Integer = 1 To table.Columns.Count
                                sb.Append(table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim())
                                If col < table.Columns.Count Then sb.Append(vbTab)
                            Next
                            sb.AppendLine()
                        Next
                        content = sb.ToString()
                    Else
                        ' 处理普通形状
                        content = "[Выбрано фигур: " & shapeRange.Count & "]"
                        For i = 1 To shapeRange.Count
                            If shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                content &= vbCrLf & shapeRange(i).TextFrame.TextRange.Text
                            End If
                        Next
                    End If
                End If

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                ' 处理文本选择
                content = selection.TextRange.Text

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                ' 处理幻灯片选择
                content = "[Выбрано слайдов: " & selection.SlideRange.Count & "]"
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "PowerPoint",  ' 使用文档名称作为标识
                content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", "")
            )
            Else
                ' 选中没有内容，清除相同 sheetName 的引用
                ClearSelectedContentBySheetName("PowerPoint")
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取PowerPoint选中内容时出错: {ex.Message}")
        End Try
    End Sub

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    Protected Overrides Function GetVBProject() As VBProject
        Try
            Dim project = Globals.ThisAddIn.Application.VBE.ActiveVBProject
            Return project
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Проверка доступа к VBA без обращения к Application.VBE: обращение к VBE
    ''' само по себе вызывает модальный диалог Центра управления безопасностью.
    ''' </summary>
    Private Function IsVbaAllowed() As Boolean
        Try
            If Globals.ThisAddIn.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable Then
                Return False
            End If
        Catch
        End Try

        Try
            Dim officeVersion = Globals.ThisAddIn.Application.Version
            Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey($"Software\Microsoft\Office\{officeVersion}\PowerPoint\Security")
                If key IsNot Nothing Then
                    Dim value = key.GetValue("AccessVBOM")
                    If value IsNot Nothing Then Return Convert.ToInt32(value) = 1
                End If
            End Using
        Catch
        End Try

        Return False
    End Function


    Protected Overrides Function RunCode(code As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(code)
            Return True
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return False
        Catch ex As Exception
            MessageBox.Show("Ошибка при выполнении кода: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 执行前预览代码
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
        Return True
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    ' 返回Office应用类型
    Protected Overrides Function GetOfficeAppType() As String
        Return "PowerPoint"
    End Function

    ' 提供PowerPoint应用程序对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function

    Protected Overrides Sub SendChatMessage(message As String)
        ' 这里可以实现word的特殊逻辑
        Send(message, "", True, "")
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息（重写基类方法）
    ''' </summary>
    Protected Overrides Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"PPT使用意图优化提示词: {intent.IntentType}, 置信度: {intent.Confidence:F2}")

            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

    Protected Overrides Function GetCurrentOfficeContent() As String
        Try
            Dim deck = Context.PowerPointDeckReader.GetDeckText(Globals.ThisAddIn.Application, 12000)
            If Not String.IsNullOrWhiteSpace(deck) Then Return deck
        Catch ex As Exception
            Debug.WriteLine("GetCurrentOfficeContent(PowerPoint) failed: " & ex.Message)
        End Try
        Return MyBase.GetCurrentOfficeContent()
    End Function

    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' 复用当前 PPT 进程，WithWindow=msoFalse 静默打开，避免创建第二个 PPT 实例
            Dim pptApp = Globals.ThisAddIn.Application

            Dim presentation As Microsoft.Office.Interop.PowerPoint.Presentation = Nothing
            Try
                presentation = pptApp.Presentations.Open(filePath,
                    ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue,
                    Untitled:=Microsoft.Office.Core.MsoTriState.msoFalse,
                    WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)

                Dim contentBuilder As New StringBuilder()
                contentBuilder.AppendLine($"Файл: {Path.GetFileName(filePath)}")
                contentBuilder.AppendLine($"Всего слайдов: {presentation.Slides.Count}")
                contentBuilder.AppendLine()

                ' 限制处理的幻灯片数量
                Dim maxSlides As Integer = Math.Min(presentation.Slides.Count, 20)

                For slideIndex As Integer = 1 To maxSlides
                    Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                    Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
                    Try
                        slide = presentation.Slides(slideIndex)
                        contentBuilder.AppendLine($"=== Слайд {slideIndex} ===")

                        ' 遍历幻灯片中的形状
                        shapes = slide.Shapes
                        For shapeIndex As Integer = 1 To shapes.Count
                            Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                            Try
                                shape = shapes(shapeIndex)

                                ' 检查是否有文本框架
                                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    If shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                        Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                                        Try
                                            textRange = shape.TextFrame.TextRange
                                            Dim text As String = textRange.Text.Trim()
                                            If Not String.IsNullOrEmpty(text) Then
                                                ' 判断形状类型
                                                Dim shapeType As String = "Текст"
                                                If shape.PlaceholderFormat IsNot Nothing Then
                                                    Select Case shape.PlaceholderFormat.Type
                                                        Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                                                             Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                                                            shapeType = "Заголовок"
                                                        Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                                                            shapeType = "Подзаголовок"
                                                        Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody
                                                            shapeType = "Основной текст"
                                                    End Select
                                                End If
                                                contentBuilder.AppendLine($"  [{shapeType}] {text}")
                                            End If
                                        Finally
                                            ComObjectHelper.ReleaseComObject(textRange)
                                        End Try
                                    End If
                                End If

                                ' 检查是否是表格
                                If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    Dim table As Microsoft.Office.Interop.PowerPoint.Table = Nothing
                                    Try
                                        table = shape.Table
                                        contentBuilder.AppendLine($"  [Таблица {table.Rows.Count} строк × {table.Columns.Count} столбцов]")
                                        ' 读取表格内容（限制行数）
                                        Dim maxRows = Math.Min(table.Rows.Count, 10)
                                        For rowIdx = 1 To maxRows
                                            Dim rowContent As New StringBuilder("    ")
                                            For colIdx = 1 To table.Columns.Count
                                                Dim cell As Microsoft.Office.Interop.PowerPoint.Cell = Nothing
                                                Dim cellShape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                                                Dim cellTextRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                                                Try
                                                    cell = table.Cell(rowIdx, colIdx)
                                                    cellShape = cell.Shape
                                                    cellTextRange = cellShape.TextFrame.TextRange
                                                    Dim cellText = cellTextRange.Text.Trim()
                                                    If cellText.Length > 20 Then cellText = cellText.Substring(0, 17) & "..."
                                                    rowContent.Append(cellText & " | ")
                                                Catch
                                                Finally
                                                    ComObjectHelper.ReleaseComObject(cellTextRange)
                                                    ComObjectHelper.ReleaseComObject(cellShape)
                                                    ComObjectHelper.ReleaseComObject(cell)
                                                End Try
                                            Next
                                            contentBuilder.AppendLine(rowContent.ToString().TrimEnd(" |".ToCharArray()))
                                        Next
                                    Finally
                                        ComObjectHelper.ReleaseComObject(table)
                                    End Try
                                End If
                            Catch shapeEx As Exception
                                Debug.WriteLine($"处理形状时出错: {shapeEx.Message}")
                            Finally
                                ComObjectHelper.ReleaseComObject(shape)
                            End Try
                        Next
                    Finally
                        ComObjectHelper.ReleaseComObject(shapes)
                        ComObjectHelper.ReleaseComObject(slide)
                    End Try

                    contentBuilder.AppendLine()
                Next

                If presentation.Slides.Count > maxSlides Then
                    contentBuilder.AppendLine($"... всего слайдов: {presentation.Slides.Count}, показаны только первые {maxSlides}")
                End If

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "PowerPoint",
                    .ParsedContent = contentBuilder.ToString(),
                    .RawData = Nothing
                }

            Finally
                If presentation IsNot Nothing Then
                    presentation.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation)
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If
                ' 不调用 pptApp.Quit()，复用的是用户当前 PPT 进程
            End Try
        Catch ex As Exception
            Debug.WriteLine($"解析PowerPoint文件时出错: {ex.Message}")
            Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "PowerPoint",
                .ParsedContent = $"[Ошибка при разборе файла PowerPoint: {ex.Message}]"
            }
        End Try
    End Function
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' 检查是否启用了选择功能
            If Not selectedCellChecked Then
                Return message
            End If

            ' 获取当前 PowerPoint 中的选择
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return message
            End If

            ' 创建内容构建器，格式化选中内容
            Dim contentBuilder As New StringBuilder()
            contentBuilder.AppendLine(vbCrLf & "--- Выбранное пользователем содержимое PowerPoint ---")

            ' 添加演示文稿信息
            Dim activePresentation = Globals.ThisAddIn.Application.ActivePresentation
            If activePresentation IsNot Nothing Then
                contentBuilder.AppendLine($"Презентация: {Path.GetFileName(activePresentation.FullName)}")
                contentBuilder.AppendLine($"Текущий слайд: {Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex}")
                contentBuilder.AppendLine("Нумерация слайдов — с 1, как в интерфейсе. В DuplicateSlide/ReplaceSlideText/UpdateChartData передавай этот номер без вычитания 1.")
            End If

            ' 根据选择类型处理内容
            Select Case selection.Type
                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes
                    ' 处理形状选择（包括表格）
                    Dim shapeRange = selection.ShapeRange
                    contentBuilder.AppendLine($"Тип выделения: фигуры (всего {shapeRange.Count})")

                    For i = 1 To shapeRange.Count
                        contentBuilder.AppendLine($"Фигура {i}:")

                        ' 检查是否是图表：модель должна видеть категории и значения,
                        ' иначе она просит у пользователя данные, которые уже есть на слайде.
                        If shapeRange(i).HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            contentBuilder.AppendLine("  " & PowerPointChartSummary.DescribeChart(shapeRange(i).Chart, i))
                        ElseIf shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim table = shapeRange(i).Table
                            contentBuilder.AppendLine($"  Таблица: {table.Rows.Count} строк × {table.Columns.Count} столбцов")

                            ' 添加表格内容
                            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
                            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

                            ' 处理表格头部
                            Dim headerBuilder As New StringBuilder("  ")
                            Dim separatorBuilder As New StringBuilder("  ")

                            For col = 1 To maxCols
                                Try
                                    Dim cellText = table.Cell(1, col).Shape.TextFrame.TextRange.Text.Trim()
                                    ' 限制单元格文本长度
                                    If cellText.Length > 20 Then
                                        cellText = cellText.Substring(0, 17) & "..."
                                    End If

                                    If col > 1 Then
                                        headerBuilder.Append(" | ")
                                        separatorBuilder.Append("-+-")
                                    End If
                                    headerBuilder.Append(cellText)
                                    separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                                Catch ex As Exception
                                    If col > 1 Then
                                        headerBuilder.Append(" | ")
                                        separatorBuilder.Append("-+-")
                                    End If
                                    headerBuilder.Append("N/A")
                                    separatorBuilder.Append("---")
                                End Try
                            Next

                            contentBuilder.AppendLine(headerBuilder.ToString())
                            contentBuilder.AppendLine(separatorBuilder.ToString())

                            ' 处理表格数据行
                            For row = 2 To maxRows
                                Dim rowBuilder As New StringBuilder("  ")

                                For col = 1 To maxCols
                                    Try
                                        Dim cellText = table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim()
                                        ' 限制单元格文本长度
                                        If cellText.Length > 20 Then
                                            cellText = cellText.Substring(0, 17) & "..."
                                        End If

                                        If col > 1 Then
                                            rowBuilder.Append(" | ")
                                        End If
                                        rowBuilder.Append(cellText)
                                    Catch ex As Exception
                                        If col > 1 Then
                                            rowBuilder.Append(" | ")
                                        End If
                                        rowBuilder.Append("N/A")
                                    End Try
                                Next

                                contentBuilder.AppendLine(rowBuilder.ToString())
                            Next

                            ' 添加表格说明
                            If table.Rows.Count > maxRows Then
                                contentBuilder.AppendLine($"  ... всего строк: {table.Rows.Count}, показаны только первые {maxRows}")
                            End If

                            If table.Columns.Count > maxCols Then
                                contentBuilder.AppendLine($"  ... всего столбцов: {table.Columns.Count}, показаны только первые {maxCols}")
                            End If
                        ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            ' 处理文本框
                            Dim textFrame = shapeRange(i).TextFrame
                            If textFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                Dim text = textFrame.TextRange.Text.Trim()
                                ' 限制文本长度
                                If text.Length > 500 Then
                                    contentBuilder.AppendLine($"  Текст: {text.Substring(0, 500)}...")
                                    contentBuilder.AppendLine($"  [Текст слишком длинный, показаны только первые 500 символов, всего: {text.Length}]")
                                Else
                                    contentBuilder.AppendLine($"  Текст: {text}")
                                End If
                            Else
                                contentBuilder.AppendLine("  [Пустое текстовое поле]")
                            End If
                        ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                            ' 处理图片
                            contentBuilder.AppendLine("  [Изображение]")
                            If shapeRange(i).AlternativeText <> "" Then
                                contentBuilder.AppendLine($"  Альтернативный текст: {shapeRange(i).AlternativeText}")
                            End If
                        Else
                            ' 其他类型的形状
                            contentBuilder.AppendLine($"  [Тип фигуры: {shapeRange(i).Type}]")
                        End If

                        ' 在形状之间添加分隔线
                        contentBuilder.AppendLine("  ---")
                    Next

                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText
                    ' 处理文本选择
                    contentBuilder.AppendLine("Тип выделения: текст")

                    Dim textRange = selection.TextRange
                    If textRange IsNot Nothing Then
                        Dim text = textRange.Text.Trim()
                        ' 限制文本长度
                        If text.Length > 1000 Then
                            contentBuilder.AppendLine(text.Substring(0, 1000) & "...")
                            contentBuilder.AppendLine($"[Текст слишком длинный, показаны только первые 1000 символов, всего: {text.Length}]")
                        Else
                            contentBuilder.AppendLine(text)
                        End If
                    Else
                        contentBuilder.AppendLine("[Не удалось получить текстовое содержимое]")
                    End If

                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides
                    ' 处理幻灯片选择
                    Dim slideRange = selection.SlideRange
                    contentBuilder.AppendLine($"Тип выделения: слайды (всего {slideRange.Count})")

                    ' 限制处理的幻灯片数量
                    Dim maxSlides = Math.Min(slideRange.Count, 5)

                    For i = 1 To maxSlides
                        Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                        Try
                            slide = slideRange(i)
                            contentBuilder.AppendLine($"Слайд {slide.SlideIndex}:")

                            Dim title = GetSlideTitle(slide)
                            If title <> "[Без заголовка]" Then
                                contentBuilder.AppendLine($"  Заголовок: {title}")
                            Else
                                contentBuilder.AppendLine("  [Без заголовка]")
                            End If

                            contentBuilder.Append(GetSlideContent(slide))
                            contentBuilder.AppendLine("  ---")
                        Finally
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next

                    ' 如果有更多幻灯片未显示，添加提示
                    If slideRange.Count > maxSlides Then
                        contentBuilder.AppendLine($"[Выбрано слайдов: {slideRange.Count}, показаны только первые {maxSlides}]")
                    End If

                Case Else
                    contentBuilder.AppendLine($"Тип выделения: неизвестно ({selection.Type})")
                    contentBuilder.AppendLine("[Нераспознанный тип выделения]")
            End Select

            contentBuilder.AppendLine("--- Конец выбранного содержимого ---" & vbCrLf)

            ' 返回原始消息加上选中内容
            Return message & contentBuilder.ToString()

        Catch ex As Exception
            Debug.WriteLine($"处理PowerPoint选中内容时出错: {ex.Message}")
            Return message ' 出错时返回原始消息
        End Try
    End Function

    ' 处理形状选择（包括表格）
    Private Function GetSlideTitle(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            ' 检查幻灯片是否有标题占位符
            Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                    Try
                        shape = shapes(i)
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    textRange = shape.TextFrame.TextRange
                                    Return textRange.Text.Trim()
                                End If
                            End If
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(textRange)
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            ' 如果没有找到标题占位符，尝试查找任何可能的标题
            Try
                shapes = slide.Shapes
                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                    Try
                        shape = shapes(i)
                        If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            textRange = shape.TextFrame.TextRange
                            Dim text = textRange.Text.Trim()
                            If Not String.IsNullOrEmpty(text) AndAlso text.Length < 100 Then
                                Return text ' 假设第一个简短文本是标题
                            End If
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(textRange)
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            Return "[Без заголовка]"
        Catch ex As Exception
            Debug.WriteLine($"获取幻灯片标题时出错: {ex.Message}")
            Return "[Ошибка получения заголовка]"
        End Try
    End Function

    ' 获取幻灯片内容
    Private Function GetSlideContent(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            Dim contentBuilder As New StringBuilder()
            Dim processedTextShapes As Integer = 0
            Dim textLimitNoted As Boolean = False
            Dim chartCount As Integer = 0
            Dim maxTextShapes As Integer = 12 ' 限制每张幻灯片处理的文本形状数量

            ' 处理幻灯片上的形状
            Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                    Dim chart As Microsoft.Office.Interop.PowerPoint.Chart = Nothing
                    Try
                        shape = shapes(i)
                        ' 跳过标题形状，因为已经单独处理过了
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                       shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                            Continue For
                        End If

                        ' 处理文本形状
                        If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                       shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then

                            If processedTextShapes >= maxTextShapes Then
                                If Not textLimitNoted Then
                                    contentBuilder.AppendLine("  [Дополнительное текстовое содержимое не показано...]")
                                    textLimitNoted = True
                                End If
                                ' Не выходим из цикла: иначе диаграммы и таблицы после текста теряются.
                                Continue For
                            End If

                            textRange = shape.TextFrame.TextRange
                            Dim text = textRange.Text.Trim()
                            If Not String.IsNullOrEmpty(text) Then
                                ' 限制文本长度
                                If text.Length > 200 Then
                                    contentBuilder.AppendLine($"  Текст: {text.Substring(0, 200)}...")
                                Else
                                    contentBuilder.AppendLine($"  Текст: {text}")
                                End If
                                processedTextShapes += 1
                            End If
                            ' 处理表格形状
                        ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            contentBuilder.AppendLine("  [Содержит таблицу]")
                            ' 处理图片形状
                        ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                            contentBuilder.AppendLine("  [Содержит изображение]")
                            If shape.AlternativeText <> "" Then
                                contentBuilder.AppendLine($"  Описание изображения: {shape.AlternativeText}")
                            End If
                            ' 处理图表形状
                        ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoChart Then
                            chartCount += 1
                            chart = shape.Chart
                            contentBuilder.Append(PowerPointChartSummary.DescribeChart(chart, chartCount).
                                Replace(vbLf, vbLf & "  ") & vbCrLf)
                            ' 处理SmartArt形状
                        ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoSmartArt Then
                            contentBuilder.AppendLine("  [Содержит графику SmartArt]")
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(chart)
                        ComObjectHelper.ReleaseComObject(textRange)
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            ' 如果没有找到任何内容
            If contentBuilder.Length = 0 Then
                Return "  [На слайде нет извлекаемого текстового содержимого]"
            End If

            Return contentBuilder.ToString()
        Catch ex As Exception
            Debug.WriteLine($"获取幻灯片内容时出错: {ex.Message}")
            Return $"  [Ошибка получения содержимого: {ex.Message}]"
        End Try
    End Function

    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            ' 获取当前活动演示文稿的路径
            If Globals.ThisAddIn.Application.ActivePresentation IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActivePresentation.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取当前工作目录时出错: {ex.Message}")
        End Try

        ' 如果无法获取演示文稿路径，则返回应用程序目录
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function


    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)
        ' 在 AI 操作完成时创建撤销点 - 使用 ErrorHandler 包装
        ErrorHandlerExtension.SafeExecute(
            Sub()
                UndoManagerExtension.CreateAIOperationUndoPoint(
                    "PowerPoint",
                    Globals.ThisAddIn.Application,
                    "AI-операция",
                    "Содержимое, созданное ИИ")
            End Sub,
            "PowerPointAi.ChatControl",
            "Создание точки отмены")

        ' 尝试检测并生成幻灯片 - 使用 ErrorHandler 包装
        ErrorHandlerExtension.SafeExecute(
            Sub()
                If allPlainMarkdownBuffer IsNot Nothing AndAlso allPlainMarkdownBuffer.Length > 0 Then
                    Dim aiResponse As String = allPlainMarkdownBuffer.ToString()

                    ' 检测并生成幻灯片
                    If PptGenerationHandlerExtension.DetectSlideOutline(aiResponse) Then
                        System.Diagnostics.Debug.WriteLine("[ChatControl] 检测到幻灯片大纲")

                        ' 在生成幻灯片前创建撤销点
                        UndoManagerExtension.CreateAIOperationUndoPoint(
                            "PowerPoint",
                            Globals.ThisAddIn.Application,
                            "AI-генерация слайдов",
                            "Генерация слайдов ИИ")

                        PptGenerationHandlerExtension.TryGenerateSlides(aiResponse, Globals.ThisAddIn.Application)
                    End If
                End If
            End Sub,
            "PowerPointAi.ChatControl",
            "Генерация слайдов",
            "Ошибка генерации слайдов; проверьте формат структуры.")

        ' 调用基类处理续写模式
        MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
    End Sub

    ' ========== 续写功能 ==========

    Private _continuationService As PowerPointContinuationService
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
                _continuationService = New PowerPointContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 检查是否可以续写
            If Not _continuationService.CanContinue() Then
                GlobalStatusStrip.ShowWarning("Не удалось получить информацию о презентации. Убедитесь, что документ открыт")
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
                    GlobalStatusStrip.ShowWarning("Не удалось получить контекст слайда")
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
    ''' 应用续写结果到PowerPoint
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
                _continuationService = New PowerPointContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart ' 首页
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd ' 末页
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor ' 当前页
            End Select

            ' 插入续写内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("Продолжение вставлено в слайд")

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
    ''' 应用模板渲染结果到PowerPoint
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
                _continuationService = New PowerPointContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart ' 首页
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd ' 末页
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor ' 当前页
            End Select

            ' 插入模板内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("Содержимое шаблона вставлено в слайд")

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
    ''' 获取当前PowerPoint上下文快照（用于自动补全）
    ''' </summary>
    Protected Overrides Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = "PowerPoint"

        Try
            Dim pres = Globals.ThisAddIn.Application.ActivePresentation
            If pres IsNot Nothing Then
                snapshot("slidesCount") = pres.Slides.Count

                ' 获取当前幻灯片信息
                Try
                    Dim currentSlide As Object = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                    If currentSlide IsNot Nothing Then
                        snapshot("currentSlide") = CInt(CallByName(currentSlide, "SlideIndex", CallType.Get))
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"GetContextSnapshot 获取当前幻灯片失败: {ex.GetType().Name}: {ex.Message}")
                End Try
            End If

            ' 获取选中内容
            Dim selText = ""
            Try
                Dim sel = Globals.ThisAddIn.Application.ActiveWindow.Selection
                If sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                    selText = sel.TextRange.Text
                ElseIf sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                    For i = 1 To Math.Min(sel.ShapeRange.Count, 3)
                        Dim shape = sel.ShapeRange(i)
                        If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                            selText &= shape.TextFrame.TextRange.Text & " "
                        End If
                    Next
                End If
            Catch
            End Try

            If selText.Length > 300 Then
                selText = selText.Substring(0, 300) & "..."
            End If
            snapshot("selection") = selText.Trim()

        Catch ex As Exception
            Debug.WriteLine($"GetContextSnapshot 出错: {ex.Message}")
        End Try

        Return snapshot
    End Function

    ''' <summary>
    ''' 捕获 Agent 启动用 PowerPoint 上下文快照
    ''' </summary>
    Protected Overrides Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
        Try
            Return New Context.PowerPointContextProvider(Globals.ThisAddIn.Application).GetContext()
        Catch ex As Exception
            Debug.WriteLine($"CaptureOfficeContext 出错: {ex.Message}")
            Return New Agent.Context.OfficeContext With {.AppType = appType}
        End Try
    End Function

    ''' <summary>
    ''' 重写保存设置方法，同步更新PPT补全管理器状态
    ''' </summary>
    Protected Overrides Sub HandleSaveSettings(jsonDoc As JObject)
        MyBase.HandleSaveSettings(jsonDoc)
        
        ' 同步更新PPT补全管理器的启用状态
        Try
            Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
            PowerPointCompletionManager.Instance.Enabled = enableAutocomplete
            Debug.WriteLine($"[PPTChatControl] 补全设置已同步: Enabled={enableAutocomplete}")
        Catch ex As Exception
            Debug.WriteLine($"[PPTChatControl] 同步补全设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 执行JSON命令（重写基类方法）- 带严格验证
    ''' </summary>
    Private Function ExecuteJsonCommandCore(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 预览模式下跳过自动执行（排版/校对模式的JSON用于预览，由用户手动点击应用）
            If IsInPreviewMode() Then
                Debug.WriteLine($"[PPTChatControl] 预览模式({GetCurrentResponseMode()})下跳过JSON命令自动执行")
                Return True ' 返回True表示"成功处理"，避免显示错误
            End If

            ' Модель иногда присылает лишний слой экранирования кавычек внутри code/params.
            ' Снимать его заранее нельзя: в валидном JSON \" — это экранированная кавычка
            ' (VBA-код со строковыми литералами, URL в кавычках), и слепая замена ломает документ.
            ' Поэтому сначала пробуем разобрать как есть и откатываемся только при ошибке.
            Dim rawJson As String = jsonCode
            Dim escapedQuote = ChrW(92) & ChrW(34)
            Dim errorMessage As String = ""
            Dim normalizedJson As JToken = Nothing

            If Not PowerPointJsonCommandSchema.ValidateJsonStructure(rawJson, errorMessage, normalizedJson) Then
                Dim unescaped = rawJson.Replace(escapedQuote, ChrW(34))
                If unescaped <> rawJson AndAlso
                   PowerPointJsonCommandSchema.ValidateJsonStructure(unescaped, errorMessage, normalizedJson) Then
                    Debug.WriteLine("[PPTChatControl] JSON распознан после снятия лишнего экранирования")
                    rawJson = unescaped
                Else
                    ' 格式验证失败
                    Debug.WriteLine($"PPT JSON格式验证失败: {errorMessage}")
                    Debug.WriteLine($"原始JSON: {rawJson.Substring(0, Math.Min(200, rawJson.Length))}...")

                    ShareRibbon.GlobalStatusStrip.ShowWarning($"Формат JSON не соответствует спецификации: {errorMessage}")
                    Return False
                End If
            End If
            
            ' 验证通过，根据类型执行
            If normalizedJson.Type = JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, JObject)
                
                ' 命令数组格式
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecutePPTCommandsArray(jsonObj("commands"), rawJson, preview)
                End If
                
                ' 单命令格式
                Return ExecutePPTSingleCommand(jsonObj, rawJson, preview)
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

    ''' <summary>
    ''' 为 Agent 原生 PPT 工具返回真实 ToolResult，并以演示文稿/目标页快照生成最小 Observation。
    ''' </summary>
    Protected Overrides Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
        Dim envelope As JObject = Nothing
        Dim toolId As String = "PowerPointCommands"
        Try
            envelope = ParsePowerPointCommandEnvelope(jsonCode)
            Dim commandArray = TryCast(envelope?("commands"), JArray)
            If commandArray IsNot Nothing AndAlso commandArray.Count > 1 AndAlso
               commandArray.OfType(Of JObject)().Any(Function(item)
                   Return String.Equals(item("command")?.ToString(), "CreateSlides", StringComparison.OrdinalIgnoreCase)
               End Function) Then
                Return Agent.ToolResult.Failed("CreateSlides",
                                               "Professional CreateSlides must be the only command in its envelope",
                                               errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                               userMessage:="Профессиональная генерация всей презентации должна быть отдельным вызовом CreateSlides; разбейте последующие операции с изображениями, диаграммами или объектами на следующий шаг",
                                               recoverable:=True,
                                               observation:=New JObject From {
                                                   {"kind", "write"},
                                                    {"summary", "CreateSlides не выполнен: обнаружен массив смешанных команд"},
                                                   {"changed", False},
                                                   {"warnings", New JArray("mixed_create_slides_envelope")}
                                               })
            End If
            toolId = GetPowerPointEnvelopeToolId(envelope)
            If String.Equals(toolId, "DiscoverOfficeCapability", StringComparison.OrdinalIgnoreCase) Then
                Return OfficeRuntime.PowerPointApiCatalogProvider.SearchAsToolResult(
                    GetPowerPointEnvelopeParams(envelope))
            End If

            If String.Equals(toolId, "ListSlideTemplates", StringComparison.OrdinalIgnoreCase) Then
                Return ListSlideTemplatesAsToolResult(GetPowerPointEnvelopeParams(envelope))
            End If

            If String.Equals(toolId, "ExtractTemplateFromSlide", StringComparison.OrdinalIgnoreCase) Then
                Return ExtractTemplateFromSlideAsToolResult(GetPowerPointEnvelopeParams(envelope))
            End If

            If String.Equals(toolId, "ExtractTemplateSetFromDeck", StringComparison.OrdinalIgnoreCase) Then
                Return ExtractTemplateSetFromDeckAsToolResult(GetPowerPointEnvelopeParams(envelope))
            End If

            If String.Equals(toolId, "AddSlideFromTemplate", StringComparison.OrdinalIgnoreCase) Then
                Return AddSlideFromTemplateAsToolResult(GetPowerPointEnvelopeParams(envelope))
            End If

            If String.Equals(toolId, "AddSlideFromNearestSlide", StringComparison.OrdinalIgnoreCase) Then
                Return AddSlideFromNearestSlideAsToolResult(GetPowerPointEnvelopeParams(envelope))
            End If

            If String.Equals(toolId, "BuildDeckFromTemplateSet", StringComparison.OrdinalIgnoreCase) Then
                Return BuildDeckFromTemplateSetAsToolResult(GetPowerPointEnvelopeParams(envelope))
            End If
            If String.Equals(toolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) Then
                Return Design.ProfessionalDeckExecutor.ExecuteAsToolResult(
                    GetPowerPointEnvelopeParams(envelope), preview)
            End If
            If String.Equals(toolId, "OfficeObjectOperation", StringComparison.OrdinalIgnoreCase) Then
                Return OfficeRuntime.PowerPointOperationExecutor.Execute(
                    GetPowerPointEnvelopeParams(envelope))
            End If
            Dim beforeSnapshot = CapturePowerPointCommandSnapshot(envelope)
            Dim success = ExecuteJsonCommandCore(jsonCode, preview)
            Dim afterSnapshot = CapturePowerPointCommandSnapshot(envelope)
            Dim observation = BuildPowerPointCommandObservation(toolId, success, beforeSnapshot, afterSnapshot)
            Dim summary = If(observation?("summary")?.ToString(), $"{toolId}: выполнение завершено")
            Dim failureDetail = _lastPptCommandError
            _lastPptCommandError = ""

            Dim commandWarning = _lastPptCommandWarning
            _lastPptCommandWarning = ""
            Dim commandInfo = _lastPptCommandInfo
            _lastPptCommandInfo = ""

            If success Then
                If Not String.IsNullOrWhiteSpace(commandInfo) Then
                    summary = summary & " (" & commandInfo & ")"
                End If
                If Not String.IsNullOrWhiteSpace(commandWarning) Then
                    Dim warnings = TryCast(observation("warnings"), JArray)
                    If warnings IsNot Nothing Then warnings.Add(commandWarning)
                    summary = summary & " (" & commandWarning & ")"
                End If
                Return Agent.ToolResult.Succeed(toolId,
                                                summary,
                                                data:=New With {.targetRefs = observation("targetRefs")},
                                                observation:=observation)
            End If

            If Not String.IsNullOrWhiteSpace(failureDetail) Then
                Dim warnings = TryCast(observation("warnings"), JArray)
                If warnings IsNot Nothing Then warnings.Add(failureDetail)
            End If
            Dim failureMessage = If(String.IsNullOrWhiteSpace(failureDetail), summary, failureDetail)

            Return Agent.ToolResult.Failed(toolId,
                                           failureMessage,
                                           errorCode:=ExceptionClassifier.CodeUnknown,
                                           userMessage:=failureMessage,
                                           debugDetail:=failureDetail,
                                           recoverable:=True,
                                           observation:=observation)
        Catch ex As Exception
            Return Agent.ToolResult.FromException(toolId, ex)
        End Try
    End Function

    ''' <summary>
    ''' Компактный каталог сохранённых макетов для агента. Только чтение, без COM.
    ''' </summary>
    Private Shared Function ListSlideTemplatesAsToolResult(params As JObject) As Agent.ToolResult
        Try
            Dim limit As Integer = 20
            Dim limitToken = params?("limit")
            If limitToken IsNot Nothing Then
                Dim parsed As Integer
                If Integer.TryParse(limitToken.ToString(), parsed) AndAlso parsed > 0 Then limit = parsed
            End If

            Dim archetype = params?("archetype")?.ToString()
            Dim query = params?("query")?.ToString()
            Dim summaries = Agent.DeckTemplates.TemplateCatalog.ListSummaries(limit, archetype, query)
            Dim items = JArray.FromObject(summaries)
            Dim summaryText = If(summaries.Count = 0,
                                 "Сохранённых макетов не найдено",
                                 $"Найдено макетов: {summaries.Count}")
            Dim observation = New JObject From {
                {"kind", "read"},
                {"summary", summaryText},
                {"changed", False},
                {"targetRefs", New JArray()},
                {"items", items},
                {"warnings", New JArray()}
            }
            Return Agent.ToolResult.Succeed("ListSlideTemplates",
                                            summaryText,
                                            data:=New With {.templates = items},
                                            observation:=observation)
        Catch ex As Exception
            Return Agent.ToolResult.Failed("ListSlideTemplates",
                                           "Не удалось прочитать каталог макетов: " & ex.Message,
                                           errorCode:=ExceptionClassifier.CodeUnknown,
                                           userMessage:="Каталог макетов недоступен",
                                           debugDetail:=ex.Message,
                                           recoverable:=True)
        End Try
    End Function

    ''' <summary>
    ''' Извлечение одного слайда в каталог макетов. Пишет только файлы, документ не меняет.
    ''' </summary>
    Private Shared Function ExtractTemplateFromSlideAsToolResult(params As JObject) As Agent.ToolResult
        Dim warnings As New List(Of String)()
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim slideIndex = If(params?("slideIndex")?.Value(Of Integer)(), 0)
            If slideIndex <= 0 Then
                Try
                    slideIndex = CInt(app.ActiveWindow.View.Slide.SlideIndex)
                Catch
                    slideIndex = 1
                End Try
            End If

            Dim errorMessage As String = ""
            Dim manifest = Design.DeckTemplate.DeckTemplateExtractor.ExtractFromSlide(
                app, slideIndex,
                params?("name")?.ToString(),
                params?("setId")?.ToString(),
                params?("archetype")?.ToString(),
                errorMessage, warnings)

            If manifest Is Nothing Then
                Return Agent.ToolResult.Failed("ExtractTemplateFromSlide",
                                               If(String.IsNullOrWhiteSpace(errorMessage), "Макет не сохранён", errorMessage),
                                               errorCode:=ExceptionClassifier.CodeUnknown,
                                               userMessage:=If(String.IsNullOrWhiteSpace(errorMessage), "Не удалось сохранить макет", errorMessage),
                                               debugDetail:=errorMessage,
                                               recoverable:=True)
            End If

            Dim warningsArray = New JArray(warnings)
            Dim hasPreview = False
            Try
                hasPreview = Not String.IsNullOrWhiteSpace(manifest.PreviewFullPath) AndAlso IO.File.Exists(manifest.PreviewFullPath)
            Catch
            End Try
            Dim summary = $"Макет сохранён: {manifest.Name} ({manifest.Archetype}), слотов: {manifest.Slots.Count}"
            Dim observation = New JObject From {
                {"kind", "write"},
                {"summary", summary},
                {"changed", False},
                {"targetRefs", New JArray($"PowerPoint:Slide/{manifest.Source.SlideIndex}")},
                {"warnings", warningsArray}
            }
            Return Agent.ToolResult.Succeed("ExtractTemplateFromSlide",
                                            summary,
                                            data:=New With {
                                                .templateId = manifest.Id,
                                                .setId = manifest.SetId,
                                                .name = manifest.Name,
                                                .archetype = manifest.Archetype,
                                                .slotCount = manifest.Slots.Count,
                                                .hasPreview = hasPreview
                                            },
                                            observation:=observation)
        Catch ex As Exception
            Return Agent.ToolResult.Failed("ExtractTemplateFromSlide",
                                           "Извлечение макета: " & ex.Message,
                                           errorCode:=ExceptionClassifier.CodeUnknown,
                                           userMessage:="Не удалось сохранить макет",
                                           debugDetail:=ex.Message,
                                           recoverable:=True)
        End Try
    End Function

    ''' <summary>
    ''' Извлечение всех непустых слайдов колоды как набора макетов.
    ''' </summary>
    Private Shared Function ExtractTemplateSetFromDeckAsToolResult(params As JObject) As Agent.ToolResult
        Dim warnings As New List(Of String)()
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim maxSlides = If(params?("maxSlides")?.Value(Of Integer)(), 20)
            Dim skipEmpty = If(params?("skipEmpty")?.Value(Of Boolean)(), True)
            Dim errorMessage As String = ""
            Dim manifests = Design.DeckTemplate.DeckTemplateExtractor.ExtractFromDeck(
                app,
                params?("setName")?.ToString(),
                maxSlides,
                skipEmpty,
                errorMessage,
                warnings)

            If manifests Is Nothing OrElse manifests.Count = 0 Then
                Return Agent.ToolResult.Failed("ExtractTemplateSetFromDeck",
                                               If(String.IsNullOrWhiteSpace(errorMessage), "Набор макетов не сохранён", errorMessage),
                                               errorCode:=ExceptionClassifier.CodeUnknown,
                                               userMessage:=If(String.IsNullOrWhiteSpace(errorMessage), "Не удалось сохранить набор макетов", errorMessage),
                                               debugDetail:=errorMessage,
                                               recoverable:=True)
            End If

            Dim setId = manifests(0).SetId
            Dim items = New JArray()
            For Each manifest In manifests
                items.Add(New JObject From {
                    {"templateId", manifest.Id},
                    {"name", manifest.Name},
                    {"archetype", manifest.Archetype},
                    {"slotCount", manifest.Slots.Count},
                    {"sourceSlide", manifest.Source.SlideIndex}
                })
            Next
            Dim warningsArray = New JArray(warnings)
            Dim summary = $"Сохранён набор макетов {setId}: {manifests.Count}"
            Dim observation = New JObject From {
                {"kind", "write"},
                {"summary", summary},
                {"changed", False},
                {"targetRefs", New JArray()},
                {"warnings", warningsArray}
            }
            Return Agent.ToolResult.Succeed("ExtractTemplateSetFromDeck",
                                            summary,
                                            data:=New With {
                                                .setId = setId,
                                                .count = manifests.Count,
                                                .templates = items
                                            },
                                            observation:=observation)
        Catch ex As Exception
            Return Agent.ToolResult.Failed("ExtractTemplateSetFromDeck",
                                           "Извлечение набора макетов: " & ex.Message,
                                           errorCode:=ExceptionClassifier.CodeUnknown,
                                           userMessage:="Не удалось сохранить набор макетов",
                                           debugDetail:=ex.Message,
                                           recoverable:=True)
        End Try
    End Function

    ''' <summary>
    ''' Вставка сохранённого макета в активную колоду с заполнением слотов.
    ''' </summary>
    Private Shared Function AddSlideFromTemplateAsToolResult(params As JObject) As Agent.ToolResult
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim target = app.ActivePresentation
            If target Is Nothing Then Return InsertToolFailure("AddSlideFromTemplate", "Нет активной презентации")

            Dim templateId = params?("templateId")?.ToString()
            If String.IsNullOrWhiteSpace(templateId) Then
                Return InsertToolFailure("AddSlideFromTemplate", "Не указан templateId")
            End If
            Dim template = Agent.DeckTemplates.TemplateCatalog.FindById(templateId)
            If template Is Nothing Then
                Return InsertToolFailure("AddSlideFromTemplate", $"Макет не найден: {templateId}")
            End If

            Dim spec = ParseSlideSpecToken(params?("spec"))
            If spec Is Nothing Then
                Return InsertToolFailure("AddSlideFromTemplate", "Параметр spec некорректен")
            End If

            Dim insertAt = ResolveInsertAt(app, target, params)
            Dim result = Design.DeckTemplate.DeckTemplateRenderer.InsertFromTemplate(app, target, template, spec, insertAt)
            Return BuildInsertToolResult("AddSlideFromTemplate", result, template.Id, 0, 0, Nothing)
        Catch ex As Exception
            Return InsertToolFailure("AddSlideFromTemplate", "Вставка макета: " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Вставка слайда на основе ближайшего макета: явный каталог, иначе текущая колода.
    ''' </summary>
    Private Shared Function AddSlideFromNearestSlideAsToolResult(params As JObject) As Agent.ToolResult
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim target = app.ActivePresentation
            If target Is Nothing Then Return InsertToolFailure("AddSlideFromNearestSlide", "Нет активной презентации")

            Dim spec = ParseSlideSpecToken(params?("spec"))
            If spec Is Nothing Then
                Return InsertToolFailure("AddSlideFromNearestSlide", "Параметр spec некорректен")
            End If

            Dim forcedSource = If(params?("sourceSlideIndex")?.Value(Of Integer)(), 0)
            If forcedSource > 0 Then
                If forcedSource > CInt(target.Slides.Count) Then
                    Return InsertToolFailure("AddSlideFromNearestSlide",
                                             $"sourceSlideIndex {forcedSource} вне диапазона 1..{target.Slides.Count}")
                End If
                Dim forcedAt = ResolveInsertAt(app, target, params)
                Dim forcedResult = Design.DeckTemplate.DeckTemplateRenderer.InsertFromSlide(app, target, forcedSource, spec, forcedAt)
                Return BuildInsertToolResult("AddSlideFromNearestSlide", forcedResult, "", forcedSource, 0, Nothing)
            End If

            Dim catalog = Agent.DeckTemplates.TemplateCatalog.ListAll()
            Dim deckCandidates As New List(Of Agent.DeckTemplates.SlideTemplateManifest)()
            Dim limit = Math.Min(40, CInt(target.Slides.Count))
            For index = 1 To limit
                Dim manifest = Design.DeckTemplate.DeckTemplateExtractor.DescribeSlideAsManifest(target, index)
                If manifest IsNot Nothing Then deckCandidates.Add(manifest)
            Next

            Dim deckBest = Agent.DeckTemplates.SlideTemplateSelector.SelectBest(spec, deckCandidates)
            Dim catalogBest = Agent.DeckTemplates.SlideTemplateSelector.SelectBest(spec, catalog)
            Dim match As Agent.DeckTemplates.SlideTemplateMatch
            If deckBest.Template IsNot Nothing AndAlso
               (catalogBest.Template Is Nothing OrElse deckBest.Score >= catalogBest.Score - 5) Then
                match = deckBest
            Else
                match = catalogBest
            End If

            If match Is Nothing OrElse match.Template Is Nothing Then
                Return InsertToolFailure("AddSlideFromNearestSlide", "Не найдено ни одного подходящего макета")
            End If

            Dim insertAt = ResolveInsertAt(app, target, params)
            Dim templateId = ""
            Dim sourceSlide = 0
            Dim result As Design.DeckTemplate.DeckTemplateRenderResult
            If Not String.IsNullOrWhiteSpace(match.Template.StorageFolder) Then
                templateId = match.Template.Id
                result = Design.DeckTemplate.DeckTemplateRenderer.InsertFromTemplate(app, target, match.Template, spec, insertAt)
            Else
                sourceSlide = match.Template.Source.SlideIndex
                If sourceSlide <= 0 Then sourceSlide = 1
                result = Design.DeckTemplate.DeckTemplateRenderer.InsertFromSlide(app, target, sourceSlide, spec, insertAt)
            End If

            Return BuildInsertToolResult("AddSlideFromNearestSlide", result, templateId, sourceSlide, match.Score, match.MatchedBy)
        Catch ex As Exception
            Return InsertToolFailure("AddSlideFromNearestSlide", "Вставка слайда: " & ex.Message)
        End Try
    End Function

    Private Shared Function ParseSlideSpecToken(token As JToken) As Agent.DeckTemplates.SlideSpec
        If token Is Nothing OrElse token.Type <> JTokenType.Object Then Return New Agent.DeckTemplates.SlideSpec()
        Try
            Dim spec = token.ToObject(Of Agent.DeckTemplates.SlideSpec)()
            If spec Is Nothing Then Return New Agent.DeckTemplates.SlideSpec()
            spec.Normalize()
            Return spec
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function ResolveInsertAt(app As Microsoft.Office.Interop.PowerPoint.Application,
                                            target As Microsoft.Office.Interop.PowerPoint.Presentation,
                                            params As JObject) As Integer
        Dim position = If(params?("position")?.ToString(), "after").Trim().ToLowerInvariant()
        If position = "end" Then Return CInt(target.Slides.Count) + 1

        Dim afterIndex = If(params?("afterSlideIndex")?.Value(Of Integer)(), 0)
        If afterIndex <= 0 Then
            Try
                afterIndex = CInt(app.ActiveWindow.View.Slide.SlideIndex)
            Catch
                afterIndex = 1
            End Try
        End If
        Return Math.Min(CInt(target.Slides.Count) + 1, afterIndex + 1)
    End Function

    Private Shared Function BuildInsertToolResult(toolId As String,
                                                  result As Design.DeckTemplate.DeckTemplateRenderResult,
                                                  templateId As String,
                                                  sourceSlide As Integer,
                                                  matchScore As Integer,
                                                  matchedBy As List(Of String)) As Agent.ToolResult
        If result Is Nothing OrElse Not result.Success Then
            Dim message = If(result Is Nothing, "Не удалось вставить слайд", result.ErrorMessage)
            If String.IsNullOrWhiteSpace(message) Then message = "Не удалось вставить слайд"
            Return InsertToolFailure(toolId, message)
        End If

        Dim warnings = New JArray(result.Warnings)
        Dim summary = $"Вставлен слайд {result.InsertedSlideIndex}, заполнено слотов: {result.SlotsFilled}"
        Dim observation = New JObject From {
            {"kind", "write"},
            {"summary", summary},
            {"changed", True},
            {"targetRefs", New JArray($"PowerPoint:Slide/{result.InsertedSlideIndex}")},
            {"warnings", warnings}
        }
        Return Agent.ToolResult.Succeed(toolId,
                                        summary,
                                        data:=New With {
                                            .insertedSlideIndex = result.InsertedSlideIndex,
                                            .insertedSlideId = result.InsertedSlideId,
                                            .slotsFilled = result.SlotsFilled,
                                            .templateId = If(templateId, ""),
                                            .sourceSlide = sourceSlide,
                                            .matchScore = matchScore,
                                            .matchedBy = If(matchedBy, New List(Of String)())
                                        },
                                        observation:=observation)
    End Function

    Private Shared Function InsertToolFailure(toolId As String, message As String) As Agent.ToolResult
        Return Agent.ToolResult.Failed(toolId,
                                       message,
                                       errorCode:=ExceptionClassifier.CodeUnknown,
                                       userMessage:=message,
                                       debugDetail:=message,
                                       recoverable:=True)
    End Function

    ''' <summary>
    ''' Сборка колоды по outline из сохранённого набора макетов или слайдов текущей колоды.
    ''' preview=true возвращает только план, без записи.
    ''' </summary>
    Private Shared Function BuildDeckFromTemplateSetAsToolResult(params As JObject) As Agent.ToolResult
        Return BuildDeckFromTemplateSetCore(Globals.ThisAddIn.Application, params)
    End Function

    ''' <summary>Ядро сборки колоды с явным app-контекстом (тестируемость и повторное использование).</summary>
    Private Shared Function BuildDeckFromTemplateSetCore(app As Microsoft.Office.Interop.PowerPoint.Application,
                                                        params As JObject) As Agent.ToolResult
        Try
            Dim outlineToken = params?("outline")
            If outlineToken Is Nothing OrElse outlineToken.Type <> JTokenType.Object Then
                Return InsertToolFailure("BuildDeckFromTemplateSet", "Не передан outline")
            End If
            Dim outline As Agent.DeckTemplates.DeckOutline
            Try
                outline = outlineToken.ToObject(Of Agent.DeckTemplates.DeckOutline)()
            Catch ex As Exception
                Return InsertToolFailure("BuildDeckFromTemplateSet", "outline некорректен: " & ex.Message)
            End Try
            If outline Is Nothing Then Return InsertToolFailure("BuildDeckFromTemplateSet", "outline пуст")
            outline.Normalize()
            If outline.Slides Is Nothing OrElse outline.Slides.Count = 0 Then
                Return InsertToolFailure("BuildDeckFromTemplateSet", "outline.slides пуст")
            End If

            Dim sourcePres = app.ActivePresentation
            Dim candidates As New List(Of Agent.DeckTemplates.SlideTemplateManifest)()
            candidates.AddRange(Agent.DeckTemplates.TemplateCatalog.ListAll())
            If sourcePres IsNot Nothing Then
                Dim limit = Math.Min(40, CInt(sourcePres.Slides.Count))
                For index = 1 To limit
                    Dim manifest = Design.DeckTemplate.DeckTemplateExtractor.DescribeSlideAsManifest(sourcePres, index)
                    If manifest IsNot Nothing Then candidates.Add(manifest)
                Next
            End If
            If candidates.Count = 0 Then
                Return InsertToolFailure("BuildDeckFromTemplateSet", "Нет макетов: сохраните слайд или колоду как макет")
            End If

            Dim maxSlides = If(params?("maxSlides")?.Value(Of Integer)(), 20)
            Dim plan = Agent.DeckTemplates.DeckBuildPlanner.Plan(outline, candidates, maxSlides)
            If plan.Count = 0 Then
                Return InsertToolFailure("BuildDeckFromTemplateSet", "Не удалось построить план колоды")
            End If

            Dim missing As New List(Of String)()
            For Each planStep In plan
                If planStep.Match Is Nothing OrElse planStep.Match.Template Is Nothing Then
                    missing.Add(planStep.Index.ToString())
                End If
            Next
            If missing.Count > 0 Then
                Return InsertToolFailure("BuildDeckFromTemplateSet",
                                         "Нет макетов для слайдов: " & String.Join(", ", missing))
            End If

            Dim planArray = BuildDeckPlanArray(plan)
            Dim previewOnly = If(params?("preview")?.Value(Of Boolean)(), False)
            If previewOnly Then
                Dim previewSummary = $"План колоды: {plan.Count} слайдов (без записи)"
                Dim previewObservation = New JObject From {
                    {"kind", "read"},
                    {"summary", previewSummary},
                    {"changed", False},
                    {"targetRefs", New JArray()},
                    {"warnings", New JArray()}
                }
                Return Agent.ToolResult.Succeed("BuildDeckFromTemplateSet",
                                                previewSummary,
                                                data:=New With {
                                                    .previewOnly = True,
                                                    .title = outline.Title,
                                                    .slideCount = plan.Count,
                                                    .plan = planArray
                                                },
                                                observation:=previewObservation)
            End If

            Dim targetMode = If(params?("target")?.ToString(), "new").Trim().ToLowerInvariant()
            Dim targetPres As Microsoft.Office.Interop.PowerPoint.Presentation
            Dim createdNew = False
            Dim initialSlideId As Integer = 0
            If targetMode = "new" Then
                targetPres = app.Presentations.Add()
                createdNew = True
                Try
                    initialSlideId = CInt(targetPres.Slides(1).SlideID)
                Catch
                End Try
                If sourcePres IsNot Nothing Then
                    Try
                        targetPres.PageSetup.SlideWidth = sourcePres.PageSetup.SlideWidth
                        targetPres.PageSetup.SlideHeight = sourcePres.PageSetup.SlideHeight
                    Catch
                    End Try
                End If
            Else
                targetPres = sourcePres
                If targetPres Is Nothing Then
                    Return InsertToolFailure("BuildDeckFromTemplateSet", "Нет активной презентации")
                End If
            End If

            Dim build = Design.DeckTemplate.DeckTemplateRenderer.BuildDeck(app, targetPres, plan)
            If build Is Nothing OrElse Not build.Success Then
                If createdNew Then
                    Try
                        targetPres.Close()
                    Catch
                    End Try
                End If
                Return InsertToolFailure("BuildDeckFromTemplateSet",
                                         If(build Is Nothing, "Сборка колоды не вернула результат", build.ErrorMessage))
            End If

            If createdNew AndAlso initialSlideId > 0 Then
                RemoveSlideById(targetPres, initialSlideId)
                Try
                    targetPres.Windows(1).Activate()
                Catch
                End Try
            End If

            Dim summary = If(createdNew,
                             $"Создана колода: {build.InsertedSlideIndexes.Count} слайдов",
                             $"Добавлено слайдов в текущую колоду: {build.InsertedSlideIndexes.Count}")
            Dim warnings = New JArray(build.Warnings)
            Dim observation = New JObject From {
                {"kind", "write"},
                {"summary", summary},
                {"changed", True},
                {"targetRefs", If(createdNew, New JArray(), New JArray(build.InsertedSlideIndexes.Select(Function(i) $"PowerPoint:Slide/{i}")))},
                {"warnings", warnings}
            }
            Return Agent.ToolResult.Succeed("BuildDeckFromTemplateSet",
                                            summary,
                                            data:=New With {
                                                .previewOnly = False,
                                                .target = If(createdNew, "new", "current"),
                                                .title = outline.Title,
                                                .insertedCount = build.InsertedSlideIndexes.Count,
                                                .slideIndexes = build.InsertedSlideIndexes,
                                                .plan = planArray,
                                                .warnings = build.Warnings
                                            },
                                            observation:=observation)
        Catch ex As Exception
            AppLogger.Warn("PPT.TemplateDeck", ex.ToString())
            Return InsertToolFailure("BuildDeckFromTemplateSet", "Сборка колоды: " & ex.Message)
        End Try
    End Function

    Private Shared Function BuildDeckPlanArray(plan As List(Of Agent.DeckTemplates.DeckPlanStep)) As JArray
        Dim array As New JArray()
        For Each planStep In plan
            Dim template = If(planStep.Match Is Nothing, Nothing, planStep.Match.Template)
            Dim templateId = ""
            Dim sourceSlide = 0
            If template IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(template.StorageFolder) Then
                    templateId = template.Id
                ElseIf template.Source IsNot Nothing Then
                    sourceSlide = template.Source.SlideIndex
                End If
            End If
            array.Add(New JObject From {
                {"index", planStep.Index},
                {"archetype", If(template Is Nothing, "", template.Archetype)},
                {"title", If(planStep.Spec Is Nothing, "", planStep.Spec.Title)},
                {"templateId", templateId},
                {"sourceSlide", sourceSlide},
                {"score", If(planStep.Match Is Nothing, 0, planStep.Match.Score)},
                {"matchedBy", If(planStep.Match Is Nothing, New JArray(), New JArray(planStep.Match.MatchedBy))}
            })
        Next
        Return array
    End Function

    Private Shared Sub RemoveSlideById(target As Microsoft.Office.Interop.PowerPoint.Presentation, slideId As Integer)
        If target Is Nothing OrElse slideId <= 0 Then Return
        For index = CInt(target.Slides.Count) To 1 Step -1
            Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
            Try
                slide = target.Slides(index)
                If CInt(slide.SlideID) = slideId Then
                    slide.Delete()
                    Return
                End If
            Catch
            Finally
                ShareRibbon.ComObjectHelper.ReleaseComObject(slide)
            End Try
        Next
    End Sub

    ' ===== Каталог макетов: обработчики WebView2 (COM остаётся на UI-потоке) =====

    Protected Overrides Sub HandleTemplatesSaveFromSlide(jsonDoc As JObject)
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim slideIndex = If(jsonDoc?("slideIndex")?.Value(Of Integer)(), 0)
            If slideIndex <= 0 Then
                Try
                    slideIndex = CInt(app.ActiveWindow.View.Slide.SlideIndex)
                Catch
                    slideIndex = 1
                End Try
            End If

            Dim warnings As New List(Of String)()
            Dim errorMessage As String = ""
            Dim manifest = Design.DeckTemplate.DeckTemplateExtractor.ExtractFromSlide(
                app, slideIndex,
                jsonDoc?("name")?.ToString(),
                jsonDoc?("setId")?.ToString(),
                jsonDoc?("archetype")?.ToString(),
                errorMessage, warnings)
            If manifest Is Nothing Then
                RespondTemplates("saveFromSlide", False, errorMessage)
                Return
            End If
            RespondTemplates("saveFromSlide", True, $"Макет сохранён: {manifest.Name}",
                             New With {
                                 .templateId = manifest.Id,
                                 .name = manifest.Name,
                                 .archetype = manifest.Archetype,
                                 .slotCount = manifest.Slots.Count,
                                 .warnings = warnings
                             })
        Catch ex As Exception
            RespondTemplates("saveFromSlide", False, "Сохранение макета: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub HandleTemplatesSaveDeck(jsonDoc As JObject)
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim warnings As New List(Of String)()
            Dim errorMessage As String = ""
            Dim manifests = Design.DeckTemplate.DeckTemplateExtractor.ExtractFromDeck(
                app,
                jsonDoc?("setName")?.ToString(),
                If(jsonDoc?("maxSlides")?.Value(Of Integer)(), 20),
                If(jsonDoc?("skipEmpty")?.Value(Of Boolean)(), True),
                errorMessage, warnings)
            If manifests Is Nothing OrElse manifests.Count = 0 Then
                RespondTemplates("saveDeck", False, errorMessage)
                Return
            End If
            RespondTemplates("saveDeck", True, $"Сохранён набор макетов: {manifests.Count}",
                             New With {
                                 .setId = manifests(0).SetId,
                                 .count = manifests.Count,
                                 .warnings = warnings
                             })
        Catch ex As Exception
            RespondTemplates("saveDeck", False, "Сохранение набора: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub HandleTemplatesInsert(jsonDoc As JObject)
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim target = app.ActivePresentation
            If target Is Nothing Then
                RespondTemplates("insert", False, "Нет активной презентации")
                Return
            End If

            Dim templateId = jsonDoc?("templateId")?.ToString()
            Dim manifest = Agent.DeckTemplates.TemplateCatalog.FindById(templateId)
            If manifest Is Nothing Then
                RespondTemplates("insert", False, $"Макет не найден: {templateId}")
                Return
            End If

            Dim specToken = jsonDoc?("spec")
            Dim spec As Agent.DeckTemplates.SlideSpec
            If specToken Is Nothing OrElse specToken.Type <> JTokenType.Object Then
                spec = New Agent.DeckTemplates.SlideSpec()
            Else
                spec = specToken.ToObject(Of Agent.DeckTemplates.SlideSpec)()
                If spec Is Nothing Then spec = New Agent.DeckTemplates.SlideSpec()
            End If
            spec.Normalize()

            Dim position = If(jsonDoc?("position")?.ToString(), "after").Trim().ToLowerInvariant()
            Dim insertAt As Integer
            If position = "end" Then
                insertAt = CInt(target.Slides.Count) + 1
            Else
                Dim currentIndex = 1
                Try
                    currentIndex = CInt(app.ActiveWindow.View.Slide.SlideIndex)
                Catch
                End Try
                insertAt = Math.Min(CInt(target.Slides.Count) + 1, currentIndex + 1)
            End If

            Dim result = Design.DeckTemplate.DeckTemplateRenderer.InsertFromTemplate(app, target, manifest, spec, insertAt)
            If result Is Nothing OrElse Not result.Success Then
                RespondTemplates("insert", False,
                                 If(result Is Nothing, "Не удалось вставить слайд", result.ErrorMessage))
                Return
            End If
            RespondTemplates("insert", True, $"Вставлен слайд {result.InsertedSlideIndex}",
                             New With {
                                 .insertedSlideIndex = result.InsertedSlideIndex,
                                 .slotsFilled = result.SlotsFilled,
                                 .warnings = result.Warnings
                             })
        Catch ex As Exception
            RespondTemplates("insert", False, "Вставка макета: " & ex.Message)
        End Try
    End Sub

    Protected Overrides Sub HandleTemplatesPreview(jsonDoc As JObject)
        Try
            Dim templateId = jsonDoc?("templateId")?.ToString()
            Dim manifest = Agent.DeckTemplates.TemplateCatalog.FindById(templateId)
            If manifest Is Nothing Then
                RespondTemplates("preview", False, "Макет не найден")
                Return
            End If
            Dim roles = If(manifest.Slots Is Nothing,
                           New List(Of String)(),
                           manifest.Slots.
                               Where(Function(slot) slot IsNot Nothing).
                               Select(Function(slot) slot.Role).
                               Where(Function(role) Not String.IsNullOrWhiteSpace(role)).
                               Distinct(StringComparer.OrdinalIgnoreCase).
                               ToList())
            Dim hasPreview = False
            Try
                hasPreview = File.Exists(manifest.PreviewFullPath)
            Catch
            End Try
            RespondTemplates("preview", True, $"Макет: {manifest.Name}",
                             New With {
                                 .id = manifest.Id,
                                 .name = manifest.Name,
                                 .archetype = manifest.Archetype,
                                 .slots = roles,
                                 .hasPreview = hasPreview
                             })
        Catch ex As Exception
            RespondTemplates("preview", False, "Предпросмотр макета: " & ex.Message)
        End Try
    End Sub

    Private Shared Function ParsePowerPointCommandEnvelope(jsonCode As String) As JObject
        If String.IsNullOrWhiteSpace(jsonCode) Then Return Nothing

        ' Сначала как есть: \" внутри JSON-строки — корректное экранирование,
        ' и его нельзя снимать вслепую (иначе ломается code с кавычками).
        Try
            Return TryCast(JToken.Parse(jsonCode), JObject)
        Catch ex As Newtonsoft.Json.JsonException
            Dim escapedQuote = ChrW(92) & ChrW(34)
            If Not jsonCode.Contains(escapedQuote) Then Throw

            ' Лишний слой экранирования: пробуем снять один раз как fallback.
            Dim normalized = jsonCode.Replace(escapedQuote, ChrW(34))
            If normalized = jsonCode Then Throw

            Try
                Return TryCast(JToken.Parse(normalized), JObject)
            Catch
                Throw ex
            End Try
        End Try
    End Function

    Private Shared Function GetPowerPointEnvelopeToolId(envelope As JObject) As String
        If envelope Is Nothing Then Return "PowerPointCommands"
        Dim commands = TryCast(envelope("commands"), JArray)
        If commands IsNot Nothing Then
            If commands.Count = 1 AndAlso commands(0).Type = JTokenType.Object Then
                Dim singleCommand = commands(0)("command")?.ToString()
                If Not String.IsNullOrWhiteSpace(singleCommand) Then Return singleCommand.Trim()
            End If
            Return "PowerPointCommands"
        End If
        Dim command = envelope("command")?.ToString()
        Return If(String.IsNullOrWhiteSpace(command), "PowerPointCommands", command.Trim())
    End Function

    Private Shared Function GetPowerPointEnvelopeParams(envelope As JObject) As JObject
        If envelope Is Nothing Then Return Nothing
        Dim commands = TryCast(envelope("commands"), JArray)
        If commands IsNot Nothing AndAlso commands.Count = 1 AndAlso commands(0).Type = JTokenType.Object Then
            Return TryCast(commands(0)("params"), JObject)
        End If
        Return TryCast(envelope("params"), JObject)
    End Function

    Private Function CapturePowerPointCommandSnapshot(envelope As JObject) As JObject
        Dim snapshot As New JObject()
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim presentation = app.ActivePresentation
            If presentation Is Nothing Then Return snapshot

            snapshot("presentation") = If(presentation.Name, "")
            snapshot("slideCount") = CInt(presentation.Slides.Count)

            Dim slideIndex = ResolvePowerPointTargetSlideIndex(envelope, presentation.Slides.Count)
            If slideIndex <= 0 Then
                Try
                    slideIndex = CInt(app.ActiveWindow.View.Slide.SlideIndex)
                Catch
                    If presentation.Slides.Count > 0 Then slideIndex = 1
                End Try
            End If
            snapshot("slideIndex") = slideIndex

            If slideIndex > 0 AndAlso slideIndex <= presentation.Slides.Count Then
                Dim slide = presentation.Slides(slideIndex)
                snapshot("shapeCount") = CInt(slide.Shapes.Count)
                snapshot("titleHash") = ComputePowerPointObservationHash(GetPowerPointSlideTitle(slide))
                snapshot("textHash") = ComputePowerPointObservationHash(GetPowerPointSlideText(slide))
                snapshot("textPreview") = TruncatePowerPointObservationText(GetPowerPointSlideText(slide), 240)
                snapshot("notesHash") = ComputePowerPointObservationHash(GetPowerPointNotesText(slide))
                snapshot("transitionHash") = ComputePowerPointObservationHash(GetPowerPointSlideTransition(slide))
            End If
            ' Переходы, тема и заметки часто применяются ко всей презентации,
            ' поэтому нужен общий отпечаток, иначе успешная команда выглядит как «без изменений».
            snapshot("transitionHashAll") = ComputePowerPointObservationHash(GetPowerPointPresentationTransitionSignature(presentation))
        Catch ex As Exception
            snapshot("captureError") = AppLogger.Redact(ex.Message)
        End Try
        Return snapshot
    End Function

    Private Function ResolvePowerPointTargetSlideIndex(envelope As JObject, slideCount As Integer) As Integer
        ' После DuplicateSlide снимок «после» должен смотреть на копию, а не на источник.
        If _lastPptCommandTargetIndex > 0 Then
            Return Math.Max(1, Math.Min(slideCount, _lastPptCommandTargetIndex))
        End If
        If envelope Is Nothing Then Return 0
        Dim command = envelope
        If envelope("commands") IsNot Nothing AndAlso envelope("commands").Type = JTokenType.Array Then
            command = TryCast(envelope("commands").FirstOrDefault(), JObject)
        End If
        Dim rawIndex = command?("params")?("slideIndex")?.Value(Of Integer)()
        If rawIndex.HasValue Then
            Dim commandName = If(command?("command")?.ToString(), "")
            Dim oneBased = String.Equals(commandName, "DuplicateSlide", StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(commandName, "ReplaceSlideText", StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(commandName, "UpdateChartData", StringComparison.OrdinalIgnoreCase)

            If oneBased Then
                If rawIndex.Value <= 0 Then Return 0
                Return Math.Max(1, Math.Min(slideCount, rawIndex.Value))
            End If

            If rawIndex.Value < 0 Then Return 0
            ' Остальные PPT-команды исторически 0-based; Office-снимок 1-based.
            Return Math.Max(1, Math.Min(slideCount, rawIndex.Value + 1))
        End If
        Return 0
    End Function

    Private Shared Function BuildPowerPointCommandObservation(toolId As String,
                                                              success As Boolean,
                                                              beforeSnapshot As JObject,
                                                              afterSnapshot As JObject) As JObject
        Dim changed = Not JToken.DeepEquals(beforeSnapshot, afterSnapshot)
        Dim slideIndex = If(afterSnapshot?("slideIndex")?.Value(Of Integer)(), beforeSnapshot?("slideIndex")?.Value(Of Integer)())
        Dim targetRefs As New JArray()
        If slideIndex > 0 Then targetRefs.Add($"PowerPoint:Slide/{slideIndex}") Else targetRefs.Add("PowerPoint:Presentation")

        Dim warnings As New JArray()
        If success AndAlso Not changed Then warnings.Add("Команда обработана, но снимок слайдов не обнаружил изменений; возможны отмена пользователем, эквивалентный формат или noop")
        If afterSnapshot?("captureError") IsNot Nothing Then warnings.Add(afterSnapshot("captureError"))

        Dim diff As New JObject From {
            {"slideCountDelta", GetPowerPointSnapshotInteger(afterSnapshot, "slideCount") - GetPowerPointSnapshotInteger(beforeSnapshot, "slideCount")},
            {"shapeCountDelta", GetPowerPointSnapshotInteger(afterSnapshot, "shapeCount") - GetPowerPointSnapshotInteger(beforeSnapshot, "shapeCount")}
        }

        Return New JObject From {
            {"kind", "write"},
            {"summary", If(success, $"Инструмент PowerPoint {toolId} выполнен", $"Сбой выполнения инструмента PowerPoint {toolId}")},
            {"targetRefs", targetRefs},
            {"changed", changed},
            {"before", beforeSnapshot},
            {"after", afterSnapshot},
            {"diff", diff},
            {"warnings", warnings}
        }
    End Function

    Private Shared Function GetPowerPointSnapshotInteger(snapshot As JObject, name As String) As Integer
        If snapshot Is Nothing Then Return 0
        Return If(snapshot(name)?.Value(Of Integer)(), 0)
    End Function

    Private Shared Function ComputePowerPointObservationHash(value As String) As String
        Dim bytes = Encoding.UTF8.GetBytes(If(value, ""))
        Using sha = System.Security.Cryptography.SHA256.Create()
            Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Shared Function GetPowerPointSlideTitle(slide As Object) As String
        Try
            If slide.Shapes.Title IsNot Nothing AndAlso slide.Shapes.Title.HasTextFrame Then
                Return If(slide.Shapes.Title.TextFrame.TextRange.Text, "")
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function GetPowerPointSlideText(slide As Object) As String
        Dim parts As New List(Of String)()
        Try
            For Each shape As Object In slide.Shapes
                Try
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                        parts.Add(If(shape.TextFrame.TextRange.Text, ""))
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        Return String.Join(vbLf, parts)
    End Function

    Private Shared Function GetPowerPointNotesText(slide As Object) As String
        Dim parts As New List(Of String)()
        Try
            For Each shape As Object In slide.NotesPage.Shapes
                Try
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then parts.Add(If(shape.TextFrame.TextRange.Text, ""))
                Catch
                End Try
            Next
        Catch
        End Try
        Return String.Join(vbLf, parts)
    End Function

    ''' <summary>
    ''' Снимок перехода слайда: без него применение перехода не видно в наблюдении,
    ''' и успешная команда ложно помечается как «изменений нет».
    ''' </summary>
    Private Shared Function GetPowerPointSlideTransition(slide As Object) As String
        Dim parts As New List(Of String)()
        Try
            Dim transition = slide.SlideShowTransition
            If transition IsNot Nothing Then
                Try : parts.Add("entry=" & transition.EntryEffect.ToString()) : Catch : End Try
                Try : parts.Add("speed=" & transition.Speed.ToString()) : Catch : End Try
                Try : parts.Add("advance=" & transition.AdvanceOnClick.ToString()) : Catch : End Try
                Try : parts.Add("duration=" & transition.Duration.ToString()) : Catch : End Try
            End If
        Catch
        End Try
        Return String.Join(";", parts)
    End Function

    ''' <summary>
    ''' Отпечаток переходов по всей презентации (ограничен первыми 50 слайдами).
    ''' </summary>
    Private Shared Function GetPowerPointPresentationTransitionSignature(presentation As Object) As String
        Dim parts As New List(Of String)()
        Try
            Dim count As Integer = CInt(presentation.Slides.Count)
            Dim limit = Math.Min(count, 50)
            For index = 1 To limit
                Dim slide As Object = Nothing
                Try
                    slide = presentation.Slides(index)
                    parts.Add(index.ToString() & ":" & GetPowerPointSlideTransition(slide))
                Finally
                    ShareRibbon.ComObjectHelper.ReleaseComObject(slide)
                End Try
            Next
        Catch
        End Try
        Return String.Join("|", parts)
    End Function

    ''' <summary>
    ''' Показывает ли диаграмма доли в процентах (PowerPoint хранит их как 0.25 для 25%).
    ''' </summary>
    Private Shared Function HasPercentageDataLabels(series As Object) As Boolean
        Try
            If Not CBool(series.HasDataLabels) Then Return False
            Dim points = series.Points
            Dim count As Integer = Math.Min(CInt(points.Count), 3)
            For index = 1 To count
                Try
                    If CInt(points(index).DataLabel.ShowPercentage) = -1 Then Return True
                Catch
                End Try
            Next
        Catch
        End Try
        Return False
    End Function

    Private Shared Function TruncatePowerPointObservationText(value As String, maxLength As Integer) As String
        Dim text = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If text.Length <= maxLength Then Return text
        Return text.Substring(0, maxLength)
    End Function

    ''' <summary>
    ''' 执行PPT命令数组
    ''' </summary>
    Private Function ExecutePPTCommandsArray(commandsArray As JToken, originalJson As String, preview As Boolean) As Boolean
        Try
            Dim commands = CType(commandsArray, JArray)
            If commands.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("Массив команд пуст")
                Return False
            End If

            ' 预览所有命令 - 使用增强的预览表单
            If preview Then
                If Not ShareRibbon.CommandPreviewForm.ShowPreview($"Предпросмотр команд PPT — всего {commands.Count}", commandsArray) Then
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
                    If ExecutePPTCommand(cmdObj, preview) Then
                        successCount += 1
                    Else
                        failCount += 1
                        If Not String.IsNullOrWhiteSpace(_lastPptCommandError) Then
                            Dim detail = _lastPptCommandError
                            AppLogger.Warn("PPT.Commands", detail)
                            ShareRibbon.GlobalStatusStrip.ShowWarning(detail)
                        End If
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
            Debug.WriteLine($"ExecutePPTCommandsArray 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"Ошибка пакетного выполнения: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行单个PPT命令
    ''' </summary>
    Private Function ExecutePPTSingleCommand(commandJson As JObject, processedJson As String, preview As Boolean) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            
            ' 预览 - 使用增强的预览表单
            If preview Then
                If Not ShareRibbon.CommandPreviewForm.ShowPreview("Предпросмотр команд PPT", commandJson) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行命令
            Dim success = ExecutePPTCommand(commandJson, preview)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"Команда '{command}' выполнена успешно")
            Else
                Dim detail = _lastPptCommandError
                If String.IsNullOrWhiteSpace(detail) Then
                    ShareRibbon.GlobalStatusStrip.ShowWarning($"Команда '{command}' выполнена с ошибкой")
                Else
                    AppLogger.Warn("PPT.Commands", detail)
                    ShareRibbon.GlobalStatusStrip.ShowWarning(detail)
                End If
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecutePPTSingleCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行具体的PPT命令
    ''' </summary>
    Private Function ExecutePPTCommand(commandJson As JObject, Optional preview As Boolean = False) As Boolean
        _lastPptCommandError = ""
        _lastPptCommandWarning = ""
        _lastPptCommandInfo = ""
        _lastPptCommandTargetIndex = 0
        Try
            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")
            
            Dim pres = Globals.ThisAddIn.Application.ActivePresentation

            Select Case command.ToLower()
                Case "insertslide"
                    Return ExecuteInsertSlide(params, pres)
                Case "inserttext"
                    Return ExecuteInsertText(params, pres)
                Case "insertshape"
                    Return ExecuteInsertShape(params, pres)
                Case "formatslide"
                    Return ExecuteFormatSlide(params, pres)
                Case "inserttable"
                    Return ExecuteInsertTable(params, pres)
                Case "createslides"
                    If preview AndAlso params IsNot Nothing Then params("preview") = True
                    Return ExecuteCreateSlides(params, pres)
                Case "addanimation"
                    Return ExecuteAddAnimation(params, pres)
                Case "applytransition"
                    Return ExecuteApplyTransition(params, pres)
                Case "beautifyslides"
                    Return ExecuteBeautifySlides(params, pres)
                Case "deleteslide"
                    Return ExecuteDeleteSlide(params, pres)
                Case "duplicateslide"
                    Return ExecuteDuplicateSlide(params, pres)
                Case "moveslide"
                    Return ExecuteMoveSlide(params, pres)
                Case "setslidelayout"
                    Return ExecuteSetSlideLayout(params, pres)
                Case "applytheme"
                    Return ExecuteApplyTheme(params, pres)
                Case "addspeakernotes"
                    Return ExecuteAddSpeakerNotes(params, pres)
                Case "executevba"
                    Dim vbaCode = params("code")?.ToString()
                    If String.IsNullOrEmpty(vbaCode) Then
                        GlobalStatusStrip.ShowWarning("ExecuteVBA: отсутствует параметр code")
                        Return False
                    End If
                    If Not IsVbaAllowed() Then
                        GlobalStatusStrip.ShowWarning("VBA недоступен: включите «Доверие доступом к объектной модели проектов VBA» или используйте DuplicateSlide и ReplaceSlideText.")
                        Return False
                    End If
                    Return CodeExecutionService.ExecuteVBACode(vbaCode, False)
                Case "replaceslidetext"
                    Return ExecuteReplaceText(params, pres)
                Case "updatechartdata"
                    Return ExecuteUpdateChartData(params, pres)
                Case Else
                    Debug.WriteLine($"不支持的PPT命令: {command}")
                    GlobalStatusStrip.ShowWarning($"Неподдерживаемая команда PPT: {command}")
                    Return False
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ExecutePPTCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteReplaceText(params As JToken, pres As Object) As Boolean
        _lastPptCommandError = ""
        Try
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            Dim matchCase = If(params("matchCase")?.Value(Of Boolean)(), False)
            Dim replacements = TryCast(params("replacements"), JArray)
            If replacements Is Nothing OrElse replacements.Count = 0 Then
                _lastPptCommandError = "ReplaceSlideText: не передан массив replacements"
                Return False
            End If

            Dim slide As Object
            If slideIndex <= 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                If slideIndex > pres.Slides.Count Then Return False
                slide = pres.Slides(slideIndex)
            End If

            Dim compare As CompareMethod = If(matchCase, CompareMethod.Binary, CompareMethod.Text)
            Dim stats As New ReplaceTextStats()
            Dim replacedCount = ReplaceTextInShapes(slide.Shapes, replacements, compare, stats)

            If replacedCount = 0 Then
                Dim sample = String.Join(" || ", stats.Texts)
                If sample.Length > 200 Then sample = sample.Substring(0, 200) & "..."
                ' Идемпотентность: повторный перевод уже переведённого слайда — не сбой,
                ' иначе agent-repair зацикливается на шаге, который уже выполнен.
                _lastPptCommandWarning = $"ReplaceSlideText: совпадений не найдено (возможно, уже заменено; slide={slideIndex}, shapes={stats.TotalShapes}, textShapes={stats.TextShapes}). Тексты: '{sample}'"
                AppLogger.Info("PPT.ReplaceSlideText", _lastPptCommandWarning)
                GlobalStatusStrip.ShowInfo("ReplaceSlideText: совпадений не найдено — возможно, текст уже заменён.")
                Return True
            End If

            GlobalStatusStrip.ShowInfo($"Заменено текстовых блоков: {replacedCount}")
            Return True
        Catch ex As Exception
            _lastPptCommandError = "ReplaceSlideText: " & ex.Message
            Debug.WriteLine($"ExecuteReplaceText ошибка: {ex.Message}")
            Return False
        End Try
    End Function

    Private Class ReplaceTextStats
        Public Property TotalShapes As Integer
        Public Property TextShapes As Integer
        Public Property Groups As Integer
        Public Property WriteErrors As Integer
        Public Property Sample As String = ""
        Public Property Texts As New List(Of String)()
    End Class

    ''' <summary>
    ''' Рекурсивно обходит фигуры, включая группы, и использует TextFrame2
    ''' (надёжнее TextFrame для шаблонных и сгруппированных текстов).
    ''' </summary>
    Private Function ReplaceTextInShapes(shapes As Object,
                                         replacements As JArray,
                                         compare As CompareMethod,
                                         stats As ReplaceTextStats) As Integer
        Dim replacedCount As Integer = 0
        If shapes Is Nothing Then Return replacedCount

        For Each shape As Object In shapes
            stats.TotalShapes += 1
            Try
                If shape.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                    stats.Groups += 1
                    replacedCount += ReplaceTextInShapes(shape.GroupItems, replacements, compare, stats)
                    Continue For
                End If

                If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim table As Object = shape.Table
                    Try
                        For rowIndex = 1 To table.Rows.Count
                            For colIndex = 1 To table.Columns.Count
                                Dim cell As Object = Nothing
                                Dim cellShape As Object = Nothing
                                Try
                                    cell = table.Cell(rowIndex, colIndex)
                                    cellShape = cell.Shape
                                    replacedCount += ReplaceTextInSingleShape(cellShape, replacements, compare, stats)
                                Catch
                                Finally
                                    If cellShape IsNot Nothing Then ComObjectHelper.ReleaseComObject(cellShape)
                                    If cell IsNot Nothing Then ComObjectHelper.ReleaseComObject(cell)
                                End Try
                            Next
                        Next
                    Finally
                        ComObjectHelper.ReleaseComObject(table)
                    End Try
                    Continue For
                End If

                replacedCount += ReplaceTextInSingleShape(shape, replacements, compare, stats)
            Catch shapeEx As Exception
                Debug.WriteLine($"ReplaceSlideText shape skipped: {shapeEx.Message}")
            End Try
        Next

        Return replacedCount
    End Function

    ''' <summary>
    ''' Заменяет текст в одной фигуре или ячейке таблицы.
    ''' </summary>
    Private Function ReplaceTextInSingleShape(shape As Object,
                                              replacements As JArray,
                                              compare As CompareMethod,
                                              stats As ReplaceTextStats) As Integer
        If shape Is Nothing Then Return 0

        Dim original As String = ""
        Dim hasText As Boolean = False

        Try
            Dim textFrame2 = shape.TextFrame2
            If textFrame2 IsNot Nothing Then
                original = If(textFrame2.TextRange.Text, "")
                hasText = True
            End If
        Catch
        End Try

        If Not hasText Then
            Try
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                   shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    original = If(shape.TextFrame.TextRange.Text, "")
                    hasText = True
                End If
            Catch
            End Try
        End If

        If Not hasText OrElse String.IsNullOrEmpty(original) Then Return 0

        stats.TextShapes += 1
        If String.IsNullOrEmpty(stats.Sample) Then stats.Sample = original
        If stats.Texts.Count < 3 Then stats.Texts.Add(EscapeWhitespace(original))

        Dim updated = ApplyReplacements(original, replacements, compare)
        If String.Equals(original, updated, StringComparison.Ordinal) Then Return 0

        Return If(ApplyShapeText(shape, updated, stats), 1, 0)
    End Function

    Private Shared Function ApplyReplacements(text As String, replacements As JArray, compare As CompareMethod) As String
        Dim updated As String = text
        For Each item As JObject In replacements.OfType(Of JObject)()
            Dim findText = item("find")?.ToString()
            Dim replaceWith = If(item("replaceWith")?.ToString(), "")
            If String.IsNullOrEmpty(findText) Then Continue For
            updated = ReplaceTextFlexible(updated, findText, replaceWith, compare)
        Next
        Return updated
    End Function

    Private Shared Function ApplyShapeText(shape As Object, updated As String, stats As ReplaceTextStats) As Boolean
        Try
            shape.TextFrame2.TextRange.Text = updated
            Return True
        Catch
        End Try

        Try
            shape.TextFrame.TextRange.Text = updated
            Return True
        Catch ex As Exception
            stats.WriteErrors += 1
            If stats.WriteErrors <= 3 Then
                AppLogger.Warn("PPT.ReplaceSlideText", "Не удалось записать текст фигуры: " & AppLogger.Redact(ex.Message))
            End If
        End Try
        Return False
    End Function

    Private Shared Function EscapeWhitespace(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        Return value.Replace(vbCr, "\r").Replace(vbLf, "\n").Replace(ChrW(11), "\v").Replace(ChrW(160), "[NBSP]").Replace(vbTab, "\t")
    End Function

    ''' <summary>
    ''' Сначала точная замена, затем замена с игнорированием переносов строк и
    ''' повторных пробелов (PowerPoint хранит абзацы с \r, поэтому точный find
    ''' из модели часто не совпадает).
    ''' </summary>
    Private Shared Function ReplaceTextFlexible(text As String,
                                                findText As String,
                                                replaceWith As String,
                                                compare As CompareMethod) As String
        If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(findText) Then Return text

        Dim direct = Strings.Replace(text, findText, replaceWith, 1, -1, compare)
        If Not String.Equals(direct, text, StringComparison.Ordinal) Then Return direct

        Dim words = findText.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                                   StringSplitOptions.RemoveEmptyEntries)
        If words.Length < 2 Then Return text

        Try
            Dim pattern = String.Join("\s+", words.Select(Function(word) System.Text.RegularExpressions.Regex.Escape(word)))
            Dim options = System.Text.RegularExpressions.RegexOptions.Singleline Or
                          System.Text.RegularExpressions.RegexOptions.CultureInvariant
            If compare = CompareMethod.Text Then options = options Or System.Text.RegularExpressions.RegexOptions.IgnoreCase
            Dim safeReplacement = replaceWith.Replace("$", "$$")
            Return System.Text.RegularExpressions.Regex.Replace(text, pattern, safeReplacement, options)
        Catch
            Return text
        End Try
    End Function

    Private Function ExecuteUpdateChartData(params As JToken, pres As Object) As Boolean
        _lastPptCommandError = ""
        _lastPptCommandWarning = ""
        Try
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), 0)
            Dim chartIndex = If(params("chartIndex")?.Value(Of Integer)(), 1)
            Dim categories = TryCast(params("categories"), JArray)
            Dim values = TryCast(params("values"), JArray)

            If values Is Nothing OrElse values.Count = 0 Then
                _lastPptCommandError = "UpdateChartData: values обязателен и не может быть пустым"
                Return False
            End If
            If categories IsNot Nothing AndAlso categories.Count > 0 AndAlso categories.Count <> values.Count Then
                _lastPptCommandError = $"UpdateChartData: передано {categories.Count} категорий и {values.Count} значений — длины должны совпадать"
                Return False
            End If

            Dim slide As Object
            If slideIndex <= 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                If slideIndex > pres.Slides.Count Then Return False
                slide = pres.Slides(slideIndex)
            End If

            Dim chartShape As Object = Nothing
            Dim foundCharts As Integer = 0
            For Each shape As Object In slide.Shapes
                Try
                    If shape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        foundCharts += 1
                        If foundCharts = chartIndex Then
                            chartShape = shape
                            Exit For
                        End If
                    End If
                Catch
                End Try
            Next

            If chartShape Is Nothing Then
                _lastPptCommandError = $"UpdateChartData: на слайде нет диаграммы #{chartIndex}"
                Return False
            End If

            Dim chart As Object = chartShape.Chart
            Dim series As Object = chart.SeriesCollection
            If series.Count = 0 Then
                _lastPptCommandError = "UpdateChartData: у диаграммы нет рядов"
                Return False
            End If

            Dim numericValues = values.OfType(Of JValue)().
                Select(Function(item) item.Value(Of Double)()).ToList()

            Dim keepsCategories = categories Is Nothing OrElse categories.Count = 0
            Dim existingCount As Integer = 0
            Try
                Dim existingCategories = TryCast(series(1).XValues, Array)
                If existingCategories IsNot Nothing Then existingCount = existingCategories.Length
            Catch
            End Try

            ' Число сегментов задаёт диаграмма. Если оно не совпадает, НЕ применяем
            ' часть значений: иначе можно молча испортить чужой слайд/диаграмму.
            If keepsCategories AndAlso existingCount > 0 AndAlso existingCount <> numericValues.Count Then
                _lastPptCommandError = $"UpdateChartData: передано {numericValues.Count} значений, а у диаграммы на слайде {If(slideIndex <= 0, "текущем", slideIndex.ToString())} {existingCount} категорий. Передай ровно {existingCount} значений или явные categories."
                Return False
            End If

            ' Для процентных диаграмм PowerPoint хранит доли (25% = 0.25).
            ' Если модель передала 30 вместо 0.30, диаграмма показала бы 3000%.
            Dim isPercentChart = HasPercentageDataLabels(series)
            If isPercentChart AndAlso numericValues.Any(Function(value) Math.Abs(value) > 1) Then
                numericValues = numericValues.Select(Function(value) value / 100.0).ToList()
            End If

            If Not keepsCategories AndAlso categories.Count <> numericValues.Count Then
                _lastPptCommandError = $"UpdateChartData: передано {categories.Count} категорий и {numericValues.Count} значений — длины должны совпадать"
                Return False
            End If

            If keepsCategories Then
                ' Подписи не меняются: обновляем только значения, чтобы модель не требовала
                ' у пользователя категории, которые уже есть на слайде.
                series(1).Values = numericValues.ToArray()
                GlobalStatusStrip.ShowInfo($"Данные диаграммы обновлены: {numericValues.Count} значений, подписи сохранены")
            Else
                Dim categoryArray = categories.OfType(Of JValue)().
                    Select(Function(item) CObj(If(item.Value?.ToString(), ""))).ToArray()
                series(1).XValues = categoryArray
                series(1).Values = numericValues.ToArray()
                GlobalStatusStrip.ShowInfo($"Данные диаграммы обновлены: {categories.Count} категорий")
            End If

            Dim seriesName = params("seriesName")?.ToString()
            If Not String.IsNullOrWhiteSpace(seriesName) Then series(1).Name = seriesName

            Return True
        Catch ex As Exception
            _lastPptCommandError = "UpdateChartData: " & ex.Message
            Debug.WriteLine($"ExecuteUpdateChartData ошибка: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim position = If(params("position")?.ToString(), "end")
            Dim title = If(params("title")?.ToString(), "")
            Dim content = If(params("content")?.ToString(), "")
            Dim layout = If(params("layout")?.ToString(), "titleAndContent")

            Dim slideIndex As Integer
            If position.ToLower() = "end" Then
                slideIndex = pres.Slides.Count + 1
            ElseIf position.ToLower() = "current" Then
                slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex + 1
            Else
                slideIndex = pres.Slides.Count + 1
            End If

            Dim layoutType As Integer = GetLayoutType(layout)
            Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
            Try
                slide = pres.Slides.Add(slideIndex, layoutType)

                TrySetPlaceholderText(slide, title, Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)

                If Not String.IsNullOrEmpty(content) Then
                    Dim contentFilled = TrySetPlaceholderText(
                        slide,
                        content,
                        Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody,
                        Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle)

                    ' Если у макета нет текстового placeholder, добавляем текстовое поле.
                    If Not contentFilled Then
                        AddTextBoxToSlide(slide, content, 50, 150, 620, 350, 18)
                    End If
                End If
            Finally
                ComObjectHelper.ReleaseComObject(slide)
            End Try

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertText(params As JToken, pres As Object) As Boolean
        Try
            Dim content = params("content")?.ToString()
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            Dim x = If(params("x")?.Value(Of Single)(), 100)
            Dim y = If(params("y")?.Value(Of Single)(), 200)

            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(slideIndex + 1)
            End If

            Dim textBox = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                x, y, 400, 100)
            textBox.TextFrame.TextRange.Text = content

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertShape(params As JToken, pres As Object) As Boolean
        Try
            Dim shapeType = If(params("shapeType")?.ToString(), "rectangle")
            Dim x = params("x")?.Value(Of Single)()
            Dim y = params("y")?.Value(Of Single)()
            Dim width = If(params("width")?.Value(Of Single)(), 100)
            Dim height = If(params("height")?.Value(Of Single)(), 100)

            Dim slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide

            ' 根据shapeType添加不同形状
            Dim msoShapeType As Integer = 1 ' msoShapeRectangle
            Select Case shapeType.ToLower()
                Case "rectangle"
                    msoShapeType = 1
                Case "oval", "circle"
                    msoShapeType = 9 ' msoShapeOval
                Case "triangle"
                    msoShapeType = 7 ' msoShapeIsoscelesTriangle
                Case "arrow"
                    msoShapeType = 13 ' msoShapeRightArrow
            End Select

            slide.Shapes.AddShape(msoShapeType, x, y, width, height)
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertShape 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteFormatSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            
            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(slideIndex + 1)
            End If

            ' 设置背景
            Dim background = params("background")?.ToString()
            If Not String.IsNullOrEmpty(background) Then
                Try
                    ' 尝试解析颜色
                    Dim color = System.Drawing.ColorTranslator.FromHtml(background)
                    slide.FollowMasterBackground = False
                    slide.Background.Fill.Solid()
                    slide.Background.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(color)
                Catch
                End Try
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertTable(params As JToken, pres As Object) As Boolean
        Try
            Dim rows = params("rows")?.Value(Of Integer)()
            Dim cols = params("cols")?.Value(Of Integer)()
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)

            If rows <= 0 OrElse cols <= 0 Then Return False

            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(slideIndex + 1)
            End If

            Dim table = slide.Shapes.AddTable(rows, cols, 50, 150, 600, 300)

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
                            table.Table.Cell(rowIdx + 1, colIdx + 1).Shape.TextFrame.TextRange.Text = rowArr(colIdx).ToString()
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

#Region "高级PPT命令实现"

    ''' <summary>
    ''' 批量创建幻灯片
    ''' </summary>
    Private Function ExecuteCreateSlides(params As JToken, pres As Object) As Boolean
        Try
            Dim slides = params("slides")
            If slides Is Nothing OrElse slides.Type <> JTokenType.Array Then
                Return False
            End If

            Dim slidesArray = CType(slides, JArray)
            Dim startIndex = pres.Slides.Count + 1

            For i = 0 To slidesArray.Count - 1
                Dim slideData = slidesArray(i)
                Dim title = If(slideData("title")?.ToString(), "")
                Dim content = If(slideData("content")?.ToString(), "")
                Dim layout = If(slideData("layout")?.ToString(), "titleAndContent")

                ' 根据layout选择布局类型
                Dim layoutType As Integer = GetLayoutType(layout)
                Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                Try
                    slide = pres.Slides.Add(startIndex + i, layoutType)

                    TrySetPlaceholderText(slide, title, Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)

                    ' 填充内容
                    If Not String.IsNullOrEmpty(content) Then
                        Dim contentFilled = TrySetPlaceholderText(
                            slide,
                            content,
                            Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody,
                            Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle)

                        ' 如果没有内容占位符，添加文本框
                        If Not contentFilled Then
                            AddTextBoxToSlide(slide, content, 50, 150, 620, 350, 18)
                        End If
                    End If

                    ' CreateSlides 不能把 Office 默认白底文本当成可交付设计。
                    BeautifySingleSlide(slide, BuildDefaultPptTheme(i))
                Finally
                    ComObjectHelper.ReleaseComObject(slide)
                End Try
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo($"Создано слайдов: {slidesArray.Count}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreateSlides 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function TrySetPlaceholderText(
        slide As Microsoft.Office.Interop.PowerPoint.Slide,
        text As String,
        ParamArray placeholderTypes() As Microsoft.Office.Interop.PowerPoint.PpPlaceholderType) As Boolean

        If slide Is Nothing OrElse String.IsNullOrEmpty(text) Then Return False

        Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
        Try
            shapes = slide.Shapes
            For i = 1 To shapes.Count
                Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                Try
                    shape = shapes(i)
                    If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                        Dim placeholderType = shape.PlaceholderFormat.Type
                        If placeholderTypes.Contains(placeholderType) Then
                            textRange = shape.TextFrame.TextRange
                            textRange.Text = text
                            Return True
                        End If
                    End If
                Finally
                    ComObjectHelper.ReleaseComObject(textRange)
                    ComObjectHelper.ReleaseComObject(shape)
                End Try
            Next
        Finally
            ComObjectHelper.ReleaseComObject(shapes)
        End Try

        Return False
    End Function

    Private Sub AddTextBoxToSlide(
        slide As Microsoft.Office.Interop.PowerPoint.Slide,
        text As String,
        x As Single,
        y As Single,
        width As Single,
        height As Single,
        Optional fontSize As Single = 0)

        Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
        Dim textBox As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
        Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
        Try
            shapes = slide.Shapes
            textBox = shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                x, y, width, height)
            textRange = textBox.TextFrame.TextRange
            textRange.Text = text
            If fontSize > 0 Then
                textRange.Font.Size = fontSize
            End If
        Finally
            ComObjectHelper.ReleaseComObject(textRange)
            ComObjectHelper.ReleaseComObject(textBox)
            ComObjectHelper.ReleaseComObject(shapes)
        End Try
    End Sub

    ''' <summary>
    ''' 获取布局类型
    ''' </summary>
    Private Function GetLayoutType(layout As String) As Integer
        Select Case layout.ToLower()
            Case "title", "titleonly"
                Return 11 ' ppLayoutTitleOnly
            Case "titleandcontent", "content"
                Return 2 ' ppLayoutText
            Case "twocontent", "twotext"
                Return 3 ' ppLayoutTwoColumnText
            Case "blank"
                Return 7 ' ppLayoutBlank
            Case "sectionheader"
                Return 1 ' ppLayoutTitle
            Case Else
                Return 2 ' ppLayoutText (默认)
        End Select
    End Function

    ''' <summary>
    ''' 添加动画效果
    ''' </summary>
    Private Function ExecuteAddAnimation(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            Dim effect = If(params("effect")?.ToString(), "fadeIn")
            Dim targetShapes = If(params("targetShapes")?.ToString(), "all")

            Dim msoEffect = GetMsoAnimEffect(effect)
            Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
            Dim timeline As Microsoft.Office.Interop.PowerPoint.TimeLine = Nothing
            Dim sequence As Microsoft.Office.Interop.PowerPoint.Sequence = Nothing
            Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
            Try
                If slideIndex < 0 Then
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                Else
                    slide = pres.Slides(slideIndex + 1)
                End If

                timeline = slide.TimeLine
                sequence = timeline.MainSequence
                shapes = slide.Shapes

                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(i)
                        Dim shouldAnimate = False

                        If targetShapes.ToLower() = "all" Then
                            shouldAnimate = True
                        ElseIf targetShapes.ToLower() = "title" Then
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                               shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                shouldAnimate = True
                            End If
                        ElseIf targetShapes.ToLower() = "content" Then
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                               (shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody OrElse
                                shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle) Then
                                shouldAnimate = True
                            End If
                        End If

                        If shouldAnimate AndAlso shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim animationEffect As Microsoft.Office.Interop.PowerPoint.Effect = Nothing
                            Try
                                animationEffect = sequence.AddEffect(shape, msoEffect)
                            Catch
                                ' 某些形状可能不支持动画
                            Finally
                                ComObjectHelper.ReleaseComObject(animationEffect)
                            End Try
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
                ComObjectHelper.ReleaseComObject(sequence)
                ComObjectHelper.ReleaseComObject(timeline)
                ComObjectHelper.ReleaseComObject(slide)
            End Try

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteAddAnimation 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取动画效果类型
    ''' </summary>
    Private Function GetMsoAnimEffect(effect As String) As Integer
        Select Case effect.ToLower()
            Case "fadein", "fade"
                Return 10 ' msoAnimEffectFade
            Case "flyin", "fly"
                Return 2 ' msoAnimEffectFly
            Case "zoom"
                Return 53 ' msoAnimEffectGrowAndTurn
            Case "wipe"
                Return 22 ' msoAnimEffectWipe
            Case "appear"
                Return 1 ' msoAnimEffectAppear
            Case "float"
                Return 42 ' msoAnimEffectFloat
            Case Else
                Return 10 ' msoAnimEffectFade (默认)
        End Select
    End Function

    ''' <summary>
    ''' 应用幻灯片切换效果
    ''' </summary>
    Private Function ExecuteApplyTransition(params As JToken, pres As Object) As Boolean
        Try
            Dim transType = If(params("transitionType")?.ToString(), "fade")
            Dim scope = If(params("scope")?.ToString(), "all")
            Dim duration = If(params("duration")?.Value(Of Single)(), 1.0F)

            Dim transEffect = GetTransitionEffect(transType)

            Dim processedCount As Integer = 0
            If scope.ToLower() = "all" Then
                Dim slides As Microsoft.Office.Interop.PowerPoint.Slides = Nothing
                Try
                    slides = pres.Slides
                    For i = 1 To slides.Count
                        Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                        Dim transition As Microsoft.Office.Interop.PowerPoint.SlideShowTransition = Nothing
                        Try
                            slide = slides(i)
                            transition = slide.SlideShowTransition
                            transition.EntryEffect = transEffect
                            transition.Duration = duration
                            transition.AdvanceOnClick = True
                            processedCount += 1
                        Finally
                            ComObjectHelper.ReleaseComObject(transition)
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next
                Finally
                    ComObjectHelper.ReleaseComObject(slides)
                End Try
            Else
                Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                Dim transition As Microsoft.Office.Interop.PowerPoint.SlideShowTransition = Nothing
                Try
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                    transition = slide.SlideShowTransition
                    transition.EntryEffect = transEffect
                    transition.Duration = duration
                    transition.AdvanceOnClick = True
                    processedCount = 1
                Finally
                    ComObjectHelper.ReleaseComObject(transition)
                    ComObjectHelper.ReleaseComObject(slide)
                End Try
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"Переходы применены к слайдам: {processedCount}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyTransition 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取切换效果类型（PpEntryEffect 官方枚举值）。
    ''' 之前的数值无效或指向别的效果（например 257 是 ppEffectCut, а не fade），
    ''' из-за чего PowerPoint молча игнорировал переход и наблюдение не видело изменений.
    ''' </summary>
    Private Function GetTransitionEffect(transType As String) As Integer
        Select Case transType.ToLower()
            Case "fade"
                Return 1793 ' ppEffectFade
            Case "push"
                Return 3852 ' ppEffectPushDown
            Case "wipe"
                Return 2817 ' ppEffectWipeLeft
            Case "split"
                Return 3585 ' ppEffectSplitHorizontalOut
            Case "reveal"
                Return 3894 ' ppEffectRevealSmoothLeft
            Case "random"
                Return 513 ' ppEffectRandom
            Case Else
                Return 1793 ' ppEffectFade (默认)
        End Select
    End Function

    ''' <summary>
    ''' 美化幻灯片
    ''' </summary>
    Private Function ExecuteBeautifySlides(params As JToken, pres As Object) As Boolean
        Try
            Dim scope = If(params("scope")?.ToString(), "all")
            Dim theme = params("theme")

            Dim processedCount As Integer = 0
            If scope.ToLower() = "all" Then
                Dim slides As Microsoft.Office.Interop.PowerPoint.Slides = Nothing
                Try
                    slides = pres.Slides
                    For i = 1 To slides.Count
                        Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                        Try
                            slide = slides(i)
                            BeautifySingleSlide(slide, theme)
                            processedCount += 1
                        Finally
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next
                Finally
                    ComObjectHelper.ReleaseComObject(slides)
                End Try
            Else
                Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                Try
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                    BeautifySingleSlide(slide, theme)
                    processedCount = 1
                Finally
                    ComObjectHelper.ReleaseComObject(slide)
                End Try
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"Оформлено слайдов: {processedCount}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteBeautifySlides 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub BeautifySingleSlide(slide As Microsoft.Office.Interop.PowerPoint.Slide, theme As JToken)
        If slide Is Nothing Then Return
        If theme Is Nothing Then theme = BuildDefaultPptTheme(Math.Max(0, slide.SlideIndex - 1))

        ' 应用背景色
        If theme IsNot Nothing AndAlso theme("background") IsNot Nothing Then
            Try
                Dim bgColor = theme("background").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(bgColor)
                slide.FollowMasterBackground = False
                slide.Background.Fill.Solid()
                slide.Background.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(color)
            Catch
            End Try
        End If

        ' 应用字体样式
        Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
        Try
            shapes = slide.Shapes
            For i = 1 To shapes.Count
                Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                Try
                    shape = shapes(i)
                    If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim isTitle = False
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            isTitle = (shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)
                        End If

                        Dim fontTheme = If(isTitle, theme?("titleFont"), theme?("bodyFont"))
                        If fontTheme IsNot Nothing Then
                            textRange = shape.TextFrame.TextRange
                            ApplyFontStyle(textRange, fontTheme)
                        End If
                    End If
                Finally
                    ComObjectHelper.ReleaseComObject(textRange)
                    ComObjectHelper.ReleaseComObject(shape)
                End Try
            Next
        Finally
            ComObjectHelper.ReleaseComObject(shapes)
        End Try
    End Sub

    Private Function BuildDefaultPptTheme(themeVariant As Integer) As JObject
        Dim backgrounds = {"#F7F3EA", "#EEF4F7", "#F5F1F7"}
        Dim accents = {"#7A3E2E", "#235B70", "#62406F"}
        Dim index = Math.Abs(themeVariant) Mod backgrounds.Length
        Return New JObject From {
            {"background", backgrounds(index)},
            {"titleFont", New JObject From {
                {"name", "Microsoft YaHei"},
                {"size", 30},
                {"bold", True},
                {"color", accents(index)}
            }},
            {"bodyFont", New JObject From {
                {"name", "Microsoft YaHei"},
                {"size", 20},
                {"color", "#27323A"}
            }}
        }
    End Function

    ''' <summary>
    ''' 应用字体样式
    ''' </summary>
    Private Sub ApplyFontStyle(textRange As Object, fontTheme As JToken)
        Try
            If fontTheme("name") IsNot Nothing Then
                textRange.Font.Name = fontTheme("name").ToString()
            End If
            If fontTheme("size") IsNot Nothing Then
                textRange.Font.Size = fontTheme("size").Value(Of Single)()
            End If
            If fontTheme("color") IsNot Nothing Then
                Dim colorStr = fontTheme("color").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                textRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(color)
            End If
            If fontTheme("bold") IsNot Nothing Then
                textRange.Font.Bold = If(fontTheme("bold").Value(Of Boolean)(), -1, 0)
            End If
        Catch ex As Exception
            Debug.WriteLine($"ApplyFontStyle 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 删除幻灯片
    ''' </summary>
    Private Function ExecuteDeleteSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            If slideIndex = -1 OrElse slideIndex = 0 Then
                slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
            End If
            If pres.Slides.Count <= 1 Then
                GlobalStatusStrip.ShowWarning("Невозможно удалить: в презентации должен остаться хотя бы один слайд")
                Return False
            End If
            If slideIndex < 1 OrElse slideIndex > pres.Slides.Count Then
                GlobalStatusStrip.ShowWarning($"Индекс слайда {slideIndex} вне диапазона (всего {pres.Slides.Count})")
                Return False
            End If
            pres.Slides(slideIndex).Delete()
            GlobalStatusStrip.ShowInfo($"Удалён слайд {slideIndex}")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteDeleteSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 复制幻灯片
    ''' </summary>
    Private Function ExecuteDuplicateSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            If slideIndex = -1 OrElse slideIndex = 0 Then
                slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
            End If
            If slideIndex < 1 OrElse slideIndex > pres.Slides.Count Then
                _lastPptCommandError = $"DuplicateSlide: слайд {slideIndex} вне диапазона 1..{pres.Slides.Count}"
                Return False
            End If

            Dim position = If(params("position")?.ToString(), "after")
            pres.Slides(slideIndex).Copy()
            Dim newIndex As Integer
            If String.Equals(position.Trim(), "end", StringComparison.OrdinalIgnoreCase) Then
                Try
                    pres.Slides.Paste(pres.Slides.Count + 1)
                Catch
                    pres.Slides.Paste()
                End Try
                newIndex = pres.Slides.Count
            Else
                pres.Slides.Paste(slideIndex + 1)
                newIndex = slideIndex + 1
            End If

            ' Копия становится активной, иначе последующие команды в том же конверте
            ' с slideIndex=-1 правят оригинал, а не новую копию.
            Try
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(newIndex)
            Catch ex As Exception
                Debug.WriteLine($"DuplicateSlide: не удалось переключиться на копию: {ex.Message}")
            End Try

            _lastPptCommandTargetIndex = newIndex
            _lastPptCommandInfo = $"новый слайд {newIndex} (копия слайда {slideIndex}), активен; далее используй slideIndex=-1 или {newIndex}"

            GlobalStatusStrip.ShowInfo($"Скопирован слайд {slideIndex} → позиция {newIndex}")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteDuplicateSlide ошибка: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 移动幻灯片
    ''' </summary>
    Private Function ExecuteMoveSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim fromIndex As Integer = If(params("fromIndex") IsNot Nothing, params("fromIndex").Value(Of Integer)(), -1)
            Dim toIndex As Integer = If(params("toIndex") IsNot Nothing, params("toIndex").Value(Of Integer)(), -1)
            If fromIndex < 1 OrElse fromIndex > pres.Slides.Count Then Return False
            If toIndex < 1 OrElse toIndex > pres.Slides.Count Then Return False
            pres.Slides(fromIndex).MoveTo(toIndex)
            GlobalStatusStrip.ShowInfo($"Слайд перемещён с позиции {fromIndex} на {toIndex}")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteMoveSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 设置幻灯片版式
    ''' </summary>
    Private Function ExecuteSetSlideLayout(params As JToken, pres As Object) As Boolean
        Try
            Dim layoutName = If(params("layout")?.ToString(), "titleandcontent").ToLower()
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            Dim targetSlide As Object
            If slideIndex = -1 OrElse slideIndex = 0 Then
                targetSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            ElseIf slideIndex >= 1 AndAlso slideIndex <= pres.Slides.Count Then
                targetSlide = pres.Slides(slideIndex)
            Else
                Return False
            End If
            Dim layoutIndex As Integer = 2
            Select Case layoutName
                Case "title" : layoutIndex = 1
                Case "titleandcontent", "titleandbody" : layoutIndex = 2
                Case "blank" : layoutIndex = 12
                Case "titleonly" : layoutIndex = 11
            End Select
            targetSlide.Layout = layoutIndex
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteSetSlideLayout 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 应用主题
    ''' </summary>
    Private Function ExecuteApplyTheme(params As JToken, pres As Object) As Boolean
        Try
            Dim themeFile = If(params("themeFile")?.ToString(), "")
            Dim themeName = If(params("themeName")?.ToString(), "")
            If Not String.IsNullOrEmpty(themeFile) AndAlso IO.File.Exists(themeFile) Then
                pres.ApplyTheme(themeFile)
            ElseIf Not String.IsNullOrEmpty(themeName) Then
                GlobalStatusStrip.ShowInfo($"Встроенную тему '{themeName}' примените вручную через меню дизайна WPS/PPT")
            End If
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyTheme 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 添加演讲备注
    ''' </summary>
    Private Function ExecuteAddSpeakerNotes(params As JToken, pres As Object) As Boolean
        Try
            Dim notes = If(params("notes")?.ToString(), "")
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            Dim targetSlide As Object
            If slideIndex = -1 OrElse slideIndex = 0 Then
                targetSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            ElseIf slideIndex >= 1 AndAlso slideIndex <= pres.Slides.Count Then
                targetSlide = pres.Slides(slideIndex)
            Else
                Return False
            End If
            targetSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text = notes
            GlobalStatusStrip.ShowInfo("Заметки докладчика добавлены")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteAddSpeakerNotes 出错: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

#Region "排版功能"

    ''' <summary>
    ''' 处理排版JSON响应（支持规则模式）
    ''' </summary>
    Protected Overrides Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
        Try
            ' 检测是否为规则模式
            If jsonDoc("rules") IsNot Nothing AndAlso jsonDoc("rules").Type = JTokenType.Array Then
                ApplyReformatRules(jsonDoc)
                Return
            End If

            ' 无效格式
            GlobalStatusStrip.ShowWarning("Недопустимый формат ответа форматирования")

        Catch ex As Exception
            Debug.WriteLine("HandleApplyDocumentPlanItem 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("Ошибка применения форматирования: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 应用规则模式的排版
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

            ' 检查上下文
            If _reformatShapes Is Nothing OrElse _reformatShapes.Count = 0 Then
                GlobalStatusStrip.ShowWarning("Контекст форматирования потерян. Выберите содержимое и запустите форматирование заново")
                Return
            End If

            ' 基于样本分类推断规则
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

            ' 应用格式到所有形状
            Dim appliedCount As Integer = 0
            Dim defaultRule As String = If(ruleDict.ContainsKey("body"), "body", ruleDict.Keys.FirstOrDefault())

            ' 保存排版前快照（用于PPT撤销，因PPT不支持UndoRecord）
            _reformatUndoSnapshots = New List(Of ShapeFormatSnapshot)()
            For i As Integer = 0 To _reformatShapes.Count - 1
                Try
                    Dim shp = DirectCast(_reformatShapes(i), Microsoft.Office.Interop.PowerPoint.Shape)
                    _reformatUndoSnapshots.Add(ShapeFormatSnapshot.Capture(shp, i))
                Catch ex As Exception
                    Debug.WriteLine($"捕获形状{i}快照失败: {ex.Message}")
                End Try
            Next

            For i As Integer = 0 To _reformatShapes.Count - 1
                Try
                    Dim shp = DirectCast(_reformatShapes(i), Microsoft.Office.Interop.PowerPoint.Shape)
                    Dim shapeType = If(i < _reformatTypes.Count, _reformatTypes(i), "")

                    ' 确定使用哪个规则
                    Dim ruleToApply As String = defaultRule

                    If sampleRuleMap.ContainsKey(i) Then
                        ruleToApply = sampleRuleMap(i)
                    Else
                        ' 基于形状类型推断规则
                            If shapeType.Contains("Заголовок") Then
                            For Each key In ruleDict.Keys
                                If key.ToLower().Contains("title") OrElse key = "标题" Then
                                    ruleToApply = key
                                    Exit For
                                End If
                            Next
                            ElseIf shapeType.Contains("Подзаголовок") Then
                            For Each key In ruleDict.Keys
                                If key.ToLower().Contains("subtitle") OrElse key = "副标题" Then
                                    ruleToApply = key
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    ' 应用规则
                    If ruleDict.ContainsKey(ruleToApply) Then
                        Dim formatting = ruleDict(ruleToApply)
                        ApplyFormattingToShape(shp, formatting)
                        appliedCount += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"应用形状 {i} 格式失败: " & ex.Message)
                End Try
            Next

            ' 保留 _reformatShapes/_reformatTypes 用于撤销快照恢复
            ' 撤销或新排版开始时会重新设置

            GlobalStatusStrip.ShowInfo($"Форматирование завершено, обработано текстовых полей: {appliedCount}")

        Catch ex As Exception
            Debug.WriteLine("ApplyReformatRules 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("Ошибка применения правил форматирования: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 应用格式化属性到形状
    ''' </summary>
    Private Sub ApplyFormattingToShape(shp As Microsoft.Office.Interop.PowerPoint.Shape, formatting As JObject)
        Try
            If shp.HasTextFrame <> Microsoft.Office.Core.MsoTriState.msoTrue Then Return

            Dim textRange = shp.TextFrame.TextRange

            ' 中文字体
            If formatting("fontNameCN") IsNot Nothing Then
                Try
                    textRange.Font.NameFarEast = formatting("fontNameCN").ToString()
                Catch
                End Try
            End If

            ' 英文字体
            If formatting("fontNameEN") IsNot Nothing Then
                Try
                    textRange.Font.Name = formatting("fontNameEN").ToString()
                Catch
                End Try
            End If

            ' 字号
            If formatting("fontSize") IsNot Nothing Then
                Dim fontSize As Single = 0
                Single.TryParse(formatting("fontSize").ToString(), fontSize)
                If fontSize > 0 Then
                    Try
                        textRange.Font.Size = fontSize
                    Catch
                    End Try
                End If
            End If

            ' 加粗
            If formatting("bold") IsNot Nothing Then
                Try
                    Dim bold As Boolean = formatting("bold").ToObject(Of Boolean)()
                    textRange.Font.Bold = If(bold, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse)
                Catch
                End Try
            End If

            ' 对齐方式
            If formatting("alignment") IsNot Nothing Then
                Dim alignment As String = formatting("alignment").ToString().ToLower()
                Try
                    Select Case alignment
                        Case "left"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft
                        Case "center"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter
                        Case "right"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight
                        Case "justify"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignJustify
                    End Select
                Catch
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine("ApplyFormattingToShape 出错: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 从快照恢复格式（PPT专用撤销方式）
    ''' </summary>
    Private Sub RestoreFormattingSnapshots()
        If _reformatUndoSnapshots Is Nothing OrElse _reformatUndoSnapshots.Count = 0 Then
            Debug.WriteLine("没有排版快照可恢复")
            Return
        End If

        Dim restoredCount As Integer = 0
        For Each snap In _reformatUndoSnapshots
            Try
                If snap.ShapeIndex >= 0 AndAlso snap.ShapeIndex < _reformatShapes.Count Then
                    Dim shp = DirectCast(_reformatShapes(snap.ShapeIndex), Microsoft.Office.Interop.PowerPoint.Shape)
                    snap.Restore(shp)
                    restoredCount += 1
                End If
            Catch ex As Exception
                Debug.WriteLine($"恢复形状{snap.ShapeIndex}失败: {ex.Message}")
            End Try
        Next
        Debug.WriteLine($"从快照恢复了 {restoredCount} 个形状的格式")

        ' 清理快照和上下文
        _reformatUndoSnapshots = Nothing
        _reformatShapes = Nothing
        _reformatTypes = Nothing
    End Sub

    ''' <summary>
    ''' 撤销排版（PPT重写：优先从快照恢复，失败后回退到基类Undo）
    ''' </summary>
    Protected Overrides Sub HandleUndoReformat()
        ' 优先尝试从快照恢复（最可靠）
        If _reformatUndoSnapshots IsNot Nothing AndAlso _reformatUndoSnapshots.Count > 0 Then
            Try
                RestoreFormattingSnapshots()
                GlobalStatusStrip.ShowInfo("Форматирование отменено")
                Return
            Catch ex As Exception
                Debug.WriteLine($"快照恢复失败，回退到Undo: {ex.Message}")
            End Try
        End If

        ' 回退到基类Undo机制
        MyBase.HandleUndoReformat()
    End Sub

#End Region

End Class

