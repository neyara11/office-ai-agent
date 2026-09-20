Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports ShareRibbon
Public Class ThisAddIn

    ' 按文档窗口维护聊天面板：多窗口时各窗口拥有独立面板
    Public Shared ReadOnly Property chatControl As ChatControl
        Get
            Dim addIn = Globals.ThisAddIn
            If addIn Is Nothing Then Return Nothing
            Return addIn.GetActiveChatControl()
        End Get
    End Property

    Private _chatPanes As HostTaskPaneRegistry

    Private ReadOnly Property ChatPanes As HostTaskPaneRegistry
        Get
            If _chatPanes Is Nothing Then
                _chatPanes = New HostTaskPaneRegistry(Me.CustomTaskPanes, AddressOf GetWindowHandle)
            End If
            Return _chatPanes
        End Get
    End Property

    ' 翻译服务：延迟初始化，首次使用时创建

    Private captureTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Shared dataCapturePane As WebDataCapturePane

    ' 在类中添加以下变量
    Private _deepseekControl As DeepseekControl
    Private _deepseekTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _doubaoControl As DoubaoChat
    Private _doubaoTaskPane As Microsoft.Office.Tools.CustomTaskPane

    ' 模板编辑器
    Private _templateEditorControl As ReformatTemplateEditorControl
    Private _templateEditorTaskPane As Microsoft.Office.Tools.CustomTaskPane

    ' 延迟初始化：WebView2 和 SQLite 仅在首次使用时加载
    Private _lazyWebView2 As New Lazy(Of Boolean)(Function()
        WebView2Loader.EnsureWebView2Loader()
        Return True
    End Function)

    Private _lazySqlite As New Lazy(Of Boolean)(Function()
        SqliteNativeLoader.EnsureLoaded()
        Return True
    End Function)

    ' WPS 宽度修复定时器
    Private widthTimer As Timer
    Private widthTimer1 As Timer

    Private Sub WordAi_Startup() Handles Me.Startup
        ' Phase 0: 仅注册事件处理器 + 状态栏初始化（微秒级，不阻塞启动）
        PhaseStartupManager.Instance.RunCriticalPhase(Me.Application)
    End Sub

    ''' <summary>
    ''' 确保核心服务已加载（WebView2 + SQLite），首次调用时初始化
    ''' 如果 Phase 2 后台预加载已完成，则直接跳过
    ''' </summary>
    Private Sub EnsureCoreServicesLoaded()
        Try
            Dim webView2Init = _lazyWebView2.Value
        Catch ex As Exception
            MessageBox.Show($"Ошибка инициализации WebView2: {ex.Message}")
        End Try
        Try
            Dim sqliteInit = _lazySqlite.Value
        Catch ex As Exception
            MessageBox.Show($"Не удалось загрузить нативную библиотеку SQLite; функции Skills/памяти могут быть недоступны: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 确保 WPS 宽度修复定时器已初始化（仅在需要时创建）
    ''' </summary>
    Private Sub EnsureWidthTimers()
        If widthTimer Is Nothing Then
            widthTimer = New Timer()
            AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
            widthTimer.Interval = 100
        End If
        If widthTimer1 Is Nothing Then
            widthTimer1 = New Timer()
            AddHandler widthTimer1.Tick, AddressOf WidthTimer1_Tick
            widthTimer1.Interval = 100
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 清理翻译服务资源（右键菜单按钮）
        ' 清理定时器资源
        If widthTimer IsNot Nothing Then
            widthTimer.Stop()
            widthTimer.Dispose()
            widthTimer = Nothing
        End If
        If widthTimer1 IsNot Nothing Then
            widthTimer1.Stop()
            widthTimer1.Dispose()
            widthTimer1 = Nothing
        End If
    End Sub


    ''' <summary>返回当前活动文档窗口对应的聊天面板，必要时为该窗口创建。</summary>
    Private Function EnsureChatPaneEntry() As HostTaskPaneRegistry.Entry
        Try
            ChatPanes.PruneClosed(AddressOf GetLiveWindowHandles)

            Dim window = GetActiveWindowObject()
            If window Is Nothing Then Return Nothing

            Dim entry = ChatPanes.GetOrCreate(window, "ИИ-помощник Word", Function() New ChatControl())
            If entry Is Nothing Then Return Nothing

            entry.Pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            entry.Pane.Width = 420
            If entry.IsNew Then
                AddHandler entry.Pane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            End If
            Return entry
        Catch ex As Exception
            MessageBox.Show($"Не удалось инициализировать панель задач Word AI: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>当前活动窗口的聊天控件；该窗口尚未打开聊天时返回 Nothing。</summary>
    Friend Function GetActiveChatControl() As ChatControl
        Try
            Dim window = GetActiveWindowObject()
            If window Is Nothing Then Return Nothing
            Dim entry = ChatPanes.GetExisting(GetWindowHandle(window))
            If entry IsNot Nothing Then Return TryCast(entry.Control, ChatControl)
        Catch ex As Exception
            Debug.WriteLine($"[Chat] GetActiveChatControl 失败: {ex.Message}")
        End Try
        Return Nothing
    End Function

    Private Function GetActiveWindowObject() As Object
        Try
            Return Me.Application.ActiveWindow
        Catch
            Return Nothing
        End Try
    End Function

    Private Function GetWindowHandle(window As Object) As Integer
        Dim wordWindow = TryCast(window, Word.Window)
        If wordWindow Is Nothing Then Return 0
        Try
            Return wordWindow.Hwnd
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>当前所有打开的文档窗口句柄；整体枚举失败时返回 Nothing 以跳过清理。</summary>
    Private Function GetLiveWindowHandles() As HashSet(Of Integer)
        Dim result As New HashSet(Of Integer)()
        Try
            For Each w As Word.Window In Me.Application.Windows
                Try
                    result.Add(w.Hwnd)
                Catch
                End Try
            Next
        Catch ex As Exception
            Debug.WriteLine($"[Chat] 枚举窗口失败: {ex.Message}")
            Return Nothing
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 创建网页爬虫任务窗格（延迟初始化，仅在用户点击爬虫按钮时创建）
    ''' </summary>
    Private Sub EnsureDataCapturePaneCreated()
        If dataCapturePane IsNot Nothing Then Return
        Try
            dataCapturePane = New WebDataCapturePane()
            captureTaskPane = Me.CustomTaskPanes.Add(dataCapturePane, "Краулер Word")
            captureTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            captureTaskPane.Width = 420
            captureTaskPane.Visible = False
            ' WebView2 краулера создаём при первом показе панели: если создать его раньше,
            ' пока Office ещё не перевесил контрол в окно задачи, страница рисуется,
            ' но не получает мышь (клики, наведение, прокрутка).
            AddHandler captureTaskPane.VisibleChanged, AddressOf CaptureTaskPane_VisibleChanged
        Catch ex As Exception
            MessageBox.Show($"Не удалось инициализировать панель задач краулера Word: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Первый показ панели краулера — момент, когда можно безопасно создать WebView2.
    ''' </summary>
    Private Sub CaptureTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane = TryCast(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane Is Nothing OrElse Not taskPane.Visible Then Return
        dataCapturePane?.EnsureWebViewInitialized()
    End Sub

    ' 解决WPS中无法显示正常宽度的问题
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If LLMUtil.IsWpsActive() Then
                EnsureWidthTimers()
                widthTimer.Start()
            End If
        End If
    End Sub

    Private Sub DeepseekTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If LLMUtil.IsWpsActive() Then
                EnsureWidthTimers()
                widthTimer1.Start()
            End If
        End If
    End Sub

    Private Sub CreateDeepseekTaskPane()
        Try
            If _deepseekControl Is Nothing Then
                ' 为新工作簿创建任务窗格
                _deepseekControl = New DeepseekControl()
                _deepseekTaskPane = Me.CustomTaskPanes.Add(_deepseekControl, "ИИ-помощник Deepseek")
                _deepseekTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _deepseekTaskPane.Width = 420
                AddHandler _deepseekTaskPane.VisibleChanged, AddressOf DeepseekTaskPane_VisibleChanged
                _deepseekTaskPane.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show($"Не удалось инициализировать панель задач: {ex.Message}")
        End Try
    End Sub

    Private Async Function CreateDoubaoTaskPane() As Task
        Try
            If _doubaoControl Is Nothing Then
                ' 为新工作簿创建任务窗格
                _doubaoControl = New DoubaoChat()
                Await _doubaoControl.InitializeAsync()
                _doubaoTaskPane = Me.CustomTaskPanes.Add(_doubaoControl, "ИИ-помощник Doubao")
                _doubaoTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _doubaoTaskPane.Width = 420
            End If
        Catch ex As Exception
            MessageBox.Show($"Не удалось инициализировать панель задач Doubao: {ex.Message}")
        End Try
    End Function

    Private Sub WidthTimer_Tick(sender As Object, e As EventArgs)
        widthTimer.Stop()
        If LLMUtil.IsWpsActive() Then
            For Each entry In ChatPanes.Entries
                Try
                    entry.Pane.Width = 420
                Catch ex As Exception
                    Debug.WriteLine($"[Chat] 设置面板宽度失败: {ex.Message}")
                End Try
            Next
        End If
    End Sub
    Private Sub WidthTimer1_Tick(sender As Object, e As EventArgs)
        widthTimer1.Stop()
        If LLMUtil.IsWpsActive() AndAlso _deepseekTaskPane IsNot Nothing Then
            _deepseekTaskPane.Width = 420
        End If
    End Sub

    Dim loadDataCaptureHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        Await ShowChatTaskPaneAsync()
    End Sub

    Public Async Function ShowChatTaskPaneAsync() As Task
        EnsureCoreServicesLoaded()
        Dim entry = EnsureChatPaneEntry()
        If entry Is Nothing Then Return
        entry.Pane.Visible = True
        If Not entry.HtmlLoaded Then
            entry.HtmlLoaded = True
            Try
                Await CType(entry.Control, ChatControl).LoadLocalHtmlFile()
            Catch ex As Exception
                entry.HtmlLoaded = False
                Debug.WriteLine($"[Chat] 加载聊天页面失败: {ex.Message}")
            End Try
        End If
    End Function

    Public Sub ShowDataCaptureTaskPane()
        EnsureDataCapturePaneCreated()
        If captureTaskPane Is Nothing Then Return
        captureTaskPane.Visible = True
    End Sub

    Public Sub ShowDeepseekTaskPane()
        EnsureCoreServicesLoaded()
        CreateDeepseekTaskPane()
        If _deepseekTaskPane Is Nothing Then Return
        _deepseekTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDoubaoTaskPane()
        EnsureCoreServicesLoaded()
        Await CreateDoubaoTaskPane()
        If _doubaoTaskPane Is Nothing Then Return
        _doubaoTaskPane.Visible = True
    End Sub

    ''' <summary>
    ''' 显示模板编辑器任务窗格
    ''' </summary>
    Public Sub ShowTemplateEditorTaskPane(Optional template As ReformatTemplate = Nothing)
        Try
            ' 如果已存在，先关闭
            If _templateEditorTaskPane IsNot Nothing Then
                Try
                    Me.CustomTaskPanes.Remove(_templateEditorTaskPane)
                Catch
                End Try
                _templateEditorTaskPane = Nothing
                _templateEditorControl = Nothing
            End If

            ' 创建预览回调（安全处理 chatControl 可能为 Nothing 的情况）
            Dim previewCallback As PreviewStyleCallback = Nothing
            If chatControl IsNot Nothing Then
                previewCallback = AddressOf chatControl.ApplyStylePreviewToSelection
            End If

            ' 创建占位符预览回调
            Dim placeholderPreviewCallback As TemplatePlaceholderPreviewCallback = AddressOf ApplyPlaceholderPreviewToDocument

            ' 创建新的编辑器控件
            _templateEditorControl = New ReformatTemplateEditorControl(template, previewCallback, placeholderPreviewCallback)

            ' 绑定事件
            AddHandler _templateEditorControl.TemplateSaved, Sub(s, t)
                GlobalStatusStrip.ShowInfo($"Шаблон «{t.Name}» сохранён")
                HideTemplateEditorTaskPane()
                ' 刷新前端模板列表
                If chatControl IsNot Nothing Then
                    chatControl.RefreshReformatTemplates()
                End If
            End Sub

            AddHandler _templateEditorControl.EditorClosed, Sub(s, e)
                HideTemplateEditorTaskPane()
            End Sub

            ' 创建TaskPane
            _templateEditorTaskPane = Me.CustomTaskPanes.Add(_templateEditorControl, "Редактор шаблонов форматирования")
            _templateEditorTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            _templateEditorTaskPane.Width = 380
            _templateEditorTaskPane.Visible = True

        Catch ex As Exception
            MessageBox.Show($"Ошибка при открытии редактора шаблонов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"ShowTemplateEditorTaskPane 错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 隐藏模板编辑器任务窗格
    ''' </summary>
    Public Sub HideTemplateEditorTaskPane()
        Try
            If _templateEditorTaskPane IsNot Nothing Then
                _templateEditorTaskPane.Visible = False
                Me.CustomTaskPanes.Remove(_templateEditorTaskPane)
                _templateEditorTaskPane = Nothing
                _templateEditorControl = Nothing
            End If
        Catch ex As Exception
            Debug.WriteLine($"HideTemplateEditorTaskPane 错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用占位符预览到文档
    ''' </summary>
    Private Sub ApplyPlaceholderPreviewToDocument(placeholderId As String, content As String, fontConfig As FontConfig, paragraphConfig As ParagraphConfig, colorConfig As ColorConfig)
        Try
            Dim wordApp = Me.Application
            If wordApp Is Nothing OrElse wordApp.Selection Is Nothing Then Return

            Dim selRange = wordApp.Selection.Range

            ' 清除当前选择内容
            selRange.Text = ""

            ' 应用占位符内容
            selRange.Text = content

            ' 应用字体设置
            If fontConfig IsNot Nothing Then
                If Not String.IsNullOrEmpty(fontConfig.FontNameCN) Then
                    selRange.Font.NameFarEast = fontConfig.FontNameCN
                End If
                If fontConfig.FontSize > 0 Then
                    selRange.Font.Size = CSng(fontConfig.FontSize)
                End If
                selRange.Font.Bold = If(fontConfig.Bold, 0, 0)
                selRange.Font.Italic = If(fontConfig.Italic, 0, 0)
                selRange.Font.Underline = If(fontConfig.Underline, 0, 0)
            End If

            ' 应用颜色设置
            If colorConfig IsNot Nothing AndAlso Not String.IsNullOrEmpty(colorConfig.FontColor) Then
                Try
                    Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(colorConfig.FontColor)
                    selRange.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                Catch ex As Exception
                    Debug.WriteLine($"应用颜色失败: {ex.Message}")
                End Try
            End If

            ' 应用段落设置
            If paragraphConfig IsNot Nothing Then
                Select Case paragraphConfig.Alignment?.ToLower()
                    Case "center"
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Case "right"
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    Case "justify"
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    Case Else
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End Select

                If paragraphConfig.FirstLineIndent > 0 AndAlso selRange.Font.Size > 0 Then
                    selRange.ParagraphFormat.FirstLineIndent = CSng(paragraphConfig.FirstLineIndent * selRange.Font.Size)
                End If

                If paragraphConfig.LineSpacing > 0 Then
                    selRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
                    selRange.ParagraphFormat.LineSpacing = CSng(paragraphConfig.LineSpacing * 12)
                End If
            End If

            Debug.WriteLine($"应用占位符预览: {placeholderId} -> {content}")

        Catch ex As Exception
            Debug.WriteLine($"ApplyPlaceholderPreviewToDocument 错误: {ex.Message}")
        End Try
    End Sub

End Class
