Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports ShareRibbon
Public Class ThisAddIn

    ' 按演示文稿窗口维护聊天面板：多窗口时各窗口拥有独立面板
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

    ' 在类中添加以下变量
    Private _deepseekControl As DeepseekControl
    Private _deepseekTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _doubaoControl As DoubaoChat
    Private _doubaoTaskPane As Microsoft.Office.Tools.CustomTaskPane

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

    Private Sub PowerPointAi_Startup() Handles Me.Startup
        ' Phase 0: 仅注册事件处理器（微秒级，不阻塞启动）
        PhaseStartupManager.Instance.RunCriticalPhase(Me.Application)
    End Sub

    ''' <summary>
    ''' 确保核心服务已加载（WebView2 + SQLite），首次调用时初始化
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
                'Await _doubaoControl.InitializeAsync()
                _doubaoTaskPane = Me.CustomTaskPanes.Add(_doubaoControl, "ИИ-помощник Doubao")
                _doubaoTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _doubaoTaskPane.Width = 420
            End If
        Catch ex As Exception
            MessageBox.Show($"Не удалось инициализировать панель задач Doubao: {ex.Message}")
        End Try
    End Function

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
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

    ''' <summary>返回当前活动演示文稿窗口对应的聊天面板，必要时为该窗口创建。</summary>
    Private Function EnsureChatPaneEntry() As HostTaskPaneRegistry.Entry
        Try
            ChatPanes.PruneClosed(AddressOf GetLiveWindowHandles)

            Dim window = GetActiveWindowObject()
            If window Is Nothing Then Return Nothing

            Dim entry = ChatPanes.GetOrCreate(window, "ИИ-помощник PPT", Function() New ChatControl())
            If entry Is Nothing Then Return Nothing

            entry.Pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            entry.Pane.Width = 420
            If entry.IsNew Then
                AddHandler entry.Pane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            End If
            Return entry
        Catch ex As Exception
            MessageBox.Show($"Не удалось инициализировать панель задач PPT AI: {ex.Message}")
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
        Dim docWindow = TryCast(window, PowerPoint.DocumentWindow)
        If docWindow Is Nothing Then Return 0
        Try
            Return docWindow.HWND
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>当前所有打开的演示文稿窗口句柄；整体枚举失败时返回 Nothing 以跳过清理。</summary>
    Private Function GetLiveWindowHandles() As HashSet(Of Integer)
        Dim result As New HashSet(Of Integer)()
        Try
            For Each w As PowerPoint.DocumentWindow In Me.Application.Windows
                Try
                    result.Add(w.HWND)
                Catch
                End Try
            Next
        Catch ex As Exception
            Debug.WriteLine($"[Chat] 枚举窗口失败: {ex.Message}")
            Return Nothing
        End Try
        Return result
    End Function

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

    Public Async Sub ShowChatTaskPane()
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
End Class
