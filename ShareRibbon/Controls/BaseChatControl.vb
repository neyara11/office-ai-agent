' ShareRibbon\Controls\BaseChatControl.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.JSON
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseChatControl
    Inherits UserControl
    Implements IChatRoutingHost

    ' 服务类实例
    Private _fileParserService As New FileParserService()
    Protected _chatStateService As New ChatStateService()
    Private _mcpService As McpService = Nothing
    Private _selfCheckLoopController As SelfCheckLoopController = Nothing
    Private _conversationRuntime As IConversationRuntime = Nothing
    Private ReadOnly _webViewInitSemaphore As New SemaphoreSlim(1, 1)
    Private _commandRouter As ChatCommandRouter = Nothing
    Private _webViewBridge As WebViewBridge = Nothing
    Private _systemPromptResolver As ChatSystemPromptResolver = Nothing
    Private _sendValidator As ChatSendValidator = Nothing
    Private _memoryTurnRecorder As MemoryTurnRecorder = Nothing
    Private _aiNativeRuntime As Agent.IAiNativeRuntime = Nothing
    Private _chatRequestOrchestrator As ChatRequestOrchestrator = Nothing
    Private _chatRoutingOrchestrator As ChatRoutingOrchestrator = Nothing

    Private ReadOnly Property CommandRouter As ChatCommandRouter
        Get
            If _commandRouter Is Nothing Then
                _commandRouter = New ChatCommandRouter(Sub(messageType) Debug.WriteLine($"未知消息类型: {messageType}"))
                RegisterWebViewCommandHandlers(_commandRouter)
            End If
            Return _commandRouter
        End Get
    End Property

    Private ReadOnly Property WebViewBridge As WebViewBridge
        Get
            If _webViewBridge Is Nothing Then
                _webViewBridge = New WebViewBridge(ChatBrowser, AddressOf InitializeWebView2)
            End If
            Return _webViewBridge
        End Get
    End Property

    Private ReadOnly Property SystemPromptResolver As ChatSystemPromptResolver
        Get
            If _systemPromptResolver Is Nothing Then
                _systemPromptResolver = New ChatSystemPromptResolver()
            End If
            Return _systemPromptResolver
        End Get
    End Property

    Private ReadOnly Property SendValidator As ChatSendValidator
        Get
            If _sendValidator Is Nothing Then
                _sendValidator = New ChatSendValidator()
            End If
            Return _sendValidator
        End Get
    End Property

    Private ReadOnly Property MemoryTurnRecorder As MemoryTurnRecorder
        Get
            If _memoryTurnRecorder Is Nothing Then
                _memoryTurnRecorder = New MemoryTurnRecorder()
            End If
            Return _memoryTurnRecorder
        End Get
    End Property

    Protected ReadOnly Property AiNativeRuntime As Agent.IAiNativeRuntime
        Get
            If _aiNativeRuntime Is Nothing Then
                Dim toolRegistry As New Agent.ToolRegistry()
                toolRegistry.LoadFromRuntimeDirectories()
                toolRegistry.LoadSkillScriptsAsTools()
                _aiNativeRuntime = New Agent.AiNativeRuntime(toolRegistry)
            End If
            Return _aiNativeRuntime
        End Get
    End Property

    Protected ReadOnly Property ConversationRuntime As IConversationRuntime
        Get
            If _conversationRuntime Is Nothing Then
                _conversationRuntime = New DefaultConversationRuntime(
                    New DefaultContextComposer(),
                    New McpToolBroker())
            End If
            Return _conversationRuntime
        End Get
    End Property

    Private ReadOnly Property ChatRequestOrchestrator As ChatRequestOrchestrator
        Get
            If _chatRequestOrchestrator Is Nothing Then
                _chatRequestOrchestrator = New ChatRequestOrchestrator(
                    ConversationRuntime,
                    _chatStateService,
                    systemHistoryMessageData,
                    _selectionPendingMap,
                    AddressOf GetApplication,
                    AddressOf ManageHistoryMessageSize)
            End If
            Return _chatRequestOrchestrator
        End Get
    End Property

    Private ReadOnly Property ChatRoutingOrchestrator As ChatRoutingOrchestrator
        Get
            If _chatRoutingOrchestrator Is Nothing Then
                _chatRoutingOrchestrator = New ChatRoutingOrchestrator(Me)
            End If
            Return _chatRoutingOrchestrator
        End Get
    End Property

    ' 延迟初始化的排版服务
    Private _reformatService As ReformatService = Nothing
    Protected ReadOnly Property ReformatSvc As ReformatService
        Get
            If _reformatService Is Nothing Then
                _reformatService = New ReformatService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf EscapeJavaScriptString,
                    AddressOf RunUiAction,
                    AddressOf ShowTemplateEditorPane,
                    AddressOf GetStylePreviewCallback,
                    AddressOf HandleUploadDocxTemplateFromPath,
                    Function(mode, prompt) SendAndGetResponseAsync(prompt, "", New List(Of HistoryMessage)()))
            End If
            Return _reformatService
        End Get
    End Property

    ' 延迟初始化的统一 AgentKernel 服务（新架构）
    Private _agentKernelService As AgentKernelService = Nothing
    Protected ReadOnly Property AgentKernelSvc As AgentKernelService
        Get
            If _agentKernelService Is Nothing Then
                _agentKernelService = New AgentKernelService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf EscapeJavaScriptString,
                    AddressOf SendAndGetResponseAsync,
                    Function(c, l, p) CodeExecutionService.ExecuteCodeWithToolResult(c, l, p),
                    _chatStateService,
                    systemHistoryMessageData,
                    AddressOf ManageHistoryMessageSize,
                    AddressOf GetOfficeAppType)
            End If
            Return _agentKernelService
        End Get
    End Property

    ' 延迟初始化的 MCP 服务
    Protected ReadOnly Property McpService As McpService
        Get
            If _mcpService Is Nothing Then
                _mcpService = New McpService(AddressOf ExecuteJavaScriptAsyncJS, AddressOf GetApplication)
            End If
            Return _mcpService
        End Get
    End Property

    ' 延迟初始化的自动补全服务
    Private _autocompleteService As AutocompleteService = Nothing
    Protected ReadOnly Property AutocompleteSvc As AutocompleteService
        Get
            If _autocompleteService Is Nothing Then
                _autocompleteService = New AutocompleteService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf GetContextSnapshot,
                    AddressOf GetOfficeAppType)
            End If
            Return _autocompleteService
        End Get
    End Property

    ' 延迟初始化的历史会话服务
    Private _historySessionService As HistorySessionService = Nothing
    Protected ReadOnly Property HistorySessionSvc As HistorySessionService
        Get
            If _historySessionService Is Nothing Then
                _historySessionService = New HistorySessionService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    _chatStateService,
                    AddressOf GetOfficeAppType,
                    AddressOf RunUiAction)
            End If
            Return _historySessionService
        End Get
    End Property

    ' 延迟初始化的代码执行服务
    Private _codeExecutionService As CodeExecutionService = Nothing
    Protected ReadOnly Property CodeExecutionService As CodeExecutionService
        Get
            If _codeExecutionService Is Nothing Then
                _codeExecutionService = New CodeExecutionService(
                    AddressOf GetVBProject,
                    AddressOf GetOfficeApplicationObject,
                    AddressOf GetApplication,
                    AddressOf RunCode,
                    AddressOf RunCodePreview,
                    AddressOf EvaluateFormula)
                ' Agent/Loop 可在后台线程运行。所有 native JSON 命令在唯一宿主入口
                ' 统一切回 UI/STA 线程，避免各 Office Executor 自行分散 Invoke。
                _codeExecutionService.JsonCommandExecutorWithResult =
                    Function(jsonCode As String, preview As Boolean) As Agent.ToolResult
                        Return UiDispatcher.InvokeSync(Of Agent.ToolResult)(
                            Me,
                            Function() ExecuteJsonCommandWithToolResult(jsonCode, preview))
                    End Function
            End If
            Return _codeExecutionService
        End Get
    End Property

    ' 延迟初始化的自检Loop控制器
    Protected ReadOnly Property SelfCheckLoopController As SelfCheckLoopController
        Get
            If _selfCheckLoopController Is Nothing Then
                _selfCheckLoopController = New SelfCheckLoopController(
                    New PreSendChecker(),
                    New PostFlushValidator())
            End If
            Return _selfCheckLoopController
        End Get
    End Property

    ' 延迟初始化的 HTTP 流服务
    Private _httpStreamService As HttpStreamService = Nothing
    Protected ReadOnly Property HttpStreamSvc As HttpStreamService
        Get
            If _httpStreamService Is Nothing Then
                _httpStreamService = New HttpStreamService(
                    _chatStateService,
                    AddressOf GetApplication,
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf WaitForRendererMapAsync,
                    Sub(uuid, plainBuf) CheckAndCompleteProcessingHook(uuid, plainBuf),
                    Nothing,
                    Nothing)
            End If
            Return _httpStreamService
        End Get
    End Property

    ''' <summary>
    ''' 当前流的最终响应 UUID（代理到 HttpStreamSvc，供子类访问）
    ''' </summary>
    Protected ReadOnly Property _finalUuid As String
        Get
            Return HttpStreamSvc.FinalUuid
        End Get
    End Property

    ''' <summary>
    ''' Host JSON command backend used by CodeExecutionService / Agent tools (P0-2).
    ''' Not a product entry for natural-language routing; NL goes ChatRouting → AgentKernel.
    ''' </summary>
    Protected Overridable Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
        Return Agent.ToolResult.Failed("",
                                       "Текущее приложение не поддерживает выполнение JSON-команд",
                                       errorCode:=ExceptionClassifier.CodeHostUnsupported,
                                       userMessage:="Текущее приложение не поддерживает выполнение JSON-команд",
                                       recoverable:=False)
    End Function

    ' 延迟初始化的意图识别服务
    Private _intentService As IntentRecognitionService = Nothing
    Protected ReadOnly Property IntentService As IntentRecognitionService
        Get
            If _intentService Is Nothing Then
                ' 根据当前Office应用类型初始化意图识别服务
                Dim appInfo = GetApplication()
                If appInfo IsNot Nothing Then
                    _intentService = New IntentRecognitionService(appInfo.Type)
                Else
                    _intentService = New IntentRecognitionService()
                End If
            End If
            Return _intentService
        End Get
    End Property

    ' 当前意图结果（用于子类访问）
    Protected CurrentIntentResult As IntentResult = Nothing

    'settings
    Protected topicRandomness As Double
    Protected contextLimit As Integer
    Protected selectedCellChecked As Boolean = False
    Protected settingsScrollChecked As Boolean = False

    Protected stopReaderStream As Boolean = False


    ' ai的历史回复
    Protected systemHistoryMessageData As New List(Of HistoryMessage)

    Protected loadingPictureBox As PictureBox

    ' 选区对比相关字段
    Protected PendingSelectionInfo As SelectionInfo = Nothing

    ' 以下属性代理到 _chatStateService 的内部字典，保持子类兼容性
    Protected ReadOnly Property _selectionPendingMap As Dictionary(Of String, SelectionInfo)
        Get
            Return _chatStateService.SelectionPendingMap
        End Get
    End Property

    Private ReadOnly Property allPlainMarkdownBuffer As StringBuilder
        Get
            Return _chatStateService.PlainMarkdownBuffer
        End Get
    End Property

    Protected ReadOnly Property _responseToRequestMap As Dictionary(Of String, String)
        Get
            Return _chatStateService.ResponseToRequestMap
        End Get
    End Property

    Protected _revisionsMap As New Dictionary(Of String, JArray)()

    Private Async Sub RunUiAction(action As Action)
        Try
            Await UiDispatcher.InvokeAsync(Me, action)
        Catch ex As Exception
            Debug.WriteLine($"[BaseChatControl] UI action dispatch failed: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Background → UI marshal without async GetResult (P0-3). Prefer RunUiAction/async when possible.
    ''' </summary>
    Private Sub RunUiActionSync(action As Action)
        If action Is Nothing Then Return
        UiDispatcher.InvokeSync(Me, action)
    End Sub

    Protected Async Function InitializeWebView2() As Task
        ' 确保在 UI 线程上执行
        If ChatBrowser.InvokeRequired Then
            Await UiDispatcher.InvokeAsync(ChatBrowser,
                Async Function()
                    Await InitializeWebView2()
                End Function)
            Return
        End If

        Await _webViewInitSemaphore.WaitAsync()
        Try
            ' 安全检查 CoreWebView2（必须在 UI 线程）
            Dim isInitialized As Boolean = False
            Try
                isInitialized = (ChatBrowser.CoreWebView2 IsNot Nothing)
            Catch ex As InvalidOperationException
                ' CoreWebView2 只能在 UI 线程访问，如果出现此异常，重新调度到 UI 线程
                Debug.WriteLine("[InitializeWebView2] 检测到线程问题，重新调度到 UI 线程")
                _webViewInitSemaphore.Release()

                ' 使用 Control.Invoke 强制在 UI 线程执行
                If ChatBrowser.InvokeRequired Then
                    ChatBrowser.Invoke(New Action(Async Sub()
                        Await InitializeWebView2()
                    End Sub))
                    Return
                End If
                Throw
            End Try

            If isInitialized Then
                Return
            End If

            Dim userDataFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyAppWebView2Cache")
            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            Dim wwwRoot As String = ResourceExtractor.ExtractResources()
            
            ' 检查资源提取是否成功
            If String.IsNullOrEmpty(wwwRoot) Then
                Dim errMsg As String = "Не удалось извлечь ресурсы, невозможно инициализировать интерфейс чата." & Environment.NewLine
                If Not String.IsNullOrEmpty(ResourceExtractor.LastError) Then
                    errMsg &= "Подробности ошибки: " & ResourceExtractor.LastError
                End If
                MessageBox.Show(errMsg, "Ошибка инициализации", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            
            ' 检查目录是否存在
            If Not Directory.Exists(wwwRoot) Then
                MessageBox.Show($"Каталог ресурсов не найден: {wwwRoot}", "Ошибка инициализации", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            
            ' 诊断：显示资源目录
            Debug.WriteLine($"[WebView2] wwwRoot = {wwwRoot}")

            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
                .UserDataFolder = userDataFolder
            }
            Dim env = Await WebView2EnvironmentCache.GetOrCreateAsync(userDataFolder)
            Await ChatBrowser.EnsureCoreWebView2Async(env)

            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
#If DEBUG Then
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True
#Else
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = False
#End If

                ' 设置虚拟主机映射
                Try
                    ChatBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "officeai.local",
                        wwwRoot,
                        CoreWebView2HostResourceAccessKind.Allow
                    )
                    Debug.WriteLine($"[WebView2] 虚拟主机映射成功: officeai.local -> {wwwRoot}")
                Catch ex As Exception
                    Debug.WriteLine($"[WebView2] 虚拟主机映射失败: {ex.Message}")
                    MessageBox.Show($"Не удалось сопоставить виртуальный хост: {ex.Message}", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

                ' 添加 WebResourceRequested 事件处理作为备用方案
                AddHandler ChatBrowser.CoreWebView2.WebResourceRequested, AddressOf OnWebResourceRequested
                ChatBrowser.CoreWebView2.AddWebResourceRequestedFilter("http://officeai.local/*", CoreWebView2WebResourceContext.Script)
                ChatBrowser.CoreWebView2.AddWebResourceRequestedFilter("http://officeai.local/*", CoreWebView2WebResourceContext.Stylesheet)
                ChatBrowser.CoreWebView2.AddWebResourceRequestedFilter("http://officeai.local/*", CoreWebView2WebResourceContext.Image)
                
                Dim htmlContent As String = My.Resources.chat_template_refactored
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)
                
                ' 等待页面加载完成后设置应用名称
                AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                
                ' 配置 Markdown 解析器
                Await ConfigureMarked()
                
                ' 设置应用名称的延迟调用
                Await Task.Delay(500) ' 等待一点时间确保页面加载
                Await SetCurrentOfficeAppName()

            Else
                MessageBox.Show("Не удалось инициализировать WebView2: CoreWebView2 недоступен.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"Ошибка инициализации: {ex.Message}{Environment.NewLine}Тип: {ex.GetType().Name}{Environment.NewLine}Стек: {ex.StackTrace}"
            MessageBox.Show(errorMessage, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _webViewInitSemaphore.Release()
        End Try
    End Function
    
    ''' <summary>
    ''' WebResourceRequested 事件处理 - 作为虚拟主机映射的备用方案
    ''' </summary>
    Private Sub OnWebResourceRequested(sender As Object, e As CoreWebView2WebResourceRequestedEventArgs)
        Try
            Dim requestUri As Uri = New Uri(e.Request.Uri)
            Dim localPath As String = requestUri.AbsolutePath.TrimStart("/"c)
            
            ' 映射到本地资源目录
            Dim wwwRoot As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAI", "www"
            )
            Dim filePath As String = Path.Combine(wwwRoot, localPath.Replace("/"c, Path.DirectorySeparatorChar))
            
            If File.Exists(filePath) Then
                Dim mimeType As String = GetMimeType(filePath)
                Dim fileBytes As Byte() = File.ReadAllBytes(filePath)
                
                        Dim response = ChatBrowser.CoreWebView2.Environment.CreateWebResourceResponse(
                            New MemoryStream(fileBytes),
                            200,
                            "OK",
                            $"Content-Type: {mimeType}{Environment.NewLine}Cache-Control: no-store, no-cache, must-revalidate{Environment.NewLine}Access-Control-Allow-Origin: *"
                        )
                e.Response = response
                Debug.WriteLine($"[WebView2] 资源加载成功: {localPath}")
            Else
                Debug.WriteLine($"[WebView2] 资源未找到: {filePath}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[WebView2] WebResourceRequested 错误: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 获取文件的 MIME 类型
    ''' </summary>
    Private Function GetMimeType(filePath As String) As String
        Dim ext As String = Path.GetExtension(filePath).ToLower()
        Select Case ext
            Case ".js" : Return "application/javascript"
            Case ".css" : Return "text/css"
            Case ".html" : Return "text/html"
            Case ".png" : Return "image/png"
            Case ".jpg", ".jpeg" : Return "image/jpeg"
            Case ".gif" : Return "image/gif"
            Case ".svg" : Return "image/svg+xml"
            Case ".json" : Return "application/json"
            Case Else : Return "application/octet-stream"
        End Select
    End Function

    ''' <summary>
    ''' 设置当前 Office 应用名称
    ''' </summary>
    Private Async Function SetCurrentOfficeAppName() As Task
        Try
            ' 获取当前应用名称
            Dim appName As String = GetOfficeApplicationName()
            
            ' 向网页注入应用名称
            Dim script As String = $"window.currentOfficeAppName = '{appName}';"
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            
        Catch ex As Exception
            Debug.WriteLine($"设置应用名称失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overridable Function GetOfficeApplicationName() As String
        ' 默认返回 "当前应用"，子类应重写此方法
        Return "Текущее приложение"
    End Function
    Private Async Function ConfigureMarked() As Task
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim script = "
            marked.setOptions({
                highlight: function (code, lang) {
                    if (hljs.getLanguage(lang)) {
                        return hljs.highlight(lang, code).value;
                    } else {
                        return hljs.highlightAuto(code).value;
                    }
                }
            });
        "
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
        Else
            MessageBox.Show("CoreWebView2 не инициализирован, невозможно настроить Marked.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Function


    ' 动态ChatHtmlFilePath属性
    Protected ReadOnly Property ChatHtmlFilePath As String
        Get
            ' 如果已经生成过文件路径，直接返回缓存的路径
            If Not String.IsNullOrEmpty(_chatHtmlFilePath) Then
                Return _chatHtmlFilePath
            End If

            Dim baseDir As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder
        )

            Dim fileName As String
            If Not String.IsNullOrEmpty(firstQuestion) Then
                ' 简单地取前10个字符
                Dim questionPrefix As String = UtilsService.GetFirst10Characters(firstQuestion)
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}_{questionPrefix}.html"
            Else
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
            End If

            _chatHtmlFilePath = Path.Combine(baseDir, fileName)
            Return _chatHtmlFilePath
        End Get
    End Property

    Private Async Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles ChatBrowser.NavigationCompleted
        If e.IsSuccess Then
            Try
                ' Async Sub中Await后VSTO同步上下文可能丢失，
                ' 统一通过Invoke确保CoreWebView2访问在UI线程
                Await Task.Delay(100)
                Await UiDispatcher.InvokeAsync(ChatBrowser,
                    Sub()
                        InitializeSettings()
                        InitializeMcpSettings()

                        If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                            RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                        End If
                    End Sub)
            Catch ex As Exception
                Debug.WriteLine($"导航完成事件处理中出错: {ex.Message}")
                Debug.WriteLine(ex.StackTrace)
            End Try
        End If
    End Sub

    Protected Sub InitializeSettings()
        Try
            ' 确保记忆功能配置是开启的（默认）
            Debug.WriteLine($"[InitializeSettings] 初始化记忆配置...")
            Debug.WriteLine($"[InitializeSettings] UseContextBuilder: {MemoryConfig.UseContextBuilder}")
            Debug.WriteLine($"[InitializeSettings] RagTopN: {MemoryConfig.RagTopN}")
            Debug.WriteLine($"[InitializeSettings] EnableUserProfile: {MemoryConfig.EnableUserProfile}")
            
            ' 如果 UseContextBuilder 是关闭的，我们强制开启它
            If Not MemoryConfig.UseContextBuilder Then
                Debug.WriteLine($"[InitializeSettings] UseContextBuilder 未开启，强制开启...")
                MemoryConfig.UseContextBuilder = True
            End If
            
            ' 确保 RagTopN 至少是 5
            If MemoryConfig.RagTopN < 3 Then
                MemoryConfig.RagTopN = 5
            End If
            
            Debug.WriteLine($"[InitializeSettings] 记忆配置已确保开启: UseContextBuilder={MemoryConfig.UseContextBuilder}, RagTopN={MemoryConfig.RagTopN}")

            ' 加载设置
            Dim chatSettings As New ChatSettings(GetApplication())
            selectedCellChecked = ChatSettings.selectedCellChecked
            contextLimit = ChatSettings.contextLimit
            topicRandomness = ChatSettings.topicRandomness
            settingsScrollChecked = ChatSettings.settingsScrollChecked

            ' 设置Office应用类型（用于前端区分Word/PPT/Excel）
            Dim appType = GetOfficeAppType()

            ' 将设置发送到前端
            Dim js As String = $"
            window.officeAppType = '{appType}';
            document.getElementById('topic-randomness').value = '{ChatSettings.topicRandomness}';
            document.getElementById('topic-randomness-value').textContent = '{ChatSettings.topicRandomness}';
            document.getElementById('context-limit').value = '{ChatSettings.contextLimit}';
            document.getElementById('context-limit-value').textContent = '{ChatSettings.contextLimit}';
            document.getElementById('settings-scroll-checked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('settings-selected-cell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
            document.getElementById('settings-executecode-preview').checked = {ChatSettings.executecodePreviewChecked.ToString().ToLower()};
            
            // 初始化自动补全设置
            var autocompleteCheckbox = document.getElementById('settings-autocomplete-enable');
            if (autocompleteCheckbox) {{
                autocompleteCheckbox.checked = {ChatSettings.EnableAutocomplete.ToString().ToLower()};
            }}
            var shortcutSelect = document.getElementById('settings-autocomplete-shortcut');
            if (shortcutSelect) {{
                shortcutSelect.value = '{ChatSettings.AutocompleteShortcut}';
            }}
            if (typeof updateAutocompleteSettings === 'function') {{
                updateAutocompleteSettings({{ enabled: {ChatSettings.EnableAutocomplete.ToString().ToLower()}, delayMs: {ChatSettings.AutocompleteDelayMs}, shortcut: '{ChatSettings.AutocompleteShortcut}' }});
            }}
            
            // 同步到主界面的checkbox
            document.getElementById('scrollChecked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('selectedCell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
        "
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"初始化设置失败: {ex.Message}")
        End Try
    End Sub

    Private Sub RegisterWebViewCommandHandlers(router As ChatCommandRouter)
        router.Register("checkedChange", Sub(jsonDoc) HandleCheckedChange(jsonDoc))
        router.Register("sendMessage", Sub(jsonDoc) HandleSendMessage(jsonDoc))
        router.Register("stopMessage", Sub(jsonDoc) HandleStopMessage(jsonDoc))
        router.Register("executeCode", Sub(jsonDoc) HandleExecuteCode(jsonDoc))
        router.Register("saveSettings", Sub(jsonDoc) HandleSaveSettings(jsonDoc))
        router.Register("getSessionList", Sub(jsonDoc) HandleGetSessionList())
        router.Register("loadSession", Sub(jsonDoc) HandleLoadSession(jsonDoc))
        router.Register("newSession", Sub(jsonDoc) HandleNewSession())
        router.Register("getMcpConnections", Sub(jsonDoc) HandleGetMcpConnections())
        router.Register("saveMcpSettings", Sub(jsonDoc) HandleSaveMcpSettings(jsonDoc))
        router.Register("clearContext", Sub(jsonDoc) ClearChatContext())
        router.Register("acceptAnswer", Sub(jsonDoc) HandleAcceptAnswer(jsonDoc))
        router.Register("rejectAnswer", Sub(jsonDoc) HandleRejectAnswer(jsonDoc))
        router.Register("applyRevisionSegment", Sub(jsonDoc) HandleApplyRevisionSegment(jsonDoc))
        router.Register("applyDocumentPlanItem", Sub(jsonDoc) HandleApplyDocumentPlanItem(jsonDoc))
        router.Register("retryReformat", Sub(jsonDoc) HandleRetryReformat(jsonDoc))
        router.Register("applyRevisionAccept", Sub(jsonDoc) HandleApplyRevisionAccept(jsonDoc))
        router.Register("triggerContinuation", Sub(jsonDoc) HandleTriggerContinuation(jsonDoc))
        router.Register("applyContinuation", Sub(jsonDoc) HandleApplyContinuation(jsonDoc))
        router.Register("refineContinuation", Sub(jsonDoc) HandleRefineContinuation(jsonDoc))
        router.Register("applyTemplateContent", Sub(jsonDoc) HandleApplyTemplateContent(jsonDoc))
        router.Register("refineTemplateContent", Sub(jsonDoc) HandleRefineTemplateContent(jsonDoc))
        router.Register("requestCompletion", Sub(jsonDoc) HandleRequestCompletion(jsonDoc))
        router.Register("acceptCompletion", Sub(jsonDoc) HandleAcceptCompletion(jsonDoc))
        router.Register("startAgent", Sub(jsonDoc) HandleStartAgent(jsonDoc))
        router.Register("startAgentExecution", Sub(jsonDoc) HandleStartAgentExecution(jsonDoc))
        router.Register("abortAgent", Sub(jsonDoc) HandleAbortAgent())
        router.Register("refineAgentPlan", Sub(jsonDoc) HandleRefineAgentPlan(jsonDoc))
        ' Legacy Ralph protocol → AgentKernel compatibility shims (do not add new features here).
        router.Register("startLoop", Sub(jsonDoc) HandleLegacyStartLoop(jsonDoc))
        router.Register("continueLoop", Sub(jsonDoc) HandleStartAgentExecution(jsonDoc))
        router.Register("replanLoop", Sub(jsonDoc) HandleAgentRefinePlan(jsonDoc))
        router.Register("cancelLoop", Sub(jsonDoc) HandleAbortAgent())
        router.Register("agent:approvePlan", Sub(jsonDoc) HandleAgentApprovePlan(jsonDoc))
        router.Register("agent:rejectPlan", Sub(jsonDoc) HandleAgentRejectPlan(jsonDoc))
        router.Register("agent:approveStep", Sub(jsonDoc) HandleAgentApproveStep(jsonDoc))
        router.Register("agent:refinePlan", Sub(jsonDoc) HandleAgentRefinePlan(jsonDoc))
        router.Register("openFileDialog", Sub(jsonDoc) HandleOpenFileDialog())
        router.Register("openApiConfigForm", Sub(jsonDoc) HandleOpenApiConfigForm())
        router.Register("getCurrentAppInfo", Sub(jsonDoc) HandleGetCurrentAppInfo())
        router.Register("getCurrentModel", Sub(jsonDoc) HandleGetCurrentModel())
        router.Register("getReformatTemplates", Sub(jsonDoc) HandleGetReformatTemplates())
        router.Register("useReformatTemplate", Sub(jsonDoc) HandleUseReformatTemplate(jsonDoc))
        router.Register("previewTemplateInWord", Sub(jsonDoc) HandlePreviewTemplateInWord(jsonDoc))
        router.Register("saveCurrentDocumentAsTemplate", Sub(jsonDoc) HandleSaveCurrentDocumentAsTemplate())
        router.Register("importTemplate", Sub(jsonDoc) HandleImportTemplate())
        router.Register("exportTemplate", Sub(jsonDoc) HandleExportTemplate(jsonDoc))
        router.Register("duplicateTemplate", Sub(jsonDoc) HandleDuplicateTemplate(jsonDoc))
        router.Register("deleteTemplate", Sub(jsonDoc) HandleDeleteTemplate(jsonDoc))
        router.Register("openTemplateEditor", Sub(jsonDoc) HandleOpenTemplateEditor(jsonDoc))
        router.Register("getStyleGuides", Sub(jsonDoc) HandleGetStyleGuides())
        router.Register("useStyleGuide", Sub(jsonDoc) HandleUseStyleGuide(jsonDoc))
        router.Register("uploadStyleGuideDocument", Sub(jsonDoc) HandleUploadStyleGuideDocument())
        router.Register("deleteStyleGuide", Sub(jsonDoc) HandleDeleteStyleGuide(jsonDoc))
        router.Register("updateStyleGuide", Sub(jsonDoc) HandleUpdateStyleGuide(jsonDoc))
        router.Register("duplicateStyleGuide", Sub(jsonDoc) HandleDuplicateStyleGuide(jsonDoc))
        router.Register("exportStyleGuide", Sub(jsonDoc) HandleExportStyleGuide(jsonDoc))
        router.Register("uploadDocxTemplate", Sub(jsonDoc) HandleUploadDocxTemplate())
        router.Register("deleteDocxMapping", Sub(jsonDoc) HandleDeleteDocxMapping(jsonDoc))
        router.Register("saveAiTemplate", Sub(jsonDoc) HandleSaveAiTemplate(jsonDoc))
        router.Register("previewAiTemplate", Sub(jsonDoc) HandlePreviewAiTemplate(jsonDoc))
        router.Register("applySmartReformat", Sub(jsonDoc) HandleApplySmartReformat(jsonDoc))
        router.Register("undoReformat", Sub(jsonDoc) HandleUndoReformat(jsonDoc))
        router.Register("refineSmartReformat", Sub(jsonDoc) HandleRefineSmartReformat(jsonDoc))
        router.Register("switchReformatTemplate", Sub(jsonDoc) HandleSwitchReformatTemplate(jsonDoc))
        router.Register("previewReformatCompare", Sub(jsonDoc) HandlePreviewReformatCompare(jsonDoc))
        router.Register("proofread", Sub(jsonDoc) HandleProofreadFocusMode(jsonDoc))
    End Sub

    Private Sub HandleStopMessage(jsonDoc As JObject)
        stopReaderStream = True

        Dim requestUuid As String = ""
        Try
            requestUuid = If(jsonDoc("requestUuid")?.ToString(), "")
        Catch
            requestUuid = ""
        End Try

        HttpStreamSvc.CancelRequest(requestUuid)
    End Sub

    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim rawJson As String = e.WebMessageAsJson
            Dim jsonDoc As JObject = JObject.Parse(rawJson)
            CommandRouter.Dispatch(jsonDoc)
            Return

        Catch ex As Exception
            Debug.WriteLine($"[DEBUG WebMessageReceived] StackTrace: {ex.StackTrace}")
        End Try
    End Sub

    ' 添加：在基类提供可覆盖的 CaptureCurrentSelectionInfo（默认返回 Nothing，Word 子类会覆写）
    Protected Overridable Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Return Nothing
    End Function


    Protected Overridable Sub HandleApplyRevisionSegment(jsonDoc As JObject)
    End Sub

    Protected Overridable Sub HandleApplyRevisionAccept(jsonDoc As JObject)
    End Sub


    ' ========== 续写功能相关方法 ==========

    ''' <summary>
    ''' 触发续写（由子类实现具体逻辑）
    ''' </summary>
    ''' <param name="jsonDoc">包含style参数的JSON对象</param>
    Protected Overridable Sub HandleTriggerContinuation(jsonDoc As JObject)
        Debug.WriteLine("HandleTriggerContinuation 被调用（基类默认不执行）")
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает продолжение текста")
    End Sub

    ''' <summary>
    ''' 应用续写结果到文档（由子类实现）
    ''' </summary>
    Protected Overridable Sub HandleApplyContinuation(jsonDoc As JObject)
        Debug.WriteLine("HandleApplyContinuation 被调用（基类默认不执行）")
    End Sub

    ''' <summary>
    ''' 调整续写（多轮对话）
    ''' </summary>
    Protected Overridable Sub HandleRefineContinuation(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim refinement As String = If(jsonDoc("refinement") IsNot Nothing, jsonDoc("refinement").ToString(), String.Empty)

            If String.IsNullOrWhiteSpace(refinement) Then
                GlobalStatusStrip.ShowWarning("Введите направление правки")
                Return
            End If

            ' 构建调整提示
            Dim refinementPrompt As New StringBuilder()
            refinementPrompt.AppendLine("Скорректируй предыдущее продолжение текста с учётом следующих требований:")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine($"【Требования к правке】{refinement}")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("Выведи только скорректированное продолжение текста, без объяснений:")

            ' 保持 responseMode = "continuation"，发送调整请求（不使用历史记录）
            Task.Run(Async Function()
                         Await Send(refinementPrompt.ToString(), GetContinuationSystemPrompt(), False, "continuation")
                     End Function)

            GlobalStatusStrip.ShowInfo("Настраиваю продолжение текста...")
        Catch ex As Exception
            Debug.WriteLine($"HandleRefineContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("Ошибка при настройке продолжения текста")
        End Try
    End Sub

    ''' <summary>
    ''' 发送续写请求（不使用聊天历史记录）
    ''' </summary>
    Protected Sub SendContinuationRequest(context As ContinuationContext, Optional style As String = "")
        Dim systemPrompt = GetContinuationSystemPrompt()
        Dim userPrompt = BuildContinuationUserPrompt(context, style)

        Task.Run(Async Function()
                     Await Send(userPrompt, systemPrompt, False, "continuation")
                 End Function)
    End Sub

    ''' <summary>
    ''' 获取续写的系统提示词
    ''' </summary>
    Protected Function GetContinuationSystemPrompt() As String
        Return "Ты профессиональный помощник по письму. На основе предоставленного контекста естественно продолжи текст. Требования:
1. Сохраняй языковой стиль, тон и терминологию исходного текста
2. Содержание должно быть связным и естественным, не повторяй уже написанное
3. Выводи только продолжение текста, без объяснений, префиксов и пометок
4. При недостатке контекста можешь разумно домыслить, но осторожно
5. Оптимальная длина продолжения — около 100–300 символов, если пользователь не просил иначе

Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть."
    End Function

    ''' <summary>
    ''' 构建续写请求的用户提示
    ''' </summary>
    Protected Function BuildContinuationUserPrompt(context As ContinuationContext, Optional style As String = "") As String
        Dim sb As New StringBuilder()

        sb.AppendLine("Продолжи текст на основе следующего контекста:")
        sb.AppendLine()
        sb.Append(context.BuildPrompt())

        If Not String.IsNullOrWhiteSpace(style) Then
            sb.AppendLine()
            sb.AppendLine($"【Требования к стилю】{style}")
        End If

        sb.AppendLine()
        sb.AppendLine("Выведи только продолжение текста, без префиксов и пояснений:")

        Return sb.ToString()
    End Function

    ' ========== 续写功能相关方法结束 ==========

    ' ========== 模板渲染功能相关方法 ==========

    ''' <summary>
    ''' 应用模板渲染结果到文档（由子类实现）
    ''' </summary>
    Protected Overridable Sub HandleApplyTemplateContent(jsonDoc As JObject)
        Debug.WriteLine("HandleApplyTemplateContent 被调用（基类默认不执行）")
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает рендеринг шаблонов")
    End Sub

    ''' <summary>
    ''' 调整模板渲染需求（多轮对话）
    ''' </summary>
    Protected Overridable Sub HandleRefineTemplateContent(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim refinement As String = If(jsonDoc("refinement") IsNot Nothing, jsonDoc("refinement").ToString(), String.Empty)

            If String.IsNullOrWhiteSpace(refinement) Then
                GlobalStatusStrip.ShowWarning("Введите требования к правке")
                Return
            End If

            ' 构建调整提示
            Dim refinementPrompt As New StringBuilder()
            refinementPrompt.AppendLine("Скорректируй ранее сгенерированное содержимое шаблона с учётом следующих требований:")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine($"【Требования к правке】{refinement}")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("Выведи только скорректированное содержимое, без объяснений:")

            ' 保持 responseMode = "template_render"，发送调整请求（不使用历史记录）
            Task.Run(Async Function()
                         Await Send(refinementPrompt.ToString(), GetTemplateRenderSystemPrompt(""), False, "template_render")
                     End Function)

            GlobalStatusStrip.ShowInfo("Настраиваю содержимое шаблона...")
        Catch ex As Exception
            Debug.WriteLine($"HandleRefineTemplateContent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("Ошибка при настройке содержимого шаблона")
        End Try
    End Sub

    ''' <summary>
    ''' 获取模板渲染的系统提示词
    ''' </summary>
    Protected Function GetTemplateRenderSystemPrompt(templateContext As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("Ты профессиональный помощник по генерации содержимого документов. Тебе нужно создавать новое содержимое на основе предоставленной пользователем структуры шаблона (формат JSON) и стиля.")
        sb.AppendLine()
        sb.AppendLine("【Важные требования к формату】")
        sb.AppendLine("- Строго запрещено использовать формат блоков кода Markdown (символы ```)")
        sb.AppendLine("- Строго запрещено использовать любые разметки Markdown (например, #, **, -, > и т. п.)")
        sb.AppendLine("- Выводи содержимое простым текстом, не оборачивай в блоки кода")
        sb.AppendLine("- Не добавляй префиксы, суффиксы, пояснения или описания")
        sb.AppendLine("- Не выводи JSON, выводи простой текст, готовый для вставки в документ")
        sb.AppendLine()
        sb.AppendLine("【Описание структуры JSON шаблона】")
        sb.AppendLine("Шаблон предоставляется в формате JSON и содержит следующую информацию:")
        sb.AppendLine("- elements: массив элементов документа; каждый элемент содержит type (тип), text (текст), styleName (имя стиля), formatting (детали форматирования)")
        sb.AppendLine("- formatting содержит: fontName (шрифт), fontSize (размер шрифта), bold (полужирный), italic (курсив), alignment (выравнивание) и т. д.")
        sb.AppendLine("- Для шаблонов PPT: массив slides содержит макет и элементы каждого слайда")
        sb.AppendLine()
        sb.AppendLine("【Требования к генерации содержимого】")
        sb.AppendLine("1. Строго соблюдай иерархическую структуру шаблона (например, соотношение заголовка, подзаголовка и основного текста)")
        sb.AppendLine("2. Сохраняй тон, терминологию и стиль, соответствующие шаблону")
        sb.AppendLine("3. Ориентируйся на размер шрифта в шаблоне для оценки важности содержимого (крупный = заголовок, мелкий = основной текст)")
        sb.AppendLine("4. Содержимое должно быть профессиональным, связным и соответствовать реальному сценарию использования")
        sb.AppendLine("5. Организуй вывод в порядке элементов шаблона")
        sb.AppendLine("6. Разделяй содержимое абзацев или слайдов пустой строкой")

        If Not String.IsNullOrWhiteSpace(templateContext) Then
            sb.AppendLine()
            sb.AppendLine("【Структура шаблона для справки】")
            sb.AppendLine("```json")
            sb.AppendLine(templateContext)
            sb.AppendLine("```")
            sb.AppendLine()
            sb.AppendLine("На основе приведённой структуры шаблона сгенерируй содержимое документа в соответствующем формате согласно запросу пользователя. Выводи простой текст, не используй разметку Markdown.")
        End If

        sb.AppendLine()
        sb.AppendLine("Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть.")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建校对模式追问的系统提示词
    ''' 当用户在校对模式下发送消息时，让AI理解当前是校对上下文
    ''' </summary>
    Private Function BuildProofreadFollowUpPrompt(selectedText As String, issueCount As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("Ты профессиональный корректор документов. Пользователь сейчас работает в режиме вычитки и выделил текст для проверки.")
        sb.AppendLine()
        sb.AppendLine("【Текущий контекст вычитки】")

        If Not String.IsNullOrWhiteSpace(selectedText) Then
            sb.AppendLine("Выделенный пользователем текст (фрагмент):")
            sb.AppendLine(selectedText)
            sb.AppendLine()
        End If

        Dim count As Integer
        If Integer.TryParse(issueCount, count) AndAlso count > 0 Then
            sb.AppendLine($"ИИ обнаружил {count} проблем вычитки; список проблем отображается на панели вычитки справа.")
            sb.AppendLine()
        End If

        sb.AppendLine("【Правила взаимодействия】")
        sb.AppendLine("1. Сообщение пользователя отправлено в режиме вычитки: это может быть уточнение по результатам, просьба объяснить или повторно проверить")
        sb.AppendLine("2. Если пользователь спрашивает о конкретной проблеме вычитки, дай профессиональное объяснение на основе контекста выделенного текста")
        sb.AppendLine("3. Если пользователь просит дополнительно проверить какие-то аспекты (например, имена собственные, согласованность данных), дай конкретные рекомендации")
        sb.AppendLine("4. Отвечай кратко и профессионально, не повторяй проблемы, уже отображённые на панели вычитки")
        sb.AppendLine("5. Строго запрещено использовать блоки кода Markdown (символы ```), выводи простой текст")

        sb.AppendLine()
        sb.AppendLine("Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть.")

        Return sb.ToString()
    End Function

    ' ========== 模板渲染功能相关方法结束 ==========

    ' ========== 自动补全功能相关方法 ==========

    ''' <summary>
    ''' 处理自动补全请求
    ''' </summary>
    Protected Overridable Async Sub HandleRequestCompletion(jsonDoc As JObject)
        Await AutocompleteSvc.HandleRequestCompletion(jsonDoc)
    End Sub

    ''' <summary>
    ''' 处理补全采纳记录
    ''' </summary>
    Protected Overridable Sub HandleAcceptCompletion(jsonDoc As JObject)
        AutocompleteSvc.HandleAcceptCompletion(jsonDoc)
    End Sub

    ''' <summary>
    ''' 获取当前Office上下文快照（由子类重写提供具体实现）
    ''' </summary>
    Protected Overridable Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = GetOfficeAppType()
        snapshot("selection") = ""
        Return snapshot
    End Function

    ''' <summary>
    ''' 为意图识别/规划阶段丰富上下文：内容区引用摘要 + RAG 相关记忆
    ''' </summary>
    Protected Sub EnrichContextForIntent(snapshot As JObject,
                                        question As String,
                                        filePaths As List(Of String),
                                        selectedContents As List(Of SendMessageReferenceContentItem))
        AutocompleteSvc.EnrichContextForIntent(snapshot, question, filePaths, selectedContents)
    End Sub

    ' ========== 已移入 AutocompleteService ==========
    ' RequestCompletionsFromLLM, RequestCompletionsWithFIM, RequestCompletionsWithChat
    ' RecordCompletionHistory, GetCompletionSystemPrompt


    ' ========== 自动补全功能相关方法结束 ==========

    ' 新增：处理用户接受答案（收藏回答时更新 conversation.is_collected）
    Protected Sub HandleAcceptAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)

            Debug.WriteLine($"用户接受回答: UUID={uuid}")
            GlobalStatusStrip.ShowInfo("Пользователь принял ответ ИИ")

            ' 更新 conversation 表收藏状态
            Dim sid = _chatStateService.CurrentSessionId
            If Not String.IsNullOrEmpty(sid) Then
                Try
                    ConversationRepository.SetLastAssistantCollected(sid, True)
                Catch ex As Exception
                    Debug.WriteLine($"SetLastAssistantCollected 失败: {ex.Message}")
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleAcceptAnswer 出错: {ex.Message}")
        End Try
    End Sub

    ' 新增：处理用户拒绝答案并发起改进
    Protected Sub HandleRejectAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim rejectedContent As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim reason As String = If(jsonDoc("reason") IsNot Nothing, jsonDoc("reason").ToString(), String.Empty)

            Debug.WriteLine($"用户拒绝回答: UUID={uuid}; reason={reason}")


            ' 构建用于改进的大模型提示（包含用户理由）
            Dim refinementPrompt As New StringBuilder()
            refinementPrompt.AppendLine("Пользователь отметил предыдущий ответ как неприемлемый. Улучши его на основе истории текущей сессии и отклонённого ответа ниже:")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("【Пожелания пользователя по улучшению】")
            If Not String.IsNullOrWhiteSpace(reason) Then
                refinementPrompt.AppendLine(reason)
            Else
                refinementPrompt.AppendLine("[Конкретных пожеланий нет, пользователь просто отметил ответ как неприемлемый]")
            End If
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("Верни ответ в следующем формате:")
            refinementPrompt.AppendLine("1) Что улучшено (1–3 строки): как исправлено;")
            refinementPrompt.AppendLine("2) Plan: кратко перечисли шаги исправления (пунктами, максимум 6);")
            refinementPrompt.AppendLine("3) Answer: дай исправленный, максимально точный ответ (с разметкой Markdown, при необходимости с примерами/кодом);")
            refinementPrompt.AppendLine("4) Clarifying Questions: если нужно больше информации, перечисли короткие вопросы в конце и приостанови выполнение;")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("[Внимание]: ответ должен быть кратким и проверяемым; в первую очередь дай готовые к выполнению выводы и способы проверки, не повторяй длинные пояснения.")

            ' 管理历史大小，保证不会无限增长
            ManageHistoryMessageSize()

            ' 将该改进请求当作新的用户问题发起（会走你已有的 SendChatMessage 流程）
            SendChatMessage(refinementPrompt.ToString())

            GlobalStatusStrip.ShowInfo("Запрос на улучшение отправлен, начинаю новый раунд улучшения")
        Catch ex As Exception
            Debug.WriteLine($"HandleRejectAnswer 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("Ошибка при отправке запроса на улучшение")
        End Try
    End Sub

    Private Sub ClearChatContext()
        systemHistoryMessageData.Clear()
        _chatStateService.StartNewSession()
        Debug.WriteLine("已清空聊天记忆（上下文）")
    End Sub

    ' 处理获取MCP连接列表请求 - 委托给 McpService
    Protected Sub HandleGetMcpConnections()
        McpService.GetMcpConnections()
    End Sub

    ' 处理保存MCP设置请求 - 委托给 McpService
    Protected Sub HandleSaveMcpSettings(jsonDoc As JObject)
        McpService.SaveMcpSettings(jsonDoc)
    End Sub

    ' MCP初始化方法 - 委托给 McpService
    Protected Sub InitializeMcpSettings()
        McpService.InitializeMcpSettings()
    End Sub

    Protected Sub HandleGetSessionList()
        HistorySessionSvc.HandleGetSessionList()
    End Sub

    Protected Sub HandleLoadSession(jsonDoc As JObject)
        HistorySessionSvc.HandleLoadSession(jsonDoc)
    End Sub

    Protected Sub HandleNewSession()
        HistorySessionSvc.HandleNewSession()
    End Sub

    Protected Overridable Sub HandleCheckedChange(jsonDoc As JObject)
        Dim prop As String = jsonDoc("property").ToString()
        Dim isChecked As Boolean = Boolean.Parse(jsonDoc("isChecked").ToString())
        If prop = "selectedCell" Then
            selectedCellChecked = isChecked
        End If
    End Sub

    Protected Overridable Sub HandleSaveSettings(jsonDoc As JObject)
        topicRandomness = jsonDoc("topicRandomness")
        contextLimit = jsonDoc("contextLimit")
        selectedCellChecked = jsonDoc("selectedCell")
        settingsScrollChecked = jsonDoc("settingsScroll")
        Dim executeCodePreview As Boolean = jsonDoc("executeCodePreview")
        Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
        Dim autocompleteShortcut As String = If(jsonDoc("autocompleteShortcut")?.Value(Of String)(), "Ctrl+.")
        Dim chatSettings As New ChatSettings(GetApplication())
        ' 保存设置到配置文件（chatMode 已弃用，统一为智能模式）
        chatSettings.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, executeCodePreview,
                                  enableAutocomplete, 800, autocompleteShortcut)
    End Sub

    ' SendMessageReferenceContentItem 已移至 Controls/Models/HistoryMessage.vb

    ' FileContentResult 类已移至 Controls/Models/HistoryMessage.vb


    ' 添加存储第一个问题的变量
    Protected firstQuestion As String = String.Empty
    Protected isFirstMessage As Boolean = True
    Private _chatHtmlFilePath As String = String.Empty ' 缓存文件路径



    ' 在 HandleSendMessage 方法中添加文件内容解析逻辑
    Protected Overridable Sub HandleSendMessage(jsonDoc As JObject)
        Debug.WriteLine($"[DEBUG BaseChatControl.HandleSendMessage] called, jsonDoc keys={String.Join(",", jsonDoc.Properties().Select(Function(p) p.Name))}")
        Dim messageValue As JToken = jsonDoc("value")
        Dim question As String
        Dim filePaths As List(Of String) = New List(Of String)()
        Dim selectedContents As List(Of SendMessageReferenceContentItem) = New List(Of SendMessageReferenceContentItem)()

        If messageValue.Type = JTokenType.Object Then
            ' New format with text, potentially filePaths, and selectedContent
            question = messageValue("text")?.ToString()

            If messageValue("filePaths") IsNot Nothing AndAlso messageValue("filePaths").Type = JTokenType.Array Then
                filePaths = messageValue("filePaths").ToObject(Of List(Of String))()
            End If

            ' 解析 selectedContent
            If messageValue("selectedContent") IsNot Nothing AndAlso messageValue("selectedContent").Type = JTokenType.Array Then
                Try
                    selectedContents = messageValue("selectedContent").ToObject(Of List(Of SendMessageReferenceContentItem))()
                Catch ex As Exception
                    Debug.WriteLine($"Error deserializing selectedContent: {ex.Message}")
                End Try
            End If
        Else
            Debug.WriteLine("HandleSendMessage: Invalid message format for 'value'.")
            Return
        End If

        If String.IsNullOrEmpty(question) AndAlso
       (filePaths Is Nothing OrElse filePaths.Count = 0) AndAlso
       (selectedContents Is Nothing OrElse selectedContents.Count = 0) Then
            Debug.WriteLine("HandleSendMessage: Empty question, no files, and no selected content.")
            Return ' Nothing to send
        End If

        ' 保存原始用户输入（用于意图识别）
        Dim originalQuestion As String = question

        ' 保存第一个问题（仅保存一次）
        If isFirstMessage AndAlso Not String.IsNullOrEmpty(question) Then
            firstQuestion = question
            isFirstMessage = False
            ' 清空缓存的文件路径，强制重新生成
            _chatHtmlFilePath = String.Empty
            Debug.WriteLine($"保存第一个问题: {firstQuestion}")
            Debug.WriteLine($"将生成文件路径: {ChatHtmlFilePath}")
        End If

        ' --- 处理选中的内容 ---
        question = AppendCurrentSelectedContent("--- Мой текущий вопрос: " & question & " ---")

        ' 检查是否有文件需要解析
        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            ' 异步处理文件解析，避免卡死UI
            HandleSendMessageWithFilesAsync(question, originalQuestion, filePaths, selectedContents, messageValue)
        Else
            ' 没有文件，直接处理消息
            HandleSendMessageCore(question, originalQuestion, filePaths, selectedContents, messageValue, "")
        End If
    End Sub

    ''' <summary>
    ''' 异步解析文件并发送消息
    ''' </summary>
    Private Sub HandleSendMessageWithFilesAsync(question As String, originalQuestion As String,
                                                 filePaths As List(Of String),
                                                 selectedContents As List(Of SendMessageReferenceContentItem),
                                                 messageValue As JToken)
        ' 显示进度提示
        GlobalStatusStrip.ShowInfo($"Разбираю файлов: {filePaths.Count}...")
        ExecuteJavaScriptAsyncJS("showFileParsingProgress(true)")

        Task.Run(Sub()
                     Try
                         Dim fileContentBuilder As New StringBuilder()
                         Dim parsedFiles As New List(Of FileContentResult)()
                         Dim totalFiles = filePaths.Count
                         Dim processedFiles = 0

                         fileContentBuilder.AppendLine(vbCrLf & "--- Ниже приведено содержимое других файлов, на которые ссылается пользователь ---")

                         ' 获取当前工作目录（需要在主线程调用）
                         Dim currentWorkingDir As String = ""
                         RunUiActionSync(Sub()
                                       currentWorkingDir = GetCurrentWorkingDirectory()
                                   End Sub)

                         For Each filePath As String In filePaths
                             Try
                                 processedFiles += 1

                                 ' 更新进度
                                 RunUiActionSync(Sub()
                                               GlobalStatusStrip.ShowInfo($"Разбор файла ({processedFiles}/{totalFiles}): {Path.GetFileName(filePath)}")
                                               ExecuteJavaScriptAsyncJS($"updateFileParsingProgress({processedFiles}, {totalFiles}, '{EscapeJavaScriptString(Path.GetFileName(filePath))}')")
                                           End Sub)

                                 ' 确定完整文件路径
                                 Dim fullFilePath As String = filePath

                                 ' 如果是绝对路径且文件存在，直接使用
                                 If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then
                                     fullFilePath = filePath
                                     Debug.WriteLine($"使用绝对路径: {fullFilePath}")
                                 ElseIf Not String.IsNullOrEmpty(currentWorkingDir) Then
                                     ' 尝试在当前工作目录下查找
                                     Dim tryPath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                                     If File.Exists(tryPath) Then
                                         fullFilePath = tryPath
                                         Debug.WriteLine($"在工作目录找到文件: {fullFilePath}")
                                     End If
                                 End If

                                 If File.Exists(fullFilePath) Then
                                     ' 根据文件扩展名选择合适的解析方法
                                     Dim fileExtension As String = Path.GetExtension(fullFilePath).ToLower()
                                     Dim fileContentResult As FileContentResult = Nothing

                                     Select Case fileExtension
                                         Case ".xlsx", ".xls", ".xlsm", ".xlsb"
                                             ' Excel文件解析需要在主线程
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".docx", ".doc", ".wps"
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".pptx", ".ppt"
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".csv", ".txt"
                                             fileContentResult = _fileParserService.ParseTextFile(fullFilePath)
                                         Case Else
                                             fileContentResult = New FileContentResult With {
                                        .FileName = Path.GetFileName(fullFilePath),
                                        .FileType = "Unknown",
                                        .ParsedContent = $"[Неподдерживаемый тип файла: {fileExtension}]"
                                    }
                                     End Select

                                     If fileContentResult IsNot Nothing Then
                                         parsedFiles.Add(fileContentResult)
                                         fileContentBuilder.AppendLine($"Имя файла: {fileContentResult.FileName}")
                                         fileContentBuilder.AppendLine($"Содержимое файла:")
                                         fileContentBuilder.AppendLine(fileContentResult.ParsedContent)
                                         fileContentBuilder.AppendLine("---")
                                     End If
                                 Else
                                     fileContentBuilder.AppendLine($"Файл '{Path.GetFileName(filePath)}' не найден, проверенный путь: {fullFilePath}")
                                     Debug.WriteLine($"文件未找到: {fullFilePath}")
                                 End If
                             Catch ex As Exception
                                 Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                                 fileContentBuilder.AppendLine($"Ошибка обработки файла '{Path.GetFileName(filePath)}': {ex.Message}")
                                 fileContentBuilder.AppendLine("---")
                             End Try
                         Next

                         fileContentBuilder.AppendLine("--- Конец содержимого файлов ---" & vbCrLf)

                         ' 文件解析完成，先保存到记忆（同步保存确保立即可检索），再在主线程继续处理消息
                         Dim appTypeForMemory = GetOfficeAppType()
                         Dim sessionIdForMemory = If(_chatStateService?.CurrentSessionId, Guid.NewGuid().ToString())
                         MemoryService.SaveFileContentToMemory(originalQuestion, fileContentBuilder.ToString(), sessionIdForMemory, appTypeForMemory)

                         ' 文件解析完成，在主线程继续处理消息
                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowInfo($"Разбор файлов завершён, обработано файлов: {parsedFiles.Count}")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")

                                       Dim questionWithFiles = question & " Конец вопроса пользователя; последующие файлы находятся в том же каталоге, их можно безопасно читать. ---"
                                       HandleSendMessageCore(questionWithFiles, originalQuestion, filePaths, selectedContents, messageValue, fileContentBuilder.ToString())
                                   End Sub)

                     Catch ex As Exception
                         Debug.WriteLine($"HandleSendMessageWithFilesAsync 出错: {ex.Message}")
                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowWarning($"Не удалось разобрать файлы: {ex.Message}")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                                       ' 重置发送按钮状态
                                       ExecuteJavaScriptAsyncJS("changeSendButton()")
                                   End Sub)
                     End Try
                 End Sub)
    End Sub

    ''' <summary>
    ''' 处理消息发送的核心逻辑（文件解析完成后调用）
    ''' </summary>
    Private Sub HandleSendMessageCore(question As String, originalQuestion As String,
                                       filePaths As List(Of String),
                                       selectedContents As List(Of SendMessageReferenceContentItem),
                                       messageValue As JToken,
                                       fileContent As String)

        ' 构建最终发送给 LLM 的消息
        Dim finalMessageToLLM As String = question

        ' 然后添加文件内容（如果有）
        If Not String.IsNullOrEmpty(fileContent) Then
            finalMessageToLLM &= fileContent
        End If

        ' 首先保存用户消息到历史记录（确保即便是意图预览模式，用户消息也会被保存）
        If Not String.IsNullOrWhiteSpace(originalQuestion) Then
            _chatStateService?.AddMessage("user", originalQuestion)
            Debug.WriteLine($"已保存用户消息到ChatStateService: {originalQuestion}")
        End If

        stopReaderStream = False ' Reset stop flag before sending new message

        ' 检查是否为模板渲染模式
        Dim responseMode As String = If(messageValue("responseMode")?.ToString(), "")
        Dim templateContext As String = If(messageValue("templateContext")?.ToString(), "")

        If responseMode = "template_render" AndAlso Not String.IsNullOrWhiteSpace(templateContext) Then
            ' 模板渲染模式：使用专用systemPrompt，不使用历史记录
            Dim templateSystemPrompt = GetTemplateRenderSystemPrompt(templateContext)
            Task.Run(Async Function()
                         Await Send(finalMessageToLLM, templateSystemPrompt, False, "template_render")
                     End Function)
        ElseIf responseMode = "proofread" Then
            ' 校对模式：用户在校对模式下发送消息，注入校对上下文
            Dim proofreadSelectedText = If(messageValue("proofreadSelectedText")?.ToString(), "")
            Dim proofreadIssueCount = If(messageValue("proofreadIssueCount")?.ToString(), "0")
            Dim proofreadSystemPrompt = BuildProofreadFollowUpPrompt(proofreadSelectedText, proofreadIssueCount)
            Task.Run(Async Function()
                         Await Send(finalMessageToLLM, proofreadSystemPrompt, False, "proofread_followup")
                     End Function)
        Else
            ' 智能模式：分析 → 预检 → AgentKernel / 普通聊天（逻辑在 ChatRoutingOrchestrator）
            Task.Run(Async Function()
                         Await ChatRoutingOrchestrator.RouteSmartModeAsync(
                             finalMessageToLLM,
                             originalQuestion,
                             filePaths,
                             selectedContents)
                     End Function)
        End If
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息
    ''' </summary>
    ''' <param name="message">消息内容</param>
    ''' <param name="intent">意图识别结果</param>
    Protected Overridable Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        ' 默认实现：如果有有效意图，使用优化的systemPrompt
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"使用意图优化提示词: {intent.IntentType}")
            Dim intentDesc As String = If(intent.UserFriendlyDescription, "")
            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "", intentDesc)
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

#Region "IChatRoutingHost"

    Private Function IChatRoutingHost_GetHistoryMessages() As List(Of HistoryMessage) Implements IChatRoutingHost.GetHistoryMessages
        Return systemHistoryMessageData
    End Function

    Private Function IChatRoutingHost_IsFollowUpQuestionAsync(question As String, history As List(Of HistoryMessage)) As Task(Of Boolean) Implements IChatRoutingHost.IsFollowUpQuestionAsync
        Return IntentService.IsFollowUpQuestionAsync(question, history)
    End Function

    Private Function IChatRoutingHost_GetContextSnapshot() As JObject Implements IChatRoutingHost.GetContextSnapshot
        Return GetContextSnapshot()
    End Function

    Private Sub IChatRoutingHost_EnrichContextForIntent(snapshot As JObject,
                                                        originalQuestion As String,
                                                        filePaths As List(Of String),
                                                        selectedContents As List(Of SendMessageReferenceContentItem)) Implements IChatRoutingHost.EnrichContextForIntent
        EnrichContextForIntent(snapshot, originalQuestion, filePaths, selectedContents)
    End Sub

    Private Function IChatRoutingHost_GetApplicationType() As String Implements IChatRoutingHost.GetApplicationType
        Return GetApplicationType()
    End Function

    Private Function IChatRoutingHost_CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext Implements IChatRoutingHost.CaptureOfficeContext
        Return CaptureOfficeContext(appType)
    End Function

    Private Function IChatRoutingHost_AnalyzeAiNativeAsync(request As Agent.AiNativeRequest) As Task(Of Agent.AiNativeRuntimeResult) Implements IChatRoutingHost.AnalyzeAiNativeAsync
        Return AiNativeRuntime.AnalyzeAsync(request)
    End Function

    Private Function IChatRoutingHost_BuildExecutionContextAsync(originalQuestion As String,
                                                                 filePaths As List(Of String),
                                                                 selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ExecutionContext) Implements IChatRoutingHost.BuildExecutionContextAsync
        Return BuildExecutionContextAsync(originalQuestion, filePaths, selectedContents)
    End Function

    Private Function IChatRoutingHost_PreSendCheckAsync(execContext As ExecutionContext) As Task(Of ContextCheckResult) Implements IChatRoutingHost.PreSendCheckAsync
        Return SelfCheckLoopController.PreSendCheckAsync(execContext)
    End Function

    Private Sub IChatRoutingHost_ShowContextHints(ragCount As Integer, intentDescription As String, contextTrace As ChatContextTrace) Implements IChatRoutingHost.ShowContextHints
        Dim intentText = If(intentDescription, "")
        Dim intentEscaped = intentText.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, " ").Replace(vbLf, " ")
        Dim traceJson = If(contextTrace Is Nothing, "{}", JObject.FromObject(contextTrace).ToString(Formatting.None))
        ExecuteJavaScriptAsyncJS($"showContextHints({{ ragCount: {ragCount}, intent: '{intentEscaped}', trace: {traceJson} }});")
    End Sub

    Private Sub IChatRoutingHost_ShowWarning(message As String) Implements IChatRoutingHost.ShowWarning
        GlobalStatusStrip.ShowWarning(message)
    End Sub

    Private Sub IChatRoutingHost_ShowIdentifyingStatus() Implements IChatRoutingHost.ShowIdentifyingStatus
        ExecuteJavaScriptAsyncJS("showIdentifyingStatus()")
    End Sub

    Private Sub IChatRoutingHost_SendChatMessage(message As String) Implements IChatRoutingHost.SendChatMessage
        SendChatMessage(message)
    End Sub

    Private Sub IChatRoutingHost_SendChatMessageWithIntent(message As String, intent As IntentResult) Implements IChatRoutingHost.SendChatMessageWithIntent
        SendChatMessageWithIntent(message, intent)
    End Sub

    Private Sub IChatRoutingHost_StartAgentPlanningFlow(message As String,
                                                        intent As IntentResult,
                                                        analysis As Agent.AiNativeRuntimeResult) Implements IChatRoutingHost.StartAgentPlanningFlow
        StartAgentPlanningFlow(message, intent, analysis)
    End Sub

    Private Sub IChatRoutingHost_SetCurrentIntentResult(intent As IntentResult) Implements IChatRoutingHost.SetCurrentIntentResult
        CurrentIntentResult = intent
    End Sub

#End Region

    ''' <summary>
    ''' Agent模式下直接启动规划流程（不显示意图预览卡片）
    ''' Primary product path: AgentKernel planning/execution.
    ''' </summary>
    Private Sub StartAgentPlanningFlow(message As String,
                                       intent As IntentResult,
                                       Optional analysis As Agent.AiNativeRuntimeResult = Nothing)
        AgentKernelSvc.AgentFullUserMessage = message
        AgentKernelSvc.AgentOriginalUserRequest = If(intent?.OriginalInput, message)

        Task.Run(Async Function()
                     Try
                         Dim goal = If(String.IsNullOrWhiteSpace(intent.OriginalInput), intent.UserFriendlyDescription, intent.OriginalInput)
                         Dim appType = GetApplicationType()

                         GlobalStatusStrip.ShowInfo("Планирую задачу...")

                         Dim currentContent = GetCurrentOfficeContent()
                         Dim historyMessages As New List(Of Tuple(Of String, String))()
                         For Each msg In systemHistoryMessageData
                             If msg.role = "user" OrElse msg.role = "assistant" Then
                                 historyMessages.Add(New Tuple(Of String, String)(msg.role, msg.content))
                             End If
                         Next

                         Dim success = Await AgentKernelSvc.StartAgentAsync(
                             message,
                             appType,
                             currentContent,
                             historyMessages,
                             CaptureOfficeContext(appType),
                             analysis?.TaskSpec,
                             analysis?.SelectedSkills)

                         If Not success Then
                             ' StartAgentAsync 已通过 Agent 完成事件展示结构化失败。执行型请求在
                             ' 失败后再发起普通聊天会产生第二份、且可能宣称成功的回答。
                             Debug.WriteLine("[StartAgentPlanningFlow] Agent 执行失败，保留失败结果，不回退普通聊天")
                             GlobalStatusStrip.ShowWarning("Не удалось выполнить задачу, посмотрите неудавшиеся шаги")
                         End If
                     Catch ex As Exception
                         Debug.WriteLine($"[StartAgentPlanningFlow] {ex.Message}")
                         GlobalStatusStrip.ShowWarning($"Не удалось запустить планирование: {ex.Message}")
                     End Try
                 End Function)
    End Sub

    Protected Overridable Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
        Return New Agent.Context.OfficeContext With {.AppType = appType}
    End Function

    ''' <summary>
    ''' 获取应用类型（子类重写）
    ''' </summary>
    Protected Overridable Function GetApplicationType() As String
        Dim appInfo = GetApplication()
        Dim appType = If(appInfo IsNot Nothing, appInfo.Type.ToString(), "Excel")
        Return appType
    End Function

#Region "Ralph Agent 智能助手"

    ''' <summary>
    ''' 处理启动Agent请求（AgentKernel 主路径）
    ''' </summary>
    Protected Sub HandleStartAgent(jsonDoc As JObject)
        Try
            Dim request = jsonDoc("request")?.ToString()
            Dim filePathsToken = jsonDoc("filePaths")
            Dim selectedContentToken = jsonDoc("selectedContent")

            If String.IsNullOrEmpty(request) AndAlso (filePathsToken Is Nothing OrElse filePathsToken.Type <> JTokenType.Array) AndAlso (selectedContentToken Is Nothing OrElse selectedContentToken.Type <> JTokenType.Array) Then
                GlobalStatusStrip.ShowWarning("Введите описание задачи или добавьте файлы")
                Return
            End If

            Debug.WriteLine($"[AgentKernel] 启动Agent，需求: {request}")
            AgentKernelSvc.AgentOriginalUserRequest = request

            ' 解析文件路径和选中内容
            Dim filePaths As New List(Of String)()
            Dim selectedContents As New List(Of SendMessageReferenceContentItem)()

            If filePathsToken IsNot Nothing AndAlso filePathsToken.Type = JTokenType.Array Then
                filePaths = filePathsToken.ToObject(Of List(Of String))()
            End If
            If selectedContentToken IsNot Nothing AndAlso selectedContentToken.Type = JTokenType.Array Then
                Try
                    selectedContents = selectedContentToken.ToObject(Of List(Of SendMessageReferenceContentItem))()
                Catch ex As Exception
                    Debug.WriteLine($"Error deserializing selectedContent: {ex.Message}")
                End Try
            End If

            ' 显示思考状态
            AgentKernelSvc.AgentThinkingUuid = Guid.NewGuid().ToString()
            Dim now = DateTime.Now
            Dim timestamp = now.ToString("yyyy-MM-dd HH:mm:ss")
            ExecuteJavaScriptAsyncJS($"createChatSection('AI', '{timestamp}', '{AgentKernelSvc.AgentThinkingUuid}')")
            ExecuteJavaScriptAsyncJS($"var thinkingDiv = document.getElementById('content-{AgentKernelSvc.AgentThinkingUuid}'); if(thinkingDiv) thinkingDiv.innerHTML = '<div class=""thinking-indicator""><div class=""thinking-dots""><span></span><span></span><span></span></div><span style=""margin-left: 12px; color: #6c757d;"">Анализирую ваш запрос...</span></div>';")

            Dim originalQuestion As String = request
            Dim finalMessageToLLM As String = request
            finalMessageToLLM = AppendCurrentSelectedContent("--- Мой текущий вопрос: " & finalMessageToLLM & " ---")

            If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
                HandleStartAgentWithFilesAsync(finalMessageToLLM, originalQuestion, filePaths, selectedContents)
            Else
                HandleStartAgentCore(finalMessageToLLM, originalQuestion, "")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleStartAgent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось запустить агента: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Compatibility: Ralph /loop and startLoop messages are rewritten to startAgent (AgentKernel).
    ''' </summary>
    Private Sub HandleLegacyStartLoop(jsonDoc As JObject)
        Try
            Dim goal = jsonDoc("goal")?.ToString()
            If String.IsNullOrWhiteSpace(goal) Then
                goal = jsonDoc("request")?.ToString()
            End If

            Debug.WriteLine("[ExecutionPathPolicy] legacy startLoop remapped to startAgent/AgentKernel")
            Dim agentRequest As New JObject()
            agentRequest("type") = "startAgent"
            agentRequest("request") = goal
            HandleStartAgent(agentRequest)
        Catch ex As Exception
            Debug.WriteLine($"HandleLegacyStartLoop failed: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось запустить агента: {ex.Message}")
        End Try
    End Sub

    Private Sub HandleStartAgentWithFilesAsync(question As String, originalQuestion As String,
                                              filePaths As List(Of String),
                                              selectedContents As List(Of SendMessageReferenceContentItem))
        ' 显示进度提示
        GlobalStatusStrip.ShowInfo($"Разбираю файлов: {filePaths.Count}...")
        ExecuteJavaScriptAsyncJS("showFileParsingProgress(true)")

        Task.Run(Sub()
                     Try
                         Dim fileContentBuilder As New StringBuilder()
                         Dim totalFiles = filePaths.Count
                         Dim processedFiles = 0

                         fileContentBuilder.AppendLine(vbCrLf & "--- Ниже приведено содержимое других файлов, на которые ссылается пользователь ---")

                         ' 获取当前工作目录（需要在主线程调用）
                         Dim currentWorkingDir As String = ""
                         RunUiActionSync(Sub()
                                       currentWorkingDir = GetCurrentWorkingDirectory()
                                   End Sub)

                         For Each filePath As String In filePaths
                             Try
                                 processedFiles += 1

                                 ' 更新进度
                                 RunUiActionSync(Sub()
                                               GlobalStatusStrip.ShowInfo($"Разбор файла ({processedFiles}/{totalFiles}): {Path.GetFileName(filePath)}")
                                               ExecuteJavaScriptAsyncJS($"updateFileParsingProgress({processedFiles}, {totalFiles}, '{EscapeJavaScriptString(Path.GetFileName(filePath))}')")
                                           End Sub)

                                 ' 确定完整文件路径
                                 Dim fullFilePath As String = filePath

                                 ' 如果是绝对路径且文件存在，直接使用
                                 If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then
                                     fullFilePath = filePath
                                     Debug.WriteLine($"使用绝对路径: {fullFilePath}")
                                 ElseIf Not String.IsNullOrEmpty(currentWorkingDir) Then
                                     ' 尝试在当前工作目录下查找
                                     Dim tryPath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                                     If File.Exists(tryPath) Then
                                         fullFilePath = tryPath
                                         Debug.WriteLine($"在工作目录找到文件: {fullFilePath}")
                                     End If
                                 End If

                                 If File.Exists(fullFilePath) Then
                                     ' 根据文件扩展名选择合适的解析方法
                                     Dim fileExtension As String = Path.GetExtension(fullFilePath).ToLower()
                                     Dim fileContentResult As FileContentResult = Nothing

                                     Select Case fileExtension
                                         Case ".xlsx", ".xls", ".xlsm", ".xlsb"
                                             ' Excel文件解析需要在主线程
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".docx", ".doc", ".wps"
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".pptx", ".ppt"
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".csv", ".txt"
                                             fileContentResult = _fileParserService.ParseTextFile(fullFilePath)
                                         Case Else
                                             fileContentResult = New FileContentResult With {
                                        .FileName = Path.GetFileName(fullFilePath),
                                        .FileType = "Unknown",
                                        .ParsedContent = $"[Неподдерживаемый тип файла: {fileExtension}]"
                                    }
                                     End Select

                                     If fileContentResult IsNot Nothing Then
                                         fileContentBuilder.AppendLine($"Имя файла: {fileContentResult.FileName}")
                                         fileContentBuilder.AppendLine($"Содержимое файла:")
                                         fileContentBuilder.AppendLine(fileContentResult.ParsedContent)
                                         fileContentBuilder.AppendLine("---")
                                     End If
                                 Else
                                     fileContentBuilder.AppendLine($"Файл '{Path.GetFileName(filePath)}' не найден, проверенный путь: {fullFilePath}")
                                     Debug.WriteLine($"文件未找到: {fullFilePath}")
                                 End If
                             Catch ex As Exception
                                 Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                                 fileContentBuilder.AppendLine($"Ошибка обработки файла '{Path.GetFileName(filePath)}': {ex.Message}")
                                 fileContentBuilder.AppendLine("---")
                             End Try
                         Next

                         fileContentBuilder.AppendLine("--- Конец содержимого файлов ---" & vbCrLf)

                         ' 文件解析完成，先保存到记忆（同步保存确保立即可检索），再在主线程继续处理消息
                         Dim appTypeForMemory = GetOfficeAppType()
                         Dim sessionIdForMemory = If(_chatStateService?.CurrentSessionId, Guid.NewGuid().ToString())
                         MemoryService.SaveFileContentToMemory(originalQuestion, fileContentBuilder.ToString(), sessionIdForMemory, appTypeForMemory)

                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowInfo($"Разбор файлов завершён, обработано файлов: {processedFiles}")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")

                                       Dim questionWithFiles = question & " Конец вопроса пользователя; последующие файлы находятся в том же каталоге, их можно безопасно читать. ---"
                                       HandleStartAgentCore(questionWithFiles, originalQuestion, fileContentBuilder.ToString())
                                   End Sub)

                     Catch ex As Exception
                         Debug.WriteLine($"HandleStartAgentWithFilesAsync 出错: {ex.Message}")
                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowWarning($"Не удалось разобрать файлы: {ex.Message}")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                                   End Sub)
                     End Try
                 End Sub)
    End Sub

    ''' <summary>
    ''' 处理Agent启动的核心逻辑（文件解析完成后调用）
    ''' AgentKernel 启动核心（始终注入 CaptureOfficeContext）
    ''' </summary>
    Private Sub HandleStartAgentCore(question As String, originalQuestion As String, fileContent As String)
        ' 构建最终发送给 LLM 的消息
        Dim finalMessageToLLM As String = question

        ' 然后添加文件内容（如果有）
        If Not String.IsNullOrEmpty(fileContent) Then
            finalMessageToLLM &= fileContent
        End If

        AgentKernelSvc.AgentFullUserMessage = finalMessageToLLM
        AgentKernelSvc.AgentOriginalUserRequest = originalQuestion

        Task.Run(Async Function()
                     Try
                         Dim appType = GetApplicationType()
                         Dim currentContent = GetCurrentOfficeContent()

                         Dim historyMessages As New List(Of Tuple(Of String, String))()
                         For Each msg In systemHistoryMessageData
                             If msg.role = "user" OrElse msg.role = "assistant" Then
                                 historyMessages.Add(New Tuple(Of String, String)(msg.role, msg.content))
                             End If
                         Next

                         GlobalStatusStrip.ShowInfo("Анализирую ваш запрос...")
                         Dim success = Await AgentKernelSvc.StartAgentAsync(finalMessageToLLM, appType, currentContent, historyMessages, CaptureOfficeContext(appType))

                         If Not success Then
                             GlobalStatusStrip.ShowWarning("Не удалось проанализировать запрос, повторите попытку")
                             AgentKernelSvc.AgentThinkingUuid = Nothing
                         End If
                     Catch ex As Exception
                         Debug.WriteLine($"[AgentKernel] HandleStartAgentCore 出错: {ex.Message}")
                         GlobalStatusStrip.ShowWarning($"Не удалось проанализировать запрос: {ex.Message}")
                         AgentKernelSvc.AgentThinkingUuid = Nothing
                     End Try
                 End Function)
    End Sub

    ''' <summary>
    ''' 处理修改Agent计划（新架构下暂不支持重新规划，记录日志）
    ''' </summary>
    Protected Overridable Sub HandleRefineAgentPlan(jsonDoc As JObject)
        Debug.WriteLine("[AgentKernel] 收到 refineAgentPlan 请求，当前架构不支持手动重新规划")
        GlobalStatusStrip.ShowInfo("Запрос на изменение получен; опишите новое требование в следующем сообщении")
    End Sub

    ''' <summary>
    ''' 处理终止Agent（兼容新旧架构）
    ''' </summary>
    Protected Sub HandleAbortAgent()
        AgentKernelSvc.AbortAgent()
    End Sub

    ''' <summary>
    ''' 处理开始执行Agent（兼容新旧架构）
    ''' </summary>
    Protected Sub HandleStartAgentExecution(jsonDoc As JObject)
        ' 新架构中，用户点击"开始执行"相当于批准计划
        AgentKernelSvc.Approve()
    End Sub

#Region "AgentKernel 统一消息处理"

    ''' <summary>
    ''' 处理用户批准计划
    ''' </summary>
    Protected Sub HandleAgentApprovePlan(jsonDoc As JObject)
        Try
            AgentKernelSvc.Approve()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentApprovePlan 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户拒绝计划
    ''' </summary>
    Protected Sub HandleAgentRejectPlan(jsonDoc As JObject)
        Try
            AgentKernelSvc.Reject()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentRejectPlan 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户批准步骤（高风险操作）
    ''' </summary>
    Protected Sub HandleAgentApproveStep(jsonDoc As JObject)
        Try
            AgentKernelSvc.Approve()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentApproveStep 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户修改计划
    ''' </summary>
    Protected Sub HandleAgentRefinePlan(jsonDoc As JObject)
        Try
            Dim feedback = jsonDoc("payload")?("feedback")?.ToString()
            If String.IsNullOrEmpty(feedback) Then
                feedback = jsonDoc("feedback")?.ToString()
            End If
            If Not String.IsNullOrEmpty(feedback) Then
                Debug.WriteLine($"[AgentKernel] 用户请求修改计划: {feedback}")
                ExecuteJavaScriptAsyncJS("addThinkingMessage('Перепланирую с учётом ваших замечаний...')")
                Dim request = AgentKernelSvc.AgentOriginalUserRequest
                If Not String.IsNullOrEmpty(request) Then
                    Dim refinedRequest = request & vbCrLf & "[Правки пользователя] " & feedback
                    AgentKernelSvc.AgentThinkingUuid = Guid.NewGuid().ToString()
                    Dim now = DateTime.Now
                    Dim timestamp = now.ToString("yyyy-MM-dd HH:mm:ss")
                    ExecuteJavaScriptAsyncJS($"createChatSection('AI', '{timestamp}', '{AgentKernelSvc.AgentThinkingUuid}')")
                    HandleStartAgentCore(refinedRequest, request, "")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentRefinePlan 出错: {ex.Message}")
        End Try
    End Sub

#End Region

    ''' <summary>
    ''' 获取当前Office内容（子类重写以提供具体实现）
    ''' </summary>
    Protected Overridable Function GetCurrentOfficeContent() As String
        ' 基类默认实现：尝试获取选区内容
        Try
            Dim selInfo = CaptureCurrentSelectionInfo("")
            If selInfo IsNot Nothing AndAlso Not String.IsNullOrEmpty(selInfo.SelectedText) Then
                Return selInfo.SelectedText
            End If
        Catch
        End Try
        Return "(нет выделенного содержимого)"
    End Function

    ''' <summary>
    ''' 构建执行上下文（供自检Loop使用）
    ''' </summary>
    Protected Overridable Async Function BuildExecutionContextAsync(
        originalQuestion As String,
        filePaths As List(Of String),
        selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ExecutionContext)

        Dim context As New ExecutionContext()
        context.UserMessage = originalQuestion
        context.OriginalQuestion = originalQuestion
        context.FilePaths = If(filePaths, New List(Of String)())

        ' 根据意图设置请求操作类型
        If CurrentIntentResult IsNot Nothing Then
            Select Case CurrentIntentResult.OfficeIntent
                Case OfficeIntentType.FORMAT_STYLE, OfficeIntentType.TEXT_FORMAT
                    context.RequestedOperation = RequestedOperation.Reformat
                    context.ExpectedFormat = InstructionFormat.DslJson
                Case OfficeIntentType.DOCUMENT_EDIT
                    ' 文档编辑可能包含校对需求，根据描述判断
                    If CurrentIntentResult.UserFriendlyDescription IsNot Nothing AndAlso
                       CurrentIntentResult.UserFriendlyDescription.ToLower().Contains("校对") Then
                        context.RequestedOperation = RequestedOperation.Proofread
                        context.ExpectedFormat = InstructionFormat.ProofreadJson
                    Else
                        context.RequestedOperation = RequestedOperation.Reformat
                        context.ExpectedFormat = InstructionFormat.DslJson
                    End If
                Case Else
                    context.RequestedOperation = RequestedOperation.GeneralQuery
            End Select
        End If

        ' 获取Office内容信息
        Try
            Dim appInfo = GetApplication()
            If appInfo IsNot Nothing Then
                Select Case appInfo.Type.ToString().ToLower()
                    Case "word"
                        context.OfficeAppType = OfficeAppType.Word
                    Case "excel"
                        context.OfficeAppType = OfficeAppType.Excel
                    Case "powerpoint"
                        context.OfficeAppType = OfficeAppType.PowerPoint
                End Select
            End If

            Dim contentInfo As New OfficeContentInfo()
            Dim selInfo = CaptureCurrentSelectionInfo("")
            If selInfo IsNot Nothing Then
                contentInfo.SelectedParagraphs = New List(Of String) From {selInfo.SelectedText}
                context.SelectionInfo = selInfo
                context.RequiresSelection = Not String.IsNullOrEmpty(selInfo.SelectedText)
            End If
            context.OfficeContent = contentInfo
        Catch ex As Exception
            Debug.WriteLine($"[BuildExecutionContextAsync] 获取Office内容失败: {ex.Message}")
        End Try

        ' 历史上下文
        If systemHistoryMessageData IsNot Nothing Then
            context.ConversationHistory = systemHistoryMessageData.ToList()
        End If

        Return context
    End Function

    ''' <summary>
    ''' 非流式 AI 请求，直接获取完整响应（供 AgentKernel 使用）
    ''' </summary>
    Private Async Function SendAndGetResponseAsync(prompt As String, systemPrompt As String, historyMessages As List(Of HistoryMessage)) As Task(Of String)
        Try
            Dim messagesArray As New JArray()
            If Not String.IsNullOrEmpty(systemPrompt) Then
                messagesArray.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})
            End If
            If historyMessages IsNot Nothing Then
                For Each msg In historyMessages
                    If Not String.IsNullOrEmpty(msg.content) Then
                        messagesArray.Add(New JObject From {{"role", msg.role}, {"content", msg.content}})
                    End If
                Next
            End If
            messagesArray.Add(New JObject From {{"role", "user"}, {"content", prompt}})

            Dim requestObj As New JObject()
            requestObj("model") = ConfigSettings.ModelName
            requestObj("messages") = messagesArray
            requestObj("stream") = False
            ReasoningRequestHelper.ApplyReasoningOptions(requestObj, ConfigSettings.ReasoningMode, ConfigSettings.ModelName, ConfigSettings.platform, ConfigSettings.ApiUrl)

            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)
            Dim apiUrl = ConfigSettings.ApiUrl
            Dim apiKey = ConfigSettings.ApiKey

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim client = HttpClientPool.GetClient(apiUrl)
            Using request As New HttpRequestMessage(HttpMethod.Post, apiUrl)

                ' Anthropic 兼容
                If apiUrl.Contains("anthropic.com") Then
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                Else
                    request.Headers.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)
                End If

                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using timeoutCts As New System.Threading.CancellationTokenSource(TimeSpan.FromMinutes(2))
                    Using response As HttpResponseMessage = Await client.SendAsync(request, timeoutCts.Token)
                        response.EnsureSuccessStatusCode()
                        Dim jsonContent As String = Await response.Content.ReadAsStringAsync()

                    ' 提取 content
                    Dim json = JObject.Parse(jsonContent)
                    Dim content As String = ""

                    ' OpenAI 格式
                    If json("choices") IsNot Nothing Then
                        Dim choices = CType(json("choices"), JArray)
                        If choices.Count > 0 Then
                            content = If(choices(0)("message")?("content")?.ToString(), "")
                        End If
                    End If

                    ' Anthropic 格式
                    If String.IsNullOrEmpty(content) AndAlso json("content") IsNot Nothing Then
                        Dim contentArr = CType(json("content"), JArray)
                        If contentArr IsNot Nothing AndAlso contentArr.Count > 0 Then
                            content = If(contentArr(0)("text")?.ToString(), "")
                        End If
                    End If

                        Return content
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[SendAndGetResponseAsync] 请求失败: {ex.Message}")
            Return ""
        End Try
    End Function

#End Region

    ''' <summary>
    ''' 处理打开文件对话框请求
    ''' </summary>
    Protected Sub HandleOpenFileDialog()
        Try
            ' 需要在UI线程上执行
            If Me.InvokeRequired Then
                RunUiActionSync(Sub() HandleOpenFileDialog())
                Return
            End If

            Using dialog As New OpenFileDialog()
                dialog.Title = "Выберите файл для ссылки"
                dialog.Filter = "Файлы Excel|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv|" &
                               "Файлы Word|*.doc;*.docx|" &
                               "Файлы PowerPoint|*.ppt;*.pptx|" &
                               "Все поддерживаемые файлы|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv;*.doc;*.docx;*.ppt;*.pptx"
                dialog.FilterIndex = 4 ' 默认显示所有支持的文件
                dialog.Multiselect = True

                If dialog.ShowDialog() = DialogResult.OK Then
                    ' 构建文件列表JSON
                    Dim filesArray As New JArray()
                    For Each filePath In dialog.FileNames
                        Dim fileObj As New JObject()
                        fileObj("name") = Path.GetFileName(filePath)
                        fileObj("path") = filePath
                        filesArray.Add(fileObj)
                    Next

                    ' 发送给前端
                    ExecuteJavaScriptAsyncJS($"addFilesFromDialog({filesArray.ToString(Formatting.None)})")
                    Debug.WriteLine($"选择了 {dialog.FileNames.Length} 个文件")
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"HandleOpenFileDialog 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("Ошибка при открытии диалога выбора файла")
        End Try
    End Sub

    ''' <summary>
    ''' 处理打开API配置窗口请求
    ''' </summary>
    Protected Sub HandleOpenApiConfigForm()
        Try
            ' 需要在UI线程上执行
            If Me.InvokeRequired Then
                RunUiActionSync(Sub() HandleOpenApiConfigForm())
                Return
            End If

            Dim configForm As New ConfigApiForm()
            If configForm.ShowDialog() = DialogResult.OK Then
                ' 配置已更新，刷新前端显示
                UpdateModelDisplayInUI()
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleOpenApiConfigForm 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("Ошибка при открытии окна конфигурации")
        End Try
    End Sub

    ''' <summary>
    ''' 处理获取当前模型信息请求
    ''' </summary>
    Protected Sub HandleGetCurrentModel()
        Try
            UpdateModelDisplayInUI()
        Catch ex As Exception
            Debug.WriteLine($"HandleGetCurrentModel 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 更新前端的模型显示
    ''' </summary>
    Protected Sub HandleGetCurrentAppInfo()
        Try
            ExecuteJavaScriptAsyncJS($"window.currentOfficeAppName = '{EscapeJavaScriptString(GetOfficeApplicationName())}';")
        Catch ex As Exception
            Debug.WriteLine($"HandleGetCurrentAppInfo failed: {ex.Message}")
        End Try
    End Sub

    Protected Sub UpdateModelDisplayInUI()
        Try
            Dim platform = ConfigSettings.platform
            Dim modelName = ConfigSettings.ModelName

            ' 转义特殊字符
            platform = If(platform, "").Replace("'", "\'").Replace("""", "\""")
            modelName = If(modelName, "").Replace("'", "\'").Replace("""", "\""")

            Dim js = $"updateCurrentModelDisplay('{platform}', '{modelName}');"
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"UpdateModelDisplayInUI 出错: {ex.Message}")
        End Try
    End Sub

#Region "排版模板功能消息处理"

    ''' <summary>
    ''' 获取排版模板列表（含docx解析出的语义映射卡片）
    ''' </summary>
    Protected Sub HandleGetReformatTemplates()
        ReformatSvc.HandleGetReformatTemplates()
    End Sub

    ''' <summary>
    ''' 刷新排版模板列表（Public，供外部调用）
    ''' </summary>
    Public Sub RefreshReformatTemplates()
        ReformatSvc.HandleGetReformatTemplates()
    End Sub

    ''' <summary>
    ''' 使用排版模板（含docx映射识别）
    ''' </summary>
    Protected Overridable Sub HandleUseReformatTemplate(jsonDoc As JObject)
        Try
            Dim templateId = jsonDoc("templateId")?.ToString()

            ' 识别docx映射卡片（ID前缀 "docx_"）
            If templateId IsNot Nothing AndAlso templateId.StartsWith("docx_") Then
                Dim mappingId = templateId.Substring(5)
                Dim mapping = SemanticMappingManager.Instance.GetMappingById(mappingId)
                If mapping IsNot Nothing Then
                    ApplyReformatWithMapping(mapping)
                    Return
                Else
                    GlobalStatusStrip.ShowWarning("Семантическое сопоставление не найдено")
                    Return
                End If
            End If

            ' 常规模板
            Dim template = ReformatTemplateManager.Instance.GetTemplateById(templateId)
            If template Is Nothing Then
                GlobalStatusStrip.ShowWarning("Шаблон не найден")
                Return
            End If

            ApplyReformatWithTemplate(template)

        Catch ex As Exception
            Debug.WriteLine($"HandleUseReformatTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось использовать шаблон: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用模板进行排版（由子类实现）
    ''' </summary>
    Protected Overridable Sub ApplyReformatWithTemplate(template As ReformatTemplate)
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает форматирование по шаблону")
    End Sub

    ''' <summary>
    ''' 使用SemanticStyleMapping直接排版（由子类实现，用于docx解析的映射）
    ''' </summary>
    Protected Overridable Sub ApplyReformatWithMapping(mapping As SemanticStyleMapping)
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает форматирование по сопоставлению документа")
    End Sub

#Region "排版规范处理方法"

    ''' <summary>
    ''' 获取排版规范列表
    ''' </summary>
    Protected Sub HandleGetStyleGuides()
        ReformatSvc.HandleGetStyleGuides()
    End Sub

    ''' <summary>
    ''' 使用排版规范
    ''' </summary>
    Protected Overridable Sub HandleUseStyleGuide(jsonDoc As JObject)
        Try
            Dim guideId = jsonDoc("guideId")?.ToString()
            Dim guide = StyleGuideManager.Instance.GetStyleGuideById(guideId)

            If guide Is Nothing Then
                GlobalStatusStrip.ShowWarning("Стандарт не найден")
                Return
            End If

            ' 由子类实现具体的排版逻辑
            ApplyReformatWithStyleGuide(guide)

        Catch ex As Exception
            Debug.WriteLine($"HandleUseStyleGuide 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось использовать стандарт: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用规范进行排版（由子类实现）
    ''' </summary>
    Protected Overridable Sub ApplyReformatWithStyleGuide(guide As StyleGuideResource)
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает форматирование по стандарту")
    End Sub

    ''' <summary>
    ''' 上传规范文档
    ''' </summary>
    Protected Sub HandleUploadStyleGuideDocument()
        ReformatSvc.HandleUploadStyleGuideDocument()
    End Sub

    ''' <summary>
    ''' 删除规范
    ''' </summary>
    Protected Sub HandleDeleteStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleDeleteStyleGuide(jsonDoc)
    End Sub

    ''' <summary>
    ''' 更新规范内容（编辑保存）
    ''' </summary>
    Protected Sub HandleUpdateStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleUpdateStyleGuide(jsonDoc)
    End Sub

    ''' <summary>
    ''' 复制规范
    ''' </summary>
    Protected Sub HandleDuplicateStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleDuplicateStyleGuide(jsonDoc)
    End Sub

    ''' <summary>
    ''' 导出规范
    ''' </summary>
    Protected Sub HandleExportStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleExportStyleGuide(jsonDoc)
    End Sub

#End Region

    ''' <summary>
    ''' 在Word中预览模板
    ''' </summary>
    Protected Overridable Sub HandlePreviewTemplateInWord(jsonDoc As JObject)
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает предпросмотр шаблона")
    End Sub

    ''' <summary>
    ''' 保存当前文档为模板
    ''' </summary>
    Protected Overridable Sub HandleSaveCurrentDocumentAsTemplate()
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает сохранение документа как шаблона")
    End Sub

    ''' <summary>
    ''' 导入模板
    ''' </summary>
    Protected Sub HandleImportTemplate()
        ReformatSvc.HandleImportTemplate()
    End Sub

    ''' <summary>
    ''' 导出模板
    ''' </summary>
    Protected Sub HandleExportTemplate(jsonDoc As JObject)
        ReformatSvc.HandleExportTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 复制模板
    ''' </summary>
    Protected Sub HandleDuplicateTemplate(jsonDoc As JObject)
        ReformatSvc.HandleDuplicateTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 删除模板
    ''' </summary>
    Protected Sub HandleDeleteTemplate(jsonDoc As JObject)
        ReformatSvc.HandleDeleteTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 打开模板编辑器
    ''' </summary>
    Protected Sub HandleOpenTemplateEditor(jsonDoc As JObject)
        ReformatSvc.HandleOpenTemplateEditor(jsonDoc)
    End Sub

    ''' <summary>
    ''' 显示模板编辑器面板（子类可重写以使用 CustomTaskPane）
    ''' </summary>
    ''' <param name="template">要编辑的模板，为空则新建</param>
    ''' <returns>如果成功显示返回 True，否则返回 False 以使用回退的 WinForm</returns>
    Protected Overridable Function ShowTemplateEditorPane(template As ReformatTemplate) As Boolean
        Return False ' 默认不支持 CustomTaskPane
    End Function

    ''' <summary>
    ''' 获取样式预览回调（子类可重写以提供实时预览功能）
    ''' </summary>
    Protected Overridable Function GetStylePreviewCallback() As PreviewStyleCallback
        Return Nothing
    End Function

    ''' <summary>
    ''' 进入模板选择模式（供Ribbon调用）
    ''' </summary>
    Public Async Sub EnterReformatTemplateMode()
        Await ReformatSvc.EnterReformatTemplateMode()
    End Sub

    ''' <summary>
    ''' 退出模板选择模式
    ''' </summary>
    Public Sub ExitReformatTemplateMode()
        ReformatSvc.ExitReformatTemplateMode()
    End Sub

    ' ========== AI模板编辑器功能 ==========

    ''' <summary>
    ''' 进入AI模板编辑模式（供外部调用）
    ''' </summary>
    Public Sub EnterAiTemplateEditorMode(Optional template As ReformatTemplate = Nothing)
        ReformatSvc.EnterAiTemplateEditorMode(template)
    End Sub

    ''' <summary>
    ''' 处理保存AI模板
    ''' </summary>
    Protected Sub HandleSaveAiTemplate(jsonDoc As JObject)
        ReformatSvc.HandleSaveAiTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 处理预览AI模板
    ''' </summary>
    Protected Overridable Sub HandlePreviewAiTemplate(jsonDoc As JObject)
        Try
            If Me.InvokeRequired Then
                RunUiActionSync(Sub() HandlePreviewAiTemplate(jsonDoc))
                Return
            End If

            Dim templateJson As String = jsonDoc("templateJson")?.ToString()
            If String.IsNullOrWhiteSpace(templateJson) Then
                GlobalStatusStrip.ShowWarning("Нет данных шаблона для предпросмотра")
                Return
            End If

            Dim template = JsonConvert.DeserializeObject(Of ReformatTemplate)(templateJson)

            ' 调用子类实现的预览方法
            PreviewTemplateInDocument(template)

        Catch ex As Exception
            Debug.WriteLine($"HandlePreviewAiTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось выполнить предпросмотр шаблона: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 在文档中预览模板效果（由子类实现）
    ''' </summary>
    Protected Overridable Sub PreviewTemplateInDocument(template As ReformatTemplate)
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает предпросмотр шаблона")
    End Sub

#End Region

    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        Dim responseUuid As String = If(jsonDoc("responseUuid")?.ToString(), "")

        Try
            ' 执行代码
            ExecuteCode(code, language, preview)

            ' 执行成功后通知前端（清空引用区、更新按钮状态）
            If Not String.IsNullOrEmpty(responseUuid) Then
                ExecuteJavaScriptAsyncJS($"handleExecutionSuccess('{responseUuid}')")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleExecuteCode 执行出错: {ex.Message}")
            ' 执行失败后通知前端（恢复按钮可点击）
            If Not String.IsNullOrEmpty(responseUuid) Then
                Dim escapedMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, " ")
                ExecuteJavaScriptAsyncJS($"handleExecutionError('{responseUuid}', '{escapedMsg}')")
            End If
        End Try
    End Sub


    ' 抽象方法，由子类实现
    Protected MustOverride Function ParseFile(filePath As String) As FileContentResult
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    ' 文本/CSV 解析已委托给 FileParserService，请使用 _fileParserService.ParseTextFile()

    Protected MustOverride Function GetApplication() As ApplicationInfo

    ''' <summary>
    ''' 获取Office应用类型，用于前端区分Word/PowerPoint/Excel
    ''' </summary>
    Protected Overridable Function GetOfficeAppType() As String
        Return "Unknown"
    End Function

    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
    Protected MustOverride Function RunCode(vbaCode As String)

    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法 - 委托给 CodeExecutionService
    Public Sub ExecuteCode(code As String, language As String, preview As Boolean)
        Dim result = CodeExecutionService.ExecuteCodeWithToolResult(code, language, preview)
        If result IsNot Nothing AndAlso Not result.Success Then
            GlobalStatusStrip.ShowWarning(If(result.UserMessage, result.Message))
        ElseIf IsPreviewOnlyResult(result) Then
            GlobalStatusStrip.ShowInfo("Предпросмотр: изменения в документ не внесены. Чтобы применить план, снимите галочку «Предпросмотр выполнения кода» в настройках чата и запустите выполнение ещё раз.")
        End If
    End Sub

    ''' <summary>
    ''' Результат, который только проверил план и ничего не изменил (например, CreateSlides в режиме предпросмотра).
    ''' Иначе кнопка «Выполнено» вводит в заблуждение: слайды не создаются.
    ''' </summary>
    Private Shared Function IsPreviewOnlyResult(result As Agent.ToolResult) As Boolean
        If result Is Nothing Then Return False

        Dim data = TryCast(result.Data, Newtonsoft.Json.Linq.JObject)
        If data Is Nothing Then Return False

        Dim isPreview = data("preview") IsNot Nothing AndAlso data("preview").ToObject(Of Boolean)()
        Dim isRendered = data("rendered") IsNot Nothing AndAlso data("rendered").ToObject(Of Boolean)()
        Return isPreview AndAlso Not isRendered
    End Function

    ' ExecuteJavaScript 已委托给 CodeExecutionService
    ' 添加清除特定 sheetName 的方法
    Public Async Sub ClearSelectedContentBySheetName(sheetName As String)
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
        $"clearSelectedContentBySheetName({JsonConvert.SerializeObject(sheetName)})"
    )
    End Sub


    ' 抽象方法 - 获取Office应用程序对象
    Protected MustOverride Function GetOfficeApplicationObject() As Object

    ' ExecuteExcelFormula, ExecuteVBACode, ContainsProcedureDeclaration, FindFirstProcedureName 已委托给 CodeExecutionService

    ' 虚方法 - 评估Excel公式（只有Excel子类会实现）
    Protected Overridable Function EvaluateFormula(formula As String, preview As Boolean) As Boolean
        ' 默认实现返回Nothing
        Return True
    End Function

    ' _responseModeMap 代理到 _chatStateService，子类通过此属性访问
    Protected ReadOnly Property _responseModeMap As Dictionary(Of String, String)
        Get
            Return _chatStateService.ResponseModeMap
        End Get
    End Property

    ''' <summary>
    ''' 获取当前响应的模式（用于子类检查是否应该跳过某些操作）
    ''' </summary>
    Protected Function GetCurrentResponseMode() As String
        If String.IsNullOrEmpty(_finalUuid) Then Return ""
        If _responseModeMap.ContainsKey(_finalUuid) Then
            Return _responseModeMap(_finalUuid)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' 检查当前是否处于排版模式（排版模式下JSON不应自动执行命令）
    ''' </summary>
    Protected Function IsInReformatMode() As Boolean
        Return GetCurrentResponseMode() = "reformat"
    End Function

    ''' <summary>
    ''' 检查当前是否处于预览模式（预览模式下JSON用于前端显示，不应自动执行命令）
    ''' 包括：排版(reformat)、校对(proofread)
    ''' </summary>
    Protected Function IsInPreviewMode() As Boolean
        Dim mode = GetCurrentResponseMode()
        Return mode = "reformat" OrElse mode = "proofread"
    End Function

    ' 测试方法已移除，如需调试请使用单独的测试类

    ' 存储调用Send时的请求参数（requestUuid/responseUuid -> JObject）
    Protected _savedRequestParams As New Dictionary(Of String, JObject)()

    Private Function GetSendValidationWarning(failure As ChatSendValidationFailure) As String
        Select Case failure
            Case ChatSendValidationFailure.MissingApiKey
                Return "Сначала настройте ApiKey модели!"
            Case ChatSendValidationFailure.MissingApiUrl
                Return "Сначала настройте API модели!"
            Case ChatSendValidationFailure.MissingQuestion
                Return "Введите вопрос!"
            Case Else
                Return "Параметры запроса неполные"
        End Select
    End Function

    Public Async Function Send(question As String, systemPrompt As String, addHistory As Boolean, responseMode As String, Optional intentDescription As String = Nothing, Optional responseUuid As String = Nothing) As Task
        Dim apiUrl As String = ConfigSettings.ApiUrl
        Dim apiKey As String = ConfigSettings.ApiKey

        Dim validation = SendValidator.Validate(apiUrl, apiKey, question)
        If Not validation.IsValid Then
            GlobalStatusStrip.ShowWarning(GetSendValidationWarning(validation.Failure))
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        Dim uuid As String = If(responseUuid, Guid.NewGuid().ToString())
        ' 这里生成 requestUuid（用于绑定选区）
        Dim requestUuid As String = Guid.NewGuid().ToString()


        ' 将 PendingSelectionInfo 绑定到 requestUuid
        Try
            If PendingSelectionInfo Is Nothing Then
                Dim captured As SelectionInfo = Nothing
                Try
                    captured = CaptureCurrentSelectionInfo(responseMode)
                Catch ex As Exception
                    Debug.WriteLine("CaptureCurrentSelectionInfo 异常: " & ex.Message)
                End Try
                If captured IsNot Nothing Then
                    PendingSelectionInfo = captured
                End If
            End If

            ' 将 PendingSelectionInfo 绑定到 requestUuid（原有逻辑）
            If PendingSelectionInfo IsNot Nothing Then
                Try
                    _selectionPendingMap(requestUuid) = PendingSelectionInfo
                Catch ex As Exception
                    Debug.WriteLine($"绑定 PendingSelectionInfo 到 requestUuid 失败: {ex.Message}")
                End Try
                ' 清空 PendingSelectionInfo，避免被下一个请求误用
                PendingSelectionInfo = Nothing
            End If
        Catch
        End Try

        Try
            systemPrompt = SystemPromptResolver.ResolveSystemPrompt(systemPrompt, GetApplication(), CurrentIntentResult, responseMode)

            ' ragCount 由 CreateRequestBody 通过 ByRef 返回，无需再次查询记忆
            Dim ragCount As Integer = 0
            Dim contextTrace As ChatContextTrace = Nothing
            Dim requestBody As String = CreateRequestBody(requestUuid, question, systemPrompt, addHistory, ragCount, contextTrace)
            If ragCount > 0 OrElse Not String.IsNullOrEmpty(intentDescription) OrElse contextTrace IsNot Nothing Then
                Dim intentEscaped As String = If(intentDescription, "").Replace("\", "\\").Replace("'", "\'").Replace(vbCr, " ").Replace(vbLf, " ")
                Dim traceJson = If(contextTrace Is Nothing, "null", JObject.FromObject(contextTrace).ToString(Formatting.None))
                Dim js As String = $"showContextHints({{ ragCount: {ragCount}, intent: '{intentEscaped}', trace: {traceJson} }});"
                ExecuteJavaScriptAsyncJS(js)
            End If
            ' 设置历史保存回调（在 FinalizeStream ClearBuffers 前执行）
            Dim capturedQuestion = question
            Dim capturedAddHistory = addHistory
            HttpStreamSvc.FinalizeCallback = Sub(ah, oq)
                                                 Dim answerContent = _chatStateService.MarkdownBuffer.ToString()
                                                 Dim answer = New HistoryMessage() With {
                                                     .role = "assistant",
                                                     .content = answerContent
                                                 }
                                                 If ah Then
                                                     systemHistoryMessageData.Add(answer)
                                                     ManageHistoryMessageSize()
                                                     _chatStateService.AddMessage("assistant", answer.content)
                                                     MemoryService.SaveConversationTurnAsync(oq, answer.content, _chatStateService.CurrentSessionId, GetOfficeAppType())
                                                     MemoryTurnRecorder.RecordConversationTurn(oq, answer.content, _chatStateService.CurrentSessionId, responseMode, ah, GetOfficeAppType())
                                                     If systemHistoryMessageData.Count = 3 Then
                                                         Dim sid = _chatStateService.CurrentSessionId
                                                         Dim title = If(oq?.Length > 80, oq.Substring(0, 80) & "...", If(oq, ""))
                                                         Dim snippet = If(oq?.Length > 200, oq.Substring(0, 200) & "...", If(oq, ""))
                                                         If Not String.IsNullOrWhiteSpace(sid) AndAlso Not String.IsNullOrWhiteSpace(title) Then
                                                             Try
                                                                 MemoryService.SaveSessionSummary(sid, title, snippet)
                                                             Catch ex As Exception
                                                                 Debug.WriteLine("SaveSessionSummary 失败: " & ex.Message)
                                                             End Try
                                                         End If
                                                     End If
                                                 End If
                                             End Sub

            Await HttpStreamSvc.SendStreamRequestAsync(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody, question, requestUuid, addHistory, responseMode, responseUuid)
            Await SaveFullWebPageAsync()
        Catch ex As Exception
            Debug.WriteLine("Send 请求失败: " & ex.Message & vbCrLf & ex.StackTrace)
            GlobalStatusStrip.ShowWarning("Ошибка запроса: " & ex.Message)
        Finally
        End Try

    End Function

    Private Sub ManageHistoryMessageSize()
        ' 如果历史消息数超过限制，有一条system和assistant，所以+2
        While systemHistoryMessageData.Count > contextLimit + 2
            ' 保留系统消息（第一条消息）
            If systemHistoryMessageData.Count > 2 Then
                ' 移除第二条消息（最早的非系统消息）
                systemHistoryMessageData.RemoveAt(2)
            End If
        End While
    End Sub


    Private Function CreateRequestBody(uuid As String, question As String, systemPrompt As String, addHistory As Boolean, ByRef ragCountOut As Integer, ByRef contextTraceOut As ChatContextTrace) As String
        Return ChatRequestOrchestrator.CreateRequestBody(uuid, question, systemPrompt, addHistory, ragCountOut, contextTraceOut)
    End Function


    ' MCP工具调用和流处理逻辑已移至 HttpStreamService

    ' _responseSelectionMap 代理到 _chatStateService
    Protected ReadOnly Property _responseSelectionMap As Dictionary(Of String, SelectionInfo)
        Get
            Return _chatStateService.ResponseSelectionMap
        End Get
    End Property

    ' 会话完成的钩子，可自行实现
    Protected Overridable Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)
        ' 处理续写模式的完成 - 显示续写预览界面
        If _responseModeMap.ContainsKey(_finalUuid) AndAlso _responseModeMap(_finalUuid) = "continuation" Then
            ExecuteJavaScriptAsyncJS($"showContinuationPreview('{_finalUuid}');")
        End If

        ' 处理模板渲染模式的完成 - 显示模板预览界面并完全隐藏代码块
        If _responseModeMap.ContainsKey(_finalUuid) AndAlso _responseModeMap(_finalUuid) = "template_render" Then
            ExecuteJavaScriptAsyncJS($"showTemplatePreview('{_finalUuid}');")
            ExecuteJavaScriptAsyncJS($"hideAllCodeBlockActions('{_finalUuid}');") ' 完全隐藏代码块操作栏
        End If

        ' 校对/排版模式 - 隐藏代码块的编辑和执行按钮（只保留复制）
        If _responseModeMap.ContainsKey(_finalUuid) Then
            Dim mode = _responseModeMap(_finalUuid)
            If mode = "proofread" OrElse mode = "reformat" Then
                ExecuteJavaScriptAsyncJS($"hideCodeActionButtons('{_finalUuid}');")
            End If
        End If

        ' === 自检Loop: Flush后校验（排版/校对场景的AI响应格式校验）===
        Try
            If _responseModeMap.ContainsKey(_finalUuid) Then
                Dim mode = _responseModeMap(_finalUuid)
                ' semantic_reformat 使用不同的验证逻辑，在 WordAi 的 ApplySemanticTaggingResult 中处理
                If mode = "reformat" OrElse mode = "proofread" Then
                    Dim aiResponse = allPlainMarkdownBuffer.ToString()
                    If Not String.IsNullOrWhiteSpace(aiResponse) Then
                        Dim expectedFormat = If(mode = "proofread", InstructionFormat.ProofreadJson, InstructionFormat.DslJson)
                        Dim execContext = New ExecutionContext()
                        Dim ignored = RunPostFlushValidationAsync(aiResponse, expectedFormat, execContext)
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SelfCheck] PostFlush校验异常: {ex.Message}")
        End Try

    End Sub

    Private Async Function RunPostFlushValidationAsync(aiResponse As String, expectedFormat As InstructionFormat, execContext As ExecutionContext) As Task
        Try
            Dim validation = Await SelfCheckLoopController.PostFlushValidateAsync(aiResponse, expectedFormat, execContext)
            If Not validation.IsValid Then
                Debug.WriteLine($"[SelfCheck] PostFlush校验失败: {String.Join(";", validation.Errors.Select(Function(e) e.Message))}")
                If validation.Errors.Any(Function(e) e.Level = ErrorLevel.Critical) Then
                    GlobalStatusStrip.ShowWarning($"Проверка формата ответа ИИ не пройдена; возможны ошибки в инструкциях")
                End If
            Else
                Debug.WriteLine($"[SelfCheck] PostFlush校验通过，解析到 {validation.ParsedInstructions.Count} 条指令")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SelfCheck] PostFlush校验异常: {ex.Message}")
        End Try
    End Function


    ' 执行js脚本的异步方法
    Public Async Function ExecuteJavaScriptAsyncJS(js As String) As Task
        Await WebViewBridge.ExecuteScriptAsync(js)
    End Function

    Private Async Function WaitForRendererMapAsync(uuid As String) As Task
        Await WebViewBridge.WaitForRendererMapAsync(uuid)
    End Function
    Private Function EscapeJavaScriptString(input As String) As String
        Return UtilsService.EscapeJavaScriptString(input)
    End Function



    ' 共用的HTTP请求方法 - 委托给 UtilsService
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Return Await UtilsService.SendHttpRequestAsync(apiUrl, apiKey, requestBody)
    End Function

    ' 加载本地HTML文件
    Public Async Function LoadLocalHtmlFile() As Task
        Await WebViewBridge.LoadHtmlFileAsync(ChatHtmlFilePath)
    End Function

    Public Async Function SaveFullWebPageAsync() As Task
        Try
            ' 1. 创建目录（同步操作，无需异步）

            Dim dir = Path.GetDirectoryName(ChatHtmlFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' 2. 获取 HTML（异步无阻塞）
            Dim htmlContent As String = Await GetFullHtmlContentAsync()

            ' 3. 保存文件（异步后台线程）
            Await Task.Run(Sub()
                               Dim fullHtml As String = "<!DOCTYPE html>" & Environment.NewLine & htmlContent
                               File.WriteAllText(
                $"{ChatHtmlFilePath}",
                HttpUtility.HtmlDecode(fullHtml),
                System.Text.Encoding.UTF8
            )
                           End Sub)

            Debug.WriteLine("保存成功")
        Catch ex As Exception
            Debug.WriteLine($"保存失败: {ex.Message}")
        End Try
    End Function

    Private Async Function GetFullHtmlContentAsync() As Task(Of String)
        Return Await WebViewBridge.GetFullHtmlContentAsync()
    End Function
    ' HistoryMessage 类已移至 Controls/Models/HistoryMessage.vb

    ' 注入辅助脚本
    Protected Sub InitializeWebView2Script()
        Debug.WriteLine("[DEBUG InitializeWebView2Script] Registering WebMessageReceived handler")
        ' 设置 Web 消息处理器
        AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived
        ' 注入 VSTO 桥接脚本
        ChatBrowser.ExecuteScriptAsync(UtilsService.GetVstoBridgeScript())
        ' 注入快捷问题配置
        InjectQuickQuestionsConfig()
    End Sub

    ' 注入快捷问题配置到前端
    Private Async Sub InjectQuickQuestionsConfig()
        Try
            Dim questions = ConfigPromptForm.GetQuickQuestionsList()
            Dim questionsJson = JsonConvert.SerializeObject(questions)
            Dim script = $"if(typeof updateQuickQuestions === 'function') {{ updateQuickQuestions({questionsJson}); }} else {{ window.predefinedPrompts = {questionsJson}; }}"
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
        Catch ex As Exception
            Debug.WriteLine($"注入快捷问题失败: {ex.Message}")
        End Try
    End Sub

    ' 选中内容发送到聊天区
    Public Async Sub AddSelectedContentItem(sheetName As String, address As String)
        Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
    $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})"
)
    End Sub


    ' VBA 异常处理 - 委托给 UtilsService
    Protected Shared Sub VBAxceptionHandle(ex As Runtime.InteropServices.COMException)
        UtilsService.HandleVbaException(ex)
    End Sub


    Protected Overridable Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
    End Sub

    ''' <summary>
    ''' 处理排版JSON解析失败的重试请求
    ''' </summary>
    Protected Overridable Async Sub HandleRetryReformat(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid")?.ToString(), "")
            Dim errorMsg As String = If(jsonDoc("error")?.ToString(), "Формат не соответствует стандарту")

            If String.IsNullOrEmpty(uuid) Then
                GlobalStatusStrip.ShowWarning("Ошибка повтора: отсутствует uuid")
                Return
            End If

            Dim retryCount As Integer = ReformatSvc.GetRetryCount(uuid)

            If retryCount >= 1 Then
                GlobalStatusStrip.ShowWarning("Достигнут предел повторных попыток форматирования")
                Return
            End If

            ReformatSvc.IncrementRetryCount(uuid)

            ' 构建重试提示
            Dim retryPrompt As New System.Text.StringBuilder()
            retryPrompt.AppendLine("JSON в твоём предыдущем ответе содержит ошибку; исправь и верни заново.")
            retryPrompt.AppendLine()
            retryPrompt.AppendLine($"Сообщение об ошибке: {errorMsg}")
            retryPrompt.AppendLine()
            retryPrompt.AppendLine("Обрати внимание на следующие требования к формату JSON:")
            retryPrompt.AppendLine("1. Все строки должны использовать двойные кавычки("")")
            retryPrompt.AppendLine("2. Не ставь запятую после последнего элемента массива или объекта")
            retryPrompt.AppendLine("3. Имена свойств должны быть заключены в двойные кавычки")
            retryPrompt.AppendLine("4. Не включай комментарии в JSON")
            retryPrompt.AppendLine("5. Убедись, что все скобки правильно закрыты")
            retryPrompt.AppendLine()
            retryPrompt.AppendLine("Верни только исправленный чистый JSON, без пояснений и маркеров блоков кода.")

            GlobalStatusStrip.ShowWarning("Не удалось разобрать JSON, повторяю попытку...")

            ' 发送重试请求
            Await Send(retryPrompt.ToString(), "", False, "reformat")

        Catch ex As Exception
            Debug.WriteLine("HandleRetryReformat 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("Ошибка повтора: " & ex.Message)
        End Try
    End Sub


#Region "语义排版Handler"

    ''' <summary>
    ''' 上传.docx模板文件并解析为SemanticStyleMapping
    ''' </summary>
    Protected Sub HandleUploadDocxTemplate()
        ReformatSvc.HandleUploadDocxTemplate()
    End Sub

    ''' <summary>
    ''' 从指定路径解析.docx模板
    ''' </summary>
    Protected Overridable Sub HandleUploadDocxTemplateFromPath(filePath As String)
        ' 默认不支持，由WordAi子类覆盖实现
        GlobalStatusStrip.ShowWarning("Текущее приложение не поддерживает разбор шаблонов Word")
    End Sub

    ''' <summary>
    ''' 删除docx语义映射
    ''' </summary>
    Protected Sub HandleDeleteDocxMapping(jsonDoc As JObject)
        ReformatSvc.HandleDeleteDocxMapping(jsonDoc)
    End Sub

    ''' <summary>
    ''' 撤销排版（根据不同的Office应用使用正确的撤销方式）
    ''' </summary>
    Protected Overridable Sub HandleUndoReformat()
        Try
            Dim appInfo = GetApplication()
            If appInfo Is Nothing Then Return

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("Не удалось получить объект приложения Office; попробуйте отменить по Ctrl+Z")
                Return
            End If

            Dim appName = GetOfficeApplicationName()
            Debug.WriteLine($"HandleUndoReformat: appName={appName}")

            Select Case appName
                Case "Word"
                    ' Word: 支持 UndoRecord，撤销入口在 Document.Undo
                    Try
                        officeApp.ActiveDocument.Undo(1)
                        GlobalStatusStrip.ShowInfo("Форматирование отменено")
                    Catch ex As Exception
                        Debug.WriteLine($"Word Document.Undo 失败，尝试 CommandBars Undo: {ex.Message}")
                        Try
                            officeApp.CommandBars.ExecuteMso("Undo")
                            GlobalStatusStrip.ShowInfo("Форматирование отменено")
                        Catch ex2 As Exception
                            Debug.WriteLine($"Word CommandBars 撤销也失败: {ex2.Message}")
                            GlobalStatusStrip.ShowWarning("Не удалось отменить форматирование, нажмите Ctrl+Z вручную")
                        End Try
                    End Try

                Case "PowerPoint"
                    ' PowerPoint: 不支持 UndoRecord，只能撤销最近一次操作
                    Try
                        ' PowerPoint Application 没有 ActiveDocument 属性，使用 ActivePresentation
                        officeApp.ActivePresentation.Undo()
                        GlobalStatusStrip.ShowWarning("Форматирование отменено (в PPT отмена ограничена; если изменения не откатились полностью, нажмите Ctrl+Z несколько раз)")
                    Catch ex As Exception
                        Debug.WriteLine($"PowerPoint 撤销失败: {ex.Message}")
                        Try
                            ' 备选：通过 CommandBars 触发标准撤销
                            officeApp.CommandBars.ExecuteMso("Undo")
                            GlobalStatusStrip.ShowWarning("Форматирование отменено (в PPT отмена ограничена; если изменения не откатились полностью, нажмите Ctrl+Z несколько раз)")
                        Catch ex2 As Exception
                            Debug.WriteLine($"PowerPoint CommandBars 撤销也失败: {ex2.Message}")
                            GlobalStatusStrip.ShowWarning("Не удалось отменить форматирование в PPT; отменяйте вручную несколькими нажатиями Ctrl+Z")
                        End Try
                    End Try

                Case "Excel"
                    ' Excel: 排版功能尚在开发中，但保留撤销入口
                    Try
                        ' Excel 使用 ActiveWorkbook，没有 ActiveDocument
                        officeApp.ActiveWorkbook.Undo()
                        GlobalStatusStrip.ShowInfo("Форматирование отменено")
                    Catch ex As Exception
                        Debug.WriteLine($"Excel 撤销失败: {ex.Message}")
                        Try
                            officeApp.CommandBars.ExecuteMso("Undo")
                            GlobalStatusStrip.ShowInfo("Форматирование отменено")
                        Catch ex2 As Exception
                            Debug.WriteLine($"Excel CommandBars 撤销也失败: {ex2.Message}")
                            GlobalStatusStrip.ShowWarning("Не удалось отменить форматирование, нажмите Ctrl+Z вручную")
                        End Try
                    End Try

                Case Else
                    ' 未知应用类型，尝试通用方式
                    Try
                        officeApp.Undo()
                        GlobalStatusStrip.ShowInfo("Форматирование отменено")
                    Catch ex As Exception
                        Debug.WriteLine($"通用撤销失败: {ex.Message}")
                        GlobalStatusStrip.ShowWarning("Не удалось отменить форматирование; попробуйте нажать Ctrl+Z")
                    End Try
            End Select

        Catch ex As Exception
            Debug.WriteLine($"HandleUndoReformat 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"Не удалось отменить форматирование: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 智能排版 v2：应用排版方案（由派生类覆写实现具体Office应用逻辑）
    ''' </summary>
    Protected Overridable Sub HandleApplySmartReformat(jsonDoc As JObject)
        Debug.WriteLine("HandleApplySmartReformat not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：撤销排版（从快照恢复，由派生类覆写实现具体Office应用逻辑）
    ''' </summary>
    Protected Overridable Sub HandleUndoReformat(jsonDoc As JObject)
        Debug.WriteLine("HandleUndoReformat not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：微调排版方案（由派生类覆写实现具体Office应用逻辑）
    ''' </summary>
    Protected Overridable Sub HandleRefineSmartReformat(jsonDoc As JObject)
        Debug.WriteLine("HandleRefineSmartReformat not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：切换排版模板/标准（由派生类覆写实现具体逻辑）
    ''' </summary>
    Protected Overridable Sub HandleSwitchReformatTemplate(jsonDoc As JObject)
        Debug.WriteLine("HandleSwitchReformatTemplate not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：显示排版前后对比（由派生类覆写实现具体逻辑）
    ''' </summary>
    Protected Overridable Sub HandlePreviewReformatCompare(jsonDoc As JObject)
        Debug.WriteLine("HandlePreviewReformatCompare not overridden in derived class")
    End Sub


    ''' <summary>
    ''' 校对专注模式消息处理（由派生类覆写实现具体逻辑）
    ''' </summary>
    Protected Overridable Sub HandleProofreadFocusMode(jsonDoc As JObject)
        Debug.WriteLine("HandleProofreadFocusMode not overridden in derived class")
    End Sub

#End Region

End Class
