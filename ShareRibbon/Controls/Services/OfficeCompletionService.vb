Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq

''' <summary>
''' Office原生补全服务 - 提供Word和PPT的AI补全功能
''' </summary>
Public Class OfficeCompletionService
    Private Shared _instance As OfficeCompletionService
    Private Shared ReadOnly _lock As New Object()
    
    Private _isEnabled As Boolean = False
    Private _currentCompletions As List(Of String)
    
    ' 取消令牌源（用于取消进行中的请求）
    Private _cancellationTokenSource As CancellationTokenSource
    
    ''' <summary>
    ''' 获取单例实例
    ''' </summary>
    Public Shared ReadOnly Property Instance As OfficeCompletionService
        Get
            If _instance Is Nothing Then
                SyncLock _lock
                    If _instance Is Nothing Then
                        _instance = New OfficeCompletionService()
                    End If
                End SyncLock
            End If
            Return _instance
        End Get
    End Property
    
    Private Sub New()
        _currentCompletions = New List(Of String)()
        _cancellationTokenSource = New CancellationTokenSource()
    End Sub
    
    ''' <summary>
    ''' 启用/禁用补全服务
    ''' </summary>
    Public Property Enabled As Boolean
        Get
            Return _isEnabled
        End Get
        Set(value As Boolean)
            _isEnabled = value
            If Not value Then
                CancelPendingRequest()
            End If
        End Set
    End Property
    
    ''' <summary>
    ''' 取消待处理的请求
    ''' </summary>
    Public Sub CancelPendingRequest()
        ' 取消进行中的 HTTP 请求
        If _cancellationTokenSource IsNot Nothing Then
            Try
                _cancellationTokenSource.Cancel()
            Catch
                ' 忽略取消异常
            End Try
        End If
        _cancellationTokenSource = New CancellationTokenSource()
    End Sub
    
    ''' <summary>
    ''' 调用LLM获取补全
    ''' </summary>
    Private Async Function GetCompletionsFromLLM(inputText As String, appType As String, 
                                                  Optional token As CancellationToken = Nothing) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            If token = Nothing Then token = _cancellationTokenSource.Token
            
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return completions
            End If
            
            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)
            
            Dim modelName = selectedModel.modelName
            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key
            
            ' 检查是否支持FIM
            If selectedModel.fimSupported AndAlso Not String.IsNullOrEmpty(selectedModel.fimUrl) Then
                completions = Await GetCompletionsWithFIM(inputText, selectedModel, apiKey, token)
            Else
                completions = Await GetCompletionsWithChat(inputText, appType, cfg, selectedModel, apiKey, token)
            End If
            
        Catch ex As OperationCanceledException
            Debug.WriteLine("[Completion] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsFromLLM 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 使用FIM API获取补全
    ''' </summary>
    Private Async Function GetCompletionsWithFIM(inputText As String, model As ConfigManager.ConfigItemModel, 
                                                  apiKey As String, token As CancellationToken) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            token.ThrowIfCancellationRequested()
            
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("prompt") = inputText
            requestObj("suffix") = ""
            requestObj("max_tokens") = 60  ' 减少 token 数量以加快响应
            requestObj("temperature") = 0.3
            requestObj("stream") = False
            
            Dim request As New HttpRequestMessage(HttpMethod.Post, model.fimUrl)
            request.Headers.Add("Authorization", "Bearer " & apiKey)
            request.Content = New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")
            
            Dim response = Await HttpClientPool.GetClient(model.fimUrl, TimeSpan.FromSeconds(8)).SendAsync(request, token)
            response.EnsureSuccessStatusCode()
            
            Dim responseBody = Await response.Content.ReadAsStringAsync()
            Dim jObj = JObject.Parse(responseBody)
            
            Dim text = jObj("choices")?(0)?("text")?.ToString()
            If Not String.IsNullOrWhiteSpace(text) Then
                ' 取第一行
                Dim firstLine = text.Trim().Split({vbCr, vbLf, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)(0)
                If firstLine.Length <= 100 Then
                    completions.Add(firstLine)
                End If
            End If
            
        Catch ex As OperationCanceledException
            Debug.WriteLine("[FIM] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsWithFIM 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 使用Chat API获取补全
    ''' </summary>
    Private Async Function GetCompletionsWithChat(inputText As String, appType As String,
                                                   cfg As ConfigManager.ConfigItem, model As ConfigManager.ConfigItemModel,
                                                   apiKey As String, token As CancellationToken) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            token.ThrowIfCancellationRequested()
            
            Dim systemPrompt = GetSystemPrompt(appType)
            
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("stream") = False
            requestObj("temperature") = 0.3
            
            Dim messages As New JArray()
            messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject() From {{"role", "user"}, {"content", $"Дополни следующий текст (верни только продолжение, не повторяй исходный текст):{vbCrLf}{inputText}"}})
            requestObj("messages") = messages
            
            ' 使用按服务商配置的 HttpClient
            Dim request As New HttpRequestMessage(HttpMethod.Post, cfg.url)
            request.Headers.Add("Authorization", "Bearer " & apiKey)
            request.Content = New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")
            
            Dim response = Await HttpClientPool.GetClient(cfg.url, TimeSpan.FromSeconds(8)).SendAsync(request, token)
            response.EnsureSuccessStatusCode()
            
            Dim responseBody = Await response.Content.ReadAsStringAsync()
            Dim jObj = JObject.Parse(responseBody)
            
            Dim msg = jObj("choices")?(0)?("message")?("content")?.ToString()
            If Not String.IsNullOrWhiteSpace(msg) AndAlso msg.Length <= 100 Then
                completions.Add(msg.Trim())
            End If
            
        Catch ex As OperationCanceledException
            Debug.WriteLine("[Chat] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsWithChat 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 获取系统提示词
    ''' </summary>
    Private Function GetSystemPrompt(appType As String) As String
        Select Case appType.ToLower()
            Case "word"
                Return "Ты ассистент автодополнения для документов Word. По вводимому тексту предскажи и дополни продолжение.
Правила:
1. Возвращай только продолжение, не повторяй уже введённый текст
2. Продолжение должно быть естественным и соответствовать контексту
3. Обычная длина продолжения — 10–50 символов
4. Сохраняй язык и стиль исходного текста, отвечай только на русском"
            
            Case "powerpoint"
                Return "Ты ассистент автодополнения для презентаций PowerPoint. По вводимому заголовку или тексту предскажи и дополни продолжение.
Правила:
1. Возвращай только продолжение, не повторяй уже введённый текст
2. Продолжение должно быть кратким и выразительным, подходящим для слайдов
3. Обычная длина продолжения — 10–30 символов
4. Для заголовков сохраняй краткость, отвечай только на русском"
            
            Case Else
                Return "Ты ассистент автодополнения для документов Office. По вводимому тексту предскажи и дополни продолжение. Возвращай только продолжение, не повторяй исходный текст, отвечай только на русском."
        End Select
    End Function
    
    ''' <summary>
    ''' 直接获取补全（不带防抖，供外部调用）
    ''' </summary>
    Public Async Function GetCompletionsDirectAsync(inputText As String, appType As String, 
                                                     Optional token As CancellationToken = Nothing) As Task(Of List(Of String))
        If Not _isEnabled OrElse Not ChatSettings.EnableAutocomplete Then
            Return New List(Of String)()
        End If
        
        If token = Nothing Then token = _cancellationTokenSource.Token
        
        Dim completions = Await GetCompletionsFromLLM(inputText, appType, token)
        _currentCompletions = completions
        Return completions
    End Function
    
    ''' <summary>
    ''' 清除当前补全
    ''' </summary>
    Public Sub ClearCompletions()
        _currentCompletions.Clear()
    End Sub
End Class
