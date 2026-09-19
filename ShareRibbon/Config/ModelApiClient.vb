Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' 模型API客户端
''' 用于调用 /v1/models API 获取远端模型列表
''' </summary>
Public Class ModelApiClient

    ''' <summary>
    ''' 异步获取模型列表
    ''' </summary>
    ''' <param name="apiUrl">API端点URL (chat/completions 端点)</param>
    ''' <param name="apiKey">API密钥</param>
    ''' <returns>模型名称列表</returns>
    Public Shared Async Function GetModelsAsync(apiUrl As String, apiKey As String) As Task(Of List(Of String))
        Try
            ' 构造 /v1/models 端点
            Dim modelsUrl As String = GetModelsEndpoint(apiUrl)
            If String.IsNullOrEmpty(modelsUrl) Then
                Return New List(Of String)()
            End If

            ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault

            Dim client = HttpClientPool.GetClient(modelsUrl)
            Using request As New HttpRequestMessage(HttpMethod.Get, modelsUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)

                ' 某些服务商需要特殊的请求头
                If apiUrl.Contains("anthropic.com") Then
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                    request.Headers.Authorization = Nothing
                End If

                Using timeoutCts As New System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(30))
                    Using response = Await client.SendAsync(request, timeoutCts.Token)

                If response.IsSuccessStatusCode Then
                    Dim jsonContent = Await response.Content.ReadAsStringAsync()
                    Return ParseModelsResponse(jsonContent, apiUrl)
                Else
                    Debug.WriteLine($"获取模型列表失败: {response.StatusCode}")
                    End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"获取模型列表异常: {ex.Message}")
        End Try

        Return New List(Of String)()
    End Function

    ''' <summary>
    ''' 根据chat/completions端点构造models端点
    ''' </summary>
    Private Shared Function GetModelsEndpoint(apiUrl As String) As String
        If String.IsNullOrEmpty(apiUrl) Then Return ""

        ' 处理不同服务商的端点差异
        If apiUrl.Contains("/chat/completions") Then
            Return apiUrl.Replace("/chat/completions", "/models")
        ElseIf apiUrl.Contains("/v1/messages") Then
            ' Anthropic 的 models 端点
            Return apiUrl.Replace("/v1/messages", "/v1/models")
        End If

        ' 尝试直接替换为 /models
        Dim uri As New Uri(apiUrl)
        Return $"{uri.Scheme}://{uri.Host}:{uri.Port}/v1/models"
    End Function

    ''' <summary>
    ''' 解析模型列表响应
    ''' </summary>
    Private Shared Function ParseModelsResponse(jsonContent As String, apiUrl As String) As List(Of String)
        Dim modelsList As New List(Of String)()

        Try
            Dim jsonObj = JObject.Parse(jsonContent)

            ' 标准 OpenAI 格式: { "data": [...] }
            If jsonObj("data") IsNot Nothing Then
                For Each modelItem In jsonObj("data")
                    Dim modelId = modelItem("id")?.ToString()
                    If Not String.IsNullOrEmpty(modelId) Then
                        modelsList.Add(modelId)
                    End If
                Next
            End If

            ' Ollama 特殊格式: { "models": [...] }
            If jsonObj("models") IsNot Nothing AndAlso modelsList.Count = 0 Then
                For Each modelItem In jsonObj("models")
                    Dim modelName = modelItem("name")?.ToString()
                    If Not String.IsNullOrEmpty(modelName) Then
                        modelsList.Add(modelName)
                    End If
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"解析模型列表响应异常: {ex.Message}")
        End Try

        Return modelsList
    End Function

    ''' <summary>
    ''' 同步获取模型列表。P0-3: SyncOverAsync 线程池桥接，避免 UI 上下文 GetResult 死锁。
    ''' Prefer GetModelsAsync from async UI handlers when possible.
    ''' </summary>
    Public Shared Function GetModelsSync(apiUrl As String, apiKey As String) As List(Of String)
        Try
            Dim models = SyncOverAsync.Run(Function() GetModelsAsync(apiUrl, apiKey), 60000)
            Return If(models, New List(Of String)())
        Catch ex As Exception
            Debug.WriteLine($"同步获取模型列表异常: {ex.Message}")
            Return New List(Of String)()
        End Try
    End Function

End Class
