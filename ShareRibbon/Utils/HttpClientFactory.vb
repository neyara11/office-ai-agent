Imports System.Net.Http
Imports System.Threading

''' <summary>
''' 统一的 HttpClient 工厂。按服务商配置决定是否接受自签名证书。
''' </summary>
Public NotInheritable Class HttpClientFactory
    Private Sub New()
    End Sub

    Public Shared Function GetAuthority(apiUrl As String) As String
        If String.IsNullOrWhiteSpace(apiUrl) Then Return Nothing

        Try
            Dim uri = New Uri(apiUrl)
            Return $"{uri.Scheme}://{uri.Host}:{uri.Port}"
        Catch
            Return Nothing
        End Try
    End Function

    Public Shared Function IsInsecureTlsAllowed(apiUrl As String) As Boolean
        Dim authority = GetAuthority(apiUrl)
        If authority Is Nothing Then Return False

        Dim data = ConfigManager.ConfigData
        If data Is Nothing Then Return False

        For Each item In data
            If item Is Nothing OrElse Not item.allowInsecureTls Then Continue For
            If String.Equals(GetAuthority(item.url), authority, StringComparison.OrdinalIgnoreCase) Then Return True
        Next

        Return False
    End Function

    Public Shared Function AnyInsecureTlsAllowed() As Boolean
        Dim data = ConfigManager.ConfigData
        If data Is Nothing Then Return False

        For Each item In data
            If item IsNot Nothing AndAlso item.allowInsecureTls Then Return True
        Next

        Return False
    End Function

    ''' <summary>
    ''' Приводит введённый пользователем адрес к полному OpenAI-совместимому endpoint чата.
    ''' Принимает как базу ("https://host/api/v1"), так и уже полный путь (".../chat/completions").
    ''' </summary>
    Public Shared Function ResolveChatCompletionsUrl(url As String) As String
        If String.IsNullOrWhiteSpace(url) Then Return url

        Dim trimmed = url.Trim().TrimEnd("/"c)
        Dim lower = trimmed.ToLowerInvariant()

        ' Уже полный endpoint (в т.ч. Anthropic /messages) — не трогаем
        If lower.EndsWith("/chat/completions") OrElse lower.EndsWith("/messages") Then Return trimmed

        If lower.EndsWith("/v1") OrElse lower.EndsWith("/v1beta") Then
            Return trimmed & "/chat/completions"
        End If

        ' Голый хост без пути → стандартный /v1/chat/completions
        Dim uri As Uri = Nothing
        If Uri.TryCreate(trimmed, UriKind.Absolute, uri) AndAlso
           (String.IsNullOrEmpty(uri.AbsolutePath) OrElse uri.AbsolutePath = "/") Then
            Return trimmed & "/v1/chat/completions"
        End If

        ' Пути вида /api/v3, /api/paas/v4, /compatible-mode/v1 и т.п.
        Return trimmed & "/chat/completions"
    End Function

    ''' <summary>
    ''' Строит endpoint списка моделей из адреса чата или базы.
    ''' </summary>
    Public Shared Function ResolveModelsUrl(url As String) As String
        Dim chat = ResolveChatCompletionsUrl(url)
        If String.IsNullOrWhiteSpace(chat) Then Return chat

        Dim lower = chat.ToLowerInvariant()
        If lower.EndsWith("/chat/completions") Then
            Return chat.Substring(0, chat.Length - "/chat/completions".Length) & "/models"
        End If
        If lower.EndsWith("/messages") Then
            Return chat.Substring(0, chat.Length - "/messages".Length) & "/models"
        End If

        Return chat
    End Function

    ''' <summary>
    ''' Строит endpoint embeddings из адреса чата или базы.
    ''' </summary>
    Public Shared Function ResolveEmbeddingsUrl(url As String) As String
        Dim chat = ResolveChatCompletionsUrl(url)
        If String.IsNullOrWhiteSpace(chat) Then Return chat

        Dim lower = chat.ToLowerInvariant()
        If lower.EndsWith("/chat/completions") Then
            Return chat.Substring(0, chat.Length - "/chat/completions".Length) & "/embeddings"
        End If
        If lower.EndsWith("/embeddings") Then Return chat

        Return chat
    End Function

    ''' <summary>
    ''' Протокол TLS для HTTP-запросов плагина.
    ''' На .NET Framework 4.7.2 нельзя использовать SecurityProtocolType.SystemDefault: Schannel может
    ''' выбрать TLS 1.3, которого нет в перечислении SslProtocols, и запрос падает с
    ''' ArgumentException ("недействительное значение SslProtocolType") / WebException ReceiveFailure.
    ''' Поэтому версия TLS фиксируется явно — TLS 1.2. Если целевой endpoint принимает только TLS 1.3,
    ''' нужен переход на .NET Framework 4.8.1 или TLS-терминация reverse-proxy на 1.2.
    ''' </summary>
    Public Shared Function CreateHandler(apiUrl As String) As HttpClientHandler
        Dim handler As New HttpClientHandler()

        If IsInsecureTlsAllowed(apiUrl) Then
            handler.ServerCertificateCustomValidationCallback =
                Function(message, cert, chain, errors) True
        End If

        Return handler
    End Function

    Public Shared Function CreateClient(apiUrl As String) As HttpClient
        Dim client As New HttpClient(CreateHandler(apiUrl))
        client.Timeout = Timeout.InfiniteTimeSpan
        Return client
    End Function

    ''' <summary>
    ''' Полное описание исключения (тип, сообщение, WebException.Status) по всей цепочке InnerException.
    ''' </summary>
    Public Shared Function DescribeError(ex As Exception) As String
        If ex Is Nothing Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()
        Dim current As Exception = ex
        While current IsNot Nothing
            If sb.Length > 0 Then sb.Append(" -> ")
            sb.Append(current.GetType().Name).Append(": ").Append(current.Message)

            Dim webEx = TryCast(current, Net.WebException)
            If webEx IsNot Nothing Then sb.Append(" [").Append(webEx.Status.ToString()).Append("]")

            current = current.InnerException
        End While

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Каким прокси воспользуется .NET для указанного URL (важно: в процессе Office он может отличаться от терминала).
    ''' </summary>
    Public Shared Function DescribeProxy(apiUrl As String) As String
        Try
            Dim proxy = Net.WebRequest.DefaultWebProxy
            If proxy Is Nothing Then Return "proxy: none"

            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(apiUrl, UriKind.Absolute, uri) Then Return "proxy: unknown"

            Dim resolved = proxy.GetProxy(uri)
            If resolved Is Nothing Then Return "proxy: none"
            Return "proxy: " & resolved.ToString()
        Catch
            Return "proxy: error"
        End Try
    End Function
End Class
