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
End Class
