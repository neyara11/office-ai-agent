Imports System.Collections.Concurrent
Imports System.Net.Http
Imports System.Threading

Public NotInheritable Class HttpClientPool
    Private Shared ReadOnly _clients As New ConcurrentDictionary(Of String, HttpClient)(StringComparer.OrdinalIgnoreCase)

    Private Sub New()
    End Sub

    Public Shared Function GetClient(apiUrl As String) As HttpClient
        Return GetClient(apiUrl, Nothing)
    End Function

    Public Shared Function GetClient(apiUrl As String, timeout As TimeSpan?) As HttpClient
        Dim key = GetClientKey(apiUrl, timeout)
        Dim factory As Func(Of String, HttpClient) =
            Function(clientKey As String) As HttpClient
                Dim client = HttpClientFactory.CreateClient(apiUrl)
                client.Timeout = If(timeout.HasValue, timeout.Value, System.Threading.Timeout.InfiniteTimeSpan)
                Return client
            End Function
        Return _clients.GetOrAdd(key, factory)
    End Function

    Private Shared Function GetClientKey(apiUrl As String, timeout As TimeSpan?) As String
        Dim authority = HttpClientFactory.GetAuthority(apiUrl)
        If String.IsNullOrWhiteSpace(authority) Then authority = If(apiUrl, "").Trim()

        Dim policy = If(HttpClientFactory.IsInsecureTlsAllowed(apiUrl), "insecure", "secure")
        Dim timeoutTicks = If(timeout.HasValue, timeout.Value.Ticks, -1L)

        Return $"{authority}|{policy}|{timeoutTicks}"
    End Function
End Class
