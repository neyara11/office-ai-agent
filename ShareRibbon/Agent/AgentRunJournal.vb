' ShareRibbon\Agent\AgentRunJournal.vb
' Structured per-run diagnostics for failed agent executions.
' Writes JSONL next to the AppLogger log; never throws and never records visual evidence payloads.

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Append-only diagnostics journal for one agent run: prompts, raw model responses,
''' parsed tool calls, host results and repair attempts. Every string is passed through
''' AppLogger.Redact, truncated, and visual evidence / base64 payloads are stripped.
''' </summary>
Public NotInheritable Class AgentRunJournal
    Private Sub New()
    End Sub

    Private Shared ReadOnly _lock As New Object()
    Private Const MaxStringLength As Integer = 8000
    Private Const MaxLineLength As Integer = 32000

    Public Shared Sub Record(runId As String, phase As String, payload As Object)
        If Not IsEnabled() Then Return
        Try
            Dim data As JToken
            If payload Is Nothing Then
                data = JValue.CreateNull()
            ElseIf TypeOf payload Is JToken Then
                data = CType(payload, JToken)
            Else
                data = JToken.FromObject(payload)
            End If

            Sanitize(data)

            Dim entry As New JObject From {
                {"ts", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")},
                {"runId", If(runId, "")},
                {"phase", If(phase, "")},
                {"correlationId", AppLogger.CorrelationId},
                {"data", data}
            }

            Dim line = AppLogger.Redact(entry.ToString(Formatting.None))
            If line.Length > MaxLineLength Then
                line = line.Substring(0, MaxLineLength) & "...(truncated)"
            End If

            Dim logPath = AppLogger.GetLogFilePath()
            Dim dirPath = Path.GetDirectoryName(logPath)
            If String.IsNullOrWhiteSpace(dirPath) Then Return
            If Not Directory.Exists(dirPath) Then Directory.CreateDirectory(dirPath)

            Dim filePath = Path.Combine(dirPath, $"agent-runs-{DateTime.Now:yyyyMMdd}.jsonl")
            SyncLock _lock
                File.AppendAllText(filePath, line & Environment.NewLine, Encoding.UTF8)
            End SyncLock
        Catch
            ' Diagnostics must never break the run.
        End Try
    End Sub

    ''' <summary>
    ''' Diagnostics are on by default for local installs; set OFFICE_AI_AGENT_DIAGNOSTICS=0
    ''' to disable writing model responses and tool parameters to disk.
    ''' </summary>
    Public Shared Function IsEnabled() As Boolean
        Try
            Dim value = Environment.GetEnvironmentVariable("OFFICE_AI_AGENT_DIAGNOSTICS")
            If String.IsNullOrWhiteSpace(value) Then Return True
            Return Not (value.Trim() = "0" OrElse value.Trim().Equals("false", StringComparison.OrdinalIgnoreCase))
        Catch
            Return True
        End Try
    End Function

    ''' <summary>Stable short hash used to correlate prompts without storing document context.</summary>
    Public Shared Function HashText(text As String) As String
        Try
            If String.IsNullOrEmpty(text) Then Return ""
            Using sha = System.Security.Cryptography.SHA256.Create()
                Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text))
                Dim sb As New StringBuilder()
                For i = 0 To 7
                    sb.Append(bytes(i).ToString("x2"))
                Next
                Return sb.ToString()
            End Using
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Truncates long strings and drops binary/visual payload keys entirely.
    ''' </summary>
    Private Shared Sub Sanitize(token As JToken)
        If token Is Nothing Then Return

        Select Case token.Type
            Case JTokenType.Object
                Dim obj = CType(token, JObject)
                For Each propertyItem In obj.Properties().ToList()
                    If IsForbiddenKey(propertyItem.Name) Then
                        propertyItem.Remove()
                        Continue For
                    End If
                    Sanitize(propertyItem.Value)
                Next
            Case JTokenType.Array
                For Each child In CType(token, JArray)
                    Sanitize(child)
                Next
            Case JTokenType.String
                Dim value = token.ToString()
                If value.Length > MaxStringLength Then
                    CType(token, JValue).Value = value.Substring(0, MaxStringLength) & "...(truncated)"
                End If
        End Select
    End Sub

    Private Shared Function IsForbiddenKey(name As String) As Boolean
        If String.IsNullOrWhiteSpace(name) Then Return False
        Select Case name.Trim().ToLowerInvariant()
            Case "dataurl", "base64", "visualevidence", "screenshot", "imagebytes"
                Return True
            Case Else
                Return False
        End Select
    End Function
End Class
