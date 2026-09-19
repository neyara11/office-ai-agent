' ShareRibbon\Services\Memory\RuleBasedMemoryExtractor.vb
' Conservative first-pass memory extraction. This is intentionally replaceable by an LLM extractor.

Imports Newtonsoft.Json.Linq

Public Class RuleBasedMemoryExtractor
    Implements IMemoryExtractor

    Public Function ExtractMemories(events As List(Of ConversationEventRecord), Optional contextJson As String = Nothing) As List(Of MemoryItemRecord) Implements IMemoryExtractor.ExtractMemories
        Dim results As New List(Of MemoryItemRecord)()
        If events Is Nothing OrElse events.Count = 0 Then Return results

        Dim userEvent = events.FirstOrDefault(Function(e) String.Equals(e.Role, "user", StringComparison.OrdinalIgnoreCase))
        Dim assistantEvent = events.FirstOrDefault(Function(e) String.Equals(e.Role, "assistant", StringComparison.OrdinalIgnoreCase))
        Dim payload = ParsePayload(contextJson)
        Dim memory = BuildMemoryItem(userEvent, assistantEvent, payload)
        If memory IsNot Nothing Then results.Add(memory)
        Return results
    End Function

    Private Shared Function BuildMemoryItem(userEvent As ConversationEventRecord, assistantEvent As ConversationEventRecord, payload As JObject) As MemoryItemRecord
        Dim userText = If(userEvent?.Content, "").Trim()
        Dim assistantText = If(assistantEvent?.Content, "").Trim()
        If userText.Length < 4 AndAlso assistantText.Length < 10 Then Return Nothing

        Dim responseMode = GetPayloadText(payload, "response_mode")
        Dim memoryType = DetectMemoryType(userText, assistantText, responseMode)
        Dim scope = DetectScope(memoryType, responseMode)
        Dim content = BuildMemoryContent(userText, assistantText, memoryType)
        If String.IsNullOrWhiteSpace(content) OrElse content.Length < 10 Then Return Nothing

        Dim sourceEvent = If(assistantEvent, userEvent)
        Dim appType = FirstNonEmpty(sourceEvent?.AppType, GetPayloadText(payload, "app_type"))
        Dim documentId = If(sourceEvent?.DocumentId, "")
        Dim importance = UnifiedMemoryService.CalculateImportance(content, memoryType, Nothing)
        Dim confidence = If(memoryType = "preference" OrElse memoryType = "format_rule", 0.8R, 0.65R)

        Return New MemoryItemRecord With {
            .SourceEventId = sourceEvent?.EventId,
            .Scope = scope,
            .AppType = appType,
            .DocumentId = documentId,
            .ProjectId = "",
            .MemoryType = memoryType,
            .Content = content,
            .Summary = TrimForStorage(content, 240),
            .Confidence = confidence,
            .Importance = importance,
            .Status = "active"
        }
    End Function

    Private Shared Function DetectMemoryType(userText As String, assistantText As String, responseMode As String) As String
        Dim combined = (If(userText, "") & " " & If(assistantText, "") & " " & If(responseMode, "")).ToLowerInvariant()

        If combined.Contains("我希望") OrElse combined.Contains("以后") OrElse combined.Contains("默认") OrElse
           combined.Contains("不要") OrElse combined.Contains("请记住") OrElse combined.Contains("偏好") OrElse
           combined.Contains("я хочу") OrElse combined.Contains("впредь") OrElse combined.Contains("по умолчанию") OrElse
           combined.Contains("запомни") OrElse combined.Contains("предпочитаю") OrElse combined.Contains("предпочтение") OrElse
           combined.Contains("не нужно") OrElse combined.Contains("не делай") OrElse combined.Contains("не использ") Then
            Return "preference"
        End If

        If combined.Contains("reformat") OrElse combined.Contains("排版") OrElse combined.Contains("格式") OrElse
           combined.Contains("公文") OrElse combined.Contains("红头") OrElse combined.Contains("gb/t") OrElse
           combined.Contains("формат") OrElse combined.Contains("оформ") OrElse combined.Contains("стил") OrElse
           combined.Contains("заголов") OrElse combined.Contains("шрифт") OrElse combined.Contains("стандарт") Then
            Return "format_rule"
        End If

        If Not String.IsNullOrWhiteSpace(assistantText) AndAlso assistantText.Length > 80 Then
            Return "solution"
        End If

        Return "fact"
    End Function

    Private Shared Function DetectScope(memoryType As String, responseMode As String) As String
        If memoryType = "preference" Then Return "user"
        If memoryType = "format_rule" Then Return "document"
        If String.Equals(responseMode, "reformat", StringComparison.OrdinalIgnoreCase) Then Return "document"
        Return "session"
    End Function

    Private Shared Function BuildMemoryContent(userText As String, assistantText As String, memoryType As String) As String
        Dim userPart = TrimForStorage(userText, 500)
        Dim assistantPart = TrimForStorage(assistantText, 700)

        Select Case memoryType
            Case "preference"
                Return $"Предпочтение/ограничение пользователя: {userPart}"
            Case "format_rule"
                Return $"Контекст форматирования: запрос={userPart}; результат={assistantPart}"
            Case "solution"
                Return $"Проблема: {userPart}{vbLf}Решение: {assistantPart}"
            Case Else
                Return $"Факт диалога: пользователь={userPart}; ассистент={assistantPart}"
        End Select
    End Function

    Private Shared Function ParsePayload(payloadJson As String) As JObject
        If String.IsNullOrWhiteSpace(payloadJson) Then Return New JObject()
        Try
            Return JObject.Parse(payloadJson)
        Catch
            Return New JObject()
        End Try
    End Function

    Private Shared Function GetPayloadText(payload As JObject, key As String) As String
        If payload Is Nothing OrElse payload(key) Is Nothing Then Return ""
        Return payload(key).ToString()
    End Function

    Private Shared Function FirstNonEmpty(primary As String, fallback As String) As String
        If Not String.IsNullOrWhiteSpace(primary) Then Return primary
        Return If(fallback, "")
    End Function

    Private Shared Function TrimForStorage(value As String, maxLen As Integer) As String
        Dim text = If(value, "").Trim()
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "..."
    End Function
End Class
