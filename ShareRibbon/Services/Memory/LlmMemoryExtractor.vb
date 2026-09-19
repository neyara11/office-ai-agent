' ShareRibbon\Services\Memory\LlmMemoryExtractor.vb
' LLM-backed memory extraction with rule-based fallback.

Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class LlmMemoryExtractor
    Implements IMemoryExtractor

    Private ReadOnly _fallback As IMemoryExtractor
    Private Shared _lastFailureTime As DateTime? = Nothing
    Private Shared ReadOnly _failureCooldown As TimeSpan = TimeSpan.FromMinutes(10)

    Public Sub New()
        Me.New(New RuleBasedMemoryExtractor())
    End Sub

    Public Sub New(fallback As IMemoryExtractor)
        _fallback = If(fallback, New RuleBasedMemoryExtractor())
    End Sub

    Public Function ExtractMemories(events As List(Of ConversationEventRecord), Optional contextJson As String = Nothing) As List(Of MemoryItemRecord) Implements IMemoryExtractor.ExtractMemories
        Dim fallbackMemories = _fallback.ExtractMemories(events, contextJson)

        If Not CanUseLlm() Then
            Return fallbackMemories
        End If

        Try
            ' Sync IMemoryExtractor boundary: run async extractor off UI SynchronizationContext.
            Dim llmMemories = SyncOverAsync.Run(Function() ExtractMemoriesAsync(events, contextJson), 60000)
            If llmMemories Is Nothing OrElse llmMemories.Count = 0 Then Return fallbackMemories
            Return MergeMemories(llmMemories, fallbackMemories)
        Catch ex As Exception
            MarkFailure()
            Debug.WriteLine($"[LlmMemoryExtractor] LLM 提取失败，回退规则提取: {ex.Message}")
            Return fallbackMemories
        End Try
    End Function

    Private Shared Function CanUseLlm() As Boolean
        If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) Then Return False
        If String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) Then Return False
        If String.IsNullOrWhiteSpace(ConfigSettings.ModelName) Then Return False
        If _lastFailureTime.HasValue AndAlso (DateTime.UtcNow - _lastFailureTime.Value) < _failureCooldown Then Return False

        Return True
    End Function

    Private Async Function ExtractMemoriesAsync(events As List(Of ConversationEventRecord), contextJson As String) As Task(Of List(Of MemoryItemRecord))
        Dim prompt = BuildUserPrompt(events, contextJson)
        If String.IsNullOrWhiteSpace(prompt) Then Return New List(Of MemoryItemRecord)()

        Dim gatewayResponse = Await AiGateway.SendChatAsync(New AiRequestOptions With {
            .ApiUrl = ConfigSettings.ApiUrl,
            .ApiKey = ConfigSettings.ApiKey,
            .ModelName = ConfigSettings.ModelName,
            .Platform = ConfigSettings.platform,
            .ReasoningMode = ReasoningRequestHelper.ReasoningDisabled,
            .SystemPrompt = BuildSystemPrompt(),
            .UserPrompt = prompt,
            .Temperature = 0.1R,
            .MaxTokens = 1200,
            .TimeoutSeconds = 30
        })

        If gatewayResponse Is Nothing OrElse Not gatewayResponse.Success Then
            Debug.WriteLine($"[LlmMemoryExtractor] API failed: {If(gatewayResponse Is Nothing, "empty response", gatewayResponse.ErrorMessage)}")
            MarkFailure()
            Return New List(Of MemoryItemRecord)()
        End If

        Return ParseMemories(gatewayResponse.Content, events)
    End Function

    Private Shared Sub MarkFailure()
        _lastFailureTime = DateTime.UtcNow
    End Sub

    Private Shared Function BuildSystemPrompt() As String
        Return "Ты извлекатель долговременной памяти Office AI Agent. Твоя задача — извлечь из одного раунда диалога пользователя и ассистента небольшое число стабильных переиспользуемых структурированных воспоминаний." & vbCrLf &
               "Извлекай только то, что явно пригодится в будущем: предпочтения пользователя, постоянные правила оформления, подтверждённые решения, факты о проекте, привычки работы с документами Office." & vbCrLf &
               "Не сохраняй разовые светские реплики, временные вопросы, чувствительные ключи, полные приватные данные и неподтверждённые догадки." & vbCrLf &
               "Возвращай только JSON, без Markdown. Формат — массив, поля каждого элемента: " & vbCrLf &
               "[{""scope"":""user|document|project|session"",""memory_type"":""preference|format_rule|solution|fact|workflow"",""content"":""воспоминание на русском, готовое к вставке в контекст"",""summary"":""краткое резюме"",""confidence"":0.0-1.0,""importance"":0.0-1.0,""expires_at"":""""}]" & vbCrLf &
               "Если сохранять нечего, верни []."
    End Function

    Private Shared Function BuildUserPrompt(events As List(Of ConversationEventRecord), contextJson As String) As String
        If events Is Nothing OrElse events.Count = 0 Then Return ""

        Dim sb As New StringBuilder()
        If Not String.IsNullOrWhiteSpace(contextJson) Then
            sb.AppendLine("JSON контекста:")
            sb.AppendLine(Truncate(contextJson, 2000))
            sb.AppendLine()
        End If

        sb.AppendLine("События диалога:")
        For Each evt In events
            If evt Is Nothing OrElse String.IsNullOrWhiteSpace(evt.Content) Then Continue For
            sb.AppendLine($"[{If(evt.Role, evt.EventType)}] app={evt.AppType}; document={evt.DocumentId}; event_id={evt.EventId}")
            sb.AppendLine(Truncate(evt.Content, 3500))
            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function

    Private Shared Function ParseMemories(content As String, events As List(Of ConversationEventRecord)) As List(Of MemoryItemRecord)
        Dim results As New List(Of MemoryItemRecord)()
        Dim jsonText = ExtractJsonArray(content)
        If String.IsNullOrWhiteSpace(jsonText) Then Return results

        Dim arr = JArray.Parse(jsonText)
        Dim sourceEvent = If(events, New List(Of ConversationEventRecord)()).FirstOrDefault(Function(e) e IsNot Nothing AndAlso String.Equals(e.Role, "user", StringComparison.OrdinalIgnoreCase))
        If sourceEvent Is Nothing Then sourceEvent = If(events, New List(Of ConversationEventRecord)()).FirstOrDefault()

        For Each token In arr
            Dim obj = TryCast(token, JObject)
            If obj Is Nothing Then Continue For

            Dim memoryContent = obj("content")?.ToString()
            If String.IsNullOrWhiteSpace(memoryContent) Then Continue For

            results.Add(New MemoryItemRecord With {
                .SourceEventId = If(sourceEvent?.EventId, ""),
                .Scope = NormalizeScope(obj("scope")?.ToString()),
                .AppType = If(sourceEvent?.AppType, ""),
                .DocumentId = If(sourceEvent?.DocumentId, ""),
                .MemoryType = NormalizeMemoryType(obj("memory_type")?.ToString()),
                .Content = memoryContent.Trim(),
                .Summary = If(obj("summary")?.ToString(), ""),
                .Confidence = Clamp01(ReadDouble(obj("confidence"), 0.7R)),
                .Importance = Clamp01(ReadDouble(obj("importance"), 0.6R)),
                .Status = "active",
                .ExpiresAt = If(obj("expires_at")?.ToString(), "")
            })
        Next

        Return results
    End Function

    Private Shared Function ExtractJsonArray(content As String) As String
        If String.IsNullOrWhiteSpace(content) Then Return ""

        Dim cleaned = content.Trim()
        Dim codeBlock = Regex.Match(cleaned, "```(?:json)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase)
        If codeBlock.Success Then cleaned = codeBlock.Groups(1).Value.Trim()

        If cleaned.StartsWith("[") AndAlso cleaned.EndsWith("]") Then Return cleaned

        If cleaned.StartsWith("{") Then
            Try
                Dim obj = JObject.Parse(cleaned)
                Dim memories = TryCast(obj("memories"), JArray)
                If memories IsNot Nothing Then Return memories.ToString(Formatting.None)
            Catch
            End Try
        End If

        Dim startIdx = cleaned.IndexOf("[")
        Dim endIdx = cleaned.LastIndexOf("]")
        If startIdx >= 0 AndAlso endIdx > startIdx Then
            Return cleaned.Substring(startIdx, endIdx - startIdx + 1)
        End If

        Return ""
    End Function

    Private Shared Function MergeMemories(primary As List(Of MemoryItemRecord), fallback As List(Of MemoryItemRecord)) As List(Of MemoryItemRecord)
        Dim merged As New List(Of MemoryItemRecord)()
        If primary IsNot Nothing Then merged.AddRange(primary)
        If fallback Is Nothing Then Return merged

        Dim keys As New HashSet(Of String)(merged.Select(Function(m) NormalizeComparableText(m.Content)), StringComparer.OrdinalIgnoreCase)
        For Each item In fallback
            Dim key = NormalizeComparableText(item.Content)
            If Not keys.Contains(key) Then
                merged.Add(item)
                keys.Add(key)
            End If
        Next
        Return merged
    End Function

    Private Shared Function NormalizeScope(scope As String) As String
        Select Case If(scope, "").Trim().ToLowerInvariant()
            Case "document", "project", "session"
                Return scope.Trim().ToLowerInvariant()
            Case Else
                Return "user"
        End Select
    End Function

    Private Shared Function NormalizeMemoryType(memoryType As String) As String
        Select Case If(memoryType, "").Trim().ToLowerInvariant()
            Case "preference", "format_rule", "solution", "fact", "workflow"
                Return memoryType.Trim().ToLowerInvariant()
            Case Else
                Return "fact"
        End Select
    End Function

    Private Shared Function ReadDouble(token As JToken, fallback As Double) As Double
        If token Is Nothing Then Return fallback
        Dim value As Double
        If Double.TryParse(token.ToString(), value) Then Return value
        Return fallback
    End Function

    Private Shared Function Clamp01(value As Double) As Double
        If value < 0 Then Return 0
        If value > 1 Then Return 1
        Return value
    End Function

    Private Shared Function Truncate(value As String, maxLength As Integer) As String
        If String.IsNullOrEmpty(value) OrElse value.Length <= maxLength Then Return If(value, "")
        Return value.Substring(0, maxLength)
    End Function

    Private Shared Function NormalizeComparableText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Return Regex.Replace(value.Trim().ToLowerInvariant(), "\s+", "")
    End Function
End Class
