Imports System.Text
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        #Region "Planning, Repair, And Explanation Helpers"

        Private Function BuildExecutionExplanation(stepIndex As Integer,
                                                   planStep As PlanStep,
                                                   toolCall As ToolCall,
                                                   toolResult As ToolResult,
                                                   fixAttempts As Integer,
                                                   undoPoint As Core.UndoManager.UndoPoint,
                                                   observation As String,
                                                   startedAt As DateTime,
                                                   finishedAt As DateTime) As ExecutionExplanation
            Dim tool = If(toolCall Is Nothing, Nothing, _toolRegistry.GetTool(toolCall.ToolId))
            Dim toolId = If(toolCall?.ToolId, "")
            Dim toolName = If(tool?.Name, toolId)
            Dim category = If(tool?.Category, "")
            Dim risk = If(tool?.RiskLevel, "unknown")
            Dim paramsJson = If(toolCall?.Parameters Is Nothing, "{}", toolCall.Parameters.ToString(Formatting.None))
            Dim success = toolResult IsNot Nothing AndAlso toolResult.Success
            Dim message = If(toolResult?.Message, "")
            Dim skillName = ExtractResultDataValue(toolResult, "skillName")
            Dim scriptFileName = ExtractResultDataValue(toolResult, "scriptFileName")
            Dim mcpToolName = ExtractResultDataValue(toolResult, "mcpToolName")
            Dim mcpStatus = ExtractResultDataValue(toolResult, "mcpStatus")
            Dim failureReason = If(success, "", ExtractResultDataValue(toolResult, "failureReason"))
            If String.IsNullOrWhiteSpace(failureReason) AndAlso Not success Then failureReason = message
            Dim verb = If(success, "выполнен", "не выполнен")
            Dim fixText = If(fixAttempts > 0, $", авт. исправлений: {fixAttempts}", "")
            Dim skillText = If(String.IsNullOrWhiteSpace(skillName), "", $", Skill: {skillName}")
            Dim scriptText = If(String.IsNullOrWhiteSpace(scriptFileName), "", $", скрипт: {scriptFileName}")
            Dim mcpText = If(String.IsNullOrWhiteSpace(mcpToolName), "", $", MCP: {mcpToolName}")
            Dim elapsedMs = CLng(Math.Max(0, (finishedAt - startedAt).TotalMilliseconds))
            If toolResult IsNot Nothing AndAlso toolResult.ElapsedMs > 0 Then elapsedMs = toolResult.ElapsedMs
            Dim undoHint = If(undoPoint Is Nothing, "", _undoManager.GetUndoHint(If(undoPoint.AppType, "")))
            Dim undoPointName = If(undoPoint?.Name, "")
            Dim canUndo = undoPoint IsNot Nothing AndAlso undoPoint.CanUndo
            Dim beforeSummary = $"Подготовка к шагу {stepIndex + 1}: {If(planStep?.Description, "")}; инструмент: {toolId}; параметры: {paramsJson}"
            Dim afterSummary = If(observation, message)
            Dim observationJson = SerializeCompactJson(If(toolResult?.Observation, Nothing), 8192)
            Dim dataSummaryJson = SerializeCompactJson(BuildDataSummary(toolResult), 4096)
            Dim repairSummary = If(fixAttempts > 0,
                                   If(success, $"AI успешно исправил после {fixAttempts} попыток", $"AI не смог исправить после {fixAttempts} попыток"),
                                   "")
            Dim text = $"Шаг {stepIndex + 1} {verb}: вызов {toolName} ({toolId}){skillText}{scriptText}{mcpText}{fixText}. {message}"

            Return New ExecutionExplanation With {
                .StepIndex = stepIndex,
                .StepDescription = If(planStep?.Description, ""),
                .ToolId = toolId,
                .ToolName = toolName,
                .ToolCategory = category,
                .RiskLevel = risk,
                .ParametersJson = paramsJson,
                .StartedAt = startedAt,
                .FinishedAt = finishedAt,
                .ElapsedMs = elapsedMs,
                .BeforeSummary = beforeSummary,
                .AfterSummary = afterSummary,
                .ObservationJson = observationJson,
                .DataSummaryJson = dataSummaryJson,
                .Success = success,
                .Message = message,
                .SkillName = skillName,
                .ScriptFileName = scriptFileName,
                .McpToolName = mcpToolName,
                .McpStatus = mcpStatus,
                .FailureReason = failureReason,
                .UndoPointName = undoPointName,
                .UndoHint = undoHint,
                .CanUndo = canUndo,
                .AutoRepairSummary = repairSummary,
                .FixAttempts = fixAttempts,
                .ExplanationText = text
            }
        End Function

        Private Function BuildDataSummary(toolResult As ToolResult) As Object
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing Then Return Nothing

            Try
                Dim token = JToken.FromObject(toolResult.Data)
                If token.Type = JTokenType.Array Then
                    Dim arr = CType(token, JArray)
                    Return New JObject From {
                        {"type", "array"},
                        {"count", arr.Count},
                        {"preview", CloneFirstItems(arr, 5)}
                    }
                End If

                If token.Type = JTokenType.Object Then
                    Dim obj = CType(token, JObject)
                    Dim summary As New JObject From {
                        {"type", "object"},
                        {"keys", BuildStringArray(obj.Properties().Select(Function(p) p.Name).Take(20))}
                    }

                    If obj("total") IsNot Nothing Then summary("total") = obj("total")
                    If obj("returned") IsNot Nothing Then summary("returned") = obj("returned")
                    If obj("truncated") IsNot Nothing Then summary("truncated") = obj("truncated")
                    If obj("items") IsNot Nothing AndAlso obj("items").Type = JTokenType.Array Then
                        summary("itemsPreview") = CloneFirstItems(CType(obj("items"), JArray), 5)
                    End If

                    Return summary
                End If

                Return token
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", "BuildDataSummary failed", ex)
                Return Nothing
            End Try
        End Function

        Private Function CloneFirstItems(items As JArray, maxItems As Integer) As JArray
            Dim result As New JArray()
            If items Is Nothing Then Return result

            For Each item In items.Take(Math.Max(0, maxItems))
                result.Add(item.DeepClone())
            Next

            Return result
        End Function

        Private Function BuildStringArray(items As IEnumerable(Of String)) As JArray
            Dim result As New JArray()
            If items Is Nothing Then Return result

            For Each item In items
                result.Add(If(item, ""))
            Next

            Return result
        End Function

        Private Function SerializeCompactJson(value As Object, maxLength As Integer) As String
            If value Is Nothing Then Return ""

            Try
                Dim json = JsonConvert.SerializeObject(value, Formatting.None)
                If maxLength > 0 AndAlso json.Length > maxLength Then
                    Return json.Substring(0, maxLength) & "...(truncated)"
                End If
                Return json
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", "SerializeCompactJson failed", ex)
                Return ""
            End Try
        End Function

        Private Function ExtractResultDataValue(toolResult As ToolResult, key As String) As String
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""

            Try
                Dim obj = JObject.FromObject(toolResult.Data)
                Dim token = obj.SelectToken(key)
                If token Is Nothing Then Return ""
                Return token.ToString()
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", $"ExtractResultDataValue key={key} failed", ex)
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 生成任务 Spec
        ''' </summary>
        Private Async Function GenerateSpecAsync(session As AgentSession) As Task(Of AgentTaskSpec)
            Dim spec As New AgentTaskSpec()
            Try
                Dim prompt = $"Проанализируй запрос и выдели структурированную спецификацию задачи:

Запрос: {session.UserRequest}

Верни JSON:
```json
{{
  ""goal"": ""одно предложение о главной цели"",
  ""constraints"": [""ограничение 1""],
  ""success_criteria"": [""критерий успеха 1""],
  ""complexity"": ""simple|medium|complex""
}}
```

Правила complexity:
- simple: одно действие, не более 2 шагов, подтверждение пользователя не требуется
- medium: 2-5 шагов, желательно подтверждение пользователя
- complex: много шагов или сложная логика, подтверждение пользователя обязательно"

                Dim response = Await SendAIRequest(prompt,
                    "Ты специалист по анализу задач. Возвращай только JSON, без пояснений.", Nothing)

                Dim jsonStr = ExtractJson(response)
                If Not String.IsNullOrEmpty(jsonStr) Then
                    Dim obj = JObject.Parse(jsonStr)
                    spec.Goal = If(obj("goal")?.ToString(), session.UserRequest)
                    spec.Complexity = If(obj("complexity")?.ToString(), "medium")

                    Dim constraints = TryCast(obj("constraints"), JArray)
                    If constraints IsNot Nothing Then
                        For Each c In constraints
                            spec.Constraints.Add(c.ToString())
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] Spec生成失败: {ex.Message}")
            End Try
            Return spec
        End Function

        ''' <summary>
        ''' 生成执行计划
        ''' </summary>
        Private Async Function GeneratePlanAsync(session As AgentSession,
                                                  systemPrompt As String,
                                                  skill As AgentSkill,
                                                  Optional attempt As Integer = 0) As Task(Of ExecutionPlan)
            Dim plan As New ExecutionPlan()
            Dim failureReason As String = ""
            Try
                Dim prompt = _promptManager.BuildPlanningPrompt(session, systemPrompt, skill)
                Dim response = Await SendAIRequest(prompt, systemPrompt, _memory.GetRecentMessages(5))

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then
                    failureReason = "в ответе модели нет JSON"
                Else
                    Dim obj = JObject.Parse(jsonStr)
                    plan.Understanding = obj("understanding")?.ToString()
                    plan.Summary = obj("summary")?.ToString()
                    plan.CapabilityGap = obj("capabilityGap")?.ToString()

                    Dim stepsArray = TryCast(obj("steps"), JArray)
                    If stepsArray IsNot Nothing Then
                        Dim stepNum = 1
                        For Each item In stepsArray
                            plan.Steps.Add(New PlanStep With {
                                .StepNumber = stepNum,
                                .Description = item("description")?.ToString(),
                                .Code = item("code")?.ToString(),
                                .Language = If(item("language")?.ToString(), "json")
                            })
                            stepNum += 1
                        Next
                    End If
                End If
            Catch ex As Exception
                failureReason = ex.Message
            End Try

            If plan.Steps.Count = 0 AndAlso String.IsNullOrWhiteSpace(plan.CapabilityGap) Then
                If String.IsNullOrWhiteSpace(failureReason) Then failureReason = "в ответе нет steps и нет capabilityGap"
                _lastPlanFailureReason = failureReason
                AppLogger.Warn("LoopEngine", $"规划无效 attempt={attempt}: {AppLogger.Redact(failureReason)}")
                If attempt = 0 Then
                    Dim correctedPrompt = systemPrompt & vbCrLf &
                        "【Исправление плана】Предыдущий ответ не является допустимым планом. Верни строгий JSON и обязательно непустой массив steps либо явное поле capabilityGap. Отвечай только на русском языке."
                    Return Await GeneratePlanAsync(session, correctedPrompt, skill, attempt + 1)
                End If
                Return Nothing
            End If
            _lastPlanFailureReason = ""
            Return plan
        End Function

        Private Function ValidatePlanCoverage(spec As AgentTaskSpec, plan As ExecutionPlan) As String
            If spec Is Nothing OrElse plan Is Nothing Then Return ""

            Dim toolIds = GetPlannedToolIds(plan)
            If spec.ExpectedOutputs.Contains("images") AndAlso
               Not toolIds.Contains("InsertImage") AndAlso
               Not PlanContainsCreateSlidesImage(plan) Then
                Return "План не покрывает вставку реального изображения, которое запросил пользователь; при отсутствии доступного источника нужно явно сообщить о нехватке возможности."
            End If
            If spec.ExpectedOutputs.Contains("images") Then
                Dim imagePathError = ValidatePlannedImagePaths(plan)
                If Not String.IsNullOrWhiteSpace(imagePathError) Then Return imagePathError
            End If
            If spec.ExpectedSlideCount > 0 AndAlso
               Not toolIds.Contains("CreateSlides") AndAlso
               Not toolIds.Contains("InsertSlide") Then
                Return $"План не покрывает создание {spec.ExpectedSlideCount} слайдов."
            End If
            Return ""
        End Function

        Private Function ValidatePlannedImagePaths(plan As ExecutionPlan) As String
            If plan?.Steps Is Nothing Then Return "План работы с изображениями пуст."

            For Each stepItem In plan.Steps
                Try
                    Dim envelope = JObject.Parse(If(stepItem.Code, ""))
                    Dim commands As New List(Of JObject)()
                    If envelope("command") IsNot Nothing Then commands.Add(envelope)
                    Dim commandArray = TryCast(envelope("commands"), JArray)
                    If commandArray IsNot Nothing Then commands.AddRange(commandArray.OfType(Of JObject)())

                    For Each commandObj In commands
                        Dim commandName = commandObj("command")?.ToString()
                        If String.Equals(commandName, "InsertImage", StringComparison.OrdinalIgnoreCase) Then
                            Dim params = TryCast(commandObj("params"), JObject)
                            Dim imagePath = params?("imagePath")?.ToString()
                            If String.IsNullOrWhiteSpace(imagePath) Then
                                Return "В плане с изображением отсутствует imagePath; частичные изменения не начинаются."
                            End If
                            If Not IO.File.Exists(imagePath) Then
                                Return $"Файл изображения недоступен: {imagePath}. Чтобы не создавать только текстовые слайды, задача не выполнялась."
                            End If
                        ElseIf String.Equals(commandName, "CreateSlides", StringComparison.OrdinalIgnoreCase) Then
                            Dim params = TryCast(commandObj("params"), JObject)
                            Dim slides = TryCast(params?("slides"), JArray)
                            If slides Is Nothing Then Continue For
                            For Each slide In slides.OfType(Of JObject)()
                                Dim imagePath = slide("imagePath")?.ToString()
                                If Not String.IsNullOrWhiteSpace(imagePath) AndAlso Not IO.File.Exists(imagePath) Then
                                    Return $"Файл изображения недоступен: {imagePath}. Чтобы не создавать только текстовые слайды, задача не выполнялась."
                                End If
                            Next
                        End If
                    Next
                Catch
                    ' Invalid plan JSON is handled by the normal tool-call parsing path.
                End Try
            Next
            Return ""
        End Function

        Private Function PlanContainsCreateSlidesImage(plan As ExecutionPlan) As Boolean
            If plan?.Steps Is Nothing Then Return False

            For Each stepItem In plan.Steps
                Try
                    Dim envelope = JObject.Parse(If(stepItem.Code, ""))
                    Dim commands As New List(Of JObject)()
                    If envelope("command") IsNot Nothing Then commands.Add(envelope)
                    Dim commandArray = TryCast(envelope("commands"), JArray)
                    If commandArray IsNot Nothing Then commands.AddRange(commandArray.OfType(Of JObject)())

                    For Each commandObj In commands
                        If String.Equals(commandObj("command")?.ToString(), "CreateSlides", StringComparison.OrdinalIgnoreCase) AndAlso
                           CreateSlidesParametersContainImage(TryCast(commandObj("params"), JObject)) Then
                            Return True
                        End If
                    Next
                Catch
                    ' Invalid plan JSON is handled by the normal tool-call parsing path.
                End Try
            Next
            Return False
        End Function

        Private Shared Function CreateSlidesParametersContainImage(parameters As JObject) As Boolean
            Dim slides = TryCast(parameters?("slides"), JArray)
            If slides Is Nothing Then Return False
            Return slides.OfType(Of JObject)().Any(
                Function(slide) Not String.IsNullOrWhiteSpace(slide("imagePath")?.ToString()))
        End Function

        Private Function ValidateExecutionOutcome(session As AgentSession) As String
            If session Is Nothing OrElse session.Spec Is Nothing Then Return ""

            Dim successfulActions = session.Iterations.
                Where(Function(item) item IsNot Nothing AndAlso
                                     item.Explanation IsNot Nothing AndAlso
                                     item.Explanation.Success AndAlso
                                     item.Action IsNot Nothing).
                Select(Function(item) item.Action).
                ToList()
            If session.Spec.ExpectedOutputs.Contains("images") AndAlso
               Not successfulActions.Any(
                   Function(toolCall) String.Equals(toolCall.ToolId, "InsertImage", StringComparison.OrdinalIgnoreCase) OrElse
                                      (String.Equals(toolCall.ToolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) AndAlso
                                       CreateSlidesParametersContainImage(toolCall.Parameters))) Then
                Return "Задача не выполнена: пользователь просил вставить изображение, но в записи выполнения нет успешной вставки реального изображения."
            End If

            If session.Spec.ExpectedSlideCount > 0 Then
                Dim created As Integer = 0
                For Each action In successfulActions
                    If String.Equals(action.ToolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) Then
                        Dim slides = TryCast(action.Parameters?("slides"), JArray)
                        If slides IsNot Nothing Then created += slides.Count
                    ElseIf String.Equals(action.ToolId, "InsertSlide", StringComparison.OrdinalIgnoreCase) Then
                        created += 1
                    End If
                Next
                If created < session.Spec.ExpectedSlideCount Then
                    Return $"Задача не выполнена: требовалось создать {session.Spec.ExpectedSlideCount} слайдов, в записи выполнения подтверждено {created}."
                End If
            End If
            Return ""
        End Function

        Private Function GetPlannedToolIds(plan As ExecutionPlan) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If plan?.Steps Is Nothing Then Return result

            For Each stepItem In plan.Steps
                Try
                    Dim obj = JObject.Parse(If(stepItem.Code, ""))
                    Dim command = obj("command")?.ToString()
                    If Not String.IsNullOrWhiteSpace(command) Then result.Add(command)
                    Dim commands = TryCast(obj("commands"), JArray)
                    If commands Is Nothing Then Continue For
                    For Each item In commands.OfType(Of JObject)()
                        command = item("command")?.ToString()
                        If Not String.IsNullOrWhiteSpace(command) Then result.Add(command)
                    Next
                Catch
                    ' Invalid plan JSON is handled by the normal tool-call parsing path.
                End Try
            Next
            Return result
        End Function

        ''' <summary>
        ''' Think：调用LLM生成思考+行动
        ''' </summary>
        Private Async Function ThinkAsync(session As AgentSession,
                                           planStep As PlanStep,
                                           systemPrompt As String) As Task(Of String)
            Dim lastObservation = _memory.GetWorkingString("lastObservation")
            Dim prompt = _promptManager.BuildReactPrompt(planStep, _memory, lastObservation)
            Dim history = _memory.GetRecentMessages(10)
            Return Await SendAIRequest(prompt, systemPrompt, history)
        End Function

        ''' <summary>
        ''' 反思并重新规划
        ''' </summary>
        Private Async Function ReflectAndReplanAsync(session As AgentSession,
                                                      observation As String,
                                                      systemPrompt As String) As Task(Of ExecutionPlan)
            Try
                Dim prompt = _promptManager.BuildReflectionPrompt(session, observation)
                Dim response = Await SendAIRequest(prompt, systemPrompt, Nothing)

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim decision = JObject.Parse(jsonStr)
                Dim strategy = decision("strategy")?.ToString()?.ToLower()

                Select Case strategy
                    Case "retry"
                        ' 重试当前计划
                        Return session.Plan
                    Case "skip"
                        ' 跳过当前步骤继续
                        Return session.Plan
                    Case "replan"
                        ' 重新生成计划
                        Return Await GeneratePlanAsync(session, systemPrompt, session.Skill)
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 反思失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 解析工具调用
        ''' </summary>
        Private Function ParseToolCall(response As String) As ToolCall
            Try
                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim obj = JObject.Parse(jsonStr)
                Dim action = obj("action")
                If action Is Nothing Then Return Nothing

                Dim toolId = action("tool")?.ToString()
                Dim params = TryCast(action("params"), JObject)
                If String.IsNullOrEmpty(toolId) Then Return Nothing
                If params Is Nothing Then params = New JObject()

                Return New ToolCall With {
                    .ToolId = toolId,
                    .Parameters = params
                }
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 解析工具调用失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        Private Function ParseFixedToolCall(fixedObj As JObject, fallback As ToolCall) As ToolCall
            If fixedObj Is Nothing Then Return fallback

            Dim toolId = fixedObj("toolId")?.ToString()
            Dim params = TryCast(fixedObj("parameters"), JObject)

            If String.IsNullOrWhiteSpace(toolId) Then
                Dim action = TryCast(fixedObj("action"), JObject)
                If action IsNot Nothing Then
                    toolId = action("tool")?.ToString()
                    If params Is Nothing Then params = TryCast(action("params"), JObject)
                End If
            End If

            If String.IsNullOrWhiteSpace(toolId) Then toolId = fallback.ToolId
            If params Is Nothing Then params = If(fallback.Parameters, New JObject())

            Return New ToolCall With {
                .ToolId = toolId,
                .Parameters = params
            }
        End Function

        Private Function BuildAvailableToolHint(appType As String,
                                                Optional executionContext As ToolExecutionContext = Nothing) As String
            Dim tools = _toolRegistry.GetVisibleTools(appType, executionContext).
                OrderBy(Function(t) t.Category).
                ThenBy(Function(t) t.Id).
                Take(40).
                Select(Function(t) $"- {t.Id}: {t.Name}")
            Return String.Join(vbCrLf, tools)
        End Function

        ''' <summary>
        ''' 格式化观察结果
        ''' </summary>
        Private Function FormatObservation(result As ToolResult) As String
            If result Is Nothing Then
                Return "❌ [unknown] нет результата инструмента"
            End If
            If result.Success Then
                Dim summary = result.ToObserveSummary()
                Dim dataSummary = FormatResultData(result.Data)
                If Not String.IsNullOrWhiteSpace(dataSummary) Then
                    Return $"✅ [{result.ToolId}] выполнено успешно: {summary}{vbCrLf}data={dataSummary}"
                End If
                Return $"✅ [{result.ToolId}] выполнено успешно: {summary}"
            End If
            ' Structured observe payload for repair/reflect (P0-4).
            Return $"❌ [{result.ToolId}] {result.ToObserveSummary()}"
        End Function

        ''' <summary>
        ''' Sends one repair request with ephemeral visual evidence when the host supplied it.
        ''' Evidence is never copied into the textual prompt, history, memory, or trace. If the
        ''' configured provider/model rejects image input, the existing text-only path remains
        ''' the deterministic fallback for the same repair attempt.
        ''' </summary>
        Private Async Function SendRepairRequestAsync(fixPrompt As String,
                                                      systemPrompt As String,
                                                      result As ToolResult) As Task(Of String)
            If _multimodalRepairDisabledForRun OrElse
               result Is Nothing OrElse
               result.VisualEvidence Is Nothing OrElse
               result.VisualEvidence.Count = 0 OrElse
               SendAIRequestWithMessages Is Nothing Then
                Return Await SendAIRequest(fixPrompt, systemPrompt, Nothing)
            End If

            Dim content As New JArray From {
                New JObject From {
                    {"type", "text"},
                    {"text", fixPrompt & vbCrLf & vbCrLf &
                        "Диагностируй визуальную проблему по приложенному реальному скриншоту PowerPoint; верни только согласованный JSON исправленного вызова инструмента. Отвечай только на русском языке."}
                }
            }
            Dim evidenceCount As Integer = 0
            For Each item In result.VisualEvidence
                If item Is Nothing OrElse evidenceCount >= 2 Then Exit For
                If item.ByteLength > 0 AndAlso
                   item.ByteLength <= 2 * 1024 * 1024 AndAlso
                   Not String.IsNullOrWhiteSpace(item.DataUrl) AndAlso
                   item.DataUrl.Length <= 3 * 1024 * 1024 AndAlso
                   item.DataUrl.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase) Then
                    content.Add(New JObject From {
                        {"type", "image_url"},
                        {"image_url", New JObject From {
                            {"url", item.DataUrl},
                            {"detail", "low"}
                        }}
                    })
                    evidenceCount += 1
                End If
            Next
            If evidenceCount = 0 Then Return Await SendAIRequest(fixPrompt, systemPrompt, Nothing)

            Dim messages As New JArray()
            If Not String.IsNullOrWhiteSpace(systemPrompt) Then
                messages.Add(New JObject From {
                    {"role", "system"},
                    {"content", systemPrompt}
                })
            End If
            messages.Add(New JObject From {
                {"role", "user"},
                {"content", content}
            })

            Try
                Dim response = Await SendAIRequestWithMessages(messages)
                If Not String.IsNullOrWhiteSpace(response) Then Return response
                _multimodalRepairDisabledForRun = True
                AppLogger.Warn("LoopEngine", "Multimodal repair returned an empty response; falling back to text-only repair")
            Catch ex As Exception
                _multimodalRepairDisabledForRun = True
                AppLogger.Warn("LoopEngine", $"Multimodal repair unavailable; falling back to text-only repair: {AppLogger.Redact(ex.Message)}")
            End Try
            Return Await SendAIRequest(fixPrompt, systemPrompt, Nothing)
        End Function

        ''' <summary>
        ''' A host call returning Success=True is not sufficient for a write operation.
        ''' Structured write observations must confirm an actual change, and operation
        ''' batches must not contain unhandled failed/partial steps.
        ''' </summary>
        Private Function ValidateObservedOutcome(result As ToolResult, appType As String) As ToolResult
            If result Is Nothing OrElse Not result.Success OrElse result.Observation Is Nothing Then Return result

            Try
                Dim observation = TryCast(result.Observation, JToken)
                If observation Is Nothing Then observation = JToken.FromObject(result.Observation)
                If observation Is Nothing OrElse observation.Type <> JTokenType.Object Then Return result

                Dim kind = If(observation("kind")?.ToString(), "").Trim().ToLowerInvariant()
                Dim writeExpectedToken = observation("writeExpected")
                Dim writeExpected = writeExpectedToken IsNot Nothing AndAlso
                                    writeExpectedToken.Type = JTokenType.Boolean AndAlso
                                    writeExpectedToken.Value(Of Boolean)()
                If kind <> "write" AndAlso (kind <> "office_operation_batch" OrElse Not writeExpected) Then
                    Return result
                End If

                Dim operations = TryCast(observation("operations"), JArray)
                If operations IsNot Nothing Then
                    Dim failedCount = operations.OfType(Of JObject)().Count(
                        Function(item) Not String.Equals(item("status")?.ToString(), "succeeded", StringComparison.OrdinalIgnoreCase))
                    If failedCount > 0 Then
                        Return ToolResult.Failed(
                            result.ToolId,
                            $"В пакете операций не выполнено шагов: {failedCount}",
                            data:=result.Data,
                            errorCode:=ExceptionClassifier.CodePartialApply,
                            userMessage:="Часть операций Office не выполнена, пытаюсь исправить",
                            recoverable:=True,
                            observation:=result.Observation,
                            artifacts:=result.Artifacts)
                    End If
                End If

                Dim changedToken = observation("changed")
                If changedToken IsNot Nothing AndAlso
                   changedToken.Type = JTokenType.Boolean AndAlso
                   Not changedToken.Value(Of Boolean)() Then
                    ' Снимок PowerPoint/Excel не покрывает все типы правок (анимации, темы,
                    ' форматирование, вид и т.п.). Если хост уже подтвердил успех, отсутствие
                    ' изменения в снимке не должно валить всю задачу — фиксируем предупреждение.
                    If String.Equals(appType, "PowerPoint", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(appType, "Excel", StringComparison.OrdinalIgnoreCase) Then
                        AppendObservationWarning(observation,
                                                 "Наблюдение не обнаружило изменений, но хост сообщил об успехе")
                        Return result
                    End If

                    Return ToolResult.Failed(
                        result.ToolId,
                        "Хост сообщил об успехе, но наблюдение не обнаружило фактических изменений",
                        data:=result.Data,
                        errorCode:=ExceptionClassifier.CodeVerifyFailed,
                        userMessage:="Office не дал ожидаемых изменений, повторно проверяю или исправляю",
                        recoverable:=True,
                        observation:=result.Observation,
                        artifacts:=result.Artifacts)
                End If
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Validate observation failed open: {AppLogger.Redact(ex.Message)}")
            End Try
            Return result
        End Function

        Private Shared Sub AppendObservationWarning(observation As JToken, message As String)
            Dim obj = TryCast(observation, JObject)
            If obj Is Nothing Then Return
            Dim warnings = TryCast(obj("warnings"), JArray)
            If warnings Is Nothing Then
                warnings = New JArray()
                obj("warnings") = warnings
            End If
            warnings.Add(message)
        End Sub

        Private Function FormatResultData(data As Object) As String
            If data Is Nothing Then Return ""

            Try
                Dim text As String
                If TypeOf data Is JToken Then
                    text = DirectCast(data, JToken).ToString(Formatting.None)
                Else
                    text = JsonConvert.SerializeObject(data, Formatting.None)
                End If

                Const maxLen As Integer = 1800
                If text.Length > maxLen Then
                    Return text.Substring(0, maxLen) & "...(truncated)"
                End If
                Return text
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 从响应中提取JSON
        ''' </summary>
        Private Function ExtractJson(response As String) As String
            If String.IsNullOrWhiteSpace(response) Then Return Nothing

            ' 查找 ```json 代码块
            Dim start = response.IndexOf("```json")
            If start >= 0 Then
                start = response.IndexOf("{"c, start)
                If start >= 0 Then
                    Dim endIdx = response.LastIndexOf("}"c)
                    If endIdx > start Then
                        Return response.Substring(start, endIdx - start + 1)
                    End If
                End If
            End If

            ' 查找纯 JSON
            start = response.IndexOf("{"c)
            If start >= 0 Then
                Dim endIdx = response.LastIndexOf("}"c)
                If endIdx > start Then
                    Return response.Substring(start, endIdx - start + 1)
                End If
            End If

            Return Nothing
        End Function

        #End Region

    End Class

End Namespace
