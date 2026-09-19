Imports System.Text
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' ReAct 循环引擎 - 核心执行逻辑
    ''' Think -> Plan -> Act -> Observe -> Reflect
    ''' </summary>
    Public Partial Class LoopEngine
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _undoManager As Core.UndoManager
        Private _multimodalRepairDisabledForRun As Boolean
        Private _lastPlanFailureReason As String = ""

        ' 循环限制
        Private Const MaxIterations As Integer = 15
        Private Const MaxNoProgress As Integer = 3
        Private Const MaxReplanAttempts As Integer = 2

        ' 回调
        Public Property OnStatusChanged As Action(Of String)
        Public Property OnIterationUpdate As Action(Of ReActIteration)
        Public Property OnStepCompleted As Action(Of Integer, Boolean, String)
        Public Property OnExecutionExplained As Action(Of ExecutionExplanation)
        Public Property OnRequestApproval As Func(Of String, Task(Of Boolean))
        Public Property OnPlanGenerated As Action(Of ExecutionPlan)
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, Task(Of String))

        Public Sub New(toolRegistry As ToolRegistry, memory As AgentMemory, promptManager As PromptManager)
            _toolRegistry = toolRegistry
            _memory = memory
            _promptManager = promptManager
            _undoManager = New Core.UndoManager()
        End Sub

        ''' <summary>
        ''' 执行 ReAct 循环
        ''' </summary>
        Public Async Function RunAsync(session As AgentSession,
                                        systemPrompt As String,
                                        Optional skill As AgentSkill = Nothing) As Task(Of AgentResult)
            Dim noProgressCount As Integer = 0
            Dim replanAttempts As Integer = 0
            _multimodalRepairDisabledForRun = False

            Try
                ' Phase 1: 生成 Spec
                OnStatusChanged?.Invoke("Анализирую задачу...")
                If session.Spec Is Nothing Then
                    session.Spec = Await GenerateSpecAsync(session)
                Else
                    AppLogger.Info("LoopEngine", $"Use precomputed TaskSpec goal={AppLogger.Redact(session.Spec.Goal)}")
                End If

                ' Phase 2: 生成计划
                OnStatusChanged?.Invoke("Формирую план выполнения...")
                session.Plan = Await GeneratePlanAsync(session, systemPrompt, skill)
                If session.Plan Is Nothing OrElse session.Plan.Steps.Count = 0 Then
                    Dim capabilityGap = If(session.Plan?.CapabilityGap, "")
                    Dim planFailure = If(String.IsNullOrWhiteSpace(capabilityGap),
                                         "Не удалось построить план выполнения" &
                                         If(String.IsNullOrWhiteSpace(_lastPlanFailureReason), "", ": " & _lastPlanFailureReason),
                                         "Текущая способность не позволяет выполнить полностью: " & capabilityGap)
                    OnStatusChanged?.Invoke(planFailure)
                    Return AgentResult.Failed(session.Id, planFailure)
                End If

                Dim coverageError = ValidatePlanCoverage(session.Spec, session.Plan)
                If Not String.IsNullOrWhiteSpace(coverageError) Then
                    OnStatusChanged?.Invoke(coverageError)
                    Return AgentResult.Failed(session.Id, coverageError)
                End If

                ' 通知计划已生成
                OnPlanGenerated?.Invoke(session.Plan)

                session.Status = AgentStatus.Executing
                OnStatusChanged?.Invoke($"План готов (шагов: {session.Plan.Steps.Count}), переход к выполнению...")
                Dim executionContext As ToolExecutionContext = ToolExecutionContext.FromSession(session, skill)

                ' Phase 3: ReAct Loop
                Dim stepIndex As Integer = 0
                While stepIndex < session.Plan.Steps.Count AndAlso session.CurrentIteration < MaxIterations
                    Dim planStep = session.Plan.Steps(stepIndex)
                    planStep.Status = StepStatus.Running

                    ' --- THINK ---
                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"Шаг {stepIndex + 1}/{session.Plan.Steps.Count}: {planStep.Description}")
                    Dim thought = Await ThinkAsync(session, planStep, systemPrompt)

                    ' --- PARSE ACTION ---
                    Dim toolCall = ParseToolCall(thought)
                    If toolCall Is Nothing Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = "Не удалось разобрать вызов инструмента"
                        OnStepCompleted?.Invoke(stepIndex, False, "Ошибка разбора")

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If

                    Dim normalizeMessage As String = ""
                    If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, executionContext, normalizeMessage) Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = normalizeMessage
                        OnStepCompleted?.Invoke(stepIndex, False, normalizeMessage)

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If
                    If Not String.IsNullOrWhiteSpace(normalizeMessage) Then
                        Debug.WriteLine($"[LoopEngine] {normalizeMessage}")
                    End If

                    ' --- RISK NOTICE (autonomous mode) ---
                    Dim tool = _toolRegistry.GetTool(toolCall.ToolId)
                    If tool IsNot Nothing AndAlso tool.RiskLevel = "risky" Then
                        Debug.WriteLine($"[LoopEngine] 自治模式执行高风险工具: {toolCall.ToolId}")
                        OnStatusChanged?.Invoke($"Шаг {stepIndex + 1}: используется рискованный инструмент {toolCall.ToolId}; риск зафиксирован, продолжаю")
                    End If

                    ' --- ACT (增强版 - 多轮自修复 + 撤销点) ---
                    session.Status = AgentStatus.Executing

                    ' 创建撤销点（执行前保存状态）
                    Dim undoPoint As Core.UndoManager.UndoPoint = Nothing
                    If _undoManager IsNot Nothing Then
                        undoPoint = _undoManager.CreateUndoPoint(
                            If(session.AppType, "Unknown"),
                            $"Шаг {stepIndex + 1}: {toolCall.ToolId}",
                            planStep.Description)
                        If undoPoint IsNot Nothing Then
                            Debug.WriteLine($"[LoopEngine] 创建撤销点: {undoPoint.Name}")
                        End If
                    End If

                    Const MaxFixAttempts As Integer = 3
                    Dim fixAttempt As Integer = 0
                    Dim toolResult As ToolResult = Nothing
                    Dim stepStartedAt = DateTime.Now

                    ' 多轮修复循环
                    While fixAttempt < MaxFixAttempts
                        Dim normalized As String = ""
                        If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, executionContext, normalized) Then
                            toolResult = ToolResult.Failed(toolCall.ToolId,
                                                           normalized,
                                                           New With {.availableTools = BuildAvailableToolHint(session.AppType, executionContext)},
                                                           ExceptionClassifier.CodeNotFound,
                                                           normalized,
                                                           normalized,
                                                           recoverable:=True)
                        Else
                            If Not String.IsNullOrWhiteSpace(normalized) Then Debug.WriteLine($"[LoopEngine] {normalized}")
                            ' 执行工具
                            toolResult = ValidateObservedOutcome(
                                Await _toolRegistry.ExecuteToolAsync(executionContext, toolCall.ToolId, toolCall.Parameters),
                                session.AppType)
                        End If

                        If Not toolResult.Success AndAlso
                           String.Equals(toolResult.ErrorCode, ExceptionClassifier.CodeSafetyNeedsApproval, StringComparison.OrdinalIgnoreCase) AndAlso
                           OnRequestApproval IsNot Nothing Then
                            OnStatusChanged?.Invoke($"Инструмент {toolCall.ToolId} ожидает подтверждения пользователя...")
                            Dim approved = Await OnRequestApproval(If(toolResult.UserMessage, toolResult.Message))
                            If approved Then
                                executionContext.ApproveTool(toolCall.ToolId, toolCall.Parameters)
                                OnStatusChanged?.Invoke($"Пользователь подтвердил инструмент {toolCall.ToolId}, продолжаю...")
                                toolResult = ValidateObservedOutcome(
                                    Await _toolRegistry.ExecuteToolAsync(executionContext, toolCall.ToolId, toolCall.Parameters),
                                    session.AppType)
                            Else
                                toolResult = ToolResult.Failed(
                                    toolCall.ToolId,
                                    "Пользователь отклонил рискованную операцию",
                                    errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                    userMessage:="Рискованная операция отменена",
                                    recoverable:=False,
                                    observation:=New JObject From {
                                        {"kind", "approval"},
                                        {"summary", "Пользователь отклонил рискованную операцию"},
                                        {"changed", False},
                                        {"warnings", New JArray("approval_rejected")}
                                    })
                            End If
                        End If

                        If toolResult.Success Then
                            ' 成功，跳出循环
                            Exit While
                        End If

                        ' 失败：尝试自动修复
                        fixAttempt += 1
                        If fixAttempt < MaxFixAttempts Then
                            ' Non-recoverable tool failures skip AI repair and go straight to observe/reflect.
                            If Not toolResult.Recoverable Then
                                AppLogger.Warn("LoopEngine", $"Skip repair for non-recoverable tool failure: {toolResult.ToObserveSummary()}")
                                Exit While
                            End If

                            OnStatusChanged?.Invoke($"Выполнение не удалось, AI вносит исправление (попытка {fixAttempt}/{MaxFixAttempts})...")
                            AppLogger.Info("LoopEngine", $"Repair attempt {fixAttempt}/{MaxFixAttempts}: {toolResult.ToObserveSummary()}")

                            ' 构建修复提示词（含结构化错误契约）
                            Dim fixPrompt = $"Предыдущее выполнение завершилось ошибкой:

Код ошибки: {If(toolResult.ErrorCode, ExceptionClassifier.CodeUnknown)}
Пояснение для пользователя: {If(toolResult.UserMessage, toolResult.Message)}
Детали отладки: {If(toolResult.DebugDetail, toolResult.Message)}
Допускает авт. исправление: {toolResult.Recoverable}

Исходный вызов инструмента: {toolCall.ToolId}
Исходные параметры: {Newtonsoft.Json.JsonConvert.SerializeObject(toolCall.Parameters)}

Доступные инструменты для {If(session.AppType, "Office")} (используй точные ID инструментов, не придумывай snake_case и незарегистрированные команды):
{BuildAvailableToolHint(session.AppType, executionContext)}

Проанализируй причину ошибки и верни исправленный вызов инструмента. Верни только JSON в формате:
```json
{{
  ""toolId"": ""..."",
  ""parameters"": {{...}}
}}
```"

                            Try
                                ' 请求 AI 修复
                                Dim fixedResponse = Await SendRepairRequestAsync(fixPrompt, systemPrompt, toolResult)
                                Dim fixedJson = ExtractJson(fixedResponse)

                                If Not String.IsNullOrEmpty(fixedJson) Then
                                    Dim fixedObj = JObject.Parse(fixedJson)
                                    Dim fixedToolCall = ParseFixedToolCall(fixedObj, toolCall)

                                    ' 使用修复后的工具调用
                                    toolCall = fixedToolCall
                                    AppLogger.Info("LoopEngine", "AI generated repair plan")
                                Else
                                    AppLogger.Warn("LoopEngine", "Unable to parse repair response; stop repair")
                                    Exit While
                                End If
                            Catch ex As Exception
                                AppLogger.Error("LoopEngine", "Repair loop exception", ex)
                                Exit While
                            End Try
                        End If
                    End While

                    ' --- OBSERVE ---
                    session.Status = AgentStatus.Observing
                    Dim observation = FormatObservation(toolResult)

                    ' 如果失败且已达最大修复次数，追加提示
                    If Not toolResult.Success AndAlso fixAttempt >= MaxFixAttempts Then
                        observation &= $" (AI пытался исправить автоматически {fixAttempt} раз, но безуспешно)"
                    ElseIf fixAttempt > 0 AndAlso toolResult.Success Then
                        observation &= $" (AI успешно исправил с попытки {fixAttempt})"
                    End If

                    _memory.SetWorking("lastObservation", observation)
                    Dim stepFinishedAt = DateTime.Now
                    Dim explanation = BuildExecutionExplanation(stepIndex, planStep, toolCall, toolResult, fixAttempt, undoPoint, observation, stepStartedAt, stepFinishedAt)
                    planStep.LastExplanation = explanation

                    ' 记录迭代
                    Dim iteration = New ReActIteration With {
                        .Index = session.CurrentIteration,
                        .Thought = thought,
                        .Action = toolCall,
                        .Observation = observation,
                        .Explanation = explanation
                    }
                    session.Iterations.Add(iteration)
                    session.CurrentIteration += 1
                    OnExecutionExplained?.Invoke(explanation)
                    OnIterationUpdate?.Invoke(iteration)

                    ' 更新步骤状态
                    If toolResult.Success Then
                        planStep.Status = StepStatus.Completed
                        noProgressCount = 0
                        OnStepCompleted?.Invoke(stepIndex, True, toolResult.Message)
                    Else
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = toolResult.ToObserveSummary()
                        noProgressCount += 1
                        OnStepCompleted?.Invoke(stepIndex, False, If(toolResult.UserMessage, toolResult.Message))
                        AppLogger.Warn("LoopEngine", $"Step failed: {toolResult.ToObserveSummary()}")

                        If Not toolResult.Recoverable Then
                            Dim terminalFailure = $"Задача остановлена из-за невосстановимой ошибки: {toolResult.ToObserveSummary()}"
                            session.Status = AgentStatus.Failed
                            OnStatusChanged?.Invoke(terminalFailure)
                            AppLogger.Warn("LoopEngine", terminalFailure)
                            Return AgentResult.Failed(session.Id, terminalFailure)
                        End If

                        ' 失败且多轮修复失败，提示可以撤销
                        If fixAttempt >= MaxFixAttempts AndAlso undoPoint IsNot Nothing Then
                            Dim undoHint = _undoManager.GetUndoHint(If(session.AppType, "Unknown"))
                            AppLogger.Info("LoopEngine", $"Execution failed; {undoHint}")
                        End If

                        ' --- REFLECT (连续失败) ---
                        If noProgressCount >= MaxNoProgress Then
                            If replanAttempts >= MaxReplanAttempts Then
                                Dim failMsg = $"Шаг многократно завершался ошибкой, достигнут лимит повторного планирования: {toolResult.ToObserveSummary()}"
                                AppLogger.Error("LoopEngine", failMsg)
                                Return AgentResult.Failed(session.Id, failMsg)
                            End If

                            session.Status = AgentStatus.Reflecting
                            OnStatusChanged?.Invoke("Анализирую причину сбоя и перестраиваю план...")
                            replanAttempts += 1
                            AppLogger.Info("LoopEngine", $"Reflect/replan attempt {replanAttempts}: {toolResult.ToObserveSummary()}")

                            Dim newPlan = Await ReflectAndReplanAsync(session, toolResult.ToObserveSummary(), systemPrompt)
                            If newPlan IsNot Nothing AndAlso newPlan.Steps.Count > 0 Then
                                session.Plan = newPlan
                                stepIndex = 0
                                noProgressCount = 0
                                Continue While
                            Else
                                Dim replanFail = $"Не удалось перестроить план: {toolResult.ToObserveSummary()}"
                                AppLogger.Error("LoopEngine", replanFail)
                                Return AgentResult.Failed(session.Id, replanFail)
                            End If
                        End If
                    End If

                    stepIndex += 1
                End While

                If session.CurrentIteration = 0 Then
                    session.Status = AgentStatus.Failed
                    Dim failMsg = "Задача не выполнила ни одного вызова инструмента. Возможно, шаги плана не содержат разбираемого action или инструменты текущего хоста не загружены."
                    OnStatusChanged?.Invoke(failMsg)
                    AppLogger.Warn("LoopEngine", failMsg)
                    Return AgentResult.Failed(session.Id, failMsg)
                End If

                Dim incompleteSteps = session.Plan.Steps.
                    Where(Function(s) s.Status <> StepStatus.Completed).
                    ToList()
                If incompleteSteps.Count > 0 Then
                    session.Status = AgentStatus.Failed
                    Dim failMsg = $"Задача не завершена, неудачных/невыполненных шагов: {incompleteSteps.Count}: {String.Join("; ", incompleteSteps.Select(Function(s) s.ErrorMessage).Where(Function(m) Not String.IsNullOrWhiteSpace(m)).Take(3))}"
                    OnStatusChanged?.Invoke(failMsg)
                    AppLogger.Warn("LoopEngine", failMsg)
                    Return AgentResult.Failed(session.Id, failMsg)
                End If

                Dim outcomeError = ValidateExecutionOutcome(session)
                If Not String.IsNullOrWhiteSpace(outcomeError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(outcomeError)
                    AppLogger.Warn("LoopEngine", outcomeError)
                    Return AgentResult.Failed(session.Id, outcomeError)
                End If

                ' 完成
                session.Status = AgentStatus.Completed
                Dim finalMsg = $"Задача выполнена, итераций: {session.CurrentIteration}"
                OnStatusChanged?.Invoke(finalMsg)
                AppLogger.Info("LoopEngine", finalMsg)
                Return AgentResult.SuccessResult(session.Id, finalMsg)

            Catch ex As Exception
                session.Status = AgentStatus.Failed
                Dim classified = ExceptionClassifier.Classify(ex)
                OnStatusChanged?.Invoke($"Ошибка выполнения: {classified.UserMessage}")
                AppLogger.Error("LoopEngine", "RunAsync unhandled exception", ex)
                Return AgentResult.Failed(session.Id, $"Исключение при выполнении: [{classified.ErrorCode}] {classified.UserMessage}")
            End Try
        End Function


        ''' <summary>
        ''' 获取撤销管理器（供外部访问）
        ''' </summary>
        Public ReadOnly Property UndoManager As Core.UndoManager
            Get
                Return _undoManager
            End Get
        End Property

    End Class

End Namespace
