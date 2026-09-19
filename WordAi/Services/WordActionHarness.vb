' WordAi\Services\WordActionHarness.vb
' Word 专属 Action Harness：把自然语言请求路由到可执行 capability，而不是在 ChatControl 中堆条件分支。

Imports ShareRibbon
Imports Newtonsoft.Json.Linq

Namespace Services

    Public Enum WordActionKind
        None
        Proofread
        DirectFormatting
        Numbering
        SemanticReformat
    End Enum

    Public Class WordActionPlan
        Public Property Kind As WordActionKind = WordActionKind.None
        Public Property Confidence As Double
        Public Property Reason As String = ""
        Public Property Intent As IntentResult
        Public Property FormattingPlan As FormattingIntentPlan
        Public Property ProofreadPlan As ProofreadIntentPlan
        Public Property Capability As WordCapabilityDescriptor
        Public Property RunRecorder As Agent.Harness.HostCapabilityRunRecorder

        Public ReadOnly Property CapabilitySummary As String
            Get
                If Capability Is Nothing Then Return ""
                Return Capability.ToHumanReadableSummary()
            End Get
        End Property

        Public ReadOnly Property ShouldHandle As Boolean
            Get
                Return Kind <> WordActionKind.None AndAlso Confidence >= 0.55
            End Get
        End Property
    End Class

    Public Enum WordCapabilityExecutionStatus
        Accepted
        Succeeded
        Failed
        Fallback
    End Enum

    Public Class WordCapabilityExecutionResult
        Public Property CapabilityId As String = ""
        Public Property Kind As WordActionKind = WordActionKind.None
        Public Property Status As WordCapabilityExecutionStatus = WordCapabilityExecutionStatus.Accepted
        Public Property Success As Boolean
        Public Property UserMessage As String = ""
        Public Property DebugDetail As String = ""
        Public Property Recoverable As Boolean = True
        Public Property Data As Object

        Public Shared Function FromPlan(plan As WordActionPlan,
                                        status As WordCapabilityExecutionStatus,
                                        success As Boolean,
                                        userMessage As String,
                                        Optional debugDetail As String = "",
                                        Optional data As Object = Nothing,
                                        Optional recoverable As Boolean = True) As WordCapabilityExecutionResult
            Return New WordCapabilityExecutionResult With {
                .CapabilityId = If(plan?.Capability?.Id, ""),
                .Kind = If(plan Is Nothing, WordActionKind.None, plan.Kind),
                .Status = status,
                .Success = success,
                .UserMessage = If(userMessage, ""),
                .DebugDetail = If(debugDetail, ""),
                .Data = data,
                .Recoverable = recoverable
            }
        End Function

        Public Function ToObserveSummary() As String
            Dim state = If(Success, "success", "failed")
            Return $"{CapabilityId} kind={Kind} status={Status} result={state}: {UserMessage}"
        End Function

        Public Function ToToolResult() As Agent.ToolResult
            Dim observation As New JObject From {
                {"kind", "capability"},
                {"summary", If(UserMessage, ToObserveSummary())},
                {"changed", Success},
                {"targetRefs", New JArray("Word:Document")},
                {"warnings", New JArray()},
                {"capabilityId", CapabilityId},
                {"status", Status.ToString()}
            }

            If Success Then
                Return Agent.ToolResult.Succeed(CapabilityId,
                                                If(UserMessage, ToObserveSummary()),
                                                data:=Data,
                                                observation:=observation)
            End If

            Return Agent.ToolResult.Failed(CapabilityId,
                                           If(DebugDetail, UserMessage),
                                           data:=Data,
                                           errorCode:=If(Status = WordCapabilityExecutionStatus.Fallback,
                                                         ExceptionClassifier.CodeUnknown,
                                                         ExceptionClassifier.CodeCom),
                                           userMessage:=If(UserMessage, "Сбой выполнения Capability"),
                                           debugDetail:=DebugDetail,
                                           recoverable:=Recoverable,
                                           observation:=observation)
        End Function
    End Class

    Public Class WordActionHarness

        Private ReadOnly _app As Object

        Public Sub New(app As Object)
            _app = app
        End Sub

        Public Function Plan(userMessage As String, Optional intent As IntentResult = Nothing) As WordActionPlan
            Dim result As New WordActionPlan With {
                .Intent = intent
            }

            If String.IsNullOrWhiteSpace(userMessage) Then Return result

            Dim hasSelection = HasUsableSelection()
            Dim formattingCompiler As New FormattingIntentCompiler()
            Dim formattingPlan = formattingCompiler.Compile(userMessage, hasSelection)
            result.FormattingPlan = formattingPlan
            If ProofreadIntentCompiler.LooksLikeProofreadCommand(userMessage) Then
                Dim proofreadCompiler As New ProofreadIntentCompiler()
                Dim proofreadPlan = proofreadCompiler.Compile(userMessage, hasSelection)
                result.ProofreadPlan = proofreadPlan
                AssignCapability(result,
                                 WordActionKind.Proofread,
                                 Math.Max(0.86, proofreadPlan.Confidence),
                                 "Сформирован исполняемый ProofreadIntentPlan")
                Return result
            End If

            If ShouldRouteToNumbering(userMessage) Then
                AssignCapability(result,
                                 WordActionKind.Numbering,
                                 0.92,
                                 "Совпадение с capability непрерывной автонумерации Word")
                Return result
            End If

            If formattingPlan IsNot Nothing AndAlso formattingPlan.HasOperations Then
                AssignCapability(result,
                                 WordActionKind.DirectFormatting,
                                 Math.Max(0.82, formattingPlan.Confidence),
                                 "Сформирован исполняемый FormattingIntentPlan")
                Return result
            End If

            If intent IsNot Nothing Then
                Select Case intent.OfficeIntent
                    Case OfficeIntentType.PROOFREAD
                        AssignCapability(result,
                                         WordActionKind.Proofread,
                                         Math.Max(0.72, intent.Confidence),
                                         "Намерение распознано как вычитка")
                        Return result

                    Case OfficeIntentType.TEXT_FORMAT
                        AssignCapability(result,
                                         WordActionKind.DirectFormatting,
                                         Math.Max(0.62, intent.Confidence),
                                         "Намерение распознано как настройка текстового формата; передано исполнителю форматирования")
                        Return result

                    Case OfficeIntentType.FORMAT_STYLE
                        AssignCapability(result,
                                         WordActionKind.SemanticReformat,
                                         Math.Max(0.66, intent.Confidence),
                                         "Намерение распознано как настройка стиля/форматирования")
                        Return result
                End Select
            End If

            If LooksLikeStructuralReformat(userMessage) Then
                AssignCapability(result,
                                 WordActionKind.SemanticReformat,
                                 0.7,
                                 "Совпадение с запросом структурного форматирования/упорядочивания заголовков/нумерации")
                Return result
            End If

            Return result
        End Function

        Private Shared Sub AssignCapability(plan As WordActionPlan,
                                            kind As WordActionKind,
                                            confidence As Double,
                                            reason As String)
            If plan Is Nothing Then Return
            plan.Kind = kind
            plan.Confidence = confidence
            plan.Capability = WordCapabilityRegistry.Require(kind)

            Dim reasonParts As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(reason) Then reasonParts.Add(reason)
            If plan.Capability IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(plan.Capability.DisplayName) Then
                reasonParts.Add("capability=" & plan.Capability.DisplayName)
            End If
            plan.Reason = String.Join("; ", reasonParts)
        End Sub

        Private Function ShouldRouteToNumbering(message As String) As Boolean
            If WordNumberingAgent.LooksLikeSequentialNumberingCommand(message) Then Return True

            ' 这里不是靠无限关键词穷举，而是结合文档状态和用户动作意图：
            ' 文档存在自动编号，用户表达“改/整理/重排/连续/12345”等序列化目标时，优先交给编号 capability。
            If CountAutomaticNumberedParagraphs() = 0 Then Return False

            Dim text = If(message, "")
            Dim hasSequenceGoal = ContainsAny(text, {"12345", "1 2 3", "1,2,3", "1，2，3", "连续", "递增", "顺序"})
            Dim hasChangeVerb = ContainsAny(text, {"改", "变", "整理", "修正", "重排", "重置", "统一", "处理"})
            Dim refersToFrontMarker = ContainsAny(text, {"前面", "前面的", "开头", "列表", "编号", "序号"})

            Return hasSequenceGoal AndAlso hasChangeVerb AndAlso refersToFrontMarker
        End Function

        Private Function CountAutomaticNumberedParagraphs() As Integer
            Try
                If _app Is Nothing OrElse _app.ActiveDocument Is Nothing Then Return 0

                Dim count As Integer = 0
                For Each para As Object In _app.ActiveDocument.Paragraphs
                    Try
                        Dim listType = para.Range.ListFormat.ListType
                        If CInt(listType) <> 0 Then count += 1
                    Catch
                    End Try
                    If count >= 1 Then Return count
                Next
            Catch
            End Try

            Return 0
        End Function

        Private Function HasUsableSelection() As Boolean
            Try
                If _app Is Nothing OrElse _app.Selection Is Nothing OrElse _app.Selection.Range Is Nothing Then Return False
                Dim text = If(_app.Selection.Range.Text, "").Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
                Return text.Length > 0
            Catch
                Return False
            End Try
        End Function

        Private Shared Function LooksLikeStructuralReformat(message As String) As Boolean
            Dim text = If(message, "")
            Dim hasStructureTarget = ContainsAny(text, {"标题", "序号", "编号", "层级", "目录", "列表"})
            Dim hasAction = ContainsAny(text, {"重构", "整理", "规范", "优化", "调整", "排版", "格式"})
            Return hasStructureTarget AndAlso hasAction
        End Function

        Private Shared Function ContainsAny(text As String, keywords As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(text) OrElse keywords Is Nothing Then Return False
            For Each keyword In keywords
                If text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next
            Return False
        End Function

    End Class

End Namespace
