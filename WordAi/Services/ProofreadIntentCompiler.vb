' WordAi\Services\ProofreadIntentCompiler.vb
' 将自然语言校对请求编译为结构化计划; 不直接访问 Word COM。

Imports System.Text
Imports System.Collections.Generic
Imports System.Linq

Namespace Services

    Public Enum ProofreadTargetScope
        AutoScope
        Selection
        Document
        CurrentParagraph
    End Enum

    Public Enum ProofreadIssueType
        Typo
        Punctuation
        Grammar
        Wording
        Terminology
        NumberFormat
        StyleConsistency
    End Enum

    Public Enum ProofreadApplyMode
        SuggestOnly
        AutoApplyHighConfidence
    End Enum

    Public Class ProofreadIntentPlan
        Public Property OriginalText As String
        Public Property Source As String = "rules"
        Public Property Confidence As Double = 0
        Public Property Scope As ProofreadTargetScope = ProofreadTargetScope.AutoScope
        Public Property IssueTypes As New List(Of ProofreadIssueType)()
        Public Property ApplyMode As ProofreadApplyMode = ProofreadApplyMode.SuggestOnly
        Public Property ShowSidePanel As Boolean = True
        Public Property Notes As New List(Of String)()

        Public Function ToHumanReadableSummary() As String
            Dim typesText = If(IssueTypes Is Nothing OrElse IssueTypes.Count = 0,
                               "все проблемы",
                               String.Join(", ", IssueTypes.Select(Function(t) IssueTypeToText(t))))
            Return $"Диапазон: {ScopeToText(Scope)}; типы: {typesText}; режим: {If(ApplyMode = ProofreadApplyMode.AutoApplyHighConfidence, "автоисправление надёжных", "предпросмотр предложений")}"
        End Function

        Private Shared Function ScopeToText(scope As ProofreadTargetScope) As String
            Select Case scope
                Case ProofreadTargetScope.Selection : Return "Текущее выделение"
                Case ProofreadTargetScope.Document : Return "Весь документ"
                Case ProofreadTargetScope.CurrentParagraph : Return "Текущий абзац"
                Case Else : Return "Автоматически"
            End Select
        End Function

        Private Shared Function IssueTypeToText(issueType As ProofreadIssueType) As String
            Select Case issueType
                Case ProofreadIssueType.Typo : Return "Опечатки"
                Case ProofreadIssueType.Punctuation : Return "Пунктуация"
                Case ProofreadIssueType.Grammar : Return "Грамматика"
                Case ProofreadIssueType.Wording : Return "Формулировки"
                Case ProofreadIssueType.Terminology : Return "Терминология"
                Case ProofreadIssueType.NumberFormat : Return "Формат чисел"
                Case ProofreadIssueType.StyleConsistency : Return "Единообразие стиля"
                Case Else : Return issueType.ToString()
            End Select
        End Function
    End Class

    Public Class ProofreadIntentCompiler

        Public Shared Function LooksLikeProofreadCommand(message As String) As Boolean
            If String.IsNullOrWhiteSpace(message) Then Return False
            Return ContainsAny(message, {"校对", "审校", "检查错别字", "错别字", "语法错误", "拼写错误", "标点", "润色检查", "proofread"})
        End Function

        Public Function Compile(message As String, hasSelection As Boolean) As ProofreadIntentPlan
            Dim plan As New ProofreadIntentPlan With {
                .OriginalText = If(message, ""),
                .Scope = InferScope(message, hasSelection)
            }

            AddIssueTypes(plan, message)
            If plan.IssueTypes.Count = 0 Then
                plan.IssueTypes.AddRange(New ProofreadIssueType() {
                    ProofreadIssueType.Typo,
                    ProofreadIssueType.Punctuation,
                    ProofreadIssueType.Grammar,
                    ProofreadIssueType.Wording,
                    ProofreadIssueType.Terminology,
                    ProofreadIssueType.NumberFormat,
                    ProofreadIssueType.StyleConsistency
                })
            End If

            If ContainsAny(message, {"自动修正", "自动修改", "直接修正", "直接修改"}) Then
                plan.ApplyMode = ProofreadApplyMode.AutoApplyHighConfidence
                plan.Notes.Add("Автоматически применять только надёжные низкорисковые исправления; остальные проблемы по-прежнему требуют подтверждения на боковой панели.")
            End If

            plan.Confidence = 0.88
            Return plan
        End Function

        Private Function InferScope(message As String, hasSelection As Boolean) As ProofreadTargetScope
            Dim text = If(message, "")
            If ContainsAny(text, {"全文", "整篇", "整个文档", "全部", "所有"}) Then Return ProofreadTargetScope.Document
            If ContainsAny(text, {"当前段", "本段", "这一段"}) Then Return ProofreadTargetScope.CurrentParagraph
            If ContainsAny(text, {"选中", "选区", "所选", "当前选择"}) Then Return ProofreadTargetScope.Selection
            If hasSelection Then Return ProofreadTargetScope.Selection
            Return ProofreadTargetScope.Document
        End Function

        Private Sub AddIssueTypes(plan As ProofreadIntentPlan, message As String)
            If ContainsAny(message, {"错别字", "拼写", "别字"}) Then AddIssueType(plan, ProofreadIssueType.Typo)
            If ContainsAny(message, {"标点", "符号"}) Then AddIssueType(plan, ProofreadIssueType.Punctuation)
            If ContainsAny(message, {"语法", "病句"}) Then AddIssueType(plan, ProofreadIssueType.Grammar)
            If ContainsAny(message, {"表达", "润色", "不通顺", "啰嗦"}) Then AddIssueType(plan, ProofreadIssueType.Wording)
            If ContainsAny(message, {"术语", "名词", "专有名词"}) Then AddIssueType(plan, ProofreadIssueType.Terminology)
            If ContainsAny(message, {"数字", "编号", "日期", "单位"}) Then AddIssueType(plan, ProofreadIssueType.NumberFormat)
            If ContainsAny(message, {"风格", "一致", "统一"}) Then AddIssueType(plan, ProofreadIssueType.StyleConsistency)
        End Sub

        Private Sub AddIssueType(plan As ProofreadIntentPlan, issueType As ProofreadIssueType)
            If Not plan.IssueTypes.Contains(issueType) Then plan.IssueTypes.Add(issueType)
        End Sub

        Private Shared Function ContainsAny(text As String, keywords As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return False
            For Each keyword In keywords
                If text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next
            Return False
        End Function
    End Class

End Namespace
