' ShareRibbon\Services\Reformat\SmartFormattingOrchestrator.vb
' 智能排版编排器 - 连接 DocumentAnalyzer、FormattingKnowledgeEngine、SemanticRenderingEngine
' 支撑速排/对话/克隆三种排版模式。本编排器不直接调用AI，只做数据准备和规则判断。

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

' ============================================================
'  枚举
' ============================================================

''' <summary>用户排版意图类型</summary>
Public Enum IntentType
    ''' <summary>自动排版 - 系统自动检测并推荐最佳标准</summary>
    AutoFormat
    ''' <summary>按指定标准排版 - 用户明确指定排版标准</summary>
    StandardFormat
    ''' <summary>样式克隆 - 照范文格式排版</summary>
    StyleClone
    ''' <summary>具体微调 - 只修改某项格式（如"标题再大一点"）</summary>
    SpecificTweak
    ''' <summary>格式清洗 - 清理混乱空白和格式污染</summary>
    FormatCleanup
End Enum

' ============================================================
'  预览方案（JSON序列化后发送到前端卡片）
' ============================================================

''' <summary>排版预览方案 - 发送到前端展示的完整方案数据</summary>
Public Class ReformatPreviewPlan
    ''' <summary>检测到的文档类型</summary>
    Public Property DetectedType As DocumentType = DocumentType.Unknown
    ''' <summary>类型识别置信度（0-1）</summary>
    Public Property TypeConfidence As Double = 0.0
    ''' <summary>推荐标准名称（如"GB/T 9704-2012"）</summary>
    Public Property StandardName As String = ""
    ''' <summary>标准描述</summary>
    Public Property StandardDescription As String = ""
    ''' <summary>即将发生的格式变更列表</summary>
    Public Property Changes As List(Of FormatChange)
    ''' <summary>总段落数</summary>
    Public Property TotalParagraphs As Integer = 0
    ''' <summary>总样式变更数</summary>
    Public Property TotalStyleChanges As Integer = 0
    ''' <summary>将要应用的语义样式映射</summary>
    Public Property SemanticMapping As SemanticStyleMapping
    ''' <summary>页面设置</summary>
    Public Property PageSettings As PageConfig
    ''' <summary>是否需要AI语义标注（True=Changes中的标签为初步结果，需AI标注后才能渲染）</summary>
    Public Property NeedsAITagging As Boolean = True
    ''' <summary>段落类型列表("text"/"image"/"table"/"formula")，用于渲染时跳过非文本段落</summary>
    Public Property ParagraphTypes As List(Of String) = Nothing
    ''' <summary>用户可读的作用范围摘要，由具体 Office 应用侧填充</summary>
    Public Property ScopeSummary As String = ""
    ''' <summary>作用范围类型名称，如 selection/wholeDocument/section</summary>
    Public Property ScopeKindName As String = ""
    ''' <summary>文本段落数量，排除图片/表格/公式等特殊元素</summary>
    Public Property TextParagraphCount As Integer = 0

    Public Sub New()
        Changes = New List(Of FormatChange)()
        SemanticMapping = New SemanticStyleMapping()
        PageSettings = New PageConfig()
    End Sub

    ' -- consumer-compat aliases --
    Public Property DocumentTypeName As String = ""
    Public ReadOnly Property TotalChanges As Integer
        Get
            Return TotalStyleChanges
        End Get
    End Property
    Public ReadOnly Property SectionCount As Integer
        Get
            Return If(Changes IsNot Nothing, Changes.Count, 0)
        End Get
    End Property

End Class

''' <summary>单处格式变更项（一个段落的格式变化描述）</summary>
Public Class FormatChange
    ''' <summary>段落索引</summary>
    Public Property ParagraphIndex As Integer = -1
    ''' <summary>段落文本预览（前50字）</summary>
    Public Property ParagraphPreview As String = ""
    ''' <summary>变更前的语义标签</summary>
    Public Property OldTag As String = ""
    ''' <summary>变更后的语义标签</summary>
    Public Property NewTag As String = ""
    ''' <summary>新字体描述（如"仿宋_GB2312 16pt"）</summary>
    Public Property NewFont As String = ""
    ''' <summary>新对齐方式（如"居中"、"两端对齐"）</summary>
    Public Property NewAlignment As String = ""
    ''' <summary>新缩进描述（如"首行缩进2字符"）</summary>
    Public Property NewIndent As String = ""
    ''' <summary>变更描述（如"宋体→仿宋 16pt, 居中→两端对齐"）</summary>
    Public Property ChangeDescription As String = ""
End Class

' ============================================================
'  用户意图
' ============================================================

''' <summary>从自然语言中解析出的用户排版意图</summary>
Public Class FormatIntent
    ''' <summary>目标文档类型</summary>
    Public Property TargetDocumentType As DocumentType = DocumentType.Unknown
    ''' <summary>排版意图类型</summary>
    Public Property IntentType As IntentType = IntentType.AutoFormat
    ''' <summary>具体格式要求列表（如"标题用红色", "行距改成1.5"）</summary>
    Public Property SpecificRequests As List(Of String)
    ''' <summary>目标标准名称（如"GB/T 9704-2012"）</summary>
    Public Property TargetStandardName As String = ""

    Public Sub New()
        SpecificRequests = New List(Of String)()
    End Sub
End Class

' ============================================================
'  对话式微调上下文
' ============================================================

''' <summary>对话式排版的微调状态上下文，跟踪当前分析/映射/标准和对话历史</summary>
Public Class RefinementContext
    ''' <summary>当前文档分析结果</summary>
    Public Property CurrentAnalysis As DocumentAnalysisResult
    ''' <summary>当前使用的样式映射（可能已通过"标题再大一点"等指令修改）</summary>
    Public Property CurrentMapping As SemanticStyleMapping
    ''' <summary>当前使用的排版标准</summary>
    Public Property CurrentStandard As FormattingStandard
    ''' <summary>当前预览方案</summary>
    Public Property CurrentPreviewPlan As ReformatPreviewPlan
    ''' <summary>对话历史（用户微调命令列表，用于"再大一点"等上下文敏感指令）</summary>
    Public Property ConversationHistory As List(Of String)
    ''' <summary>是否已经应用到文档</summary>
    Public Property IsApplied As Boolean = False
    ''' <summary>消费者兼容别名：最后一次预览方案</summary>
    Public Property LastPlan As ReformatPreviewPlan
        Get
            Return CurrentPreviewPlan
        End Get
        Set(value As ReformatPreviewPlan)
            CurrentPreviewPlan = value
        End Set
    End Property

    Public Sub New()
        CurrentAnalysis = New DocumentAnalysisResult()
        CurrentMapping = New SemanticStyleMapping()
        CurrentStandard = Nothing
        CurrentPreviewPlan = Nothing
        ConversationHistory = New List(Of String)()
    End Sub

    ''' <summary>一次性更新上下文中的分析/映射/标准和预览方案</summary>
    Public Sub Update(analysis As DocumentAnalysisResult,
                     mapping As SemanticStyleMapping,
                     standard As FormattingStandard,
                     previewPlan As ReformatPreviewPlan)
        CurrentAnalysis = analysis
        CurrentMapping = mapping
        CurrentStandard = standard
        CurrentPreviewPlan = previewPlan
    End Sub

    ''' <summary>记录一条用户微调命令到对话历史</summary>
    Public Sub AddConversation(userCommand As String)
        ConversationHistory.Add(userCommand)
    End Sub

    ''' <summary>清空所有上下文状态</summary>
    Public Sub Clear()
        CurrentAnalysis = New DocumentAnalysisResult()
        CurrentMapping = New SemanticStyleMapping()
        CurrentStandard = Nothing
        CurrentPreviewPlan = Nothing
        ConversationHistory.Clear()
    End Sub
End Class

' ============================================================
'  主编排器
' ============================================================

''' <summary>
''' 智能排版编排器 —— 中央调度器。
''' 上层调用方（ChatFormatterAgent / ReformatService）编排完整流程：
'''   DocumentAnalyzer → FormattingKnowledgeEngine → 本编排器 → SemanticRenderingEngine
''' 本编排器不直接调用 AI，只做数据准备和规则判断。
''' </summary>
Public Class SmartFormattingOrchestrator

    Private ReadOnly _analyzer As DocumentAnalyzer
    Private ReadOnly _knowledgeEngine As FormattingKnowledgeEngine
    Private ReadOnly _standardRegistry As FormattingStandardRegistry
    Private ReadOnly _refinementContext As RefinementContext

    ' ---- 标准名称关键词映射（用于 ParseUserIntent） ----
    Private Shared ReadOnly StandardKeywords As Dictionary(Of String, String) =
        New Dictionary(Of String, String) From {
            {"公文", "GB/T 9704-2012"},
            {"GB/T 9704", "GB/T 9704-2012"},
            {"国标", "GB/T 9704-2012"},
            {"党政", "GB/T 9704-2012"},
            {"学术", "学术论文通用格式"},
            {"论文", "学术论文通用格式"},
            {"参考文献", "GB/T 7714-2015"},
            {"GB/T 7714", "GB/T 7714-2015"},
            {"商务", "商务报告通用规范"},
            {"报告", "商务报告通用规范"},
            {"商业", "商务报告通用规范"},
            {"официальный документ", "GB/T 9704-2012"},
            {"научная статья", "学术论文通用格式"},
            {"бизнес-отчёт", "商务报告通用规范"},
            {"бизнес-отчет", "商务报告通用规范"}
        }

    ''' <summary>获取当前对话式微调的上下文状态</summary>
    Public ReadOnly Property RefinementContext As RefinementContext
        Get
            Return _refinementContext
        End Get
    End Property

    ''' <summary>使用默认 DocumentAnalyzer 和 FormattingKnowledgeEngine 构造</summary>
    Public Sub New()
        _analyzer = New DocumentAnalyzer()
        _knowledgeEngine = New FormattingKnowledgeEngine()
        _standardRegistry = New FormattingStandardRegistry(_knowledgeEngine)
        _refinementContext = New RefinementContext()
    End Sub

    ''' <summary>注入自定义分析器和知识引擎</summary>
    Public Sub New(docAnalyzer As DocumentAnalyzer, knowledgeEngine As FormattingKnowledgeEngine)
        _analyzer = docAnalyzer
        _knowledgeEngine = If(knowledgeEngine, New FormattingKnowledgeEngine())
        _standardRegistry = New FormattingStandardRegistry(_knowledgeEngine)
        _refinementContext = New RefinementContext()
    End Sub

    ' ============================================================
    '  模式A：速排（一键分析 → 推荐 → 预览）
    ' ============================================================

    ''' <summary>
    ''' 分析文档并推荐排版方案。
    ''' 内部流程：Analyze → DetectType → GetStandard → GeneratePreviewPlan
    ''' </summary>
    ''' <param name="paragraphTexts">文档各段落的纯文本列表</param>
    Public Function AnalyzeAndRecommend(paragraphTexts As List(Of String)) As ReformatPreviewPlan
        If paragraphTexts Is Nothing OrElse paragraphTexts.Count = 0 Then
            Return New ReformatPreviewPlan()
        End If

        ' 1. 分析文档类型和结构
        Dim analysis = _analyzer.Analyze(paragraphTexts)

        ' 2. 获取匹配标准
        Dim standard = _standardRegistry.SelectBest(Nothing, analysis)

        If standard Is Nothing Then
            Return New ReformatPreviewPlan With {
                .DetectedType = analysis.DocumentType,
                .TypeConfidence = analysis.Confidence,
                .TotalParagraphs = analysis.ParagraphCount
            }
        End If

        ' 3. 生成预览方案
        Dim plan = GeneratePreviewPlan(analysis, standard, paragraphTexts)

        ' 4. 保存到对话式微调上下文
        _refinementContext.Update(analysis, plan.SemanticMapping, standard, plan)

        Return plan
    End Function

    ''' <summary>
    ''' 分析文档并推荐排版方案（增强版，接受Word富文本格式信息）
    ''' </summary>
    ''' <param name="paragraphTexts">文档各段落的纯文本列表</param>
    ''' <param name="paragraphStyles">段落样式名称列表</param>
    ''' <param name="paragraphFontSizes">段落字号列表（pt）</param>
    ''' <param name="paragraphIsBold">段落是否加粗列表</param>
    Public Function AnalyzeAndRecommend(paragraphTexts As List(Of String),
                                        paragraphStyles As List(Of String),
                                        paragraphFontSizes As List(Of Single),
                                        paragraphIsBold As List(Of Boolean)) As ReformatPreviewPlan
        If paragraphTexts Is Nothing OrElse paragraphTexts.Count = 0 Then
            Return New ReformatPreviewPlan()
        End If

        ' 使用增强版分析器
        Dim analysis = _analyzer.Analyze(paragraphTexts, paragraphStyles, paragraphFontSizes, paragraphIsBold)

        Dim standard = _standardRegistry.SelectBest(Nothing, analysis)

        If standard Is Nothing Then
            Return New ReformatPreviewPlan With {
                .DetectedType = analysis.DocumentType,
                .TypeConfidence = analysis.Confidence,
                .TotalParagraphs = analysis.ParagraphCount
            }
        End If

        Dim plan = GeneratePreviewPlan(analysis, standard, paragraphTexts)
        _refinementContext.Update(analysis, plan.SemanticMapping, standard, plan)
        Return plan
    End Function

    ''' <summary>
    ''' 生成备选排版方案。用于“换一种”，按候选标准循环，而不是重复返回同一推荐。
    ''' </summary>
    Public Function GenerateAlternativePlan(paragraphTexts As List(Of String),
                                            paragraphStyles As List(Of String),
                                            paragraphFontSizes As List(Of Single),
                                            paragraphIsBold As List(Of Boolean),
                                            currentStandardName As String,
                                            variantIndex As Integer) As ReformatPreviewPlan
        If paragraphTexts Is Nothing OrElse paragraphTexts.Count = 0 Then
            Return New ReformatPreviewPlan()
        End If

        Dim analysis = _analyzer.Analyze(paragraphTexts, paragraphStyles, paragraphFontSizes, paragraphIsBold)
        Dim candidates = GetAlternativeStandardCandidates(analysis, currentStandardName)
        If candidates.Count = 0 Then
            Return AnalyzeAndRecommend(paragraphTexts, paragraphStyles, paragraphFontSizes, paragraphIsBold)
        End If

        Dim safeIndex = Math.Abs(variantIndex) Mod candidates.Count
        Dim selected = candidates(safeIndex)
        Dim plan = GeneratePreviewPlan(analysis, selected.Standard, paragraphTexts)
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        plan.StandardDescription = BuildAlternativeDescription(selected, safeIndex + 1, candidates.Count)
        _refinementContext.Update(analysis, plan.SemanticMapping, selected.Standard, plan)
        Return plan
    End Function

    Private Function GetAlternativeStandardCandidates(analysis As DocumentAnalysisResult,
                                                      currentStandardName As String) As List(Of FormattingStandardCandidate)
        Dim docTypeName = If(analysis Is Nothing, DocumentType.GeneralDocument.ToString(), analysis.DocumentType.ToString())
        Dim allCandidates = _standardRegistry.GetAllCandidates().
            Where(Function(c) c IsNot Nothing AndAlso
                              c.Standard IsNot Nothing AndAlso
                              c.Standard.IsActive).
            Where(Function(c) Not String.Equals(c.Standard.Name, currentStandardName, StringComparison.OrdinalIgnoreCase)).
            ToList()

        Dim sameType = allCandidates.
            Where(Function(c) c.Standard.ApplicableDocumentTypes IsNot Nothing AndAlso
                              c.Standard.ApplicableDocumentTypes.Contains(docTypeName)).
            OrderByDescending(Function(c) c.Confidence).
            ToList()

        Dim general = allCandidates.
            Where(Function(c) c.Standard.ApplicableDocumentTypes IsNot Nothing AndAlso
                              c.Standard.ApplicableDocumentTypes.Contains(DocumentType.GeneralDocument.ToString())).
            OrderByDescending(Function(c) c.Confidence).
            ToList()

        Dim merged As New List(Of FormattingStandardCandidate)()
        For Each candidate In sameType.Concat(general).Concat(allCandidates.OrderByDescending(Function(c) c.Confidence))
            If Not merged.Any(Function(x) String.Equals(x.Standard.Name, candidate.Standard.Name, StringComparison.OrdinalIgnoreCase)) Then
                merged.Add(candidate)
            End If
        Next

        Return merged
    End Function

    Private Shared Function BuildAlternativeDescription(candidate As FormattingStandardCandidate,
                                                        displayIndex As Integer,
                                                        totalCount As Integer) As String
        If candidate Is Nothing OrElse candidate.Standard Is Nothing Then Return ""

        Dim parts As New List(Of String)()
        parts.Add($"Вариант {displayIndex}/{totalCount}")
        If Not String.IsNullOrWhiteSpace(candidate.Reason) Then parts.Add(candidate.Reason)
        If Not String.IsNullOrWhiteSpace(candidate.Standard.Description) Then parts.Add(candidate.Standard.Description)
        Return String.Join("; ", parts)
    End Function

    ' ============================================================
    '  模式B：对话式排版 —— 自然语言指令解析
    ' ============================================================

    ''' <summary>
    ''' 解析用户自然语言排版指令。
    ''' 纯规则实现（不调用AI），返回结构化的 FormatIntent。
    ''' 调用方（ChatFormatterAgent）可进一步用 AI 增强解析结果。
    ''' </summary>
    ''' <param name="userMessage">用户输入的自然语言消息</param>
    ''' <param name="analysis">当前文档分析结果（用于上下文推断）</param>
    Public Function ParseUserIntent(userMessage As String, analysis As DocumentAnalysisResult) As FormatIntent
        Dim intent As New FormatIntent()
        If String.IsNullOrWhiteSpace(userMessage) Then Return intent

        Dim message = userMessage.Trim()

        ' ---- 1. 识别意图类型 ----

        ' 克隆意图
        If ContainsAny(message, {"克隆", "照这个", "范文", "模仿", "参照", "格式克隆", "скопируй формат", "как образец", "по образцу", "скопировать стиль", "повтори оформление", "клонировать", "образец"}) Then
            intent.IntentType = IntentType.StyleClone
            Return intent
        End If

        ' 清洗意图
        If ContainsAny(message, {"清洗", "清理", "清除格式", "去除格式", "格式清理", "очисти формат", "убери формат", "сбрось формат", "убрать форматирование", "очистить форматирование", "сбросить формат"}) Then
            intent.IntentType = IntentType.FormatCleanup
            Return intent
        End If

        ' 标准指定意图（如 "按公文标准"、"用GB/T 9704"）
        For Each kvp In StandardKeywords
            If message.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0 Then
                intent.TargetStandardName = kvp.Value
                intent.IntentType = IntentType.StandardFormat
                Exit For
            End If
        Next

        ' 微调意图（"大一点"、"红色"、"加粗" 等）
        If Not HasTweakIntent(message) Then
            ' 未匹配到其他目的时默认为自动排版
            If intent.IntentType = IntentType.AutoFormat AndAlso
               ContainsAny(message, {"排版", "格式", "段落", "字体", "行距", "整理", "формат", "оформи", "отформатируй", "стиль", "абзац", "шрифт", "интервал", "выравнивание", "оформление", "отформатировать"}) Then
                intent.IntentType = IntentType.AutoFormat
            End If
        Else
            intent.IntentType = IntentType.SpecificTweak
        End If

        ' ---- 2. 提取具体格式请求 ----

        ' 字号调整
        If Regex.IsMatch(message, "(大|小).{0,4}(点|一些|一点)") Then
            intent.SpecificRequests.Add("adjust_font_size")
        End If
        If Regex.IsMatch(message, "(标题|正文).{0,4}(大|小)") Then
            intent.SpecificRequests.Add("adjust_font_size")
        End If

        ' 颜色
        Dim colors = {"红色", "蓝色", "黑色", "绿色", "白色", "красный", "синий", "чёрный", "черный", "зелёный", "зеленый", "белый", "голубой"}
        For Each color In colors
            If message.Contains(color) Then
                intent.SpecificRequests.Add($"color_{color}")
            End If
        Next

        ' 行距
        Dim lineSpacingMatch = Regex.Match(message, "行距.{0,4}([\d.]+)")
        If lineSpacingMatch.Success Then
            intent.SpecificRequests.Add($"line_spacing_{lineSpacingMatch.Groups(1).Value}")
        End If

        ' 对齐
        If ContainsAny(message, {"居中", "居中对齐", "по центру", "выровнять по центру"}) Then
            intent.SpecificRequests.Add("align_center")
        ElseIf ContainsAny(message, {"左对齐", "靠左", "по левому краю", "по левому"}) Then
            intent.SpecificRequests.Add("align_left")
        ElseIf ContainsAny(message, {"右对齐", "靠右", "по правому краю", "по правому"}) Then
            intent.SpecificRequests.Add("align_right")
        ElseIf ContainsAny(message, {"两端对齐", "по ширине", "выровнять по ширине", "выровнять"}) Then
            intent.SpecificRequests.Add("align_justify")
        End If

        ' 加粗
        If ContainsAny(message, {"加粗", "粗体", "粗一点", "жирный", "полужирный", "жирнее", "жирным"}) Then
            intent.SpecificRequests.Add("bold")
        End If

        ' 缩进
        If ContainsAny(message, {"缩进", "отступ", "отступ первой строки", "красная строка"}) Then
            Dim indentMatch = Regex.Match(message, "缩进.{0,4}([\d.]+)")
            If indentMatch.Success Then
                intent.SpecificRequests.Add($"indent_{indentMatch.Groups(1).Value}")
            Else
                intent.SpecificRequests.Add("indent")
            End If
        End If

        ' ---- 3. 识别目标文档类型 ----

        If ContainsAny(message, {"公文", "通知", "决定", "批复", "请示", "函", "официальный", "приказ", "распоряжение", "уведомление", "письмо"}) Then
            intent.TargetDocumentType = DocumentType.OfficialDocument
        ElseIf ContainsAny(message, {"论文", "学术", "期刊", "学报", "научная", "научный", "статья", "диссертация"}) Then
            intent.TargetDocumentType = DocumentType.AcademicPaper
        ElseIf ContainsAny(message, {"报告", "商务", "商业", "汇报", "总结", "отчёт", "отчет", "доклад", "бизнес"}) Then
            intent.TargetDocumentType = DocumentType.BusinessReport
        ElseIf ContainsAny(message, {"合同", "协议", "合约", "договор", "контракт"}) Then
            intent.TargetDocumentType = DocumentType.Contract
        ElseIf ContainsAny(message, {"简历", "履历", "резюме"}) Then
            intent.TargetDocumentType = DocumentType.[Resume]
        Else
            ' 从标准名称反推文档类型
            If Not String.IsNullOrEmpty(intent.TargetStandardName) Then
                Dim standard = _knowledgeEngine.GetStandardByName(intent.TargetStandardName)
                If standard IsNot Nothing AndAlso standard.ApplicableDocumentTypes.Count > 0 Then
                    Dim parsed As DocumentType
                    If [Enum].TryParse(Of DocumentType)(standard.ApplicableDocumentTypes(0), parsed) Then
                        intent.TargetDocumentType = parsed
                    End If
                End If
            End If
        End If

        Return intent
    End Function

    ' ============================================================
    '  预览方案生成
    ' ============================================================

    ''' <summary>
    ''' 根据文档分析结果和排版标准生成完整的预览方案。
    ''' 遍历段落结构，为标题和正文分别生成 FormatChange 条目。
    ''' </summary>
    ''' <param name="analysis">DocumentAnalyzer 的分析结果</param>
    ''' <param name="standard">目标排版标准</param>
    ''' <param name="paragraphTexts">原始段落文本（可选，用于生成段落预览）</param>
    Public Function GeneratePreviewPlan(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        Optional paragraphTexts As List(Of String) = Nothing) As ReformatPreviewPlan

        Dim plan As New ReformatPreviewPlan()
        If analysis Is Nothing OrElse standard Is Nothing Then Return plan

        plan.DetectedType = analysis.DocumentType
        plan.TypeConfidence = analysis.Confidence
        plan.StandardName = standard.Name
        plan.StandardDescription = standard.Description
        plan.TotalParagraphs = analysis.ParagraphCount
        If standard.SemanticMapping Is Nothing Then standard.SemanticMapping = New SemanticStyleMapping()
        standard.SemanticMapping.EnsureBaselineTags()
        plan.SemanticMapping = standard.SemanticMapping
        plan.PageSettings = standard.SemanticMapping.PageConfig

        ' --- 标题变更 ---
        Dim headingChanges = BuildHeadingChanges(analysis, standard, paragraphTexts)
        plan.Changes.AddRange(headingChanges)

        ' --- 标题/编号兜底：分析器保守时，仍按常见编号形态生成结构候选 ---
        Dim structuralHeadingChanges = BuildHeuristicHeadingChanges(analysis, standard, paragraphTexts)
        plan.Changes.AddRange(structuralHeadingChanges)

        ' --- 正文变更（跳过已处理的标题段落） ---
        Dim bodyChanges = BuildBodyChanges(analysis, standard, paragraphTexts)
        plan.Changes.AddRange(bodyChanges)

        ' --- 去重（同一段落不出现两次） ---
        Dim seenIndices As New HashSet(Of Integer)()
        Dim distinctChanges As New List(Of FormatChange)()
        For Each ch In plan.Changes
            If seenIndices.Add(ch.ParagraphIndex) Then
                distinctChanges.Add(ch)
            End If
        Next
        plan.Changes = distinctChanges

        plan.TotalStyleChanges = plan.Changes.Count

        Return plan
    End Function

    ' ============================================================
    '  微调处理
    ' ============================================================

    ''' <summary>
    ''' 应用用户的自然语言微调指令（如"标题再大一点"、"正文用红色"）。
    ''' 直接修改 RefinementContext 中的 SemanticStyleMapping，然后返回更新后的预览方案。
    ''' </summary>
    ''' <param name="refinementCommand">用户微调指令</param>
    Public Function ApplyRefinement(refinementCommand As String) As ReformatPreviewPlan
        If String.IsNullOrWhiteSpace(refinementCommand) Then
            Return _refinementContext.CurrentPreviewPlan
        End If

        ' 记录命令到历史
        _refinementContext.AddConversation(refinementCommand)

        Dim command = refinementCommand.Trim()
        Dim mapping = _refinementContext.CurrentMapping
        If mapping Is Nothing Then Return Nothing

        ' 1. 确定要修改的目标标签
        Dim targetTags = GetTargetTags(command, mapping)

        ' 2. 应用微调
        ApplyTweakToTags(command, targetTags, _refinementContext)

        ' 3. 更新时间戳
        mapping.LastModified = DateTime.Now

        ' 4. 重新生成预览方案
        If _refinementContext.CurrentAnalysis IsNot Nothing AndAlso
           _refinementContext.CurrentStandard IsNot Nothing Then

            Dim updatedPlan = GeneratePreviewPlan(
                _refinementContext.CurrentAnalysis,
                _refinementContext.CurrentStandard)

            ' 覆盖为标准映射（微调后的映射才是实际要用的）
            updatedPlan.SemanticMapping = mapping

            _refinementContext.CurrentPreviewPlan = updatedPlan
            Return updatedPlan
        End If

        Return _refinementContext.CurrentPreviewPlan
    End Function

    ' ============================================================
    '  模式C：范文克隆
    ' ============================================================

    ''' <summary>消费者兼容方法：是否有活动的排版上下文</summary>
    Public Function HasActiveContext() As Boolean
        Return _refinementContext.CurrentPreviewPlan IsNot Nothing AndAlso Not _refinementContext.IsApplied
    End Function

    ''' <summary>消费者兼容方法：一键速排</summary>
    Public Async Function QuickReformatAsync(paragraphs As List(Of String),
                                              wordParagraphs As List(Of Object)) As Task(Of ReformatPreviewPlan)
        Dim plan = AnalyzeAndRecommend(paragraphs)
        _refinementContext.CurrentPreviewPlan = plan
        _refinementContext.IsApplied = False
        plan.NeedsAITagging = True
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        Return plan
    End Function

    ''' <summary>消费者兼容方法：对话式排版（意图驱动）</summary>
    Public Async Function ChatReformatAsync(userMessage As String,
                                             paragraphs As List(Of String),
                                             wordParagraphs As List(Of Object)) As Task(Of ReformatPreviewPlan)

        ' 解析用户意图
        Dim analysis = If(_refinementContext.CurrentAnalysis, New DocumentAnalysisResult())
        Dim intent = ParseUserIntent(userMessage, analysis)
        NormalizeStructuralIntent(userMessage, intent)

        Dim plan As ReformatPreviewPlan

        Select Case intent.IntentType
            Case IntentType.SpecificTweak
                ' 微调模式：在当前映射基础上做增量修改，不需要重新AI标注
                plan = ApplyRefinement(userMessage)
                If plan Is Nothing Then
                    ' 没有活动上下文，降级为自动排版
                    plan = AnalyzeAndRecommend(paragraphs)
                End If

            Case IntentType.StandardFormat
                ' 指定标准排版：用用户指定的标准（如"按公文排版"→GB/T 9704-2012）
                Dim standard As FormattingStandard = Nothing
                If Not String.IsNullOrEmpty(intent.TargetStandardName) Then
                    standard = _standardRegistry.FindStandardByName(intent.TargetStandardName)
                ElseIf intent.TargetDocumentType <> DocumentType.Unknown Then
                    standard = _standardRegistry.GetStandardForDocumentType(intent.TargetDocumentType)
                End If

                If standard IsNot Nothing Then
                    ' 先做文档分析获取结构，再用指定标准生成预览
                    Dim docAnalysis = _analyzer.Analyze(paragraphs)
                    plan = GeneratePreviewPlan(docAnalysis, standard, paragraphs)
                    ' 用户明确指定了标准时，用用户意图覆盖分析器猜测的文档类型
                    If intent.TargetDocumentType <> DocumentType.Unknown Then
                        plan.DetectedType = intent.TargetDocumentType
                    End If
                    plan.DocumentTypeName = standard.Name
                    _refinementContext.Update(docAnalysis, plan.SemanticMapping, standard, plan)
                Else
                    ' 标准未找到，降级为自动排版
                    plan = AnalyzeAndRecommend(paragraphs)
                End If

            Case IntentType.StyleClone
                ' 格式克隆：由上层HandleMirrorFormat处理，此处降级为自动排版
                plan = AnalyzeAndRecommend(paragraphs)

            Case IntentType.FormatCleanup
                ' 格式清洗：先自动分析，清洗逻辑在渲染时处理
                plan = AnalyzeAndRecommend(paragraphs)

            Case Else
                ' AutoFormat 或 Unknown：自动检测文档类型并推荐标准
                plan = AnalyzeAndRecommend(paragraphs)
        End Select

        _refinementContext.CurrentPreviewPlan = plan
        _refinementContext.IsApplied = False
        plan.NeedsAITagging = True
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        Return plan
    End Function

    ''' <summary>消费者兼容方法：对话式排版（意图驱动，增强版）</summary>
    Public Async Function ChatReformatAsync(userMessage As String,
                                             paragraphs As List(Of String),
                                             wordParagraphs As List(Of Object),
                                             paragraphStyles As List(Of String),
                                             paragraphFontSizes As List(Of Single),
                                             paragraphIsBold As List(Of Boolean),
                                             Optional recognizedIntent As FormatIntent = Nothing) As Task(Of ReformatPreviewPlan)

        ' 解析用户意图
        Dim analysis = If(_refinementContext.CurrentAnalysis, New DocumentAnalysisResult())
        Dim intent As FormatIntent = recognizedIntent
        If intent Is Nothing Then
            intent = ParseUserIntent(userMessage, analysis)
        End If
        NormalizeStructuralIntent(userMessage, intent)

        Dim plan As ReformatPreviewPlan

        Select Case intent.IntentType
            Case IntentType.SpecificTweak
                plan = ApplyRefinement(userMessage)
                If plan Is Nothing Then
                    plan = AnalyzeAndRecommend(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
                End If

            Case IntentType.StandardFormat
                Dim standard As FormattingStandard = Nothing
                If Not String.IsNullOrEmpty(intent.TargetStandardName) Then
                    standard = _standardRegistry.FindStandardByName(intent.TargetStandardName)
                ElseIf intent.TargetDocumentType <> DocumentType.Unknown Then
                    standard = _standardRegistry.GetStandardForDocumentType(intent.TargetDocumentType)
                End If
                If standard IsNot Nothing Then
                    Dim docAnalysis = _analyzer.Analyze(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
                    plan = GeneratePreviewPlan(docAnalysis, standard, paragraphs)
                    ' 用户明确指定了标准时，用用户意图覆盖分析器猜测的文档类型
                    If intent.TargetDocumentType <> DocumentType.Unknown Then
                        plan.DetectedType = intent.TargetDocumentType
                    End If
                    plan.DocumentTypeName = standard.Name
                    _refinementContext.Update(docAnalysis, plan.SemanticMapping, standard, plan)
                Else
                    plan = AnalyzeAndRecommend(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
                End If

            Case Else
                plan = AnalyzeAndRecommend(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
        End Select

        _refinementContext.CurrentPreviewPlan = plan
        _refinementContext.IsApplied = False
        plan.NeedsAITagging = True
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        Return plan
    End Function

    ''' <summary>获取文档类型的中文名称</summary>
    Private Function GetDocumentTypeName(docType As DocumentType) As String
        Select Case docType
            Case DocumentType.OfficialDocument : Return "Официальный документ"
            Case DocumentType.AcademicPaper : Return "Научная статья"
            Case DocumentType.BusinessReport : Return "Бизнес-отчёт"
            Case DocumentType.Contract : Return "Договор/соглашение"
            Case DocumentType.[Resume] : Return "Резюме"
            Case DocumentType.GeneralDocument : Return "Обычный документ"
            Case Else : Return "Неизвестно"
        End Select
    End Function

    Private Shared Sub NormalizeStructuralIntent(userMessage As String, intent As FormatIntent)
        If intent Is Nothing OrElse Not IsHeadingNumberingRequest(userMessage) Then Return

        If intent.IntentType = IntentType.SpecificTweak Then
            intent.IntentType = IntentType.AutoFormat
        End If
        If intent.TargetDocumentType = DocumentType.Unknown Then
            intent.TargetDocumentType = DocumentType.GeneralDocument
        End If
        If Not intent.SpecificRequests.Contains("rebuild_heading_numbering") Then
            intent.SpecificRequests.Add("rebuild_heading_numbering")
        End If
    End Sub

    ' ============================================================
    '  内部辅助 —— 生成变更
    ' ============================================================

    ''' <summary>
    ''' 根据段落位置和内容推断合理的默认语义标签（公开版本，供ChatFormatterAgent等外部调用）
    ''' </summary>
    Public Shared Function InferDefaultTagPublic(
        paraIndex As Integer,
        totalParagraphs As Integer,
        text As String,
        standard As FormattingStandard) As String
        Return InferDefaultTag(paraIndex, totalParagraphs, text, standard)
    End Function

    ''' <summary>
    ''' 根据段落位置和内容推断合理的默认语义标签
    ''' </summary>
    Private Shared Function InferDefaultTag(
        paraIndex As Integer,
        totalParagraphs As Integer,
        text As String,
        standard As FormattingStandard) As String

        Dim trimmed = text.Trim()
        Dim mapping = standard?.SemanticMapping
        If mapping Is Nothing Then Return "body.normal"

        ' 文档末尾附近（最后3个非空段落）的短段落可能是署名/日期
        If paraIndex >= totalParagraphs - 4 Then
            If trimmed.Length <= 30 Then
                ' 匹配日期格式：2024年1月15日
                If System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d{4}年\d{1,2}月\d{1,2}日") OrElse
                   System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d{2}\.\d{2}\.\d{4}") Then
                    If mapping.FindTag("footer.date") IsNot Nothing Then Return "footer.date"
                End If
                ' 匹配机构署名：以"政府""办公室""局""部""厅""委员会"结尾
                If System.Text.RegularExpressions.Regex.IsMatch(trimmed, "(政府|办公室|局|部|厅|委员会|党委)$") Then
                    If mapping.FindTag("footer.signature") IsNot Nothing Then Return "footer.signature"
                End If
                ' 抄送
                If trimmed.StartsWith("抄送") OrElse trimmed.StartsWith("Копия") Then
                    If mapping.FindTag("footer.cc") IsNot Nothing Then Return "footer.cc"
                End If
            End If
        End If

        ' 包含发文字号格式
        If System.Text.RegularExpressions.Regex.IsMatch(trimmed, ".+发〔\d{4}〕\d*号") Then
            If mapping.FindTag("header.refno") IsNot Nothing Then Return "header.refno"
        End If

        ' 包含"签发人"
        If trimmed.StartsWith("签发人") Then
            If mapping.FindTag("header.signer") IsNot Nothing Then Return "header.signer"
        End If

        ' 主送机关（以"各"开头且以"："结尾）
        If trimmed.StartsWith("各") AndAlso trimmed.EndsWith("：") Then
            If mapping.FindTag("title.recipient") IsNot Nothing Then Return "title.recipient"
        End If

        ' 附件说明
        If trimmed.StartsWith("附件") OrElse trimmed.StartsWith("Приложение") Then
            If mapping.FindTag("body.attachment") IsNot Nothing Then Return "body.attachment"
        End If

        If IsOfficialDocumentStandard(standard) Then
            If System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^[一二三四五六七八九十]+、") Then
                If mapping.FindTag("title.1") IsNot Nothing Then Return "title.1"
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^（[一二三四五六七八九十]+）") OrElse
               System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\([一二三四五六七八九十]+\)") Then
                If mapping.FindTag("title.2") IsNot Nothing Then Return "title.2"
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d+[\.．、]") OrElse
               System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d+[.)]") Then
                If mapping.FindTag("title.3") IsNot Nothing Then Return "title.3"
            End If
            If trimmed.StartsWith("附注") OrElse
               (trimmed.StartsWith("（") AndAlso (trimmed.Contains("联系人") OrElse trimmed.Contains("联系电话") OrElse trimmed.Contains("电话"))) Then
                If mapping.FindTag("footer.note") IsNot Nothing Then Return "footer.note"
            End If
        End If

        ' 摘要区域（文档开头附近的非标题段落）
        If paraIndex < 5 Then
            If trimmed.StartsWith("摘要") OrElse trimmed.StartsWith("摘　要") Then
                If mapping.FindTag("title.abstract") IsNot Nothing Then Return "title.abstract"
            End If
            If trimmed.StartsWith("关键词") OrElse trimmed.StartsWith("关键字") Then
                If mapping.FindTag("title.keywords") IsNot Nothing Then Return "title.keywords"
            End If
        End If

        ' 默认：正文
        Return "body.normal"
    End Function

    ''' <summary>根据标题结构生成格式变更条目</summary>
    Private Function BuildHeadingChanges(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        paragraphTexts As List(Of String)) As List(Of FormatChange)

        Dim changes As New List(Of FormatChange)()
        If analysis?.DocStructure Is Nothing Then Return changes

        For Each heading In analysis.DocStructure.Headings
            Dim change As New FormatChange()
            change.ParagraphIndex = heading.ParagraphIndex
            change.ParagraphPreview = TruncateText(heading.Text, 50)
            change.OldTag = $"heading.{heading.Level}"
            change.NewTag = GetStandardTagForHeading(heading.Level, standard)

            Dim semanticTag = standard?.SemanticMapping?.FindTag(change.NewTag)
            If semanticTag IsNot Nothing Then
                ResolveChangeFromTag(change, semanticTag)
            End If

            changes.Add(change)
        Next

        Return changes
    End Function

    ''' <summary>根据正文段落列表生成格式变更条目</summary>
    Private Function BuildBodyChanges(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        paragraphTexts As List(Of String)) As List(Of FormatChange)

        Dim changes As New List(Of FormatChange)()
        If analysis?.DocStructure Is Nothing OrElse paragraphTexts Is Nothing Then Return changes

        ' 收集所有非正文段落索引（标题/列表/表格）
        Dim excludeIndices As New HashSet(Of Integer)()
        For Each h In analysis.DocStructure.Headings
            excludeIndices.Add(h.ParagraphIndex)
        Next
        For Each idx In analysis.DocStructure.ListParagraphIndices
            excludeIndices.Add(idx)
        Next
        For Each idx In analysis.DocStructure.TableParagraphIndices
            excludeIndices.Add(idx)
        Next

        ' 每个正文段落尝试映射到合适的语义标签
        Dim bodyTag = standard?.SemanticMapping?.FindTag("body.normal")

        For i = 0 To paragraphTexts.Count - 1
            Dim text = paragraphTexts(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For
            If excludeIndices.Contains(i) Then Continue For

            Dim change As New FormatChange()
            change.ParagraphIndex = i
            change.ParagraphPreview = TruncateText(text, 50)
            change.OldTag = "body.normal"

            ' 根据段落位置和内容推断合理的默认标签
            Dim inferredTag = InferDefaultTag(i, paragraphTexts.Count, text, standard)
            change.NewTag = inferredTag

            ' 用推断的标签解析格式信息
            Dim targetTag = standard?.SemanticMapping?.FindTag(inferredTag)
            If targetTag Is Nothing Then targetTag = bodyTag
            If targetTag IsNot Nothing Then
                ResolveChangeFromTag(change, targetTag)
            End If

            changes.Add(change)
        Next

        Return changes
    End Function

    ''' <summary>在分析器没有识别出标题时，用编号/标题文本形态补充结构候选。</summary>
    Private Function BuildHeuristicHeadingChanges(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        paragraphTexts As List(Of String)) As List(Of FormatChange)

        Dim changes As New List(Of FormatChange)()
        If paragraphTexts Is Nothing OrElse paragraphTexts.Count = 0 Then Return changes

        Dim existingIndices As New HashSet(Of Integer)()
        If analysis?.DocStructure IsNot Nothing Then
            For Each h In analysis.DocStructure.Headings
                existingIndices.Add(h.ParagraphIndex)
            Next
        End If

        For i = 0 To paragraphTexts.Count - 1
            If existingIndices.Contains(i) Then Continue For

            Dim text = If(paragraphTexts(i), "")
            Dim level = InferHeadingLevelFromText(text, standard)
            If level <= 0 Then Continue For

            Dim change As New FormatChange()
            change.ParagraphIndex = i
            change.ParagraphPreview = TruncateText(text.Trim(), 50)
            change.OldTag = "body.normal"
            change.NewTag = GetStandardTagForHeading(level, standard)

            Dim semanticTag = standard?.SemanticMapping?.FindTag(change.NewTag)
            If semanticTag IsNot Nothing Then
                ResolveChangeFromTag(change, semanticTag)
            End If

            changes.Add(change)
        Next

        Return changes
    End Function

    Private Shared Function InferHeadingLevelFromText(text As String, standard As FormattingStandard) As Integer
        If String.IsNullOrWhiteSpace(text) Then Return 0

        Dim trimmed = text.Trim()
        If trimmed.Length <= 1 OrElse trimmed.Length > 100 Then Return 0

        If Regex.IsMatch(trimmed, "^第[一二三四五六七八九十百千万\d]+[章节篇条]\s*") Then Return 1
        If Regex.IsMatch(trimmed, "^[一二三四五六七八九十]+[、.．]\s*\S+") Then Return 1
        If Regex.IsMatch(trimmed, "^[（(][一二三四五六七八九十\d]+[）)]\s*\S+") Then Return 2
        If Regex.IsMatch(trimmed, "^\d+[.．]\d+[.．]\d+\s*\S*") Then Return 3
        If Regex.IsMatch(trimmed, "^\d+[.．]\d+\s*\S*") Then Return 2
        If Regex.IsMatch(trimmed, "^\d+[.．、]\s*\S+") Then
            Return If(IsOfficialDocumentStandard(standard), 3, 1)
        End If
        If Regex.IsMatch(trimmed, "^\d+[.)]\s*\S+") Then
            Return If(IsOfficialDocumentStandard(standard), 3, 1)
        End If

        Return 0
    End Function

    ''' <summary>将语义标签的格式信息填入 FormatChange</summary>
    Private Shared Sub ResolveChangeFromTag(change As FormatChange, tag As SemanticTag)
        If tag Is Nothing Then Return

        change.NewFont = FormatFontDescription(tag.Font)
        change.NewAlignment = GetAlignmentDisplayName(tag.Paragraph?.Alignment)

        If tag.Paragraph IsNot Nothing AndAlso tag.Paragraph.FirstLineIndent > 0 Then
            change.NewIndent = $"首行缩进{tag.Paragraph.FirstLineIndent}字符"
        End If

        change.ChangeDescription = BuildFormatDescription(tag)
    End Sub

    ''' <summary>获取标准中对应标题级别的最佳语义标签ID，逐级回退</summary>
    Private Shared Function GetStandardTagForHeading(level As Integer, standard As FormattingStandard) As String
        If standard?.SemanticMapping Is Nothing Then Return $"heading.{level}"

        If IsOfficialDocumentStandard(standard) Then
            Dim officialTagId = $"title.{level}"
            If standard.SemanticMapping.FindTag(officialTagId) IsNot Nothing Then
                Return officialTagId
            End If
        End If

        Dim tagId = $"heading.{level}"
        If standard.SemanticMapping.FindTag(tagId) IsNot Nothing Then
            Return tagId
        End If

        Dim titleTagId = $"title.{level}"
        If standard.SemanticMapping.FindTag(titleTagId) IsNot Nothing Then
            Return titleTagId
        End If

        ' 逐级回退
        For fallback = level - 1 To 1 Step -1
            Dim fallbackId = $"heading.{fallback}"
            If standard.SemanticMapping.FindTag(fallbackId) IsNot Nothing Then
                Return fallbackId
            End If
            Dim fallbackTitleId = $"title.{fallback}"
            If standard.SemanticMapping.FindTag(fallbackTitleId) IsNot Nothing Then
                Return fallbackTitleId
            End If
        Next

        Return "body.normal"
    End Function

    Private Shared Function IsOfficialDocumentStandard(standard As FormattingStandard) As Boolean
        If standard Is Nothing Then Return False
        Dim text = $"{standard.Id} {standard.Name} {standard.Description}".ToLowerInvariant()
        If text.Contains("9704") OrElse text.Contains("公文") OrElse text.Contains("党政") Then Return True
        Return standard.ApplicableDocumentTypes IsNot Nothing AndAlso
               standard.ApplicableDocumentTypes.Contains(DocumentType.OfficialDocument.ToString())
    End Function

    ' ============================================================
    '  内部辅助 —— 微调
    ' ============================================================

    ''' <summary>根据命令语义确定要修改的目标标签列表</summary>
    Private Shared Function GetTargetTags(command As String, mapping As SemanticStyleMapping) As List(Of SemanticTag)
        Dim tags As New List(Of SemanticTag)()
        If mapping Is Nothing Then Return tags

        Dim hasHeading = ContainsAny(command, {"标题", "题目", "章", "节", "заголовок", "заголовки", "название", "глава", "раздел"})
        Dim hasBody = ContainsAny(command, {"正文", "内容", "段落", "文字", "текст", "содержание", "абзац", "основной текст"})
        Dim hasAll = ContainsAny(command, {"全部", "所有", "整体", "全局", "все", "весь", "всё", "целиком", "полностью"})

        If hasAll OrElse (Not hasHeading AndAlso Not hasBody) Then
            tags.AddRange(mapping.SemanticTags)
        Else
            If hasHeading Then
                tags.AddRange(mapping.SemanticTags.Where(Function(t) t.TagId.StartsWith("heading") OrElse
                                                                     t.TagId.StartsWith("title")))
            End If
            If hasBody Then
                tags.AddRange(mapping.SemanticTags.Where(Function(t) t.TagId.StartsWith("body")))
            End If
        End If

        ' 无命中时退回到全部标签
        If tags.Count = 0 Then
            tags.AddRange(mapping.SemanticTags)
        End If

        Return tags
    End Function

    ''' <summary>对目标标签应用实际的格式数值修改</summary>
    Private Shared Sub ApplyTweakToTags(command As String, targetTags As List(Of SemanticTag), context As RefinementContext)
        If String.IsNullOrWhiteSpace(command) OrElse targetTags Is Nothing Then Return

        ' 判断是否为重复操作（"再大一点" → 沿用上一次操作）
        Dim isRepeat = ContainsAny(command, {"再", "更", "还", "继续", "进一步", "ещё", "еще", "больше", "продолжай", "дальше"})

        For Each tag In targetTags
            ' ---- 字号 ----
            Dim increaseMatch = Regex.Match(command, "(大|增大|加大|放大|调大)\s*([一二两三四五六七八九十\d]+)?\s*(号|磅|pt|点|一些|一点|字号|字体)?")
            If increaseMatch.Success OrElse
               (isRepeat AndAlso Regex.IsMatch(command, "大.{0,3}(点|一些)")) Then
                tag.Font.FontSize = Math.Min(tag.Font.FontSize + ParseFontDeltaAmount(increaseMatch.Groups(2).Value), 72)
            End If

            Dim decreaseMatch = Regex.Match(command, "(小|减小|缩小|调小)\s*([一二两三四五六七八九十\d]+)?\s*(号|磅|pt|点|一些|一点|字号|字体)?")
            If decreaseMatch.Success OrElse
               (isRepeat AndAlso Regex.IsMatch(command, "小.{0,3}(点|一些)")) Then
                tag.Font.FontSize = Math.Max(tag.Font.FontSize - ParseFontDeltaAmount(decreaseMatch.Groups(2).Value), 8)
            End If

            ' 直接指定字号
            Dim sizeMatch = Regex.Match(command, "(设为|设置为|改为|改成|统一为).{0,8}(\d+)\s*(pt|磅|点)")
            If sizeMatch.Success Then
                tag.Font.FontSize = Double.Parse(sizeMatch.Groups(2).Value)
            End If

            ' ---- 颜色 ----
            If ContainsAny(command, {"红色", "красный"}) Then
                tag.Color.FontColor = "#C00000"
            ElseIf ContainsAny(command, {"蓝色", "синий", "голубой"}) Then
                tag.Color.FontColor = "#2E5090"
            ElseIf ContainsAny(command, {"黑色", "чёрный", "черный"}) Then
                tag.Color.FontColor = "#000000"
            ElseIf ContainsAny(command, {"绿色", "зелёный", "зеленый"}) Then
                tag.Color.FontColor = "#008000"
            End If

            ' ---- 加粗 ----
            If ContainsAny(command, {"加粗", "粗体", "粗一点", "жирный", "полужирный", "жирнее", "жирным"}) Then
                tag.Font.Bold = True
            End If
            If ContainsAny(command, {"取消加粗", "不加粗", "细体", "нежирный", "снять жирный", "убрать жирный", "обычный шрифт"}) Then
                tag.Font.Bold = False
            End If

            ' ---- 对齐 ----
            If ContainsAny(command, {"居中", "居中对齐", "по центру", "выровнять по центру"}) Then
                tag.Paragraph.Alignment = "center"
            ElseIf ContainsAny(command, {"左对齐", "靠左", "по левому краю", "по левому"}) Then
                tag.Paragraph.Alignment = "left"
            ElseIf ContainsAny(command, {"右对齐", "靠右", "по правому краю", "по правому"}) Then
                tag.Paragraph.Alignment = "right"
            ElseIf ContainsAny(command, {"两端对齐", "по ширине", "выровнять по ширине", "выровнять"}) Then
                tag.Paragraph.Alignment = "justify"
            End If

            ' ---- 行距 ----
            Dim lineSpacingMatch = Regex.Match(command, "行距.{0,4}([\d.]+)")
            If lineSpacingMatch.Success Then
                Dim newSpacing As Double
                If Double.TryParse(lineSpacingMatch.Groups(1).Value, newSpacing) Then
                    tag.Paragraph.LineSpacing = Math.Max(0.5, Math.Min(newSpacing, 3.0))
                End If
            End If

            ' ---- 缩进 ----
            If ContainsAny(command, {"缩进", "отступ", "отступ первой строки", "красная строка"}) Then
                Dim indentMatch = Regex.Match(command, "缩进.{0,4}([\d.]+)")
                If indentMatch.Success Then
                    Dim newIndent As Double
                    If Double.TryParse(indentMatch.Groups(1).Value, newIndent) Then
                        tag.Paragraph.FirstLineIndent = newIndent
                    End If
                Else
                    tag.Paragraph.FirstLineIndent = 2
                End If
            End If

            ' ---- 字体 ----
            If command.Contains("宋体") Then
                tag.Font.FontNameCN = "宋体"
            ElseIf command.Contains("仿宋") Then
                tag.Font.FontNameCN = "仿宋_GB2312"
            ElseIf command.Contains("黑体") Then
                tag.Font.FontNameCN = "黑体"
            ElseIf command.Contains("楷体") Then
                tag.Font.FontNameCN = "楷体_GB2312"
            ElseIf command.Contains("微软雅黑") Then
                tag.Font.FontNameCN = "微软雅黑"
            End If
        Next
    End Sub

    ' ============================================================
    '  内部辅助 —— 格式化与工具
    ' ============================================================

    ''' <summary>根据语义标签构建格式变更的文字描述</summary>
    Private Shared Function BuildFormatDescription(tag As SemanticTag) As String
        Dim parts As New List(Of String)()

        If tag.Font IsNot Nothing Then
            parts.Add(FormatFontDescription(tag.Font))
        End If

        If tag.Paragraph IsNot Nothing Then
            Dim alignName = GetAlignmentDisplayName(tag.Paragraph.Alignment)
            If Not String.IsNullOrEmpty(alignName) Then
                parts.Add(alignName)
            End If
            If tag.Paragraph.FirstLineIndent > 0 Then
                parts.Add($"отступ первой строки {tag.Paragraph.FirstLineIndent} симв.")
            End If
            If tag.Paragraph.LineSpacing > 0 AndAlso
               Math.Abs(tag.Paragraph.LineSpacing - 1.5) > 0.01 Then
                If IsOfficialDocumentLineSpacing(tag) Then
                    parts.Add("фиксированный ~28 пт")
                Else
                    parts.Add($"межстрочный {tag.Paragraph.LineSpacing}")
                End If
            End If
        End If

        Return String.Join(", ", parts)
    End Function

    Private Shared Function IsOfficialDocumentLineSpacing(tag As SemanticTag) As Boolean
        If tag Is Nothing OrElse tag.Font Is Nothing OrElse tag.Paragraph Is Nothing Then Return False
        If Math.Abs(tag.Font.FontSize - 16) > 0.01 Then Return False
        If Math.Abs(tag.Paragraph.LineSpacing - 1.75) > 0.01 Then Return False
        Return String.Equals(tag.Font.FontNameCN, "仿宋_GB2312", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(tag.Font.FontNameCN, "黑体", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(tag.Font.FontNameCN, "楷体_GB2312", StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>格式化字体描述字符串</summary>
    Private Shared Function FormatFontDescription(font As FontConfig) As String
        If font Is Nothing Then Return ""
        Dim parts As New List(Of String)()

        ' В русскоязычном контуре показываем основной (латинский) шрифт, а не восточноазиатский.
        If Not String.IsNullOrEmpty(font.FontNameEN) Then
            parts.Add(font.FontNameEN)
        ElseIf Not String.IsNullOrEmpty(font.FontNameCN) Then
            parts.Add(font.FontNameCN)
        End If
        If font.FontSize > 0 Then
            parts.Add($"{font.FontSize}pt")
        End If
        If font.Bold Then
            parts.Add("полужирный")
        End If
        If font.Italic Then
            parts.Add("курсив")
        End If

        Return String.Join(" ", parts)
    End Function

    ''' <summary>Отображаемое название выравнивания</summary>
    Private Shared Function GetAlignmentDisplayName(alignment As String) As String
        If String.IsNullOrEmpty(alignment) Then Return ""
        Select Case alignment.ToLower()
            Case "center" : Return "по центру"
            Case "right" : Return "по правому краю"
            Case "justify" : Return "по ширине"
            Case Else : Return "по левому краю"
        End Select
    End Function

    ''' <summary>截断文本并在末尾加省略号</summary>
    Private Shared Function TruncateText(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "…"
    End Function

    ''' <summary>检查文本是否包含任一关键词（忽略大小写）</summary>
    Private Shared Function ContainsAny(text As String, keywords As String()) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each kw In keywords
            If text.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Shared Function IsHeadingNumberingRequest(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        Dim hasStructureVerb = ContainsAny(message, {"重构", "整理", "规范", "优化", "调整", "梳理", "统一", "перестрой", "переструктурировать", "упорядочить", "нормализовать", "оптимизировать"})
        Dim hasStructureTarget = ContainsAny(message, {"序号", "编号", "标题", "层级", "章节", "нумерация", "нумерацию", "заголовки", "уровни", "разделы", "номера"})

        Return hasStructureVerb AndAlso hasStructureTarget
    End Function

    Private Shared Function ParseFontDeltaAmount(value As String) As Double
        If String.IsNullOrWhiteSpace(value) Then Return 1

        Dim numeric As Double
        If Double.TryParse(value, numeric) Then Return Math.Max(0.5, numeric)

        Select Case value.Trim()
            Case "一" : Return 1
            Case "二", "两" : Return 2
            Case "三" : Return 3
            Case "四" : Return 4
            Case "五" : Return 5
            Case "六" : Return 6
            Case "七" : Return 7
            Case "八" : Return 8
            Case "九" : Return 9
            Case "十" : Return 10
            Case Else : Return 1
        End Select
    End Function

    ''' <summary>判断消息是否表达微调意图</summary>
    Private Shared Function HasTweakIntent(message As String) As Boolean
        Return Regex.IsMatch(message, "(再|更|调|改|设|换).{0,4}(大|小|颜色|字体|行距|对齐|加粗|缩进)") OrElse
               Regex.IsMatch(message, "(大|小|加大|增大|减小|缩小|颜色|字体|字号|行距|对齐|加粗|缩进).{0,6}(号|磅|pt|点|一些|一点)") OrElse
               ContainsAny(message, {"红色", "蓝色", "黑色", "居中", "加粗", "красный", "синий", "чёрный", "черный", "зелёный", "зеленый", "белый", "по центру", "по левому краю", "по правому краю", "по ширине", "жирный", "полужирный", "отступ"}) OrElse
               Regex.IsMatch(message, "(больше|меньше|жирнее|крупнее|мельче).{0,4}(шрифт|текст|заголовок|абзац)")
    End Function

End Class
