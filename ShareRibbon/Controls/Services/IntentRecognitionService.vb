' ShareRibbon\Controls\Services\IntentRecognitionService.vb
' 意图识别服务：分析用户输入并识别操作意图

Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 通用Office操作意图类型枚举 - 支持Excel/Word/PowerPoint
''' </summary>
Public Enum OfficeIntentType
    ' === 通用意图 ===
    GENERAL_QUERY       ' 一般查询/问答
    FORMAT_STYLE        ' 格式样式调整

    ' === Excel特有意图 ===
    DATA_ANALYSIS       ' 数据分析（统计、汇总、透视表）
    FORMULA_CALC        ' 公式计算
    CHART_GEN           ' 图表生成
    DATA_CLEANING       ' 数据清洗（去重、填充、格式化）
    REPORT_GEN          ' 报表生成
    DATA_TRANSFORMATION ' 数据转换（合并、拆分、转置）

    ' === Word特有意图 ===
    DOCUMENT_EDIT       ' 文档编辑（插入、删除、替换）
    TEXT_FORMAT         ' 文本格式化（字体、段落、样式）
    TABLE_OPERATION     ' 表格操作
    IMAGE_INSERT        ' 图片插入
    TOC_GENERATION      ' 目录生成
    REVIEW_COMMENT      ' 审阅批注
    PROOFREAD           ' 校对（语法、拼写、表达检查）

    ' === PowerPoint特有意图 ===
    SLIDE_CREATE        ' 创建幻灯片
    SLIDE_LAYOUT        ' 幻灯片布局
    ANIMATION_EFFECT    ' 动画效果
    TRANSITION_EFFECT   ' 切换效果
    TEMPLATE_APPLY      ' 应用模板
    SPEAKER_NOTES       ' 演讲者备注
End Enum

''' <summary>
''' Excel操作意图类型枚举（兼容旧代码）
''' </summary>
Public Enum ExcelIntentType
    DATA_ANALYSIS       ' 数据分析（统计、汇总、透视表）
    FORMULA_CALC        ' 公式计算
    CHART_GEN           ' 图表生成
    DATA_CLEANING       ' 数据清洗（去重、填充、格式化）
    REPORT_GEN          ' 报表生成
    DATA_TRANSFORMATION ' 数据转换（合并、拆分、转置）
    FORMAT_STYLE        ' 格式样式调整
    GENERAL_QUERY       ' 一般查询
End Enum

''' <summary>
''' 意图识别结果
''' </summary>
Public Class IntentResult
    ''' <summary>
    ''' 通用Office意图类型（支持Excel/Word/PowerPoint）
    ''' </summary>
    Public Property OfficeIntent As OfficeIntentType = OfficeIntentType.GENERAL_QUERY

    ''' <summary>
    ''' 主要意图类型（兼容旧代码，映射到ExcelIntentType）
    ''' </summary>
    Public Property IntentType As ExcelIntentType = ExcelIntentType.GENERAL_QUERY

    ''' <summary>
    ''' 次要意图（可能有多个操作）
    ''' </summary>
    Public Property SecondaryIntents As List(Of ExcelIntentType) = New List(Of ExcelIntentType)()

    ''' <summary>
    ''' 意图置信度 (0-1)
    ''' </summary>
    Public Property Confidence As Double = 0.5

    ''' <summary>
    ''' 响应模式
    ''' </summary>
    Public Property ResponseMode As String = ""

    ''' <summary>
    ''' 是否需要VBA代码
    ''' </summary>
    Public Property RequiresVBA As Boolean = True

    ''' <summary>
    ''' 是否可以使用直接操作命令
    ''' </summary>
    Public Property CanUseDirectCommand As Boolean = False

    ''' <summary>
    ''' 提取的关键实体（如范围、列名等）
    ''' </summary>
    Public Property ExtractedEntities As Dictionary(Of String, String) = New Dictionary(Of String, String)()

    ''' <summary>
    ''' 用户友好的意图描述
    ''' </summary>
    Public Property UserFriendlyDescription As String = ""

    ''' <summary>
    ''' 执行计划步骤列表
    ''' </summary>
    Public Property ExecutionPlan As List(Of ExecutionStep) = New List(Of ExecutionStep)()

    ''' <summary>
    ''' 原始用户输入
    ''' </summary>
    Public Property OriginalInput As String = ""

    ''' <summary>
    ''' 是否为追问/延续性问题（由上下文话题连续性分析设置，避免每次调用 LLM）
    ''' </summary>
    Public Property IsFollowUp As Boolean = False

    Public Property RequestedOutputs As List(Of String) = New List(Of String)()
End Class

''' <summary>
''' 意图识别服务 - 支持Excel/Word/PowerPoint
''' </summary>
Public Class IntentRecognitionService

    ''' <summary>
    ''' 当前Office应用类型
    ''' </summary>
    Public Property AppType As OfficeApplicationType = OfficeApplicationType.Excel

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New()
        Me.AppType = OfficeApplicationType.Excel
    End Sub

    ''' <summary>
    ''' 带应用类型的构造函数
    ''' </summary>
    Public Sub New(appType As OfficeApplicationType)
        Me.AppType = appType
    End Sub

    ''' <summary>
    ''' 带应用类型名称的构造函数，避免 VB 在 Option Strict Off 下把 "Word" 强转为枚举整数。
    ''' </summary>
    Public Sub New(appType As String)
        Me.AppType = ParseOfficeApplicationType(appType)
    End Sub

    Private Shared Function ParseOfficeApplicationType(appType As String) As OfficeApplicationType
        Select Case If(appType, "").Trim().ToLowerInvariant()
            Case "word"
                Return OfficeApplicationType.Word
            Case "powerpoint", "ppt"
                Return OfficeApplicationType.PowerPoint
            Case Else
                Return OfficeApplicationType.Excel
        End Select
    End Function

#Region "Excel关键词映射"

    ' 数据分析关键词
    Private Shared ReadOnly DataAnalysisKeywords As String() = {
        "统计", "分析", "汇总", "求和", "平均", "最大", "最小", "计数",
        "透视表", "数据透视", "分组", "聚合", "占比", "百分比", "增长率",
        "趋势", "对比", "排名", "top", "前几", "后几"
    }

    ' 公式计算关键词
    Private Shared ReadOnly FormulaCalcKeywords As String() = {
        "公式", "计算", "求", "加", "减", "乘", "除", "sum", "average",
        "vlookup", "if", "countif", "sumif", "index", "match"
    }

    ' 图表生成关键词
    Private Shared ReadOnly ChartGenKeywords As String() = {
        "图表", "柱状图", "折线图", "饼图", "条形图", "散点图", "面积图",
        "chart", "graph", "可视化", "画图", "生成图", "做个图"
    }

    ' 数据清洗关键词
    Private Shared ReadOnly DataCleaningKeywords As String() = {
        "清洗", "去重", "删除重复", "填充", "空值", "缺失", "替换",
        "格式化", "规范", "trim", "清理", "整理", "修复"
    }

    ' 报表生成关键词
    Private Shared ReadOnly ReportGenKeywords As String() = {
        "报表", "报告", "表格", "生成表", "导出", "输出", "创建表",
        "周报", "月报", "日报", "汇报", "模板"
    }

    ' 数据转换关键词
    Private Shared ReadOnly DataTransformKeywords As String() = {
        "合并", "拆分", "转置", "行列转换", "连接", "vlookup", "关联",
        "join", "merge", "split", "transpose", "提取", "截取"
    }

    ' 格式样式关键词（通用）
    Private Shared ReadOnly FormatStyleKeywords As String() = {
        "格式", "样式", "颜色", "字体", "边框", "对齐", "加粗",
        "斜体", "底色", "高亮", "条件格式", "美化"
    }

#End Region

#Region "Word关键词映射"

    ' 文档编辑关键词
    Private Shared ReadOnly DocumentEditKeywords As String() = {
        "插入", "删除", "替换", "复制", "粘贴", "剪切", "撤销",
        "查找", "搜索", "定位", "跳转", "选中", "全选"
    }

    ' 文本格式关键词
    Private Shared ReadOnly TextFormatKeywords As String() = {
        "字体", "字号", "加粗", "斜体", "下划线", "删除线", "上标", "下标",
        "段落", "行距", "缩进", "首行缩进", "对齐", "两端对齐"
    }

    ' 表格操作关键词
    Private Shared ReadOnly TableOperationKeywords As String() = {
        "表格", "插入表格", "删除表格", "合并单元格", "拆分单元格",
        "添加行", "添加列", "删除行", "删除列", "表格样式"
    }

    ' 图片插入关键词
    Private Shared ReadOnly ImageInsertKeywords As String() = {
        "图片", "插入图片", "添加图片", "图像", "截图", "照片",
        "调整图片", "裁剪", "旋转图片"
    }

    ' 目录生成关键词
    Private Shared ReadOnly TocGenerationKeywords As String() = {
        "目录", "生成目录", "更新目录", "插入目录", "大纲",
        "标题样式", "章节", "页码"
    }

    ' 审阅批注关键词
    Private Shared ReadOnly ReviewCommentKeywords As String() = {
        "批注", "注释", "评论", "修订", "审阅", "接受修订", "拒绝修订",
        "比较文档", "合并文档"
    }

    ' 校对关键词
    Private Shared ReadOnly ProofreadKeywords As String() = {
        "校对", "检查语法", "拼写检查", "语法错误", "拼写错误", "错别字",
        "语病", "用词错误", "标点错误", "表达问题", "检查拼写", "语法检查",
        "校验文本", "文字校对", "审校", "润色检查", "proofread"
    }

#End Region

#Region "PowerPoint关键词映射"

    ' 幻灯片创建关键词
    Private Shared ReadOnly SlideCreateKeywords As String() = {
        "幻灯片", "新建幻灯片", "添加幻灯片", "插入幻灯片",
        "删除幻灯片", "复制幻灯片", "ppt", "演示文稿"
    }

    ' 幻灯片布局关键词
    Private Shared ReadOnly SlideLayoutKeywords As String() = {
        "布局", "版式", "标题幻灯片", "标题和内容", "空白幻灯片",
        "两栏内容", "比较", "仅标题"
    }

    ' 动画效果关键词
    Private Shared ReadOnly AnimationEffectKeywords As String() = {
        "动画", "进入动画", "退出动画", "强调动画", "路径动画",
        "淡入", "飞入", "缩放", "旋转动画", "动画顺序"
    }

    ' 切换效果关键词
    Private Shared ReadOnly TransitionEffectKeywords As String() = {
        "切换", "过渡", "幻灯片切换", "淡出", "推入", "擦除",
        "百叶窗", "棋盘", "切换时间"
    }

    ' 模板应用关键词
    Private Shared ReadOnly TemplateApplyKeywords As String() = {
        "模板", "主题", "应用模板", "设计", "配色方案",
        "背景", "更换背景", "幻灯片母版"
    }

    ' 演讲者备注关键词
    Private Shared ReadOnly SpeakerNotesKeywords As String() = {
        "备注", "演讲者备注", "笔记", "提示", "演讲稿",
        "演示者视图", "备注页"
    }

#End Region

#Region "Русские ключевые слова (P0 русификация)"

    Private Shared ReadOnly DataAnalysisKeywordsRu As String() = {
        "статистик", "аналитик", "анализ", "сводк", "сводн", "сумм", "итог",
        "средн", "максимум", "минимум", "количеств", "группиров", "агрегир",
        "доля", "процент", "тренд", "сравнен", "рейтинг", "топ-", "первые", "последние"
    }

    Private Shared ReadOnly FormulaCalcKeywordsRu As String() = {
        "формул", "вычисл", "рассчит", "расчёт", "расчет", "сложить", "сложи",
        "вычесть", "умнож", "разделить", "поделить", "сумма", "среднее значение",
        "впр", "счёт", "счет", "посчитай"
    }

    Private Shared ReadOnly ChartGenKeywordsRu As String() = {
        "диаграмм", "график", "круговая", "круговую", "столбчатая", "столбчат",
        "линейная", "линейный", "точечная", "гистограмм", "визуализ",
        "построй", "нарисуй", "графики"
    }

    Private Shared ReadOnly DataCleaningKeywordsRu As String() = {
        "очист", "очисть", "дубликат", "дубл", "повторяющ", "заполн",
        "пустые значения", "пропуск", "замен", "унифиц", "унифицир",
        "обрезать", "пробел", "удали пустые"
    }

    Private Shared ReadOnly ReportGenKeywordsRu As String() = {
        "отчёт", "отчет", "доклад", "сводк", "экспорт", "выгруз",
        "недельный", "месячный", "дневной", "шаблон", "сформируй отч"
    }

    Private Shared ReadOnly DataTransformKeywordsRu As String() = {
        "объедин", "объедини", "разъедин", "разделить", "транспонир",
        "транспон", "преобраз", "соедини", "сопостав", "извлеч", "разбить"
    }

    Private Shared ReadOnly FormatStyleKeywordsRu As String() = {
        "формат", "стиль", "цвет", "шрифт", "границ", "выравн", "полужирн",
        "жирн", "курсив", "заливк", "подсвет", "условное форм", "оформ", "красив"
    }

    Private Shared ReadOnly DocumentEditKeywordsRu As String() = {
        "вставить", "вставь", "удалить", "удали", "заменить", "замени",
        "скопир", "вырезать", "отмен", "найти", "поиск", "перейти",
        "выдели", "выделить"
    }

    Private Shared ReadOnly TextFormatKeywordsRu As String() = {
        "шрифт", "кегл", "размер шрифта", "жирн", "полужирн", "курсив",
        "подчёрк", "подчерк", "зачёрк", "зачерк", "надстроч", "подстроч",
        "абзац", "интервал", "отступ", "выравн"
    }

    Private Shared ReadOnly TableOperationKeywordsRu As String() = {
        "таблиц", "ячейк", "строк", "столбц", "объедин", "разъедин",
        "добавь строк", "добавь столбц", "удали строк", "удали столбц", "стиль таблиц"
    }

    Private Shared ReadOnly ImageInsertKeywordsRu As String() = {
        "изображен", "картинк", "рисунок", "фото", "скриншот", "снимок",
        "вставь картин", "обрез", "поверн"
    }

    Private Shared ReadOnly TocGenerationKeywordsRu As String() = {
        "оглавлен", "содержан", "навигац", "обнови оглав", "вставь оглав",
        "заголовк", "разделы", "нумерац"
    }

    Private Shared ReadOnly ReviewCommentKeywordsRu As String() = {
        "примечан", "коммент", "реценз", "отслежив", "принять правк",
        "отклонить правк", "режим правок", "сравнить документ", "объединить документ"
    }

    Private Shared ReadOnly ProofreadKeywordsRu As String() = {
        "вычит", "вычита", "проверь текст", "проверить текст", "орфограф",
        "пунктуац", "грамматик", "опечат", "ошибк", "стилист", "редактир",
        "правописан", "запят"
    }

    Private Shared ReadOnly SlideCreateKeywordsRu As String() = {
        "слайд", "презентац", "добавь слайд", "новый слайд", "удали слайд",
        "скопируй слайд", "дублируй слайд", "ppt", "ппт"
    }

    Private Shared ReadOnly SlideLayoutKeywordsRu As String() = {
        "макет", "разметк", "титульн", "заголовок и содержан", "пустой слайд",
        "две колонк", "сравнен", "только заголовок"
    }

    Private Shared ReadOnly AnimationEffectKeywordsRu As String() = {
        "анимац", "анимир", "появлен", "исчезн", "постепен", "вращен",
        "траектор", "порядок анимац"
    }

    Private Shared ReadOnly TransitionEffectKeywordsRu As String() = {
        "переход", "смену слайд", "затухан", "сдвиг", "растворен", "жалюзи",
        "шахмат", "время переход"
    }

    Private Shared ReadOnly TemplateApplyKeywordsRu As String() = {
        "шаблон", "дизайн", "палитр", "цветов", "фонов", "мастер слайд",
        "тема оформл", "тему оформл", "стиль презентац"
    }

    Private Shared ReadOnly SpeakerNotesKeywordsRu As String() = {
        "заметк", "заметки докладч", "подсказк", "текст выступлен",
        "режим докладч", "заметках"
    }

    ''' <summary>
    ''' Скоринг с учётом русского и китайского словарей: нормализация считается
    ''' отдельно для каждого языка, итог — максимум. Китайское поведение не меняется.
    ''' </summary>
    Private Function ScoreKeywords(text As String, chineseKeywords As String(), russianKeywords As String()) As Double
        Return Math.Max(CalculateKeywordScore(text, chineseKeywords), CalculateKeywordScore(text, russianKeywords))
    End Function

#End Region

#Region "公共方法"

    ''' <summary>
    ''' 识别用户意图
    ''' </summary>
    ''' <param name="question">用户问题</param>
    ''' <param name="context">上下文信息（可选）</param>
    ''' <returns>意图识别结果</returns>
    Public Function IdentifyIntent(question As String, Optional context As JObject = Nothing) As IntentResult
        Dim result As New IntentResult()

        If String.IsNullOrWhiteSpace(question) Then
            Return result
        End If

        Dim lowerQuestion = question.ToLower()

        ' 根据AppType使用不同的关键词映射计算意图分数
        Select Case AppType
            Case OfficeApplicationType.Excel
                IdentifyExcelIntent(lowerQuestion, result)
            Case OfficeApplicationType.Word
                IdentifyWordIntent(lowerQuestion, result)
            Case OfficeApplicationType.PowerPoint
                IdentifyPowerPointIntent(lowerQuestion, result)
        End Select

        ' 通用格式样式意图检测
        Dim formatScore = ScoreKeywords(lowerQuestion, FormatStyleKeywords, FormatStyleKeywordsRu)
        If formatScore > result.Confidence Then
            result.IntentType = ExcelIntentType.FORMAT_STYLE
            result.OfficeIntent = OfficeIntentType.FORMAT_STYLE
            result.Confidence = Math.Min(formatScore, 1.0)
        End If

        ' 提取关键实体
        ExtractEntities(question, result)

        ' 判断是否可以使用直接命令
        DetermineExecutionMethod(result)

        Debug.WriteLine($"[{AppType}] 意图识别结果: {result.OfficeIntent}, 置信度: {result.Confidence:F2}")
        Return result
    End Function

    ''' <summary>
    ''' 识别Excel特有意图
    ''' </summary>
    Private Sub IdentifyExcelIntent(lowerQuestion As String, result As IntentResult)
        Dim scores As New Dictionary(Of OfficeIntentType, Double)()
        scores(OfficeIntentType.DATA_ANALYSIS) = ScoreKeywords(lowerQuestion, DataAnalysisKeywords, DataAnalysisKeywordsRu)
        scores(OfficeIntentType.FORMULA_CALC) = ScoreKeywords(lowerQuestion, FormulaCalcKeywords, FormulaCalcKeywordsRu)
        scores(OfficeIntentType.CHART_GEN) = ScoreKeywords(lowerQuestion, ChartGenKeywords, ChartGenKeywordsRu)
        scores(OfficeIntentType.DATA_CLEANING) = ScoreKeywords(lowerQuestion, DataCleaningKeywords, DataCleaningKeywordsRu)
        scores(OfficeIntentType.REPORT_GEN) = ScoreKeywords(lowerQuestion, ReportGenKeywords, ReportGenKeywordsRu)
        scores(OfficeIntentType.DATA_TRANSFORMATION) = ScoreKeywords(lowerQuestion, DataTransformKeywords, DataTransformKeywordsRu)

        Dim maxScore As Double = 0
        Dim maxIntent = OfficeIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        If maxScore > 0.1 Then
            result.OfficeIntent = maxIntent
            result.IntentType = MapToExcelIntentType(maxIntent)
            result.Confidence = Math.Min(maxScore, 1.0)
        End If
    End Sub

    ''' <summary>
    ''' 识别Word特有意图
    ''' </summary>
    Private Sub IdentifyWordIntent(lowerQuestion As String, result As IntentResult)
        Dim scores As New Dictionary(Of OfficeIntentType, Double)()
        scores(OfficeIntentType.DOCUMENT_EDIT) = ScoreKeywords(lowerQuestion, DocumentEditKeywords, DocumentEditKeywordsRu)
        scores(OfficeIntentType.TEXT_FORMAT) = ScoreKeywords(lowerQuestion, TextFormatKeywords, TextFormatKeywordsRu)
        scores(OfficeIntentType.TABLE_OPERATION) = ScoreKeywords(lowerQuestion, TableOperationKeywords, TableOperationKeywordsRu)
        scores(OfficeIntentType.IMAGE_INSERT) = ScoreKeywords(lowerQuestion, ImageInsertKeywords, ImageInsertKeywordsRu)
        scores(OfficeIntentType.TOC_GENERATION) = ScoreKeywords(lowerQuestion, TocGenerationKeywords, TocGenerationKeywordsRu)
        scores(OfficeIntentType.REVIEW_COMMENT) = ScoreKeywords(lowerQuestion, ReviewCommentKeywords, ReviewCommentKeywordsRu)
        scores(OfficeIntentType.PROOFREAD) = ScoreKeywords(lowerQuestion, ProofreadKeywords, ProofreadKeywordsRu)

        Dim maxScore As Double = 0
        Dim maxIntent = OfficeIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        If maxScore > 0.1 Then
            result.OfficeIntent = maxIntent
            result.IntentType = ExcelIntentType.GENERAL_QUERY ' Word意图映射到通用查询
            result.Confidence = Math.Min(maxScore, 1.0)
        End If
    End Sub

    ''' <summary>
    ''' 识别PowerPoint特有意图
    ''' </summary>
    Private Sub IdentifyPowerPointIntent(lowerQuestion As String, result As IntentResult)
        Dim scores As New Dictionary(Of OfficeIntentType, Double)()
        scores(OfficeIntentType.SLIDE_CREATE) = ScoreKeywords(lowerQuestion, SlideCreateKeywords, SlideCreateKeywordsRu)
        scores(OfficeIntentType.SLIDE_LAYOUT) = ScoreKeywords(lowerQuestion, SlideLayoutKeywords, SlideLayoutKeywordsRu)
        scores(OfficeIntentType.ANIMATION_EFFECT) = ScoreKeywords(lowerQuestion, AnimationEffectKeywords, AnimationEffectKeywordsRu)
        scores(OfficeIntentType.TRANSITION_EFFECT) = ScoreKeywords(lowerQuestion, TransitionEffectKeywords, TransitionEffectKeywordsRu)
        scores(OfficeIntentType.TEMPLATE_APPLY) = ScoreKeywords(lowerQuestion, TemplateApplyKeywords, TemplateApplyKeywordsRu)
        scores(OfficeIntentType.SPEAKER_NOTES) = ScoreKeywords(lowerQuestion, SpeakerNotesKeywords, SpeakerNotesKeywordsRu)

        Dim maxScore As Double = 0
        Dim maxIntent = OfficeIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        ' This is only the deterministic fallback when the LLM response is malformed.
        ' One exact PowerPoint capability term must remain actionable instead of being
        ' diluted by the total number of keywords in the category.
        If maxScore > 0 Then
            result.OfficeIntent = maxIntent
            result.IntentType = ExcelIntentType.GENERAL_QUERY ' PPT意图映射到通用查询
            result.Confidence = Math.Min(Math.Max(maxScore, 0.55), 1.0)
        End If
    End Sub

    ''' <summary>
    ''' 将通用意图映射到Excel意图类型（兼容旧代码）
    ''' </summary>
    Private Function MapToExcelIntentType(intent As OfficeIntentType) As ExcelIntentType
        Select Case intent
            Case OfficeIntentType.DATA_ANALYSIS
                Return ExcelIntentType.DATA_ANALYSIS
            Case OfficeIntentType.FORMULA_CALC
                Return ExcelIntentType.FORMULA_CALC
            Case OfficeIntentType.CHART_GEN
                Return ExcelIntentType.CHART_GEN
            Case OfficeIntentType.DATA_CLEANING
                Return ExcelIntentType.DATA_CLEANING
            Case OfficeIntentType.REPORT_GEN
                Return ExcelIntentType.REPORT_GEN
            Case OfficeIntentType.DATA_TRANSFORMATION
                Return ExcelIntentType.DATA_TRANSFORMATION
            Case OfficeIntentType.FORMAT_STYLE
                Return ExcelIntentType.FORMAT_STYLE
            Case Else
                Return ExcelIntentType.GENERAL_QUERY
        End Select
    End Function

    ''' <summary>
    ''' 异步识别意图（始终使用LLM进行置信度评分）
    ''' </summary>
    Public Async Function IdentifyIntentAsync(question As String, Optional context As JObject = Nothing) As Task(Of IntentResult)
        ' 首先使用关键词匹配进行初步分类（但不使用其置信度）
        Dim result = IdentifyIntent(question, context)

        ' 始终调用LLM进行置信度评分（用户要求置信度由大模型打分）
        If Not String.IsNullOrWhiteSpace(question) Then
            Try
                Dim llmResult = Await IdentifyIntentWithLLMAsync(question, context)
                If llmResult IsNot Nothing Then
                    ' 使用LLM的置信度（这是核心改动）
                    result.Confidence = llmResult.Confidence

                    ' 如果LLM的意图类型判断更可信，也使用LLM的意图
                    If llmResult.Confidence > 0.3 Then
                        ' OfficeIntent 才是跨宿主路由的权威字段。IntentType 仅保留给旧版 Excel
                        ' 调用方兼容；只覆盖 IntentType 会把 Word/PPT 的 LLM 结果悄悄退回
                        ' GENERAL_QUERY。
                        result.OfficeIntent = llmResult.OfficeIntent
                        result.IntentType = llmResult.IntentType
                        result.UserFriendlyDescription = llmResult.UserFriendlyDescription
                    End If

                    Debug.WriteLine($"LLM意图识别结果: {result.OfficeIntent}, 置信度: {result.Confidence:F2}")
                End If
            Catch ex As Exception
                Debug.WriteLine($"LLM意图识别失败，使用默认置信度0.5: {ex.Message}")
                ' 如果LLM调用失败，使用默认中等置信度
                result.Confidence = 0.5
            End Try
        End If

        Return result
    End Function

    ''' <summary>
    ''' 调用大模型识别意图 - 增强版，包含记忆上下文
    ''' </summary>
    Private Async Function IdentifyIntentWithLLMAsync(question As String, context As JObject) As Task(Of IntentResult)
        Dim result As IntentResult = Nothing

        Try
            ' 获取API配置
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return result
            End If

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)

            Dim apiUrl = HttpClientFactory.ResolveChatCompletionsUrl(cfg.url)
            Dim apiKey = cfg.key
            Dim modelName = selectedModel.modelName

            ' 构建上下文信息
            Dim contextInfo As String = ""
            If context IsNot Nothing Then
                If context("sheetName") IsNot Nothing Then
                    contextInfo &= $"当前工作表: {context("sheetName")}" & vbCrLf
                End If
                If context("selectionAddress") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context("selectionAddress").ToString()) Then
                    contextInfo &= $"选中区域: {context("selectionAddress")}" & vbCrLf
                End If
                If context("selection") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context("selection").ToString()) Then
                    contextInfo &= $"选中内容预览:" & vbCrLf & context("selection").ToString() & vbCrLf
                End If
                ' 阶段四：内容区引用摘要与 RAG 记忆
                If context("referenceSummary") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context("referenceSummary").ToString()) Then
                    contextInfo &= "用户引用: " & context("referenceSummary").ToString() & vbCrLf
                End If
                If context("ragSnippets") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context("ragSnippets").ToString()) Then
                    contextInfo &= "相关记忆:" & vbCrLf & context("ragSnippets").ToString() & vbCrLf
                End If
            End If

            ' 增强版提示词：构建更智能的意图识别系统提示词
            Dim systemPrompt = GetEnhancedIntentRecognitionSystemPrompt()
            Dim userMessage = $"用户问题: {question}"
            If Not String.IsNullOrEmpty(contextInfo) Then
                userMessage &= vbCrLf & vbCrLf & "当前Office上下文信息:" & vbCrLf & contextInfo
            End If

            ' 构建请求体 - 包含历史对话（作为正确的 role 消息）
            Dim messages As New JArray()
            messages.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})

            ' 将历史对话作为 user/assistant 消息注入，让大模型理解上下文后再做意图识别
            If context IsNot Nothing AndAlso context("conversationHistory") IsNot Nothing Then
                Dim historyArr = TryCast(context("conversationHistory"), JArray)
                If historyArr IsNot Nothing AndAlso historyArr.Count > 0 Then
                    For Each hMsg In historyArr
                        Dim hRole = hMsg("role")?.ToString()
                        Dim hContent = hMsg("content")?.ToString()
                        If Not String.IsNullOrEmpty(hRole) AndAlso Not String.IsNullOrEmpty(hContent) Then
                            messages.Add(New JObject From {{"role", hRole}, {"content", hContent}})
                        End If
                    Next
                    Debug.WriteLine($"[IntentRecognition] 注入 {historyArr.Count} 条历史消息用于意图识别")
                End If
            End If

            messages.Add(New JObject From {{"role", "user"}, {"content", userMessage}})

            Dim gatewayResponse = Await AiGateway.SendChatAsync(New AiRequestOptions With {
                .ApiUrl = apiUrl,
                .ApiKey = apiKey,
                .ModelName = modelName,
                .Platform = cfg.platform,
                .Messages = messages,
                .Temperature = 0.3R,
                .MaxTokens = 500,
                .TimeoutSeconds = 30
            })

            If gatewayResponse IsNot Nothing AndAlso gatewayResponse.Success Then
                result = ParseLLMIntentResponse(gatewayResponse.Content, question)
            Else
                Debug.WriteLine($"IdentifyIntentWithLLMAsync API失败: {If(gatewayResponse Is Nothing, "empty response", gatewayResponse.ErrorMessage)}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"IdentifyIntentWithLLMAsync 出错: {ex.Message}")
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 获取增强版意图识别系统提示词 - 根据AppType返回不同的提示词
    ''' </summary>
    Private Function GetEnhancedIntentRecognitionSystemPrompt() As String
        Select Case AppType
            Case OfficeApplicationType.Word
                Return GetEnhancedWordIntentRecognitionPrompt()
            Case OfficeApplicationType.PowerPoint
                Return GetEnhancedPowerPointIntentRecognitionPrompt()
            Case Else ' Excel
                Return GetEnhancedExcelIntentRecognitionPrompt()
        End Select
    End Function

    ''' <summary>
    ''' 增强版Excel意图识别提示词 - 更智能，支持记忆上下文
    ''' </summary>
    Private Function GetEnhancedExcelIntentRecognitionPrompt() As String
        Return "Ты — эксперт по распознаванию намерений Excel. Глубоко анализируй вопрос пользователя, контекст и релевантную память, точно определяй истинное намерение пользователя.

【Ключевые аспекты анализа】
1. Семантика запроса: понять, что пользователь действительно хочет сделать
2. Связь с контекстом: учитывать выделенный диапазон, лист и релевантную память
3. Сложность операции: оценить, нужны ли несколько шагов
4. Оценка риска: определить, безопасна ли операция и нужно ли подтверждение

【Формат возвращаемого JSON】
```json
{
  ""intentType"": ""DATA_ANALYSIS"",
  ""confidence"": 0.92,
  ""description"": ""Пользователь хочет статистически проанализировать выделенный диапазон: среднее и сумму"",
  ""requiresConfirmation"": false,
  ""suggestedAction"": ""Выполнить статистический расчёт"",
  ""executionPriority"": ""high"",
  ""suggestedSteps"": [
    ""Определить диапазон данных"",
    ""Вычислить среднее"",
    ""Вычислить сумму"",
    ""Вывести результат""
  ]
}
```

【Допустимые значения intentType】
- DATA_ANALYSIS: анализ данных (статистика, сводки, сводные таблицы)
- FORMULA_CALC: вычисления по формулам
- CHART_GEN: построение диаграмм
- DATA_CLEANING: очистка данных (удаление дубликатов, заполнение)
- REPORT_GEN: формирование отчётов
- DATA_TRANSFORMATION: преобразование данных (объединение, разделение)
- FORMAT_STYLE: настройка оформления
- GENERAL_QUERY: обычный вопрос (без операций с Excel)
- UNCLEAR: намерение неясно, требуется уточнение
- MULTI_STEP_TASK: сложная многошаговая задача (нужен Ralph Loop)

【Правила обоснованного выбора】
1. **Оценка уверенности**:
   - 0.9-1.0: очень ясно, можно выполнять сразу
   - 0.7-0.89: достаточно ясно, можно выполнять
   - 0.5-0.69: не совсем ясно, рекомендуется подтверждение
   - <0.5: неясно, нужно уточнить

2. **Использование памяти**:
   - предпочитай прежние привычки и предпочтения пользователя
   - учитывай прежние проблемы пользователя и их решения
   - продолжай предыдущий контекст диалога

3. **Распознавание многошаговых задач**:
   - если запрос сложный и требует нескольких операций, используй MULTI_STEP_TASK
   - перечисли ключевые шаги в suggestedSteps

4. **Безопасность операций**:
   - операции только для чтения (статистика, запросы) requiresConfirmation=false
   - изменяющие операции (удаление, перезапись) requiresConfirmation=true
   - массовое изменение данных requiresConfirmation=true

【Примеры сценариев】
- «посчитай продажи» + выделен диапазон → DATA_ANALYSIS, confidence=0.95
- «сделай прошлую диаграмму ещё раз» + в памяти есть диаграмма → CHART_GEN, confidence=0.88
- «приведи данные в порядок и сделай отчёт» → MULTI_STEP_TASK, confidence=0.90

Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть."
    End Function

    ''' <summary>
    ''' Excel意图识别提示词
    ''' </summary>
    Private Function GetEnhancedWordIntentRecognitionPrompt() As String
        Return "Ты — эксперт по распознаванию намерений Word. Глубоко анализируй вопрос пользователя, контекст и релевантную память, точно определяй истинное намерение пользователя.

【Формат возвращаемого JSON】
```json
{
  ""intentType"": ""DOCUMENT_EDIT"",
  ""confidence"": 0.88,
  ""description"": ""Пользователь хочет вставить текст в позицию курсора"",
  ""requiresConfirmation"": false,
  ""suggestedAction"": ""Вставить текст"",
  ""executionPriority"": ""medium"",
  ""suggestedSteps"": [
    ""Подтвердить позицию вставки"",
    ""Вставить указанный текст""
  ]
}
```

【Допустимые значения intentType】
- DOCUMENT_EDIT: редактирование документа (вставка, удаление, замена текста)
- TEXT_FORMAT: форматирование текста (шрифт, абзац, стиль)
- TABLE_OPERATION: работа с таблицами (создание, редактирование)
- IMAGE_INSERT: вставка и обработка изображений
- TOC_GENERATION: создание и обновление оглавления
- REVIEW_COMMENT: рецензирование и примечания
- PROOFREAD: вычитка (грамматика, орфография, словоупотребление, стиль)
- FORMAT_STYLE: настройка оформления
- GENERAL_QUERY: обычный вопрос (без операций с Word)
- UNCLEAR: намерение неясно, требуется уточнение
- MULTI_STEP_TASK: сложная многошаговая задача

【Ключевые аспекты анализа】
1. Понять реальную потребность с учётом выделенного текста и памяти
2. Определить, безопасна ли операция и нужно ли подтверждение пользователя
3. Распознать необходимость нескольких шагов (например, улучшение оформления документа)

【Обоснованные решения】
- простое приветствие → GENERAL_QUERY, confidence=0.9
- существенное изменение документа → requiresConfirmation=true
- учитывай предпочтения пользователя из памяти

Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть."
    End Function

    ''' <summary>
    ''' Word意图识别提示词
    ''' </summary>
    Private Function GetEnhancedPowerPointIntentRecognitionPrompt() As String
        Return "Ты — эксперт по распознаванию намерений PowerPoint. Глубоко анализируй вопрос пользователя, контекст и релевантную память, точно определяй истинное намерение пользователя.

【Формат возвращаемого JSON】
```json
{
  ""intentType"": ""SLIDE_CREATE"",
  ""confidence"": 0.90,
  ""description"": ""Пользователь хочет создать презентацию из 3 слайдов"",
  ""interactionMode"": ""execute"",
  ""requestedOutputs"": [""slides"", ""images""],
  ""requiresConfirmation"": false,
  ""suggestedAction"": ""Пакетно создать слайды"",
  ""executionPriority"": ""high"",
  ""suggestedSteps"": [
    ""Определить число слайдов"",
    ""Создать 1-й слайд-титул"",
    ""Создать 2-й слайд с содержимым"",
    ""Создать 3-й слайд-итог""
  ]
}
```

【Допустимые значения intentType】
- SLIDE_CREATE: создание слайдов
- SLIDE_LAYOUT: компоновка и макет слайдов
- ANIMATION_EFFECT: анимация
- TRANSITION_EFFECT: переходы
- TEMPLATE_APPLY: применение шаблона и темы
- SPEAKER_NOTES: заметки докладчика
- FORMAT_STYLE: настройка оформления
- GENERAL_QUERY: обычный вопрос (без операций с PPT)
- UNCLEAR: намерение неясно, требуется уточнение
- MULTI_STEP_TASK: сложная многошаговая задача

【Обоснованные решения】
- interactionMode должен принимать одно из значений execute, answer, clarify; он не зависит от intentType.
- requestedOutputs перечисляет артефакты, которые нужно проверить после выполнения, например slides, images, chart, table, formatted_content.
- если пользователь просит изменить текущую презентацию — execute; если спрашивает о причине/использовании — answer; если не хватает блокирующей информации — clarify.
- «сделай презентацию» → MULTI_STEP_TASK, confidence=0.85
- «добавь анимацию» → ANIMATION_EFFECT, confidence=0.92
- учитывай предпочтения по шаблонам из памяти

Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть."
    End Function

    ''' <summary>
    ''' PowerPoint意图识别提示词
    ''' </summary>
    Private Function ParseLLMIntentResponse(responseContent As String, originalQuestion As String) As IntentResult
        Dim result As New IntentResult()
        result.OriginalInput = originalQuestion

        Try
            Dim content = ExtractAssistantContentOrRaw(responseContent)
            If String.IsNullOrEmpty(content) Then Return Nothing

            ' 提取JSON部分
            Dim jsonMatch = Regex.Match(content, "\{[\s\S]*\}")
            If Not jsonMatch.Success Then Return Nothing

            Dim intentJson = JObject.Parse(jsonMatch.Value)

            ' 解析意图类型 - 根据AppType处理不同的意图
            Dim intentTypeStr = intentJson("intentType")?.ToString()?.ToUpper()
            ParseIntentTypeByApp(intentTypeStr, result)

            ' 解析置信度
            If intentJson("confidence") IsNot Nothing Then
                result.Confidence = CDbl(intentJson("confidence"))
            End If

            ' 解析描述
            If intentJson("description") IsNot Nothing Then
                result.UserFriendlyDescription = intentJson("description").ToString()
            End If

            ' 与封闭意图枚举解耦：interactionMode 决定回答、澄清还是执行。
            Dim interactionMode = intentJson("interactionMode")?.ToString()
            If String.IsNullOrWhiteSpace(interactionMode) Then interactionMode = intentJson("responseMode")?.ToString()
            result.ResponseMode = If(interactionMode, "").Trim().ToLowerInvariant()

            Dim requestedOutputs = TryCast(intentJson("requestedOutputs"), JArray)
            If requestedOutputs IsNot Nothing Then
                result.RequestedOutputs = requestedOutputs.Select(Function(item) item.ToString().Trim().ToLowerInvariant()).Where(Function(item) Not String.IsNullOrWhiteSpace(item)).Distinct().ToList()
            End If

            ' 解析是否需要确认
            If intentJson("requiresConfirmation") IsNot Nothing Then
                Dim needsConfirm = CBool(intentJson("requiresConfirmation"))
                If needsConfirm Then
                    result.Confidence = Math.Min(result.Confidence, 0.5) ' 降低置信度以触发确认
                End If
            End If

            Debug.WriteLine($"[{AppType}] LLM意图解析: {result.OfficeIntent}, 置信度: {result.Confidence:F2}, 描述: {result.UserFriendlyDescription}")

        Catch ex As Exception
            Debug.WriteLine($"ParseLLMIntentResponse 出错: {ex.Message}")
            Return Nothing
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 根据AppType解析意图类型
    ''' </summary>
    Private Sub ParseIntentTypeByApp(intentTypeStr As String, result As IntentResult)
        ' 通用意图
        Select Case intentTypeStr
            Case "FORMAT_STYLE"
                result.OfficeIntent = OfficeIntentType.FORMAT_STYLE
                result.IntentType = ExcelIntentType.FORMAT_STYLE
                Return
            Case "GENERAL_QUERY"
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
                result.IntentType = ExcelIntentType.GENERAL_QUERY
                Return
            Case "UNCLEAR"
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
                result.IntentType = ExcelIntentType.GENERAL_QUERY
                result.Confidence = 0.3
                Return
        End Select

        ' 根据AppType解析特定意图
        Select Case AppType
            Case OfficeApplicationType.Excel
                ParseExcelIntentType(intentTypeStr, result)
            Case OfficeApplicationType.Word
                ParseWordIntentType(intentTypeStr, result)
            Case OfficeApplicationType.PowerPoint
                ParsePowerPointIntentType(intentTypeStr, result)
        End Select
    End Sub

    ''' <summary>
    ''' 解析Excel意图类型
    ''' </summary>
    Private Sub ParseExcelIntentType(intentTypeStr As String, result As IntentResult)
        Select Case intentTypeStr
            Case "DATA_ANALYSIS"
                result.OfficeIntent = OfficeIntentType.DATA_ANALYSIS
                result.IntentType = ExcelIntentType.DATA_ANALYSIS
            Case "FORMULA_CALC"
                result.OfficeIntent = OfficeIntentType.FORMULA_CALC
                result.IntentType = ExcelIntentType.FORMULA_CALC
            Case "CHART_GEN"
                result.OfficeIntent = OfficeIntentType.CHART_GEN
                result.IntentType = ExcelIntentType.CHART_GEN
            Case "DATA_CLEANING"
                result.OfficeIntent = OfficeIntentType.DATA_CLEANING
                result.IntentType = ExcelIntentType.DATA_CLEANING
            Case "REPORT_GEN"
                result.OfficeIntent = OfficeIntentType.REPORT_GEN
                result.IntentType = ExcelIntentType.REPORT_GEN
            Case "DATA_TRANSFORMATION"
                result.OfficeIntent = OfficeIntentType.DATA_TRANSFORMATION
                result.IntentType = ExcelIntentType.DATA_TRANSFORMATION
            Case Else
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
                result.IntentType = ExcelIntentType.GENERAL_QUERY
        End Select
    End Sub

    ''' <summary>
    ''' 解析Word意图类型
    ''' </summary>
    Private Sub ParseWordIntentType(intentTypeStr As String, result As IntentResult)
        Select Case intentTypeStr
            Case "DOCUMENT_EDIT"
                result.OfficeIntent = OfficeIntentType.DOCUMENT_EDIT
            Case "TEXT_FORMAT"
                result.OfficeIntent = OfficeIntentType.TEXT_FORMAT
            Case "TABLE_OPERATION"
                result.OfficeIntent = OfficeIntentType.TABLE_OPERATION
            Case "IMAGE_INSERT"
                result.OfficeIntent = OfficeIntentType.IMAGE_INSERT
            Case "TOC_GENERATION"
                result.OfficeIntent = OfficeIntentType.TOC_GENERATION
            Case "REVIEW_COMMENT"
                result.OfficeIntent = OfficeIntentType.REVIEW_COMMENT
            Case "PROOFREAD"
                result.OfficeIntent = OfficeIntentType.PROOFREAD
            Case Else
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
        End Select
        result.IntentType = ExcelIntentType.GENERAL_QUERY ' Word意图映射到通用
    End Sub

    ''' <summary>
    ''' 解析PowerPoint意图类型
    ''' </summary>
    Private Sub ParsePowerPointIntentType(intentTypeStr As String, result As IntentResult)
        Select Case intentTypeStr
            Case "SLIDE_CREATE"
                result.OfficeIntent = OfficeIntentType.SLIDE_CREATE
            Case "SLIDE_LAYOUT"
                result.OfficeIntent = OfficeIntentType.SLIDE_LAYOUT
            Case "ANIMATION_EFFECT"
                result.OfficeIntent = OfficeIntentType.ANIMATION_EFFECT
            Case "TRANSITION_EFFECT"
                result.OfficeIntent = OfficeIntentType.TRANSITION_EFFECT
            Case "TEMPLATE_APPLY"
                result.OfficeIntent = OfficeIntentType.TEMPLATE_APPLY
            Case "SPEAKER_NOTES"
                result.OfficeIntent = OfficeIntentType.SPEAKER_NOTES
            Case Else
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
        End Select
        result.IntentType = ExcelIntentType.GENERAL_QUERY ' PPT意图映射到通用
    End Sub

    ''' <summary>
    ''' 获取优化后的系统提示词 - 使用PromptManager统一管理
    ''' </summary>
    Public Function GetOptimizedSystemPrompt(intent As IntentResult) As String
        ' 使用PromptManager获取组合后的提示词
        Dim context As New PromptContext With {
            .ApplicationType = AppType.ToString(),
            .IntentResult = intent,
            .FunctionMode = String.Empty
        }

        Return PromptManager.Instance.GetCombinedPrompt(context)
    End Function

    ''' <summary>
    ''' Excel专用JSON Schema约束（内置默认值）
    ''' </summary>
    Public Function GenerateUserFriendlyDescription(intent As IntentResult) As String
        Dim description As String

        Select Case AppType
            Case OfficeApplicationType.Word
                description = GetWordIntentDescription(intent.OfficeIntent)
            Case OfficeApplicationType.PowerPoint
                description = GetPowerPointIntentDescription(intent.OfficeIntent)
            Case Else ' Excel
                description = GetExcelIntentDescription(intent.IntentType)
        End Select

        ' 如果有提取到的实体，补充描述
        If intent.ExtractedEntities.ContainsKey("range") Then
            description &= $" (диапазон: {intent.ExtractedEntities("range")})"
        ElseIf intent.ExtractedEntities.ContainsKey("column") Then
            description &= $" (столбец: {intent.ExtractedEntities("column")})"
        End If

        intent.UserFriendlyDescription = description
        Return description
    End Function

    ''' <summary>
    ''' 获取Excel意图描述
    ''' </summary>
    Private Function GetExcelIntentDescription(intentType As ExcelIntentType) As String
        Select Case intentType
            Case ExcelIntentType.DATA_ANALYSIS
                Return "Статистический анализ данных"
            Case ExcelIntentType.FORMULA_CALC
                Return "Вычисление по формулам"
            Case ExcelIntentType.CHART_GEN
                Return "Построение диаграммы"
            Case ExcelIntentType.DATA_CLEANING
                Return "Очистка и упорядочивание данных"
            Case ExcelIntentType.REPORT_GEN
                Return "Формирование отчёта"
            Case ExcelIntentType.DATA_TRANSFORMATION
                Return "Преобразование данных"
            Case ExcelIntentType.FORMAT_STYLE
                Return "Настройка формата и стиля"
            Case Else
                Return "Обработка запроса в Excel"
        End Select
    End Function

    ''' <summary>
    ''' 获取Word意图描述
    ''' </summary>
    Private Function GetWordIntentDescription(intentType As OfficeIntentType) As String
        Select Case intentType
            Case OfficeIntentType.DOCUMENT_EDIT
                Return "Редактирование содержимого документа"
            Case OfficeIntentType.TEXT_FORMAT
                Return "Форматирование текста"
            Case OfficeIntentType.TABLE_OPERATION
                Return "Работа с таблицами документа"
            Case OfficeIntentType.IMAGE_INSERT
                Return "Вставка и обработка изображений"
            Case OfficeIntentType.TOC_GENERATION
                Return "Создание или обновление оглавления"
            Case OfficeIntentType.REVIEW_COMMENT
                Return "Добавление примечаний и правок"
            Case OfficeIntentType.FORMAT_STYLE
                Return "Настройка оформления документа"
            Case Else
                Return "Обработка запроса в Word"
        End Select
    End Function

    ''' <summary>
    ''' 获取PowerPoint意图描述
    ''' </summary>
    Private Function GetPowerPointIntentDescription(intentType As OfficeIntentType) As String
        Select Case intentType
            Case OfficeIntentType.SLIDE_CREATE
                Return "Создание новых слайдов"
            Case OfficeIntentType.SLIDE_LAYOUT
                Return "Изменение макета слайдов"
            Case OfficeIntentType.ANIMATION_EFFECT
                Return "Добавление анимации"
            Case OfficeIntentType.TRANSITION_EFFECT
                Return "Настройка переходов"
            Case OfficeIntentType.TEMPLATE_APPLY
                Return "Применение темы/шаблона"
            Case OfficeIntentType.SPEAKER_NOTES
                Return "Редактирование заметок докладчика"
            Case OfficeIntentType.FORMAT_STYLE
                Return "Настройка оформления слайдов"
            Case Else
                Return "Обработка запроса в PowerPoint"
        End Select
    End Function

    ''' <summary>
    ''' 构建执行计划预览
    ''' </summary>
    Public Function BuildExecutionPlanPreview(intent As IntentResult) As List(Of ExecutionStep)
        Dim plan As New List(Of ExecutionStep)()

        Select Case intent.IntentType
            Case ExcelIntentType.DATA_ANALYSIS
                plan.Add(New ExecutionStep(1, "Определение области данных", "search"))
                plan.Add(New ExecutionStep(2, "Анализ структуры и типов данных", "data"))
                plan.Add(New ExecutionStep(3, "Статистические вычисления", "formula"))
                plan.Add(New ExecutionStep(4, "Вывод результатов анализа", "data"))

            Case ExcelIntentType.FORMULA_CALC
                plan.Add(New ExecutionStep(1, "Определение целевой ячейки", "search"))
                plan.Add(New ExecutionStep(2, "Построение формулы", "formula"))
                plan.Add(New ExecutionStep(3, "Применение формулы к диапазону", "formula"))

            Case ExcelIntentType.CHART_GEN
                plan.Add(New ExecutionStep(1, "Определение источника данных для диаграммы", "search"))
                plan.Add(New ExecutionStep(2, "Выбор подходящего типа диаграммы", "chart"))
                plan.Add(New ExecutionStep(3, "Создание и настройка диаграммы", "chart"))
                plan.Add(New ExecutionStep(4, "Настройка положения и оформления", "format"))

            Case ExcelIntentType.DATA_CLEANING
                plan.Add(New ExecutionStep(1, "Сканирование области данных", "search"))
                plan.Add(New ExecutionStep(2, "Выявление требующих очистки данных", "data"))
                plan.Add(New ExecutionStep(3, "Выполнение очистки", "clean"))
                plan.Add(New ExecutionStep(4, "Проверка результата очистки", "data"))

            Case ExcelIntentType.REPORT_GEN
                plan.Add(New ExecutionStep(1, "Сбор данных отчёта", "search"))
                plan.Add(New ExecutionStep(2, "Проектирование структуры отчёта", "data"))
                plan.Add(New ExecutionStep(3, "Заполнение данными", "data"))
                plan.Add(New ExecutionStep(4, "Применение оформления отчёта", "format"))

            Case ExcelIntentType.DATA_TRANSFORMATION
                plan.Add(New ExecutionStep(1, "Чтение исходных данных", "search"))
                plan.Add(New ExecutionStep(2, "Преобразование данных", "data"))
                plan.Add(New ExecutionStep(3, "Вывод результата", "data"))

            Case ExcelIntentType.FORMAT_STYLE
                plan.Add(New ExecutionStep(1, "Выбор целевой области", "search"))
                plan.Add(New ExecutionStep(2, "Применение форматирования", "format"))

            Case Else
                plan.Add(New ExecutionStep(1, "Анализ запроса", "search"))
                plan.Add(New ExecutionStep(2, "Формирование решения", "data"))
                plan.Add(New ExecutionStep(3, "Выполнение операции", "default"))
        End Select

        ' 根据提取的实体更新步骤描述
        If intent.ExtractedEntities.ContainsKey("range") Then
            For Each execStep In plan
                If execStep.Description.Contains("област") OrElse execStep.Description.Contains("диапазон") Then
                    execStep.WillModify = intent.ExtractedEntities("range")
                End If
            Next
        End If

        intent.ExecutionPlan = plan
        Return plan
    End Function

#End Region

#Region "提示词模板"

    Private Function CalculateKeywordScore(text As String, keywords As String()) As Double
        Dim matchCount As Integer = 0
        Dim totalWeight As Double = 0

        For Each keyword In keywords
            If text.Contains(keyword.ToLower()) Then
                matchCount += 1
                ' 关键词越长，权重越高
                totalWeight += keyword.Length / 10.0
            End If
        Next

        ' 归一化分数
        If keywords.Length > 0 Then
            Return (matchCount / keywords.Length * 0.5) + (totalWeight / keywords.Length * 0.5)
        End If

        Return 0
    End Function

    ''' <summary>
    ''' 提取关键实体
    ''' </summary>
    Private Sub ExtractEntities(question As String, result As IntentResult)
        ' 提取单元格范围 (如 A1:B10, A1, Sheet1!A1:B10)
        Dim rangePattern As New Regex("([A-Za-z]+\d+)(:[A-Za-z]+\d+)?", RegexOptions.IgnoreCase)
        Dim rangeMatch = rangePattern.Match(question)
        If rangeMatch.Success Then
            result.ExtractedEntities("range") = rangeMatch.Value
        End If

        ' 提取列名 (如 A列, B列)
        Dim columnPattern As New Regex("([A-Za-z])列", RegexOptions.IgnoreCase)
        Dim columnMatch = columnPattern.Match(question)
        If columnMatch.Success Then
            result.ExtractedEntities("column") = columnMatch.Groups(1).Value.ToUpper()
        End If

        ' 提取工作表名 (如 Sheet1, 工作表1)
        Dim sheetPattern As New Regex("(Sheet\d+|工作表\d+)", RegexOptions.IgnoreCase)
        Dim sheetMatch = sheetPattern.Match(question)
        If sheetMatch.Success Then
            result.ExtractedEntities("sheet") = sheetMatch.Value
        End If

        ' 提取数字 (可能是行数、数量等)
        Dim numberPattern As New Regex("\b(\d+)\b")
        Dim numberMatch = numberPattern.Match(question)
        If numberMatch.Success Then
            result.ExtractedEntities("number") = numberMatch.Value
        End If
    End Sub

    ''' <summary>
    ''' 判断执行方式
    ''' </summary>
    Private Sub DetermineExecutionMethod(result As IntentResult)
        ' 以下意图可以使用直接命令
        Dim directCommandIntents = {
            ExcelIntentType.FORMULA_CALC,
            ExcelIntentType.FORMAT_STYLE,
            ExcelIntentType.DATA_CLEANING,
            ExcelIntentType.CHART_GEN
        }

        result.CanUseDirectCommand = directCommandIntents.Contains(result.IntentType)

        ' 复杂操作仍需要VBA
        result.RequiresVBA = Not result.CanUseDirectCommand OrElse
                            result.SecondaryIntents.Count > 1 OrElse
                            result.Confidence < 0.3
    End Sub

#End Region

#Region "上下文相关性检查"

    ''' <summary>
    ''' 异步检查新问题是否与历史对话相关（是追问或继续之前的话题）
    ''' </summary>
    ''' <param name="newQuestion">新问题</param>
    ''' <param name="historyMessages">历史消息列表</param>
    ''' <returns>True表示相关（追问），False表示无关（新话题）</returns>
    Public Async Function IsFollowUpQuestionAsync(newQuestion As String, historyMessages As List(Of HistoryMessage)) As Task(Of Boolean)
        ' 如果没有历史记录，则认为是新话题
        If historyMessages Is Nothing OrElse historyMessages.Count < 2 Then
            Return False
        End If

        ' 获取最近的对话上下文（排除system消息，最多取最近4条）
        Dim filteredHistory = historyMessages.Where(Function(m) m.role <> "system").ToList()
        Dim takeCount = Math.Min(4, filteredHistory.Count)
        Dim recentHistory = filteredHistory.Skip(filteredHistory.Count - takeCount).ToList()

        If recentHistory.Count = 0 Then
            Return False
        End If

        Try
            ' 获取API配置
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return False
            End If

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)

            Dim apiUrl = HttpClientFactory.ResolveChatCompletionsUrl(cfg.url)
            Dim apiKey = cfg.key
            Dim modelName = selectedModel.modelName

            ' 构建上下文摘要
            Dim contextSummary As New StringBuilder()
            For Each msg In recentHistory
                Dim roleLabel = If(msg.role = "user", "Пользователь", "AI")
                Dim contentPreview = If(msg.content?.Length > 200, msg.content.Substring(0, 200) & "...", msg.content)
                contextSummary.AppendLine($"{roleLabel}: {contentPreview}")
            Next

            ' 构建判断提示词
            Dim systemPrompt = "Ты помощник по анализу контекста диалога. Определи, связан ли новый вопрос пользователя с предыдущим диалогом. Отвечай только на русском языке.

Возвращай только JSON:
```json
{""isFollowUp"": true, ""reason"": ""краткая причина""}
```

Критерии:
- isFollowUp=true: новый вопрос — уточнение, дополнение, разъяснение или продолжение предыдущей темы
- isFollowUp=false: новый вопрос — совершенно новая тема, не связанная с предыдущим диалогом

Примеры:
- Ранее обсуждали формулы Excel, новый вопрос «есть другие способы?» → true
- Ранее обсуждали формулы Excel, новый вопрос «построй диаграмму» → false (новая тема)
- Ранее обсуждали анализ данных, новый вопрос «результат неверный» → true (уточнение)
- Ранее обсуждали сложение столбцов A и B в столбец C, новый вопрос «ещё раз» → true (повтор прошлой операции)
- Ранее обсуждали анализ данных, новый вопрос «сделай ещё раз» → true (повтор прошлой операции)
- Ранее обсуждали форматирование таблицы, новый вопрос «продолжай» → true (продолжение прошлой задачи)
- Ранее обсуждали создание диаграммы, новый вопрос «сделай ещё одну» → true (повтор похожей операции)

Особое внимание:
- Короткие запросы вроде «ещё раз», «сделай ещё раз», «продолжай», «сделай ещё одну» обычно являются уточнением и означают повтор или продолжение прошлой операции
- Только явно совершенно посторонняя новая тема считается false"

            Dim userMessage = $"Предыдущий диалог:
{contextSummary}

Новый вопрос: {newQuestion}

Связан ли новый вопрос с предыдущим диалогом?"

            ' 构建请求
            Dim messages As New JArray()
            messages.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject From {{"role", "user"}, {"content", userMessage}})

            Dim gatewayResponse = Await AiGateway.SendChatAsync(New AiRequestOptions With {
                .ApiUrl = apiUrl,
                .ApiKey = apiKey,
                .ModelName = modelName,
                .Platform = cfg.platform,
                .Messages = messages,
                .Temperature = 0.2R,
                .MaxTokens = 100,
                .TimeoutSeconds = 15
            })

            If gatewayResponse IsNot Nothing AndAlso gatewayResponse.Success Then
                Return ParseFollowUpResponse(gatewayResponse.Content)
            End If

            Debug.WriteLine($"IsFollowUpQuestionAsync API失败: {If(gatewayResponse Is Nothing, "empty response", gatewayResponse.ErrorMessage)}")

        Catch ex As Exception
            Debug.WriteLine($"IsFollowUpQuestionAsync 出错: {ex.Message}")
        End Try

        ' 默认认为可能相关（避免误判导致每次都弹框）
        Return True
    End Function

    ''' <summary>
    ''' 解析追问判断的响应
    ''' </summary>
    Private Function ParseFollowUpResponse(responseContent As String) As Boolean
        Try
            Dim content = ExtractAssistantContentOrRaw(responseContent)
            If String.IsNullOrEmpty(content) Then Return True

            ' 提取JSON部分
            Dim jsonMatch = Regex.Match(content, "\{[\s\S]*\}")
            If Not jsonMatch.Success Then Return True

            Dim resultJson = JObject.Parse(jsonMatch.Value)
            Dim isFollowUpToken = resultJson("isFollowUp")
            If isFollowUpToken Is Nothing Then Return True

            Dim isFollowUp As Boolean = isFollowUpToken.Value(Of Boolean)()

            Debug.WriteLine($"追问判断结果: isFollowUp={isFollowUp}, 原因={resultJson("reason")}")
            Return isFollowUp

        Catch ex As Exception
            Debug.WriteLine($"ParseFollowUpResponse 出错: {ex.Message}")
            Return True ' 默认认为相关
        End Try
    End Function

    Private Function ExtractAssistantContentOrRaw(responseContent As String) As String
        If String.IsNullOrWhiteSpace(responseContent) Then Return ""

        Dim trimmed = responseContent.TrimStart()
        ' 普通文本/Markdown 是合法响应，不应先用 JObject.Parse 制造 first-chance
        ' JsonReaderException。只有看起来像 API envelope 时才尝试解析。
        If Not trimmed.StartsWith("{") Then Return responseContent

        Try
            Dim responseJson = JObject.Parse(responseContent)
            Dim choices = responseJson("choices")
            If choices IsNot Nothing AndAlso choices.Count > 0 Then
                Dim content = choices(0)("message")?("content")?.ToString()
                If Not String.IsNullOrEmpty(content) Then Return content
            End If
        Catch
        End Try

        Return responseContent
    End Function

#End Region

End Class
