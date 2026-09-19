' ShareRibbon\Services\Reformat\SemanticPromptBuilder.vb
' 统一构建语义标注提示词

Imports System.Text
Imports ShareRibbon.Services.Reformat

''' <summary>
''' 语义提示词构建器 - 为AI构建语义标注的系统提示词
''' 模板排版和规范排版共用同一构建逻辑
''' </summary>
Public Class SemanticPromptBuilder

    ' 单例场景管理器
    Private Shared _scenarioManager As ScenarioManager

    Shared Sub New()
        _scenarioManager = New ScenarioManager()
    End Sub

    ''' <summary>
    ''' 构建语义标注提示词（带样式上下文）
    ''' </summary>
    ''' <param name="mapping">语义样式映射（包含可用标签）</param>
    ''' <param name="paragraphs">段落文本列表（仅文本段落，非文本已过滤）</param>
    ''' <param name="paragraphStyles">段落样式名称（仅文本段落，与paragraphs一一对应）</param>
    ''' <param name="originalParaIndices">原文档中的段落索引（仅文本段落，用于映射回正确位置）</param>
    ''' <param name="detectedHeadings">DocumentAnalyzer检测到的标题信息</param>
    Public Shared Function BuildSemanticTaggingPrompt(
        mapping As SemanticStyleMapping,
        paragraphs As List(Of String),
        Optional paragraphStyles As List(Of String) = Nothing,
        Optional originalParaIndices As List(Of Integer) = Nothing,
        Optional detectedHeadings As String = Nothing,
        Optional documentTypeContext As String = Nothing,
        Optional paragraphFontSizes As List(Of Single) = Nothing,
        Optional paragraphIsBold As List(Of Boolean) = Nothing) As String

        Dim sb As New StringBuilder()

        ' ===== 1. 角色定义 =====
        sb.AppendLine("Ты эксперт по анализу структуры документов, хорошо распознаёшь структуру и семантические роли абзацев. Отвечай только на русском языке.")
        sb.AppendLine("Твоя задача — проанализировать содержимое документа, определить семантическую роль каждого абзаца, чтобы система автоматически применила соответствующий стандартный формат.")
        sb.AppendLine()

        ' ===== 2. 任务说明 =====
        sb.AppendLine("【Задача】")
        sb.AppendLine("Проанализируй документ в следующем порядке:")
        sb.AppendLine("Шаг 1: определи тип документа и общую структуру")
        sb.AppendLine("Шаг 2: распознай ключевые структурные элементы (заголовки, основной текст, подпись, дата и т. п.)")
        sb.AppendLine("Шаг 3: назначь каждому абзацу подходящую семантическую метку")
        sb.AppendLine()

        ' ===== 3. 文档类型上下文 =====
        If Not String.IsNullOrEmpty(documentTypeContext) Then
            sb.AppendLine("【Стандарт оформления】")
            sb.AppendLine(documentTypeContext)
            sb.AppendLine()
        End If

        ' ===== 4. 输出格式 =====
        sb.AppendLine("【Формат вывода】")
        sb.AppendLine("Выводи только чистый JSON-массив, без другого содержимого (без markdown-блоков кода, без пояснений).")
        sb.AppendLine("[")
        sb.AppendLine("  {""paraIndex"":0, ""tag"":""header.org"", ""reason"":""в начале документа, текст соответствует шаблону обозначения органа""},")
        sb.AppendLine("  {""paraIndex"":1, ""tag"":""header.refno"", ""reason"":""содержит формат номера документа""},")
        sb.AppendLine("  ...")
        sb.AppendLine("]")
        sb.AppendLine("Требования:")
        sb.AppendLine("- поле reason кратко объясняет основание (не более 30 символов)")
        sb.AppendLine("- у каждого абзаца должна быть ровно одна метка")
        sb.AppendLine("- paraIndex используй из индексов, указанных в 【Абзацы документа】")
        sb.AppendLine()

        ' ===== 5. 可用标签 =====
        sb.AppendLine("【Доступные семантические метки】")
        For Each tag In mapping.SemanticTags
            sb.Append($"- {tag.TagId}: {tag.DisplayName}")
            If Not String.IsNullOrEmpty(tag.MatchHint) Then
                sb.Append($". Подсказка распознавания: {tag.MatchHint}")
            End If
            sb.AppendLine()
        Next
        sb.AppendLine()

        ' ===== 6. 结构识别指南 =====
        sb.AppendLine("【Руководство по распознаванию структуры】")
        sb.AppendLine("Определяя роль абзаца, учитывай следующие признаки:")
        sb.AppendLine("1. Содержимое текста: есть ли характерный шаблон (номер документа, дата, номер и т. п.)")
        sb.AppendLine("2. Позиция абзаца: в начале, середине или конце документа")
        sb.AppendLine("3. Контекстные связи: отношение к предыдущему и следующему абзацу (например, после заголовка обычно идёт основной текст)")
        sb.AppendLine("4. Форматные признаки: короткий абзац с увеличенным кеглем и полужирным начертанием обычно является заголовком")
        sb.AppendLine("5. При неуверенности используй наиболее общую метку body.normal")
        sb.AppendLine()

        ' ===== 场景化结构识别（从 JSON 加载） =====
        Dim scenario As FormattingScenario = Nothing
        If Not String.IsNullOrEmpty(documentTypeContext) Then
            scenario = _scenarioManager.MatchScenario(documentTypeContext)
        End If

        If scenario IsNot Nothing Then
            ' 使用场景的识别模式（详细）
            Dim patternsText = ScenarioManager.BuildIdentificationPatternsText(scenario)
            If Not String.IsNullOrEmpty(patternsText) Then
                sb.Append(patternsText)
            End If

            ' 使用场景的结构指导
            Dim guidanceText = ScenarioManager.BuildStructureGuidanceText(scenario)
            If Not String.IsNullOrEmpty(guidanceText) Then
                sb.Append(guidanceText)
            End If
        End If

        ' ===== 7. 标注示例（按文档类型） =====
        sb.AppendLine("【标注示例】")
        If scenario IsNot Nothing Then
            ' 使用场景的示例
            Dim examplesText = ScenarioManager.BuildExamplesText(scenario)
            If Not String.IsNullOrEmpty(examplesText) Then
                sb.Append(examplesText)
            Else
                ' 回退到旧方法
                Dim examples As String = GetExamplesByDocumentType(documentTypeContext, mapping)
                sb.Append(examples)
            End If
        Else
            ' 无场景，使用旧方法
            Dim examples As String = GetExamplesByDocumentType(documentTypeContext, mapping)
            sb.Append(examples)
        End If
        sb.AppendLine()

        ' ===== 8. 严格要求 =====
        sb.AppendLine("【Строгие требования】")
        sb.AppendLine("1. Используй только перечисленные выше метки, самодельные метки запрещены")
        sb.AppendLine("2. Возвращай чистый JSON-массив, без markdown-обёртки")
        sb.AppendLine("3. У каждого абзаца должна быть ровно одна метка")
        sb.AppendLine("4. Соблюдай иерархию: после title.1 может идти title.2 или body, нельзя сразу перескакивать на title.3")
        sb.AppendLine()

        ' ===== 9. 自动检测结果 =====
        If Not String.IsNullOrEmpty(detectedHeadings) Then
            sb.AppendLine("【Автоматически определённая система структура заголовков (справочно, можно исправить)】")
            sb.AppendLine(detectedHeadings)
            sb.AppendLine()
        End If

        ' ===== 10. 文档段落（完整文本+上下文） =====
        sb.AppendLine("【Абзацы документа】")
        Dim hasStyles = paragraphStyles IsNot Nothing AndAlso paragraphStyles.Count = paragraphs.Count
        Dim hasOrigIdx = originalParaIndices IsNot Nothing AndAlso originalParaIndices.Count = paragraphs.Count
        Dim hasFontSizes = paragraphFontSizes IsNot Nothing AndAlso paragraphFontSizes.Count = paragraphs.Count
        Dim hasBold = paragraphIsBold IsNot Nothing AndAlso paragraphIsBold.Count = paragraphs.Count

        For i = 0 To paragraphs.Count - 1
            Dim origIdx = If(hasOrigIdx, originalParaIndices(i), i)
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 位置标签
            Dim positionLabel = ""
            If i = 0 Then
                positionLabel = " [начало документа]"
            ElseIf i >= paragraphs.Count - 3 Then
                positionLabel = " [конец документа]"
            End If

            ' 样式提示（简洁）
            Dim styleHint As String = ""
            If hasStyles AndAlso Not String.IsNullOrEmpty(paragraphStyles(i)) Then
                styleHint = $" [стиль:{paragraphStyles(i)}]"
            End If

            ' 格式线索（简洁）
            Dim formatHint As String = ""
            If hasFontSizes Then
                formatHint = $" {paragraphFontSizes(i):F0}pt"
            End If
            If hasBold AndAlso paragraphIsBold(i) Then
                formatHint &= " полужирный"
            End If
            If formatHint <> "" Then
                formatHint = $" [формат:{formatHint.Trim()}]"
            End If

            ' 上下文：显示前一段落的最后20字
            Dim contextBefore As String = ""
            If i > 0 AndAlso Not String.IsNullOrWhiteSpace(paragraphs(i - 1)) Then
                Dim prevText = paragraphs(i - 1).Trim()
                If prevText.Length > 20 Then prevText = "..." & prevText.Substring(prevText.Length - 20)
                contextBefore = $"  ↑выше: {prevText}" & vbCrLf
            End If

            ' 不截断段落文本，但超长段落只取前300字+后缀
            If text.Length > 300 Then
                text = text.Substring(0, 300) & $"...[всего {text.Length} симв.]"
            End If

            sb.Append(contextBefore)
            sb.AppendLine($"[{origIdx}]{positionLabel}{styleHint}{formatHint} {text}")
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 根据文档类型获取对应的标注示例
    ''' </summary>
    ''' <param name="documentTypeContext">文档类型上下文（标准名称）</param>
    ''' <param name="mapping">语义样式映射</param>
    Private Shared Function GetExamplesByDocumentType(documentTypeContext As String, mapping As SemanticStyleMapping) As String
        Dim sb As New StringBuilder()
        Dim ctx = If(documentTypeContext, "").ToLowerInvariant()

        ' Официальные документы
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("公文") OrElse documentTypeContext.Contains("GB/T 9704") OrElse
            ctx.Contains("официальн") OrElse ctx.Contains("служебн") OrElse
            ctx.Contains("приказ") OrElse ctx.Contains("распоряжен")) Then
            sb.AppendLine("Примеры разметки официального документа:")
            sb.AppendLine("«Федеральная служба» → header.org (основание: в начале документа, соответствует шаблону обозначения органа)")
            sb.AppendLine("「×政发〔2024〕15号」 → header.refno (основание: содержит формат номера документа 〔〕X号)")
            sb.AppendLine("«Подписал: Иванов И.И.» → header.signer (основание: содержит отметку о подписанте)")
            sb.AppendLine("«О мерах по обеспечению безопасности» → title.main (основание: заголовок документа, формат «О … »)")
            sb.AppendLine("«Руководителям структурных подразделений:» → title.recipient (основание: адресат, заканчивается двоеточием)")
            sb.AppendLine("«1. Общие положения» → title.1 (основание: заголовок первого уровня)")
            sb.AppendLine("«(1) Основные принципы» → title.2 (основание: заголовок второго уровня)")
            sb.AppendLine("«1.1. Усилить контроль» → title.3 (основание: заголовок третьего уровня)")
            sb.AppendLine("«В целях повышения качества работы...» → body.normal (основание: абзац основного текста)")
            sb.AppendLine("«Приложение: 1. План мероприятий» → body.attachment (основание: отметка о приложении)")
            sb.AppendLine("«Руководитель службы» → footer.signature (основание: наименование органа в конце документа)")
            sb.AppendLine("«15 января 2024 г.» → footer.date (основание: формат даты в конце документа)")
            sb.AppendLine("«(Контактное лицо: Петров, тел. 123-45-67)» → footer.note (основание: примечание документа)")
            sb.AppendLine("«Копия: всем подразделениям» → footer.cc (основание: начинается с отметки о рассылке)")
            Return sb.ToString()
        End If

        ' Научная статья
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("学术") OrElse documentTypeContext.Contains("论文") OrElse
            ctx.Contains("научн") OrElse ctx.Contains("статья") OrElse ctx.Contains("диссертац") OrElse
            ctx.Contains("курсов") OrElse ctx.Contains("вкр")) Then
            sb.AppendLine("Примеры разметки научной статьи:")
            sb.AppendLine("«Исследование методов распознавания изображений на основе глубокого обучения» → title.main (основание: заголовок статьи)")
            sb.AppendLine("«Аннотация» → title.abstract (основание: заголовок аннотации)")
            sb.AppendLine("«В работе предложен новый метод...» → body.abstract (основание: текст аннотации)")
            sb.AppendLine("«Ключевые слова» → title.keywords (основание: заголовок ключевых слов)")
            sb.AppendLine("«глубокое обучение; распознавание изображений» → body.keywords (основание: содержание ключевых слов)")
            sb.AppendLine("«Глава 1. Введение» → heading.1 (основание: заголовок главы)")
            sb.AppendLine("«1.1. Актуальность исследования» → heading.2 (основание: заголовок второго уровня)")
            sb.AppendLine("«В последние годы с развитием технологий...» → body.normal (основание: абзац основного текста)")
            sb.AppendLine("«Список литературы» → title.references (основание: заголовок списка литературы)")
            Return sb.ToString()
        End If

        ' Бизнес-отчёт
        If Not String.IsNullOrEmpty(documentTypeContext) AndAlso
           (documentTypeContext.Contains("商务") OrElse documentTypeContext.Contains("报告") OrElse
            ctx.Contains("бизнес") OrElse ctx.Contains("отчёт") OrElse ctx.Contains("отчет") OrElse
            ctx.Contains("доклад") OrElse ctx.Contains("аналитическ")) Then
            sb.AppendLine("Примеры разметки бизнес-отчёта:")
            sb.AppendLine("«Отчёт об итогах работы за 2024 год» → title.main (основание: заголовок отчёта)")
            sb.AppendLine("«1. Обзор результатов года» → heading.1 (основание: заголовок первого уровня)")
            sb.AppendLine("«(1) Анализ выручки» → heading.2 (основание: заголовок второго уровня)")
            sb.AppendLine("«В 2024 году выручка компании выросла на 15%...» → body.normal (основание: абзац основного текста)")
            sb.AppendLine("«Таким образом...» → body.summary (основание: абзац с выводами)")
            Return sb.ToString()
        End If

        ' Универсальный документ (по умолчанию)
        sb.AppendLine("Примеры разметки универсального документа:")
        sb.AppendLine("«Глава 1. Общие положения» → heading.1 (основание: заголовок главы)")
        sb.AppendLine("«1.1. Цели и основания» → heading.2 (основание: заголовок второго уровня)")
        sb.AppendLine("«1.1.1. В целях регулирования...» → heading.3 (основание: заголовок третьего уровня)")
        sb.AppendLine("«Настоящий документ определяет...» → body.normal (основание: абзац основного текста)")

        ' Если в mapping есть пользовательские метки, показать и их
        If mapping IsNot Nothing AndAlso mapping.SemanticTags.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("Особые метки, поддерживаемые текущим стандартом:")
            For Each tag In mapping.SemanticTags.Take(6)
                If tag.TagId.StartsWith("header.") OrElse tag.TagId.StartsWith("title.") OrElse tag.TagId.StartsWith("footer.") Then
                    sb.AppendLine($"- {tag.TagId}: {tag.DisplayName}")
                End If
            Next
        End If

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建带重试提示的标注提示词（当校验失败时使用）
    ''' </summary>
    Public Shared Function BuildRetryPrompt(
        mapping As SemanticStyleMapping,
        paragraphs As List(Of String),
        errors As List(Of String)) As String

        Dim sb As New StringBuilder()

        ' 原始提示词
        sb.Append(BuildSemanticTaggingPrompt(mapping, paragraphs))
        sb.AppendLine()

        ' 错误反馈
        sb.AppendLine("【В предыдущем выводе есть следующие ошибки, исправь их】")
        For Each errMsg In errors
            sb.AppendLine($"- {errMsg}")
        Next
        sb.AppendLine()
        sb.AppendLine("Повторно выведи корректный JSON-массив (с полями paraIndex, tag и reason).")

        Return sb.ToString()
    End Function
End Class
