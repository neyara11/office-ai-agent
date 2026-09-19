' ShareRibbon\Services\Proofread\ProofreadPromptBuilder.vb
' 校对Prompt构建器 - 构建AI校对Prompt，解析校对结果

Imports System.Collections.Generic
Imports System.Text

''' <summary>
''' 校对Prompt构建器
''' </summary>
Public Class ProofreadPromptBuilder

    ''' <summary>
    ''' 构建全文校对Prompt
    ''' </summary>
    Public Shared Function BuildFullDocumentPrompt(paragraphs As List(Of String)) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("Ты профессиональный корректор русского документа. Внимательно проверь документ ниже и найди проблемы, требующие исправления. Отвечай только на русском языке.")
        sb.AppendLine()
        sb.AppendLine("Выводи только JSON-массив, без любого другого содержимого (без markdown-блоков кода, без пояснений).")
        sb.AppendLine()

        sb.AppendLine("【Область проверки】")
        sb.AppendLine("1. Орфографические ошибки и опечатки")
        sb.AppendLine("2. Ошибки словоупотребления, в том числе:")
        sb.AppendLine("   - смешение похожих слов и паронимов")
        sb.AppendLine("   - неверное употребление предлогов и падежных форм")
        sb.AppendLine("   - прочие частые лексические ошибки")
        sb.AppendLine("3. Ошибки пунктуации:")
        sb.AppendLine("   - пропущенные или лишние знаки препинания")
        sb.AppendLine("   - тире и дефис, кавычки, скобки")
        sb.AppendLine("   - несогласованные парные знаки")
        sb.AppendLine("4. Грамматические ошибки и нарушения согласования")
        sb.AppendLine("5. Неудачные или двусмысленные формулировки")
        sb.AppendLine()

        sb.AppendLine("【Принцип минимальных правок】")
        sb.AppendLine("- Исправляй только то, что действительно ошибочно; не переписывай корректный текст ради стиля")
        sb.AppendLine("- suggestion должен сохранять смысл оригинала и исправлять только ошибку")
        sb.AppendLine("- original должен точно совпадать с исходным текстом, включая пунктуацию и пробелы")
        sb.AppendLine()

        sb.AppendLine("【Содержание документа】")
        For i = 0 To paragraphs.Count - 1
            Dim para = paragraphs(i)
            If Not String.IsNullOrWhiteSpace(para) Then
                sb.AppendLine($"[абзац{i}] {para}")
            End If
        Next
        sb.AppendLine()

        sb.AppendLine("【Формат вывода】")
        sb.AppendLine("Выведи чистый JSON-массив (без обёртки в markdown-блок кода):")
        sb.AppendLine("[")
        sb.AppendLine("  {")
        sb.AppendLine("    ""paragraphIndex"": 0,")
        sb.AppendLine("    ""original"": ""фрагмент оригинала, требующий исправления"",")
        sb.AppendLine("    ""suggestion"": ""исправленный текст"",")
        sb.AppendLine("    ""issueType"": ""SpellingError"",")
        sb.AppendLine("    ""severity"": ""High"",")
        sb.AppendLine("    ""explanation"": ""краткое пояснение причины правки""")
        sb.AppendLine("  }")
        sb.AppendLine("]")
        sb.AppendLine()

        sb.AppendLine("【Допустимые значения issueType (единый lowerCamelCase)】")
        sb.AppendLine("- SpellingError: орфографическая ошибка")
        sb.AppendLine("- WordUsageError: ошибка словоупотребления")
        sb.AppendLine("- PunctuationError: ошибка пунктуации")
        sb.AppendLine("- GrammaticalError: грамматическая ошибка")
        sb.AppendLine("- ExpressionError: неудачное выражение")
        sb.AppendLine()

        sb.AppendLine("【Допустимые значения severity】")
        sb.AppendLine("- High: обязательно исправить (например, орфографическая или грубая грамматическая ошибка)")
        sb.AppendLine("- Medium: рекомендуется исправить (например, неточное словоупотребление, лёгкая стилистическая ошибка)")
        sb.AppendLine("- Low: необязательное улучшение (например, формулировку можно сделать точнее)")
        sb.AppendLine()

        sb.AppendLine("【Примечания】")
        sb.AppendLine("1. original должен точно совпадать с текстом документа, включая пунктуацию и пробелы")
        sb.AppendLine("2. Если в одном абзаце несколько проблем, верни несколько записей")
        sb.AppendLine("3. Возвращай только фрагменты, требующие правки; абзацы без проблем в результат не включай")
        sb.AppendLine("4. Если исправлять нечего, верни пустой массив: []")
        sb.AppendLine("5. Не пропускай очевидные проблемы, но и не переусердствуй с правками")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 分析AI返回的校对结果，保留解析失败状态。
    ''' </summary>
    Public Shared Function AnalyzeProofreadResponse(
        aiResponse As String,
        Optional paragraphs As List(Of String) = Nothing) As ProofreadAnalysisResult

        Try
            Dim parseResult = ProofreadJsonParser.Parse(aiResponse)
            Dim analysis As New ProofreadAnalysisResult With {
                .RawResponsePreview = BuildRawResponsePreview(aiResponse),
                .Summary = parseResult.Summary,
                .FormatDetected = parseResult.FormatDetected
            }

            If Not parseResult.Success Then
                Debug.WriteLine($"[ProofreadPromptBuilder] 解析校对结果失败: {parseResult.ErrorMessage}")
                analysis.Status = ProofreadAnalysisStatus.ParseFailed
                analysis.ErrorMessage = parseResult.ErrorMessage
                Return analysis
            End If

            analysis.Issues = If(parseResult.Issues, New List(Of ProofreadIssue)())
            analysis.Status = If(analysis.Issues.Count > 0, ProofreadAnalysisStatus.HasIssues, ProofreadAnalysisStatus.NoIssues)
            Return analysis

        Catch ex As Exception
            Debug.WriteLine($"[ProofreadPromptBuilder] 解析校对结果异常: {ex.Message}")
            Return New ProofreadAnalysisResult With {
                .Status = ProofreadAnalysisStatus.ModelFailed,
                .ErrorMessage = ex.Message,
                .RawResponsePreview = BuildRawResponsePreview(aiResponse)
            }
        End Try
    End Function

    ''' <summary>
    ''' 解析AI返回的校对结果（兼容旧调用；失败时仍返回空列表）。
    ''' 新代码应使用 AnalyzeProofreadResponse 区分状态。
    ''' </summary>
    Public Shared Function ParseProofreadResponse(
        aiResponse As String,
        Optional paragraphs As List(Of String) = Nothing) As List(Of ProofreadIssue)

        Dim analysis = AnalyzeProofreadResponse(aiResponse, paragraphs)
        If analysis Is Nothing OrElse Not analysis.HasIssues Then
            Return New List(Of ProofreadIssue)()
        End If
        Return analysis.Issues
    End Function

    Private Shared Function BuildRawResponsePreview(aiResponse As String) As String
        If String.IsNullOrWhiteSpace(aiResponse) Then Return ""
        Dim normalized = aiResponse.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If normalized.Length <= 500 Then Return normalized
        Return normalized.Substring(0, 500) & "..."
    End Function

End Class
