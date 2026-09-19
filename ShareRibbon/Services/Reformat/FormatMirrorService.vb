' ShareRibbon\Services\Reformat\FormatMirrorService.vb
' 格式克隆：分析现有文档/选区的实际格式，提取规则，让AI生成SemanticStyleMapping

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
' 注意：此服务专门用于 Word，使用 Object 类型和后期绑定以避免 ShareRibbon 直接依赖 Word Interop

''' <summary>从文档段落提取的格式信息</summary>
Public Class ExtractedParagraphFormat
    Public Property StyleName As String = ""
    Public Property SampleText As String = ""
    Public Property FontNameCN As String = ""
    Public Property FontNameEN As String = ""
    Public Property FontSize As Double = 12
    Public Property Bold As Boolean = False
    Public Property Italic As Boolean = False
    Public Property AlignmentStr As String = "left"   ' left/center/right/justify
    Public Property FirstLineIndentCm As Double = 0
    Public Property LineSpacingPt As Double = 0
    Public Property SpaceBeforePt As Double = 0
    Public Property SpaceAfterPt As Double = 0
    Public Property OccurrenceCount As Integer = 1
End Class

Public Class FormatMirrorService

    Private Const MaxSamplesToExtract As Integer = 200   ' 最多采样段落数
    Private Const PointsPerCm As Double = 28.35

    ''' <summary>
    ''' 从 Word 文档中提取段落格式信息（全文或当前选区）
    ''' </summary>
    ''' <param name="wordApp">Word Application 对象（使用 Object 类型避免直接依赖 Word Interop）</param>
    ''' <param name="selectionOnly">是否仅处理选区</param>
    Public Shared Function ExtractFormattingFromDocument(
        wordApp As Object,
        selectionOnly As Boolean) As List(Of ExtractedParagraphFormat)

        Dim result As New List(Of ExtractedParagraphFormat)()
        Dim styleMap As New Dictionary(Of String, ExtractedParagraphFormat)()

        If wordApp Is Nothing Then Return result

        Try
            ' 使用动态绑定访问 Word 对象模型
            Dim documents As Object = wordApp.[GetType]().InvokeMember("Documents", Reflection.BindingFlags.GetProperty, Nothing, wordApp, Nothing)
            Dim docCount As Integer = CInt(documents.[GetType]().InvokeMember("Count", Reflection.BindingFlags.GetProperty, Nothing, documents, Nothing))
            If docCount = 0 Then Return result

            Dim doc As Object = wordApp.[GetType]().InvokeMember("ActiveDocument", Reflection.BindingFlags.GetProperty, Nothing, wordApp, Nothing)

            ' 根据selectionOnly参数选择段落来源
            Dim paragraphs As Object
            If selectionOnly Then
                Dim selection As Object = wordApp.[GetType]().InvokeMember("Selection", Reflection.BindingFlags.GetProperty, Nothing, wordApp, Nothing)
                paragraphs = selection.[GetType]().InvokeMember("Paragraphs", Reflection.BindingFlags.GetProperty, Nothing, selection, Nothing)
            Else
                paragraphs = doc.[GetType]().InvokeMember("Paragraphs", Reflection.BindingFlags.GetProperty, Nothing, doc, Nothing)
            End If
            Dim paraCount As Integer = CInt(paragraphs.[GetType]().InvokeMember("Count", Reflection.BindingFlags.GetProperty, Nothing, paragraphs, Nothing))

            Dim count As Integer = 0
            For i As Integer = 1 To Math.Min(paraCount, MaxSamplesToExtract)
                Dim p As Object = paragraphs.[GetType]().InvokeMember("Item", Reflection.BindingFlags.GetProperty, Nothing, paragraphs, New Object() {i})
                Dim rangeObj As Object = p.[GetType]().InvokeMember("Range", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                Dim txt As String = CStr(rangeObj.[GetType]().InvokeMember("Text", Reflection.BindingFlags.GetProperty, Nothing, rangeObj, Nothing))
                txt = txt.Replace(Chr(13), "").Replace(Chr(7), "").Trim()
                If String.IsNullOrWhiteSpace(txt) Then Continue For

                count += 1
                Dim styleObj As Object = p.[GetType]().InvokeMember("Style", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                Dim styleName As String = styleObj.ToString()
                Dim fontObj As Object = rangeObj.[GetType]().InvokeMember("Font", Reflection.BindingFlags.GetProperty, Nothing, rangeObj, Nothing)
                Dim fontSize As Object = fontObj.[GetType]().InvokeMember("Size", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                Dim fontBold As Object = fontObj.[GetType]().InvokeMember("Bold", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                Dim alignment As Object = p.[GetType]().InvokeMember("Alignment", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)

                Dim fmt As New ExtractedParagraphFormat()
                fmt.StyleName = styleName
                fmt.SampleText = If(txt.Length > 60, txt.Substring(0, 60) & "…", txt)

                ' 字体信息
                Try
                    Dim nameFarEast As Object = fontObj.[GetType]().InvokeMember("NameFarEast", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                    Dim fontName As Object = fontObj.[GetType]().InvokeMember("Name", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                    fmt.FontNameCN = If(nameFarEast Is Nothing, "", nameFarEast.ToString())
                    fmt.FontNameEN = If(fontName Is Nothing, "", fontName.ToString())
                    fmt.FontSize = If(fontSize Is Nothing, 12, Convert.ToDouble(fontSize))
                    fmt.Bold = Convert.ToInt32(fontBold) = -1 OrElse Convert.ToInt32(fontBold) = 1
                    Dim fontItalic As Object = fontObj.[GetType]().InvokeMember("Italic", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                    fmt.Italic = Convert.ToInt32(fontItalic) = -1 OrElse Convert.ToInt32(fontItalic) = 1
                Catch
                End Try

                ' 段落信息
                Try
                    Dim alignVal As Integer = Convert.ToInt32(alignment)
                    ' WdParagraphAlignment: 0=left, 1=center, 2=right, 3=justify
                    Select Case alignVal
                        Case 1 : fmt.AlignmentStr = "center"
                        Case 2 : fmt.AlignmentStr = "right"
                        Case 3 : fmt.AlignmentStr = "justify"
                        Case Else : fmt.AlignmentStr = "left"
                    End Select
                    Dim firstLineIndent As Object = p.[GetType]().InvokeMember("FirstLineIndent", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                    Dim lineSpacing As Object = p.[GetType]().InvokeMember("LineSpacing", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                    Dim spaceBefore As Object = p.[GetType]().InvokeMember("SpaceBefore", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                    Dim spaceAfter As Object = p.[GetType]().InvokeMember("SpaceAfter", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)

                    fmt.FirstLineIndentCm = If(firstLineIndent IsNot Nothing, Math.Round(Convert.ToDouble(firstLineIndent) / PointsPerCm, 2), 0)
                    fmt.LineSpacingPt = If(lineSpacing IsNot Nothing, Math.Round(Convert.ToDouble(lineSpacing), 1), 0)
                    fmt.SpaceBeforePt = If(spaceBefore IsNot Nothing, Math.Round(Convert.ToDouble(spaceBefore), 1), 0)
                    fmt.SpaceAfterPt = If(spaceAfter IsNot Nothing, Math.Round(Convert.ToDouble(spaceAfter), 1), 0)
                Catch
                End Try

                ' 使用改进的分组key：包含字体名和缩进，区分更多格式差异
                Dim key As String = String.Join("|", {
                    styleName,
                    fmt.FontSize.ToString(),
                    fmt.Bold.ToString(),
                    fmt.AlignmentStr,
                    fmt.FontNameCN,
                    fmt.FirstLineIndentCm.ToString("F1")
                })

                If styleMap.ContainsKey(key) Then
                    styleMap(key).OccurrenceCount += 1
                    Continue For
                End If

                styleMap(key) = fmt
            Next

            ' 按出现次数排序，频率高的排前面
            result = styleMap.Values.OrderByDescending(Function(f) f.OccurrenceCount).ToList()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"[FormatMirrorService] 提取失败: {ex.Message}")
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 构建让 AI 将提取的格式规则转换为 SemanticStyleMapping 的提示词
    ''' </summary>
    ''' <param name="extracted">从文档提取的格式信息列表</param>
    ''' <param name="availableTags">可用语义标签列表（从SemanticStyleMapping动态获取）</param>
    Public Shared Function BuildClonePrompt(extracted As List(Of ExtractedParagraphFormat),
                                            Optional availableTags As List(Of SemanticTag) = Nothing) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("Ты — эксперт по вёрстке. На основе извлечённых из реального документа параметров абзацев сформируй JSON SemanticStyleMapping.")
        sb.AppendLine()

        ' 可用语义标签（动态生成，不再硬编码）
        sb.AppendLine("【Доступные семантические теги】")
        If availableTags IsNot Nothing AndAlso availableTags.Count > 0 Then
            For Each tag In availableTags
                sb.Append($"- {tag.TagId}: {tag.DisplayName}")
                If Not String.IsNullOrEmpty(tag.MatchHint) Then
                    sb.Append($" ({tag.MatchHint})")
                End If
                sb.AppendLine()
            Next
        Else
            ' 降级：使用基础标签集
            sb.AppendLine("title.main (главный заголовок), title.1 (заголовок 1), title.2 (заголовок 2), title.3 (заголовок 3)")
            sb.AppendLine("heading.1 (заголовок 1), heading.2 (заголовок 2), heading.3 (заголовок 3)")
            sb.AppendLine("body.normal (основной текст), body.emphasis (акцент в тексте)")
            sb.AppendLine("body.abstract (аннотация), body.keywords (ключевые слова)")
            sb.AppendLine("list.ordered (нумерованный список), list.unordered (маркированный список)")
            sb.AppendLine("quote (цитата/аннотация), caption (подпись к рисунку или таблице)")
            sb.AppendLine("header.org (реквизит организации), header.refno (номер документа)")
            sb.AppendLine("footer.signature (подпись), footer.date (дата)")
        End If
        sb.AppendLine()

        sb.AppendLine("【Извлечённые из документа правила форматирования (по убыванию частоты)】")

        For Each f In extracted.Take(30)
            sb.AppendLine($"- Имя стиля: {f.StyleName} | вхождений: {f.OccurrenceCount} | образец: «{f.SampleText}»")
            sb.AppendLine($"  Шрифт: CN={f.FontNameCN} EN={f.FontNameEN} размер={f.FontSize}pt Bold={f.Bold} Italic={f.Italic}")
            sb.AppendLine($"  Абзац: выравнивание={f.AlignmentStr} первая строка={f.FirstLineIndentCm}см межстрочный={f.LineSpacingPt}pt до={f.SpaceBeforePt}pt после={f.SpaceAfterPt}pt")
        Next

        sb.AppendLine()
        sb.AppendLine("【Требования】")
        sb.AppendLine("1. Сопоставь каждый формат наиболее подходящему семантическому тегу (body.normal обязателен)")
        sb.AppendLine("2. Верни только следующую JSON-структуру, без пояснений:")
        sb.AppendLine("{""name"":""Клонированный формат"",""semanticTags"":[{""tagId"":""title.1"",""font"":{""fontNameCN"":""..."",")
        sb.AppendLine("""fontNameEN"":""..."",""fontSize"":16,""bold"":true},""paragraph"":{""alignment"":""center""}}]}")
        sb.AppendLine("Формат полей полностью совпадает с выводом StyleGuideConverter (используется поле semanticTags).")

        Return sb.ToString()
    End Function

End Class
