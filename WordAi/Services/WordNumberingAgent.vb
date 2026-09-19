' WordAi\Services\WordNumberingAgent.vb
' Word 专属编号 Agent：处理自动编号连续化、重置为 1..n 等确定性任务。

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Word = Microsoft.Office.Interop.Word

Namespace Services

    Public Class WordNumberingResult
        Public Property Success As Boolean
        Public Property ScopeSummary As String = ""
        Public Property DetectedCount As Integer
        Public Property AppliedCount As Integer
        Public Property ObservedPreview As New List(Of String)()
        Public Property ErrorMessage As String = ""

        Public Function ToHumanReadableSummary() As String
            If Success Then
                Dim parts As New List(Of String) From {
                    $"Диапазон: {ScopeSummary}",
                    $"Автонумерованных абзацев перестроено в непрерывную последовательность 1,2,3...: {AppliedCount}"
                }

                If ObservedPreview IsNot Nothing AndAlso ObservedPreview.Count > 0 Then
                    parts.Add("Наблюдение: " & String.Join(", ", ObservedPreview.Take(5)))
                End If

                Return String.Join("; ", parts)
            End If

            If Not String.IsNullOrWhiteSpace(ErrorMessage) Then Return ErrorMessage
            Return "Не найдено автонумерованных абзацев для перестройки"
        End Function
    End Class

    Public Class WordNumberingAgent

        Private ReadOnly _app As Object

        Public Sub New(app As Object)
            _app = app
        End Sub

        Public Shared Function LooksLikeSequentialNumberingCommand(message As String) As Boolean
            If String.IsNullOrWhiteSpace(message) Then Return False

            Dim text = message.Trim()
            Dim hasNumberTopic = ContainsAny(text, {"序号", "编号", "自动编号", "列表编号", "前面的号", "前面的数字"})
            If Not hasNumberTopic Then Return False

            Dim hasSequenceIntent = ContainsAny(text, {
                "12345", "1 2 3 4 5", "1,2,3", "1，2，3",
                "连续", "递增", "重排", "重置", "重新编号", "顺序",
                "改为", "改成", "变成", "整理", "修正", "统一"
            })

            Return hasSequenceIntent
        End Function

        Public Function RebuildSequentialNumbering(message As String) As WordNumberingResult
            Dim result As New WordNumberingResult()

            Try
                If _app Is Nothing OrElse _app.ActiveDocument Is Nothing Then
                    result.ErrorMessage = "Нет документа Word для обработки"
                    Return result
                End If

                Dim scope = ResolveScope(message)
                result.ScopeSummary = If(IsExplicitSelectionScope(message) AndAlso HasUsableSelection(), "Текущее выделение", "Весь документ")

                Dim targets = CollectNumberedParagraphs(scope)
                result.DetectedCount = targets.Count
                If targets.Count = 0 Then
                    result.ErrorMessage = $"{result.ScopeSummary}: автонумерованные абзацы Word не обнаружены"
                    Return result
                End If

                Dim undoStarted As Boolean = False
                Try
                    If _app.UndoRecord IsNot Nothing Then
                        _app.UndoRecord.StartCustomRecord("AI-перестановка автонумерации")
                        undoStarted = True
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[WordNumberingAgent] UndoRecord start failed: {ex.Message}")
                End Try

                Try
                    result.AppliedCount = ApplyContinuousNumbering(targets)
                    ObserveNumbering(targets, result)
                    result.Success = result.AppliedCount > 0
                    If Not result.Success AndAlso String.IsNullOrWhiteSpace(result.ErrorMessage) Then
                        result.ErrorMessage = "Автонумерованные абзацы обнаружены, но применить непрерывную нумерацию не удалось"
                    End If
                Finally
                    If undoStarted Then
                        Try
                            _app.UndoRecord.EndCustomRecord()
                        Catch ex As Exception
                            Debug.WriteLine($"[WordNumberingAgent] UndoRecord end failed: {ex.Message}")
                        End Try
                    End If
                End Try

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = $"Не удалось перестроить автонумерацию: {ex.Message}"
                Debug.WriteLine($"[WordNumberingAgent] RebuildSequentialNumbering failed: {ex}")
            End Try

            Return result
        End Function

        Private Function ResolveScope(message As String) As Object
            If IsExplicitSelectionScope(message) AndAlso HasUsableSelection() Then
                Return _app.Selection.Range
            End If

            If HasUsableSelection() AndAlso Not IsExplicitDocumentScope(message) Then
                Return _app.Selection.Range
            End If

            Return _app.ActiveDocument.Content
        End Function

        Private Function CollectNumberedParagraphs(scope As Object) As List(Of Object)
            Dim result As New List(Of Object)()
            If scope Is Nothing Then Return result

            For Each para As Object In scope.Paragraphs
                If IsAutomaticNumberedParagraph(para) Then
                    result.Add(para)
                End If
            Next

            Return result
        End Function

        Private Function IsAutomaticNumberedParagraph(para As Object) As Boolean
            Try
                If para Is Nothing OrElse para.Range Is Nothing Then Return False
                Dim listType = para.Range.ListFormat.ListType
                Return listType = Word.WdListType.wdListSimpleNumbering OrElse
                       listType = Word.WdListType.wdListOutlineNumbering OrElse
                       listType = Word.WdListType.wdListMixedNumbering OrElse
                       listType = Word.WdListType.wdListListNumOnly
            Catch
                Return False
            End Try
        End Function

        Private Function ApplyContinuousNumbering(targets As List(Of Object)) As Integer
            Dim applied As Integer = 0
            If targets Is Nothing OrElse targets.Count = 0 Then Return applied

            Dim template = _app.ListGalleries(Word.WdListGalleryType.wdNumberGallery).ListTemplates(1)

            For i = 0 To targets.Count - 1
                Dim para = targets(i)
                Try
                    Dim rng = para.Range
                    rng.ListFormat.RemoveNumbers(Word.WdNumberType.wdNumberParagraph)
                    rng.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate:=template,
                        ContinuePreviousList:=(i > 0),
                        ApplyTo:=Word.WdListApplyTo.wdListApplyToSelection,
                        DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior,
                        ApplyLevel:=1)
                    applied += 1
                Catch ex As Exception
                    Debug.WriteLine($"[WordNumberingAgent] paragraph {i} numbering apply failed: {ex.Message}")
                End Try
            Next

            Return applied
        End Function

        Private Sub ObserveNumbering(targets As List(Of Object), result As WordNumberingResult)
            If targets Is Nothing OrElse result Is Nothing Then Return

            For i = 0 To Math.Min(targets.Count, 5) - 1
                Try
                    Dim listString = targets(i).Range.ListFormat.ListString
                    Dim text = CleanParagraphText(targets(i).Range.Text)
                    result.ObservedPreview.Add($"{listString} {text}")
                Catch ex As Exception
                    Debug.WriteLine($"[WordNumberingAgent] observe paragraph {i} failed: {ex.Message}")
                End Try
            Next
        End Sub

        Private Shared Function CleanParagraphText(text As String) As String
            If text Is Nothing Then Return ""
            Dim cleaned = text.Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
            If cleaned.Length > 24 Then cleaned = cleaned.Substring(0, 24) & "..."
            Return cleaned
        End Function

        Private Function HasUsableSelection() As Boolean
            Try
                If _app Is Nothing OrElse _app.Selection Is Nothing OrElse _app.Selection.Range Is Nothing Then Return False
                Dim selectedText = CleanParagraphText(_app.Selection.Range.Text)
                Return selectedText.Length > 0
            Catch
                Return False
            End Try
        End Function

        Private Shared Function IsExplicitDocumentScope(message As String) As Boolean
            Return ContainsAny(If(message, ""), {"全文", "整篇", "整个文档", "全部", "所有"})
        End Function

        Private Shared Function IsExplicitSelectionScope(message As String) As Boolean
            Return ContainsAny(If(message, ""), {"选中", "选区", "所选", "当前选择"})
        End Function

        Private Shared Function ContainsAny(text As String, words As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(text) OrElse words Is Nothing Then Return False
            Return words.Any(Function(word) text.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0)
        End Function

    End Class

End Namespace
