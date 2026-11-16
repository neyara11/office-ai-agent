' PowerPointAi\Context\PowerPointDeckReader.vb
' Reads a compact, token-budgeted text map of the active presentation.
' COM stays in the PowerPoint host project; the shared Harness only receives text.

Imports System.Diagnostics
Imports System.Text
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports PowerPointAi.Extensions
Imports ShareRibbon

Namespace Context
    Public NotInheritable Class PowerPointDeckReader
        Private Const DefaultBudgetChars As Integer = 12000
        Private Const CacheSeconds As Integer = 2

        Private Shared ReadOnly _cacheLock As New Object()
        Private Shared _cacheKey As String = ""
        Private Shared _cacheText As String = ""
        Private Shared _cacheAt As DateTime = DateTime.MinValue

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Returns the current slide in full and a compact map of the remaining slides,
        ''' capped by budgetChars. Never throws to the caller.
        ''' </summary>
        Public Shared Function GetDeckText(app As PowerPoint.Application,
                                           Optional budgetChars As Integer = DefaultBudgetChars) As String
            If app Is Nothing Then Return ""
            Dim presentation As PowerPoint.Presentation = Nothing
            Try
                presentation = app.ActivePresentation
                If presentation Is Nothing Then Return ""

                Dim currentIndex = GetCurrentSlideIndex(app)
                Dim key = $"{presentation.FullName}|{presentation.Slides.Count}|{currentIndex}|{budgetChars}"

                SyncLock _cacheLock
                    If key = _cacheKey AndAlso (DateTime.UtcNow - _cacheAt).TotalSeconds < CacheSeconds Then
                        Return _cacheText
                    End If
                End SyncLock

                Dim text = BuildDeckText(presentation, currentIndex, Math.Max(2000, budgetChars))

                SyncLock _cacheLock
                    _cacheKey = key
                    _cacheText = text
                    _cacheAt = DateTime.UtcNow
                End SyncLock
                Return text
            Catch ex As Exception
                Debug.WriteLine("PowerPointDeckReader.GetDeckText: " & ex.Message)
                Return ""
            Finally
                If presentation IsNot Nothing Then ComObjectHelper.ReleaseComObject(presentation)
            End Try
        End Function

        ''' <summary>Invalidate the cache after the agent changes the deck.</summary>
        Public Shared Sub Invalidate()
            SyncLock _cacheLock
                _cacheKey = ""
                _cacheText = ""
                _cacheAt = DateTime.MinValue
            End SyncLock
        End Sub

        Public Shared Function GetCurrentSlideIndex(app As PowerPoint.Application) As Integer
            Try
                If app Is Nothing OrElse app.ActiveWindow Is Nothing OrElse app.ActiveWindow.View Is Nothing Then Return 0
                Dim currentSlide As Object = app.ActiveWindow.View.Slide
                If currentSlide Is Nothing Then Return 0
                Return CInt(CallByName(currentSlide, "SlideIndex", CallType.Get))
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function BuildDeckText(presentation As PowerPoint.Presentation,
                                              currentIndex As Integer,
                                              budgetChars As Integer) As String
            Dim sb As New StringBuilder()
            Dim slides As PowerPoint.Slides = Nothing
            Try
                slides = presentation.Slides
                sb.AppendLine($"Презентация: {IO.Path.GetFileName(presentation.FullName)}; слайдов: {slides.Count}; текущий слайд: {currentIndex}")
                Dim budget = budgetChars - sb.Length

                If currentIndex >= 1 AndAlso currentIndex <= slides.Count Then
                    Dim currentText = ReadSlide(slides(currentIndex), currentIndex, True)
                    AppendBlock(sb, currentText, budget)
                    budget = budgetChars - sb.Length
                End If

                For index = 1 To slides.Count
                    If index = currentIndex Then Continue For
                    If budget <= 200 Then
                        sb.AppendLine("... (остальные слайды не показаны из-за лимита контекста)")
                        Exit For
                    End If
                    Dim block = ReadSlide(slides(index), index, False)
                    AppendBlock(sb, block, budget)
                    budget = budgetChars - sb.Length
                Next
            Finally
                If slides IsNot Nothing Then ComObjectHelper.ReleaseComObject(slides)
            End Try
            Return sb.ToString()
        End Function

        Private Shared Sub AppendBlock(sb As StringBuilder, block As String, remaining As Integer)
            If String.IsNullOrWhiteSpace(block) Then Return
            If remaining <= 0 Then
                sb.AppendLine("... (лимит контекста исчерпан)")
                Return
            End If
            If block.Length <= remaining Then
                sb.AppendLine(block.TrimEnd())
            Else
                sb.AppendLine(block.Substring(0, Math.Max(0, remaining)).TrimEnd())
                sb.AppendLine("... (слайд обрезан из-за лимита контекста)")
            End If
        End Sub

        Private Shared Function ReadSlide(slide As PowerPoint.Slide, index As Integer, full As Boolean) As String
            Dim sb As New StringBuilder()
            Dim shapes As PowerPoint.Shapes = Nothing
            Dim title As String = ""
            Dim body As New List(Of String)()
            Dim tables As New List(Of String)()
            Dim charts As New List(Of String)()
            Dim chartCount As Integer = 0

            Try
                shapes = slide.Shapes
                For shapeIndex = 1 To shapes.Count
                    Dim shape As PowerPoint.Shape = Nothing
                    Dim textRange As PowerPoint.TextRange = Nothing
                    Dim table As PowerPoint.Table = Nothing
                    Dim chart As PowerPoint.Chart = Nothing
                    Try
                        shape = shapes(shapeIndex)
                        If shape.HasChart = MsoTriState.msoTrue Then
                            chart = shape.Chart
                            chartCount += 1
                            charts.Add("  " & PowerPointChartSummary.DescribeChart(chart, chartCount).Replace(vbLf, vbLf & "  ").TrimEnd())
                            Continue For
                        End If

                        If shape.HasTable = MsoTriState.msoTrue Then
                            table = shape.Table
                            tables.Add(ReadTable(table, full))
                            Continue For
                        End If

                        If shape.HasTextFrame = MsoTriState.msoTrue AndAlso
                           shape.TextFrame.HasText = MsoTriState.msoTrue Then
                            textRange = shape.TextFrame.TextRange
                            Dim text = NormalizeText(textRange.Text)
                            If text.Length > 0 Then
                                If title.Length = 0 AndAlso IsTitlePlaceholder(shape) Then
                                    title = text
                                Else
                                    body.Add(text)
                                End If
                            End If
                        End If
                    Catch shapeEx As Exception
                        Debug.WriteLine($"PowerPointDeckReader shape skipped: {shapeEx.Message}")
                    Finally
                        If chart IsNot Nothing Then ComObjectHelper.ReleaseComObject(chart)
                        If table IsNot Nothing Then ComObjectHelper.ReleaseComObject(table)
                        If textRange IsNot Nothing Then ComObjectHelper.ReleaseComObject(textRange)
                        If shape IsNot Nothing Then ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                If shapes IsNot Nothing Then ComObjectHelper.ReleaseComObject(shapes)
            End Try

            If title.Length = 0 AndAlso body.Count > 0 Then
                title = body(0)
                body.RemoveAt(0)
            End If

            sb.Append($"=== Слайд {index} ===")
            If title.Length > 0 Then sb.Append(" " & Truncate(title, If(full, 300, 120)))
            sb.AppendLine()

            If full Then
                For Each part In body
                    sb.AppendLine("  " & Truncate(part, 2000))
                Next
                For Each tableText In tables
                    sb.AppendLine(tableText)
                Next
                For Each chartText In charts
                    sb.AppendLine(chartText)
                Next
                AppendNotes(sb, slide, full:=True)
            Else
                If body.Count > 0 Then
                    Dim joined = String.Join(" / ", body)
                    sb.AppendLine("  " & Truncate(joined, 220))
                End If
                For Each tableText In tables
                    sb.AppendLine("  " & FirstLine(tableText))
                Next
                For Each chartText In charts
                    sb.AppendLine("  " & FirstLine(chartText))
                Next
                AppendNotes(sb, slide, full:=False)
            End If

            Return sb.ToString()
        End Function

        Private Shared Function ReadTable(table As PowerPoint.Table, full As Boolean) As String
            Dim sb As New StringBuilder()
            Try
                Dim rowLimit = If(full, Math.Min(table.Rows.Count, 15), Math.Min(table.Rows.Count, 4))
                Dim colLimit = Math.Min(table.Columns.Count, If(full, 8, 5))
                sb.AppendLine($"  [Таблица {table.Rows.Count}×{table.Columns.Count}]")
                For rowIndex = 1 To rowLimit
                    Dim cells As New List(Of String)()
                    For colIndex = 1 To colLimit
                        Dim cell As PowerPoint.Cell = Nothing
                        Dim cellShape As PowerPoint.Shape = Nothing
                        Dim cellText As PowerPoint.TextRange = Nothing
                        Try
                            cell = table.Cell(rowIndex, colIndex)
                            cellShape = cell.Shape
                            If cellShape.HasTextFrame = MsoTriState.msoTrue AndAlso
                               cellShape.TextFrame.HasText = MsoTriState.msoTrue Then
                                cellText = cellShape.TextFrame.TextRange
                                cells.Add(Truncate(NormalizeText(cellText.Text), If(full, 80, 30)))
                            Else
                                cells.Add("")
                            End If
                        Catch
                            cells.Add("")
                        Finally
                            If cellText IsNot Nothing Then ComObjectHelper.ReleaseComObject(cellText)
                            If cellShape IsNot Nothing Then ComObjectHelper.ReleaseComObject(cellShape)
                            If cell IsNot Nothing Then ComObjectHelper.ReleaseComObject(cell)
                        End Try
                    Next
                    sb.AppendLine("    " & String.Join(" | ", cells))
                Next
                If table.Rows.Count > rowLimit Then sb.AppendLine($"    ... ещё {table.Rows.Count - rowLimit} строк")
            Catch ex As Exception
                Debug.WriteLine($"PowerPointDeckReader table: {ex.Message}")
            End Try
            Return sb.ToString()
        End Function

        Private Shared Sub AppendNotes(sb As StringBuilder, slide As PowerPoint.Slide, full As Boolean)
            Dim notesPage As PowerPoint.Slide = Nothing
            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                notesPage = slide.NotesPage
                If notesPage Is Nothing Then Return
                shapes = notesPage.Shapes
                Dim best As String = ""
                For shapeIndex = 1 To shapes.Count
                    Dim shape As PowerPoint.Shape = Nothing
                    Dim textRange As PowerPoint.TextRange = Nothing
                    Try
                        shape = shapes(shapeIndex)
                        If shape.HasTextFrame = MsoTriState.msoTrue AndAlso
                           shape.TextFrame.HasText = MsoTriState.msoTrue Then
                            textRange = shape.TextFrame.TextRange
                            Dim text = NormalizeText(textRange.Text)
                            If text.Length > best.Length Then best = text
                        End If
                    Catch
                    Finally
                        If textRange IsNot Nothing Then ComObjectHelper.ReleaseComObject(textRange)
                        If shape IsNot Nothing Then ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
                If best.Length > 0 AndAlso Not IsSlideNumberOnly(best) Then
                    sb.AppendLine("  [Заметки] " & Truncate(best, If(full, 1500, 200)))
                End If
            Catch ex As Exception
                Debug.WriteLine($"PowerPointDeckReader notes: {ex.Message}")
            Finally
                If shapes IsNot Nothing Then ComObjectHelper.ReleaseComObject(shapes)
                If notesPage IsNot Nothing Then ComObjectHelper.ReleaseComObject(notesPage)
            End Try
        End Sub

        Private Shared Function IsTitlePlaceholder(shape As PowerPoint.Shape) As Boolean
            Try
                If shape.PlaceholderFormat Is Nothing Then Return False
                Select Case shape.PlaceholderFormat.Type
                    Case PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                         PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        Return True
                    Case Else
                        Return False
                End Select
            Catch
                Return False
            End Try
        End Function

        Private Shared Function IsSlideNumberOnly(text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return True
            Dim trimmed = text.Trim()
            If trimmed.Length > 6 Then Return False
            Dim number As Integer
            Return Integer.TryParse(trimmed, number)
        End Function

        Private Shared Function NormalizeText(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            Dim text = value.Replace(vbCrLf, " / ")
            text = text.Replace(vbCr, " / ")
            text = text.Replace(vbLf, " / ")
            text = text.Replace(vbTab, " ")
            While text.Contains("  ")
                text = text.Replace("  ", " ")
            End While
            Return text.Trim()
        End Function

        Private Shared Function Truncate(value As String, maxLength As Integer) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If maxLength <= 0 OrElse value.Length <= maxLength Then Return value
            Return value.Substring(0, maxLength).TrimEnd() & "..."
        End Function

        Private Shared Function FirstLine(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim lines = value.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
            For Each line In lines
                Dim trimmed = line.Trim()
                If trimmed.Length > 0 Then Return trimmed
            Next
            Return ""
        End Function
    End Class
End Namespace
