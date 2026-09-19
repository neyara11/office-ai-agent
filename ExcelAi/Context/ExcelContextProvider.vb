Imports Excel = Microsoft.Office.Interop.Excel
Imports ShareRibbon.Agent.Context
Imports System.Diagnostics
Imports System.Text

Namespace Context
    Public Class ExcelContextProvider
        Implements IContextProvider

        Private ReadOnly _app As Excel.Application

        Public Sub New(app As Excel.Application)
            _app = app
        End Sub

        Public Function GetContext() As OfficeContext Implements IContextProvider.GetContext
            Dim ctx As New OfficeContext With {.AppType = "Excel"}

            Try
                Dim workbookName As String = ""
                Try
                    If _app.ActiveWorkbook IsNot Nothing Then workbookName = _app.ActiveWorkbook.Name
                Catch
                End Try

                Dim worksheet As Excel.Worksheet = TryCast(_app.ActiveSheet, Excel.Worksheet)
                Dim usedRange As Excel.Range = Nothing
                If worksheet IsNot Nothing Then usedRange = worksheet.UsedRange

                Dim selectedRange As Excel.Range = TryCast(_app.Selection, Excel.Range)
                Dim primaryRange As Excel.Range = If(selectedRange, usedRange)

                If primaryRange IsNot Nothing Then
                    Dim preview = BuildRangePreview(primaryRange, 8, 8)
                    Dim profile = BuildRangeProfile(primaryRange)
                    ctx.Selection = New SelectionInfo With {
                        .Address = GetRangeAddress(primaryRange),
                        .ItemCount = SafeCount(primaryRange),
                        .DataType = InferRangeDataType(primaryRange),
                        .Preview = preview & vbCrLf & profile
                    }
                End If

                Dim summary As New StringBuilder()
                summary.AppendLine($"Рабочая книга: {If(String.IsNullOrWhiteSpace(workbookName), "(не сохранена или неизвестна)", workbookName)}")
                If worksheet IsNot Nothing Then summary.AppendLine($"Текущий лист: {worksheet.Name}")
                If usedRange IsNot Nothing Then
                    summary.AppendLine($"Используемый диапазон: {GetRangeAddress(usedRange)}")
                    summary.AppendLine($"Размер используемого диапазона: {usedRange.Rows.Count} строк x {usedRange.Columns.Count} столбцов")
                    summary.AppendLine(BuildHeaderSummary(usedRange))
                    summary.AppendLine(BuildFormulaSummary(usedRange))
                    summary.AppendLine(BuildQualitySummary(usedRange))
                    summary.AppendLine($"Рекомендуемый диапазон по умолчанию: {If(selectedRange IsNot Nothing, GetRangeAddress(selectedRange), GetRangeAddress(usedRange))}")
                    summary.AppendLine($"Рекомендуемая позиция вывода: {GetSuggestedOutputCell(usedRange)}")
                End If

                ctx.DocStructure = New DocumentStructure With {
                    .Summary = summary.ToString()
                }
            Catch ex As Exception
                Debug.WriteLine("获取Excel上下文失败: " & ex.Message)
            End Try

            Return ctx
        End Function

        Private Function GetRangeAddress(range As Excel.Range) As String
            Try
                Dim ws As Excel.Worksheet = TryCast(range.Worksheet, Excel.Worksheet)
                Dim address = range.Address(False, False)
                If ws IsNot Nothing Then Return $"{ws.Name}!{address}"
                Return address
            Catch
                Return "(неизвестный диапазон)"
            End Try
        End Function

        Private Function SafeCount(range As Excel.Range) As Integer
            Try
                Dim count = CLng(range.Rows.Count) * CLng(range.Columns.Count)
                If count > Integer.MaxValue Then Return Integer.MaxValue
                Return CInt(count)
            Catch
                Return 0
            End Try
        End Function

        Private Function BuildRangePreview(range As Excel.Range, maxRows As Integer, maxCols As Integer) As String
            Dim sb As New StringBuilder()
            Try
                sb.AppendLine("Предпросмотр данных:")
                Dim rows = Math.Min(range.Rows.Count, maxRows)
                Dim cols = Math.Min(range.Columns.Count, maxCols)

                For r As Integer = 1 To rows
                    Dim values As New List(Of String)()
                    For c As Integer = 1 To cols
                        Dim valueText = GetCellText(range.Cells(r, c))
                        If valueText.Length > 40 Then valueText = valueText.Substring(0, 40) & "..."
                        values.Add(valueText)
                    Next
                    sb.AppendLine(String.Join(vbTab, values))
                Next

                If range.Rows.Count > rows OrElse range.Columns.Count > cols Then
                    sb.AppendLine($"...показаны только первые {rows} строк x {cols} столбцов")
                End If
            Catch ex As Exception
                sb.AppendLine($"Не удалось построить предпросмотр данных: {ex.Message}")
            End Try
            Return sb.ToString().TrimEnd()
        End Function

        Private Function BuildRangeProfile(range As Excel.Range) As String
            Dim sb As New StringBuilder()
            Try
                sb.AppendLine()
                sb.AppendLine("Профиль диапазона:")
                sb.AppendLine($"- Адрес: {GetRangeAddress(range)}")
                sb.AppendLine($"- Размер: {range.Rows.Count} строк x {range.Columns.Count} столбцов")
                sb.AppendLine($"- Тип: {InferRangeDataType(range)}")

                Dim headers = GetHeaders(range, Math.Min(range.Columns.Count, 12))
                If headers.Count > 0 Then sb.AppendLine($"- Заголовки: {String.Join(", ", headers)}")
                sb.AppendLine(BuildColumnTypeSummary(range, Math.Min(range.Columns.Count, 8), Math.Min(range.Rows.Count, 30)))
            Catch ex As Exception
                sb.AppendLine($"- Не удалось построить профиль диапазона: {ex.Message}")
            End Try
            Return sb.ToString().TrimEnd()
        End Function

        Private Function BuildHeaderSummary(range As Excel.Range) As String
            Dim headers = GetHeaders(range, Math.Min(range.Columns.Count, 20))
            If headers.Count = 0 Then Return "Заголовки: не распознаны"
            Return "Заголовки: " & String.Join(", ", headers)
        End Function

        Private Function GetHeaders(range As Excel.Range, maxCols As Integer) As List(Of String)
            Dim headers As New List(Of String)()
            Try
                If range.Rows.Count < 1 Then Return headers
                For c As Integer = 1 To Math.Min(range.Columns.Count, maxCols)
                    Dim text = GetCellText(range.Cells(1, c)).Trim()
                    If Not String.IsNullOrWhiteSpace(text) Then
                        headers.Add($"{ColumnLetter(range.Column + c - 1)}={text}")
                    End If
                Next
            Catch
            End Try
            Return headers
        End Function

        Private Function BuildColumnTypeSummary(range As Excel.Range, maxCols As Integer, maxRows As Integer) As String
            Dim parts As New List(Of String)()
            Try
                For c As Integer = 1 To Math.Min(range.Columns.Count, maxCols)
                    Dim numericCount As Integer = 0
                    Dim textCount As Integer = 0
                    Dim dateCount As Integer = 0
                    Dim formulaCount As Integer = 0
                    Dim sampleCount As Integer = 0

                    For r As Integer = 2 To Math.Min(range.Rows.Count, maxRows)
                        Dim cell As Excel.Range = Nothing
                        Try
                            cell = range.Cells(r, c)
                            Dim raw = cell.Value2
                            If raw Is Nothing Then Continue For
                            sampleCount += 1
                            If cell.HasFormula Then formulaCount += 1
                            Dim text = raw.ToString()
                            Dim number As Double
                            If Double.TryParse(text, number) Then
                                numericCount += 1
                            ElseIf IsDate(text) Then
                                dateCount += 1
                            Else
                                textCount += 1
                            End If
                        Catch
                        End Try
                    Next

                    Dim kind = "empty"
                    If sampleCount > 0 Then
                        If numericCount >= textCount AndAlso numericCount >= dateCount Then
                            kind = "numeric"
                        ElseIf dateCount >= numericCount AndAlso dateCount >= textCount Then
                            kind = "date"
                        Else
                            kind = "text"
                        End If
                    End If
                    If formulaCount > 0 Then kind &= $"+formula({formulaCount})"
                    parts.Add($"{ColumnLetter(range.Column + c - 1)}:{kind}")
                Next
            Catch
            End Try

            If parts.Count = 0 Then Return "- Типы столбцов: не распознаны"
            Return "- Типы столбцов: " & String.Join(", ", parts)
        End Function

        Private Function BuildFormulaSummary(range As Excel.Range) As String
            Try
                Dim formulaCells As Integer = 0
                Dim errorCells As Integer = 0
                Dim rows = Math.Min(range.Rows.Count, 80)
                Dim cols = Math.Min(range.Columns.Count, 40)

                For r As Integer = 1 To rows
                    For c As Integer = 1 To cols
                        Dim cell As Excel.Range = Nothing
                        Try
                            cell = range.Cells(r, c)
                            If cell.HasFormula Then formulaCells += 1
                            Dim text = GetCellText(cell)
                            If text.StartsWith("#") Then errorCells += 1
                        Catch
                        End Try
                    Next
                Next

                Return $"Формулы/ошибки: ячеек с формулами {formulaCells}, предположительно ошибочных значений {errorCells}"
            Catch ex As Exception
                Return $"Формулы/ошибки: не удалось подсчитать ({ex.Message})"
            End Try
        End Function

        Private Function BuildQualitySummary(range As Excel.Range) As String
            Try
                Dim blankCount As Integer = 0
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim duplicateRows As Integer = 0
                Dim rows = Math.Min(range.Rows.Count, 120)
                Dim cols = Math.Min(range.Columns.Count, 30)

                For r As Integer = 1 To rows
                    Dim rowValues As New List(Of String)()
                    For c As Integer = 1 To cols
                        Dim text = GetCellText(range.Cells(r, c)).Trim()
                        If String.IsNullOrWhiteSpace(text) Then blankCount += 1
                        rowValues.Add(text)
                    Next

                    Dim key = String.Join("|", rowValues)
                    If Not String.IsNullOrWhiteSpace(key.Replace("|", "")) Then
                        If seen.Contains(key) Then duplicateRows += 1 Else seen.Add(key)
                    End If
                Next

                Return $"Качество данных: пустых ячеек в выборке {blankCount}, предположительно дублирующихся строк {duplicateRows}"
            Catch ex As Exception
                Return $"Качество данных: не удалось подсчитать ({ex.Message})"
            End Try
        End Function

        Private Function InferRangeDataType(range As Excel.Range) As String
            Try
                If range.Rows.Count >= 2 AndAlso range.Columns.Count >= 2 Then
                    Dim headers = GetHeaders(range, Math.Min(range.Columns.Count, 8))
                    If headers.Count >= Math.Min(range.Columns.Count, 2) Then Return "Таблица (вероятно, с заголовками)"
                    Return "Таблица"
                End If
                If range.Rows.Count = 1 AndAlso range.Columns.Count > 1 Then Return "Данные строки"
                If range.Columns.Count = 1 AndAlso range.Rows.Count > 1 Then Return "Данные столбца"
                Return "Ячейка"
            Catch
                Return "Данные Excel"
            End Try
        End Function

        Private Function GetSuggestedOutputCell(usedRange As Excel.Range) As String
            Try
                Dim ws As Excel.Worksheet = TryCast(usedRange.Worksheet, Excel.Worksheet)
                Dim row = usedRange.Row
                Dim col = usedRange.Column + usedRange.Columns.Count + 2
                Dim addr = ColumnLetter(col) & row.ToString()
                If ws IsNot Nothing Then Return $"{ws.Name}!{addr}"
                Return addr
            Catch
                Return "Пустая область справа от текущего листа"
            End Try
        End Function

        Private Function GetCellText(cell As Excel.Range) As String
            Try
                Dim value = cell.Text
                If value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(value.ToString()) Then Return value.ToString()
                value = cell.Value2
                If value IsNot Nothing Then Return value.ToString()
            Catch
            End Try
            Return ""
        End Function

        Private Function ColumnLetter(columnNumber As Integer) As String
            Dim dividend = columnNumber
            Dim columnName As String = ""
            While dividend > 0
                Dim modulo = (dividend - 1) Mod 26
                columnName = Chr(65 + modulo) & columnName
                dividend = CInt((dividend - modulo) / 26)
            End While
            Return columnName
        End Function
    End Class
End Namespace
