' PowerPointAi/PowerPointChartSummary.vb
' Чтение данных диаграмм для контекста запроса и наблюдений.

Imports System.Text

''' <summary>
''' Компактное текстовое описание диаграммы: тип, категории и значения первого ряда.
''' Модель должна видеть реальные подписи слайда, а не просить их у пользователя.
''' </summary>
Friend Module PowerPointChartSummary

    Friend Function DescribeChart(chart As Object, Optional index As Integer = 1) As String
        Dim sb As New StringBuilder()
        Try
            Try
                sb.AppendLine($"Диаграмма {index} (тип {chart.ChartType}):")
            Catch
                sb.AppendLine($"Диаграмма {index}:")
            End Try

            Dim series = chart.SeriesCollection
            Dim seriesCount As Integer = CInt(series.Count)
            If seriesCount = 0 Then
                sb.AppendLine("  нет рядов данных")
                Return sb.ToString().TrimEnd()
            End If

            Dim first = series(1)
            Try
                Dim name = Convert.ToString(first.Name)
                If Not String.IsNullOrWhiteSpace(name) Then sb.AppendLine("  Ряд 1: " & name)
            Catch
            End Try
            Try
                Dim categories = TryCast(first.XValues, Array)
                If categories IsNot Nothing AndAlso categories.Length > 0 Then
                    sb.AppendLine($"  Категории ({categories.Length}): " & JoinValues(categories, 40))
                End If
            Catch
            End Try
            Try
                Dim values = TryCast(first.Values, Array)
                If values IsNot Nothing AndAlso values.Length > 0 Then
                    sb.AppendLine("  Значения: " & JoinValues(values, 40))
                End If
            Catch
            End Try

            If seriesCount > 1 Then sb.AppendLine($"  Всего рядов: {seriesCount}, показан первый")
        Catch ex As Exception
            sb.AppendLine($"  Диаграмма {index}: не удалось прочитать данные ({ex.Message})")
        End Try
        Return sb.ToString().TrimEnd()
    End Function

    Friend Function JoinValues(values As Array, limit As Integer) As String
        Dim parts As New List(Of String)()
        Dim count = Math.Min(values.Length, limit)
        For index = 0 To count - 1
            Try
                parts.Add(Convert.ToString(values.GetValue(index)))
            Catch
            End Try
        Next
        If values.Length > count Then parts.Add("…")
        Return String.Join(", ", parts.ToArray())
    End Function

End Module
