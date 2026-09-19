' ShareRibbon\Controls\Services\ExecutionPlanRenderer.vb
' 执行计划渲染服务：将JSON命令转换为用户友好的执行步骤

Imports Newtonsoft.Json.Linq

''' <summary>
''' 执行计划渲染服务
''' 将大模型返回的JSON命令转换为用户可理解的执行步骤
''' </summary>
Public Class ExecutionPlanRenderer

#Region "命令描述映射"

    ' 操作类型中文描述
    Private Shared ReadOnly OperationDescriptions As New Dictionary(Of String, String) From {
        {"removeDuplicates", "Удаление дубликатов"},
        {"fillEmpty", "Заполнение пустых значений"},
        {"trim", "Удаление пробелов"},
        {"replace", "Замена содержимого"},
        {"transpose", "Транспонирование данных"},
        {"split", "Разделение столбцов"},
        {"merge", "Объединение столбцов"},
        {"summary", "Создание сводки"},
        {"pivot", "Создание сводной таблицы"},
        {"groupby", "Группировка и итоги"},
        {"ranking", "Анализ рейтинга"}
    }

    ' 图表类型中文描述
    Private Shared ReadOnly ChartTypeDescriptions As New Dictionary(Of String, String) From {
        {"Column", "Гистограмма"},
        {"Line", "Линейчатая диаграмма"},
        {"Pie", "Круговая диаграмма"},
        {"Bar", "Полосчатая диаграмма"},
        {"Scatter", "Точечная диаграмма"},
        {"Area", "Диаграмма с областями"}
    }

#End Region

#Region "公共方法"

    ''' <summary>
    ''' 将JSON命令解析为执行计划
    ''' </summary>
    Public Function ParseJsonToExecutionPlan(jsonCommand As String) As List(Of ExecutionStep)
        Dim plan As New List(Of ExecutionStep)()

        Try
            Dim json = JObject.Parse(jsonCommand)
            Dim command = json("command")?.ToString()
            Dim params = json("params")

            If String.IsNullOrEmpty(command) Then
                Return plan
            End If

            ' 根据命令类型生成步骤
            Select Case command.ToLower()
                Case "applyformula", "formula", "calculate"
                    plan.AddRange(GenerateFormulaSteps(params))
                Case "writedata", "write", "setvalue"
                    plan.AddRange(GenerateWriteDataSteps(params))
                Case "formatrange", "format", "style"
                    plan.AddRange(GenerateFormatSteps(params))
                Case "createchart", "chart"
                    plan.AddRange(GenerateChartSteps(params))
                Case "cleandata", "clean"
                    plan.AddRange(GenerateCleanDataSteps(params))
                Case "dataanalysis", "analyze"
                    plan.AddRange(GenerateAnalysisSteps(params))
                Case "transformdata", "transform"
                    plan.AddRange(GenerateTransformSteps(params))
                Case "generatereport", "report"
                    plan.AddRange(GenerateReportSteps(params))
                Case Else
                    ' 通用处理
                    plan.Add(New ExecutionStep(1, $"Выполнить команду {command}", "default"))
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ParseJsonToExecutionPlan 出错: {ex.Message}")
            plan.Add(New ExecutionStep(1, "Не удалось разобрать команду", "default"))
        End Try

        Return plan
    End Function

#End Region

#Region "步骤生成方法"

    ''' <summary>
    ''' 生成公式应用步骤
    ''' </summary>
    Private Function GenerateFormulaSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim targetRange = If(params?("targetRange")?.ToString(), "Целевой диапазон")
        Dim formula = If(params?("formula")?.ToString(), "")
        Dim fillDown = If(params?("fillDown")?.Value(Of Boolean)(), False)

        steps.Add(New ExecutionStep(1, $"Применить формулу к {targetRange}", "formula") With {
            .WillModify = targetRange,
            .EstimatedTime = "1 сек"
        })

        If Not String.IsNullOrEmpty(formula) Then
            Dim formulaDesc = GetFormulaDescription(formula)
            steps.Add(New ExecutionStep(2, $"Содержимое формулы: {formulaDesc}", "formula"))
        End If

        If fillDown Then
            steps.Add(New ExecutionStep(3, "Автоматически заполнить формулу вниз", "formula"))
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据写入步骤
    ''' </summary>
    Private Function GenerateWriteDataSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim targetRange = If(params?("targetRange")?.ToString(), "Целевой диапазон")
        
        steps.Add(New ExecutionStep(1, $"Записать данные в {targetRange}", "data") With {
            .WillModify = targetRange,
            .EstimatedTime = "1 сек"
        })

        Return steps
    End Function

    ''' <summary>
    ''' 生成格式化步骤
    ''' </summary>
    Private Function GenerateFormatSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim range = If(params?("range")?.ToString(), If(params?("targetRange")?.ToString(), "Целевой диапазон"))
        Dim style = If(params?("style")?.ToString(), "")
        
        steps.Add(New ExecutionStep(1, $"Выбрать диапазон {range}", "search") With {
            .EstimatedTime = "1 сек"
        })

        Dim formatDesc = "Применить форматирование"
        If Not String.IsNullOrEmpty(style) Then
            formatDesc = $"Применить стиль {style}"
        End If

        Dim formatDetails As New List(Of String)()
        If params?("bold")?.Value(Of Boolean)() = True Then formatDetails.Add("Полужирный")
        If params?("italic")?.Value(Of Boolean)() = True Then formatDetails.Add("Курсив")
        If params?("borders")?.Value(Of Boolean)() = True Then formatDetails.Add("Границы")
        
        If formatDetails.Count > 0 Then
            formatDesc &= $" ({String.Join(", ", formatDetails)})"
        End If

        steps.Add(New ExecutionStep(2, formatDesc, "format") With {
            .WillModify = range,
            .EstimatedTime = "1 сек"
        })

        Return steps
    End Function

    ''' <summary>
    ''' 生成图表创建步骤
    ''' </summary>
    Private Function GenerateChartSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim chartType = If(params?("type")?.ToString(), "Column")
        Dim dataRange = If(params?("dataRange")?.ToString(), "Диапазон данных")
        Dim title = If(params?("title")?.ToString(), "")
        Dim position = If(params?("position")?.ToString(), "")

        Dim chartTypeName = If(ChartTypeDescriptions.ContainsKey(chartType), ChartTypeDescriptions(chartType), chartType)

        steps.Add(New ExecutionStep(1, $"Прочитать {dataRange} как источник данных диаграммы", "search") With {
            .EstimatedTime = "1 сек"
        })

        steps.Add(New ExecutionStep(2, $"Создать {chartTypeName}", "chart") With {
            .EstimatedTime = "2 сек"
        })

        If Not String.IsNullOrEmpty(title) Then
            steps.Add(New ExecutionStep(3, $"Задать заголовок диаграммы: {title}", "chart"))
        End If

        If Not String.IsNullOrEmpty(position) Then
            steps.Add(New ExecutionStep(4, $"Разместить диаграмму в {position}", "chart") With {
                .WillModify = position
            })
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据清洗步骤
    ''' </summary>
    Private Function GenerateCleanDataSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim operation = If(params?("operation")?.ToString(), "clean")
        Dim range = If(params?("range")?.ToString(), "Диапазон данных")

        Dim operationDesc = If(OperationDescriptions.ContainsKey(operation), OperationDescriptions(operation), operation)

        steps.Add(New ExecutionStep(1, $"Сканировать диапазон {range}", "search") With {
            .EstimatedTime = "1 сек"
        })

        steps.Add(New ExecutionStep(2, $"Выполнить очистку: {operationDesc}", "clean") With {
            .WillModify = range,
            .EstimatedTime = "2 сек"
        })

        steps.Add(New ExecutionStep(3, "Проверить результат очистки", "data"))

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据分析步骤
    ''' </summary>
    Private Function GenerateAnalysisSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim analysisType = If(params?("type")?.ToString(), "summary")
        Dim sourceRange = If(params?("sourceRange")?.ToString(), "Диапазон данных")
        Dim targetRange = If(params?("targetRange")?.ToString(), "")

        Dim analysisDesc = If(OperationDescriptions.ContainsKey(analysisType), OperationDescriptions(analysisType), analysisType)

        steps.Add(New ExecutionStep(1, $"Прочитать данные {sourceRange}", "search") With {
            .EstimatedTime = "1 сек"
        })

        steps.Add(New ExecutionStep(2, $"Выполнить анализ: {analysisDesc}", "data") With {
            .EstimatedTime = "3 сек"
        })

        If Not String.IsNullOrEmpty(targetRange) Then
            steps.Add(New ExecutionStep(3, $"Вывести результат в {targetRange}", "data") With {
                .WillModify = targetRange
            })
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成数据转换步骤
    ''' </summary>
    Private Function GenerateTransformSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim operation = If(params?("operation")?.ToString(), "transform")
        Dim sourceRange = If(params?("sourceRange")?.ToString(), "Исходный диапазон")
        Dim targetRange = If(params?("targetRange")?.ToString(), "")

        Dim operationDesc = If(OperationDescriptions.ContainsKey(operation), OperationDescriptions(operation), operation)

        steps.Add(New ExecutionStep(1, $"Прочитать данные {sourceRange}", "search"))
        steps.Add(New ExecutionStep(2, $"Выполнить преобразование: {operationDesc}", "data") With {
            .EstimatedTime = "2 сек"
        })

        If Not String.IsNullOrEmpty(targetRange) Then
            steps.Add(New ExecutionStep(3, $"Вывести в {targetRange}", "data") With {
                .WillModify = targetRange
            })
        End If

        Return steps
    End Function

    ''' <summary>
    ''' 生成报表生成步骤
    ''' </summary>
    Private Function GenerateReportSteps(params As JToken) As List(Of ExecutionStep)
        Dim steps As New List(Of ExecutionStep)()
        
        Dim sourceRange = If(params?("sourceRange")?.ToString(), "Диапазон данных")
        Dim targetSheet = If(params?("targetSheet")?.ToString(), "Новый лист")
        Dim title = If(params?("title")?.ToString(), "Отчёт")
        Dim includeChart = If(params?("includeChart")?.Value(Of Boolean)(), False)

        steps.Add(New ExecutionStep(1, $"Собрать данные {sourceRange}", "search"))
        steps.Add(New ExecutionStep(2, $"Создать лист отчёта: {targetSheet}", "data") With {
            .EstimatedTime = "1 сек"
        })
        steps.Add(New ExecutionStep(3, $"Заполнить данные и задать заголовок: {title}", "data"))
        steps.Add(New ExecutionStep(4, "Применить формат отчёта", "format") With {
            .EstimatedTime = "2 сек"
        })

        If includeChart Then
            steps.Add(New ExecutionStep(5, "Добавить диаграмму данных", "chart") With {
                .EstimatedTime = "2 сек"
            })
        End If

        Return steps
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 获取公式的友好描述
    ''' </summary>
    Private Function GetFormulaDescription(formula As String) As String
        If String.IsNullOrEmpty(formula) Then Return ""

        ' 移除开头的=
        formula = formula.TrimStart("="c)

        ' 识别常见公式
        Dim upperFormula = formula.ToUpper()
        
        If upperFormula.StartsWith("SUM(") Then
            Return "Сумма"
        ElseIf upperFormula.StartsWith("AVERAGE(") Then
            Return "Среднее значение"
        ElseIf upperFormula.StartsWith("COUNT(") Then
            Return "Количество"
        ElseIf upperFormula.StartsWith("MAX(") Then
            Return "Максимум"
        ElseIf upperFormula.StartsWith("MIN(") Then
            Return "Минимум"
        ElseIf upperFormula.StartsWith("VLOOKUP(") Then
            Return "Вертикальный поиск"
        ElseIf upperFormula.StartsWith("IF(") Then
            Return "Условная проверка"
        ElseIf upperFormula.StartsWith("SUMIF(") Then
            Return "Сумма по условию"
        ElseIf upperFormula.StartsWith("COUNTIF(") Then
            Return "Количество по условию"
        ElseIf upperFormula.Contains("+") Then
            Return "Сложение"
        ElseIf upperFormula.Contains("-") Then
            Return "Вычитание"
        ElseIf upperFormula.Contains("*") Then
            Return "Умножение"
        ElseIf upperFormula.Contains("/") Then
            Return "Деление"
        Else
            ' 截断过长的公式
            If formula.Length > 30 Then
                Return formula.Substring(0, 27) & "..."
            End If
            Return formula
        End If
    End Function

#End Region

End Class
