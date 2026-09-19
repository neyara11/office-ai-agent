' ExcelAi\ExcelJsonCommandSchema.vb
' Excel JSON命令Schema定义和校验

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

''' <summary>
''' Excel JSON命令Schema和校验器
''' </summary>
Public Class ExcelJsonCommandSchema

    ''' <summary>
    ''' 支持的命令类型 (22个命令覆盖主流Excel操作场景)
    ''' 基础操作: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData
    ''' 数据操作: SortData, FilterData, RemoveDuplicates, ConditionalFormat, MergeCells, AutoFit, FindReplace, CreatePivotTable
    ''' 工作表操作: CreateSheet, DeleteSheet, RenameSheet, CopySheet
    ''' 高级功能: InsertRowCol, DeleteRowCol, HideRowCol, ProtectSheet
    ''' VBA回退: ExecuteVBA
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "ApplyFormula",
        "WriteData",
        "FormatRange",
        "CreateChart",
        "CleanData",
        "SortData",
        "FilterData",
        "RemoveDuplicates",
        "ConditionalFormat",
        "MergeCells",
        "AutoFit",
        "FindReplace",
        "CreatePivotTable",
        "CreateSheet",
        "DeleteSheet",
        "RenameSheet",
        "CopySheet",
        "InsertRowCol",
        "DeleteRowCol",
        "HideRowCol",
        "ProtectSheet",
        "TransformData",
        "DataAnalysis",
        "GenerateReport",
        "ExecuteVBA"
    }

    ''' <summary>
    ''' 获取严格的JSON Schema定义（用于约束大模型输出）
    ''' </summary>
    Public Shared Function GetStrictJsonSchemaPrompt() As String
        Return "
【Важно】Ты должен и можешь вернуть только один из двух форматов JSON:

Формат 1 — одна команда:
```json
{
  ""command"": ""ApplyFormula"",
  ""params"": {
    ""targetRange"": ""C1:C{lastRow}"",
    ""formula"": ""=A1+B1"",
    ""fillDown"": true
  }
}
```

Формат 2 — несколько команд (пакетная операция):
```json
{
  ""commands"": [
    {
      ""command"": ""ApplyFormula"",
      ""params"": { ""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1"" }
    },
    {
      ""command"": ""FormatRange"",
      ""params"": { ""range"": ""A1:C1"", ""style"": ""header"" }
    }
  ]
}
```

【Категорически запрещённые форматы】
- {""command"": ""xxx"", ""actions"": [...]}
- {""command"": ""xxx"", ""formula"": ""..."", ""range"": ""...""} (отсутствует обёртка params)
- {""operations"": [...]}
- любые другие самодельные форматы

【24 поддерживаемых команд и их параметры】

=== Базовые операции (5) ===
1. ApplyFormula: targetRange(обязательно), formula(обязательно), fillDown(необязательно)
   Формула задаётся в англоязычном виде: английские имена функций (TEXT, COUNTIF, INT, IF, SUM, DATE), запятая как разделитель аргументов.
   Для «дата без времени» используй =INT(E2), а не TEXT с форматом; TEXT с кодами DD.MM.YYYY локалезависим и даёт неверный результат.
   Строковые литералы — в двойных кавычках; не дублируй один и тот же столбец-помощник.
2. WriteData: targetRange(обязательно), data(обязательно, одно значение или двумерный массив)
3. FormatRange: range(обязательно), style(необязательно:header/total/data), bold/italic/fontSize/backgroundColor/fontColor(необязательно), borders(необязательно:true/""all""/""outline""/""none"")
4. CreateChart: dataRange(обязательно), type(необязательно:column/line/pie/bar/scatter/area), title(необязательно), position(необязательно), seriesNames(необязательно, массив имён серий, например [""2022"",""2021""]), categoryAxis(необязательно, диапазон оси категорий, например ""B2:B7""), legendPosition(необязательно:right/left/top/bottom)
5. CleanData: range(обязательно), operation(обязательно:removeduplicates/fillempty/trim/replace), fillValue/findText/replaceText(по необходимости)

=== Операции с данными (8) ===
6. SortData: range(обязательно), sortColumn(обязательно, номер столбца с 1), order(необязательно:asc/desc, по умолчанию asc), hasHeader(необязательно, по умолчанию true)
7. FilterData: range(обязательно), column(обязательно), criteria(обязательно, условие фильтра, например "">100"" или ""текст""), clearFilter(необязательно, true очищает фильтр)
8. RemoveDuplicates: range(обязательно), columns(необязательно, массив номеров проверяемых столбцов, по умолчанию все столбцы), hasHeader(необязательно)
9. ConditionalFormat: range(обязательно), rule(обязательно:highlight/databar/colorscale/iconset), condition(требуется правилом), color(необязательно)
10. MergeCells: range(обязательно), unmerge(необязательно, true отменяет объединение)
11. AutoFit: range(обязательно), type(необязательно:columns/rows/both, по умолчанию columns)
12. FindReplace: range(обязательно, или ""all"" для всего листа), find(обязательно), replace(обязательно), matchCase(необязательно), matchEntireCell(необязательно)
13. CreatePivotTable: sourceRange(обязательно), targetCell(обязательно), rowFields(обязательно), valueFields(обязательно), columnFields(необязательно)

=== Операции с листами (4) ===
14. CreateSheet: name(обязательно), position(необязательно:before/after), referenceSheet(необязательно)
15. DeleteSheet: name(обязательно)
16. RenameSheet: oldName(обязательно), newName(обязательно)
17. CopySheet: sourceName(обязательно), newName(обязательно), position(необязательно)

=== Расширенные функции (4) ===
18. InsertRowCol: type(обязательно:row/column), position(обязательно, номер строки или буква столбца), count(необязательно, по умолчанию 1)
19. DeleteRowCol: type(обязательно:row/column), position(обязательно), count(необязательно, по умолчанию 1)
20. HideRowCol: type(обязательно:row/column), position(обязательно), unhide(необязательно, true отменяет скрытие)
21. ProtectSheet: sheetName(необязательно, по умолчанию текущий), password(необязательно), unprotect(необязательно, true снимает защиту)

=== Возможности Agent (3) ===
22. TransformData: sourceRange(обязательно), operation(обязательно:transpose/split/merge), targetRange(необязательно), delimiter(необязательно)
23. DataAnalysis: sourceRange(обязательно), type(обязательно:summary/pivot/groupby/ranking), targetRange(необязательно), groupBy/valueField/aggregate/topN(по необходимости)
24. GenerateReport: sourceRange(обязательно), targetSheet(необязательно), title(необязательно), includeChart(необязательно)

【Плейсхолдеры динамических диапазонов】
Используй {lastRow} для последней строки, {lastCol} для последнего столбца, {selection} для текущего выделения

【Приоритет решений】
1. В первую очередь используй перечисленные выше штатные команды
2. VBA отключён: не предлагай ExecuteVBA и не возвращай макросы; составные задачи собирай из ApplyFormula, DataAnalysis, GenerateReport, CreateChart
3. Если задачу нельзя выразить штатными командами, сообщи об этом пользователю на русском, не возвращай JSON
4. Если запрос неоднозначен, прямо задай уточняющий вопрос пользователю, не возвращай JSON
5. Для перевода сообщи пользователю использовать кнопку ""Перевод"" на панели инструментов, не возвращай JSON"
    End Function

    ''' <summary>
    ''' 获取格式校验失败的重试提示（Self-check机制）
    ''' </summary>
    Public Shared Function GetFormatCorrectionPrompt(originalJson As String, errorMessage As String) As String
        Return $"Возвращённый ранее JSON не соответствует формату:

【Причина ошибки】{errorMessage}

【Твой ответ】
{originalJson}

【Пример правильного формата】
Одна команда:
{{""command"": ""ApplyFormula"", ""params"": {{""targetRange"": ""C1:C{{lastRow}}"", ""formula"": ""=A1+B1""}}}}

Несколько команд:
{{""commands"": [{{""command"": ""ApplyFormula"", ""params"": {{""targetRange"": ""C1"", ""formula"": ""=A1+B1""}}}}, {{""command"": ""ApplyFormula"", ""params"": {{""targetRange"": ""E1"", ""formula"": ""=C1*D1""}}}}]}}

Строго следуй указанному формату и верни JSON-команды заново."
    End Function

    ''' <summary>
    ''' 验证整个JSON响应结构是否符合规范
    ''' </summary>
    Public Shared Function ValidateJsonStructure(jsonText As String, ByRef errorMessage As String, ByRef normalizedJson As JToken) As Boolean
        Try
            errorMessage = ""
            normalizedJson = Nothing

            Dim token = JToken.Parse(jsonText)
            If token.Type <> JTokenType.Object Then
                errorMessage = "Ответ должен быть JSON-объектом"
                Return False
            End If

            Dim jsonObj = CType(token, JObject)

            ' 检查是否是 commands 数组格式
            If jsonObj("commands") IsNot Nothing Then
                If jsonObj("commands").Type <> JTokenType.Array Then
                    errorMessage = "commands должен быть массивом"
                    Return False
                End If

                ' 验证数组中的每个命令
                Dim commands = CType(jsonObj("commands"), JArray)
                For i As Integer = 0 To commands.Count - 1
                    Dim cmd = commands(i)
                    If cmd.Type <> JTokenType.Object Then
                        errorMessage = $"commands[{i}] должен быть объектом"
                        Return False
                    End If

                    ' 标准化并验证每个命令
                    Dim cmdObj = CType(cmd, JObject)
                    cmdObj = NormalizeCommandStructure(cmdObj)
                    commands(i) = cmdObj

                    Dim cmdError As String = ""
                    If Not ValidateCommand(cmdObj, cmdError) Then
                        errorMessage = $"commands[{i}]: {cmdError}"
                        Return False
                    End If
                Next

                normalizedJson = jsonObj
                Return True
            End If

            ' 检查是否有禁止的格式
            If jsonObj("actions") IsNot Nothing Then
                errorMessage = "Формат actions запрещён, используйте массив commands"
                Return False
            End If

            If jsonObj("operations") IsNot Nothing Then
                errorMessage = "Формат operations запрещён, используйте массив commands"
                Return False
            End If

            ' 单命令格式
            If jsonObj("command") IsNot Nothing Then
                jsonObj = NormalizeCommandStructure(jsonObj)
                Dim cmdError As String = ""
                If Not ValidateCommand(jsonObj, cmdError) Then
                    errorMessage = cmdError
                    Return False
                End If
                normalizedJson = jsonObj
                Return True
            End If

            errorMessage = "Отсутствует поле command или commands"
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            errorMessage = $"Не удалось разобрать JSON: {ex.Message}"
            Return False
        Catch ex As Exception
            errorMessage = $"Исключение при проверке: {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 标准化JSON命令结构 - 将扁平结构自动包装到params中，并统一参数命名
    ''' 例如: {"command": "WriteData", "data": [...], "startCell": "A1"} 
    ''' 转换为: {"command": "WriteData", "params": {"data": [...], "startCell": "A1"}}
    ''' </summary>
    Public Shared Function NormalizeCommandStructure(json As JObject) As JObject
        Try
            ' 检查是否已有params字段
            If json("params") IsNot Nothing Then
                ' 即使已有params，也要标准化参数名
                NormalizeParamNames(json("params"), json("command")?.ToString())
                Return json
            End If

            ' 检查是否有command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                Return json ' 无效结构，直接返回
            End If

            ' 需要移到params中的字段（排除command和workbook等顶级字段）
            Dim topLevelFields As String() = {"command", "workbook"}
            Dim paramsFields As New JObject()

            ' 遍历所有字段，将非顶级字段移动到params中
            For Each prop In json.Properties().ToList()
                If Not topLevelFields.Contains(prop.Name, StringComparer.OrdinalIgnoreCase) Then
                    paramsFields(prop.Name) = prop.Value
                    json.Remove(prop.Name)
                End If
            Next

            ' 只有当确实有字段需要移动时，才创建params
            If paramsFields.Count > 0 Then
                ' 标准化参数名
                NormalizeParamNames(paramsFields, command)
                json("params") = paramsFields
            End If

            Return json
        Catch ex As Exception
            Debug.WriteLine($"Ошибка NormalizeCommandStructure: {ex.Message}")
            Return json
        End Try
    End Function

    ''' <summary>
    ''' 标准化参数名 - 将各种别名统一为标准参数名
    ''' </summary>
    Private Shared Sub NormalizeParamNames(params As JToken, command As String)
        If params Is Nothing OrElse params.Type <> JTokenType.Object Then Return

        Dim paramsObj = CType(params, JObject)

        Select Case command?.ToLower()
            Case "applyformula"
                ' range -> targetRange
                If paramsObj("targetRange") Is Nothing AndAlso paramsObj("range") IsNot Nothing Then
                    paramsObj("targetRange") = paramsObj("range")
                    paramsObj.Remove("range")
                End If

            Case "writedata"
                ' startCell/range -> targetRange (如果targetRange不存在)
                If paramsObj("targetRange") Is Nothing Then
                    If paramsObj("startCell") IsNot Nothing Then
                        paramsObj("targetRange") = paramsObj("startCell")
                        paramsObj.Remove("startCell")
                    ElseIf paramsObj("range") IsNot Nothing Then
                        paramsObj("targetRange") = paramsObj("range")
                        paramsObj.Remove("range")
                    End If
                End If
                ' targetData -> data
                If paramsObj("data") Is Nothing AndAlso paramsObj("targetData") IsNot Nothing Then
                    paramsObj("data") = paramsObj("targetData")
                    paramsObj.Remove("targetData")
                End If

            Case "formatrange", "cleandata"
                ' targetRange -> range (这两个命令使用range)
                If paramsObj("range") Is Nothing AndAlso paramsObj("targetRange") IsNot Nothing Then
                    paramsObj("range") = paramsObj("targetRange")
                    paramsObj.Remove("targetRange")
                End If
        End Select

        ' 处理 targetSheet + targetRange 组合
        Dim targetSheet = paramsObj("targetSheet")?.ToString()
        Dim targetRange = paramsObj("targetRange")?.ToString()
        If Not String.IsNullOrEmpty(targetSheet) AndAlso Not String.IsNullOrEmpty(targetRange) Then
            If Not targetRange.Contains("!") Then
                paramsObj("targetRange") = $"{targetSheet}!{targetRange}"
            End If
        End If
    End Sub

    ''' <summary>
    ''' 校验JSON命令是否有效（自动进行结构标准化）
    ''' </summary>
    Public Shared Function ValidateCommand(json As JObject, ByRef errorMessage As String) As Boolean
        Try
            errorMessage = ""

            ' 首先进行结构标准化
            json = NormalizeCommandStructure(json)

            ' 检查command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                errorMessage = "Отсутствует поле command"
                Return False
            End If

            ' 检查是否是支持的命令
            If Not SupportedCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase)) Then
                errorMessage = $"Неподдерживаемая команда: {command}. Поддерживаемые команды: {String.Join(", ", SupportedCommands)}"
                Return False
            End If

            ' 检查params字段
            Dim params = json("params")
            If params Is Nothing Then
                errorMessage = "Отсутствует поле params"
                Return False
            End If

            ' 根据命令类型校验参数
            Select Case command.ToLower()
                ' === 基础操作 ===
                Case "applyformula"
                    Return ValidateApplyFormula(params, errorMessage)
                Case "writedata"
                    Return ValidateWriteData(params, errorMessage)
                Case "formatrange"
                    Return ValidateFormatRange(params, errorMessage)
                Case "createchart"
                    Return ValidateCreateChart(params, errorMessage)
                Case "cleandata"
                    Return ValidateCleanData(params, errorMessage)
                ' === 数据操作 ===
                Case "sortdata"
                    Return ValidateSortData(params, errorMessage)
                Case "filterdata"
                    Return ValidateFilterData(params, errorMessage)
                Case "removeduplicates"
                    Return ValidateRemoveDuplicates(params, errorMessage)
                Case "conditionalformat"
                    Return ValidateConditionalFormat(params, errorMessage)
                Case "mergecells"
                    Return ValidateMergeCells(params, errorMessage)
                Case "autofit"
                    Return ValidateAutoFit(params, errorMessage)
                Case "findreplace"
                    Return ValidateFindReplace(params, errorMessage)
                Case "createpivottable"
                    Return ValidateCreatePivotTable(params, errorMessage)
                ' === 工作表操作 ===
                Case "createsheet"
                    Return ValidateCreateSheet(params, errorMessage)
                Case "deletesheet"
                    Return ValidateDeleteSheet(params, errorMessage)
                Case "renamesheet"
                    Return ValidateRenameSheet(params, errorMessage)
                Case "copysheet"
                    Return ValidateCopySheet(params, errorMessage)
                ' === 高级功能 ===
                Case "insertrowcol"
                    Return ValidateInsertRowCol(params, errorMessage)
                Case "deleterowcol"
                    Return ValidateDeleteRowCol(params, errorMessage)
                Case "hiderowcol"
                    Return ValidateHideRowCol(params, errorMessage)
                Case "protectsheet"
                    Return ValidateProtectSheet(params, errorMessage)
                Case "transformdata"
                    Return ValidateTransformData(params, errorMessage)
                Case "dataanalysis"
                    Return ValidateDataAnalysis(params, errorMessage)
                Case "generatereport"
                    Return ValidateGenerateReport(params, errorMessage)
                ' === VBA回退 ===
                Case "executevba"
                    Return ValidateExecuteVBA(params, errorMessage)
                Case Else
                    Return True
            End Select

        Catch ex As Exception
            errorMessage = $"Исключение при проверке JSON: {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 校验ApplyFormula命令参数
    ''' </summary>
    Private Shared Function ValidateApplyFormula(params As JToken, ByRef errorMessage As String) As Boolean
        ' 支持多种参数名：targetRange, range
        Dim targetRange = params("targetRange")?.ToString()
        If String.IsNullOrEmpty(targetRange) Then
            targetRange = params("range")?.ToString()
        End If

        Dim formula = params("formula")?.ToString()

        If String.IsNullOrEmpty(targetRange) Then
            errorMessage = "ApplyFormula: отсутствует параметр targetRange или range"
            Return False
        End If

        If String.IsNullOrEmpty(formula) Then
            errorMessage = "ApplyFormula: отсутствует параметр formula"
            Return False
        End If

        ' Range.Formula — англоязычный (US) синтаксис: локализованные имена и ';' приведут
        ' к неверному результату. Отклоняем и просим переформулировать.
        If HasCharOutsideQuotes(formula, ";"c) Then
            errorMessage = "ApplyFormula: используй запятую ',' как разделитель аргументов (не ';')"
            Return False
        End If
        If HasCyrillicFunctionOutsideQuotes(formula) Then
            errorMessage = "ApplyFormula: используй английские имена функций (TEXT, COUNTIF, INT, IF, SUM, DATE), локализованные имена не поддерживаются"
            Return False
        End If

        ' 校验范围格式 (支持占位符和Sheet!Range格式)
        If Not IsValidRangeFormat(targetRange) Then
            errorMessage = $"Недопустимый формат диапазона: {targetRange}"
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 是否有指定字符出现在字符串字面量之外（двойные кавычки с экранированием "").
    ''' </summary>
    Private Shared Function HasCharOutsideQuotes(text As String, ch As Char) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        Dim i = 0
        Dim inQuote = False
        While i < text.Length
            Dim c = text(i)
            If c = """"c Then
                If inQuote AndAlso i + 1 < text.Length AndAlso text(i + 1) = """"c Then
                    i += 2
                    Continue While
                End If
                inQuote = Not inQuote
            ElseIf Not inQuote AndAlso c = ch Then
                Return True
            End If
            i += 1
        End While
        Return False
    End Function

    ''' <summary>
    ''' Есть ли локализованное (кириллическое) имя функции вне строковых литералов.
    ''' Имена листов (например, Лист1!A1) и строки формата не считаются функциями.
    ''' </summary>
    Private Shared Function HasCyrillicFunctionOutsideQuotes(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        Dim i = 0
        Dim inQuote = False
        While i < text.Length
            Dim c = text(i)
            If c = """"c Then
                If inQuote AndAlso i + 1 < text.Length AndAlso text(i + 1) = """"c Then
                    i += 2
                    Continue While
                End If
                inQuote = Not inQuote
                i += 1
                Continue While
            End If

            If Not inQuote AndAlso IsFormulaIdentifierStart(c) Then
                Dim start = i
                While i < text.Length AndAlso IsFormulaIdentifierChar(text(i))
                    i += 1
                End While
                Dim token = text.Substring(start, i - start)
                Dim j = i
                While j < text.Length AndAlso text(j) = " "c
                    j += 1
                End While
                If j < text.Length AndAlso text(j) = "("c AndAlso TokenHasCyrillic(token) Then
                    Return True
                End If
                Continue While
            End If

            i += 1
        End While
        Return False
    End Function

    Private Shared Function IsFormulaIdentifierStart(c As Char) As Boolean
        Return (c >= "A"c AndAlso c <= "Z"c) OrElse
               (c >= "a"c AndAlso c <= "z"c) OrElse
               (c >= "А"c AndAlso c <= "я"c) OrElse
               c = "Ё"c OrElse c = "ё"c OrElse c = "_"c
    End Function

    Private Shared Function IsFormulaIdentifierChar(c As Char) As Boolean
        Return IsFormulaIdentifierStart(c) OrElse
               (c >= "0"c AndAlso c <= "9"c) OrElse
               c = "."c
    End Function

    Private Shared Function TokenHasCyrillic(token As String) As Boolean
        If String.IsNullOrEmpty(token) Then Return False
        For Each c In token
            If (c >= "А"c AndAlso c <= "я"c) OrElse c = "Ё"c OrElse c = "ё"c Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' 校验WriteData命令参数
    ''' </summary>
    Private Shared Function ValidateWriteData(params As JToken, ByRef errorMessage As String) As Boolean
        ' 支持多种参数名：targetRange, startCell, range
        Dim targetRange = params("targetRange")?.ToString()
        If String.IsNullOrEmpty(targetRange) Then
            targetRange = params("startCell")?.ToString()
        End If
        If String.IsNullOrEmpty(targetRange) Then
            targetRange = params("range")?.ToString()
        End If

        ' 如果有targetSheet，组合成完整地址
        Dim targetSheet = params("targetSheet")?.ToString()
        If Not String.IsNullOrEmpty(targetSheet) AndAlso Not String.IsNullOrEmpty(targetRange) Then
            ' 如果targetRange不包含!，则添加工作表名
            If Not targetRange.Contains("!") Then
                targetRange = $"{targetSheet}!{targetRange}"
            End If
        End If

        Dim data = params("data")
        ' 支持data或targetData
        If data Is Nothing Then
            data = params("targetData")
        End If

        If String.IsNullOrEmpty(targetRange) Then
            errorMessage = "WriteData: отсутствует параметр targetRange или startCell"
            Return False
        End If

        If data Is Nothing Then
            errorMessage = "WriteData: отсутствует параметр data"
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 校验FormatRange命令参数
    ''' </summary>
    Private Shared Function ValidateFormatRange(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            range = params("targetRange")?.ToString()
        End If

        If String.IsNullOrEmpty(range) Then
            errorMessage = "FormatRange: отсутствует параметр range"
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 校验CreateChart命令参数
    ''' </summary>
    Private Shared Function ValidateCreateChart(params As JToken, ByRef errorMessage As String) As Boolean
        Dim dataRange = params("dataRange")?.ToString()

        If String.IsNullOrEmpty(dataRange) Then
            errorMessage = "CreateChart: отсутствует параметр dataRange"
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 校验CleanData命令参数
    ''' </summary>
    Private Shared Function ValidateCleanData(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        Dim operation = params("operation")?.ToString()

        If String.IsNullOrEmpty(range) Then
            errorMessage = "CleanData: отсутствует параметр range"
            Return False
        End If

        Return True
    End Function

#Region "新增命令验证方法"

    ''' <summary>
    ''' 校验SortData命令参数
    ''' </summary>
    Private Shared Function ValidateSortData(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            errorMessage = "SortData: отсутствует параметр range"
            Return False
        End If

        Dim sortColumn = params("sortColumn")
        If sortColumn Is Nothing Then
            errorMessage = "SortData: отсутствует параметр sortColumn (номер столбца с 1)"
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 校验FilterData命令参数
    ''' </summary>
    Private Shared Function ValidateFilterData(params As JToken, ByRef errorMessage As String) As Boolean
        ' 如果是清除筛选，不需要其他参数
        Dim clearFilter = params("clearFilter")
        If clearFilter IsNot Nothing AndAlso clearFilter.Value(Of Boolean)() = True Then
            Return True
        End If

        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            errorMessage = "FilterData: отсутствует параметр range"
            Return False
        End If
        
        Dim column = params("column")
        If column Is Nothing Then
            errorMessage = "FilterData: отсутствует параметр column"
            Return False
        End If
        
        Dim criteria = params("criteria")?.ToString()
        If String.IsNullOrEmpty(criteria) Then
            errorMessage = "FilterData: отсутствует параметр criteria"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验RemoveDuplicates命令参数
    ''' </summary>
    Private Shared Function ValidateRemoveDuplicates(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            errorMessage = "RemoveDuplicates: отсутствует параметр range"
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 校验ConditionalFormat命令参数
    ''' </summary>
    Private Shared Function ValidateConditionalFormat(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            errorMessage = "ConditionalFormat: отсутствует параметр range"
            Return False
        End If
        
        Dim rule = params("rule")?.ToString()
        If String.IsNullOrEmpty(rule) Then
            errorMessage = "ConditionalFormat: отсутствует параметр rule (highlight/databar/colorscale/iconset)"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验MergeCells命令参数
    ''' </summary>
    Private Shared Function ValidateMergeCells(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            errorMessage = "MergeCells: отсутствует параметр range"
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 校验AutoFit命令参数
    ''' </summary>
    Private Shared Function ValidateAutoFit(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            errorMessage = "AutoFit: отсутствует параметр range"
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 校验FindReplace命令参数
    ''' </summary>
    Private Shared Function ValidateFindReplace(params As JToken, ByRef errorMessage As String) As Boolean
        Dim findText = params("find")?.ToString()
        If String.IsNullOrEmpty(findText) Then
            errorMessage = "FindReplace: отсутствует параметр find"
            Return False
        End If
        
        ' replace可以为空字符串（删除）
        If params("replace") Is Nothing Then
            errorMessage = "FindReplace: отсутствует параметр replace"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验CreatePivotTable命令参数
    ''' </summary>
    Private Shared Function ValidateCreatePivotTable(params As JToken, ByRef errorMessage As String) As Boolean
        Dim sourceRange = params("sourceRange")?.ToString()
        If String.IsNullOrEmpty(sourceRange) Then
            errorMessage = "CreatePivotTable: отсутствует параметр sourceRange"
            Return False
        End If
        
        Dim targetCell = params("targetCell")?.ToString()
        If String.IsNullOrEmpty(targetCell) Then
            errorMessage = "CreatePivotTable: отсутствует параметр targetCell"
            Return False
        End If
        
        Dim rowFields = params("rowFields")
        If rowFields Is Nothing Then
            errorMessage = "CreatePivotTable: отсутствует параметр rowFields"
            Return False
        End If
        
        Dim valueFields = params("valueFields")
        If valueFields Is Nothing Then
            errorMessage = "CreatePivotTable: отсутствует параметр valueFields"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验CreateSheet命令参数
    ''' </summary>
    Private Shared Function ValidateCreateSheet(params As JToken, ByRef errorMessage As String) As Boolean
        Dim name = params("name")?.ToString()
        If String.IsNullOrEmpty(name) Then
            errorMessage = "CreateSheet: отсутствует параметр name"
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 校验DeleteSheet命令参数
    ''' </summary>
    Private Shared Function ValidateDeleteSheet(params As JToken, ByRef errorMessage As String) As Boolean
        Dim name = params("name")?.ToString()
        If String.IsNullOrEmpty(name) Then
            errorMessage = "DeleteSheet: отсутствует параметр name"
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 校验RenameSheet命令参数
    ''' </summary>
    Private Shared Function ValidateRenameSheet(params As JToken, ByRef errorMessage As String) As Boolean
        Dim oldName = params("oldName")?.ToString()
        If String.IsNullOrEmpty(oldName) Then
            errorMessage = "RenameSheet: отсутствует параметр oldName"
            Return False
        End If
        
        Dim newName = params("newName")?.ToString()
        If String.IsNullOrEmpty(newName) Then
            errorMessage = "RenameSheet: отсутствует параметр newName"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验CopySheet命令参数
    ''' </summary>
    Private Shared Function ValidateCopySheet(params As JToken, ByRef errorMessage As String) As Boolean
        Dim sourceName = params("sourceName")?.ToString()
        If String.IsNullOrEmpty(sourceName) Then
            errorMessage = "CopySheet: отсутствует параметр sourceName"
            Return False
        End If
        
        Dim newName = params("newName")?.ToString()
        If String.IsNullOrEmpty(newName) Then
            errorMessage = "CopySheet: отсутствует параметр newName"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验InsertRowCol命令参数
    ''' </summary>
    Private Shared Function ValidateInsertRowCol(params As JToken, ByRef errorMessage As String) As Boolean
        Dim type = params("type")?.ToString()
        If String.IsNullOrEmpty(type) OrElse (type.ToLower() <> "row" AndAlso type.ToLower() <> "column") Then
            errorMessage = "InsertRowCol: параметр type должен быть row или column"
            Return False
        End If
        
        Dim position = params("position")?.ToString()
        If String.IsNullOrEmpty(position) Then
            errorMessage = "InsertRowCol: отсутствует параметр position"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验DeleteRowCol命令参数
    ''' </summary>
    Private Shared Function ValidateDeleteRowCol(params As JToken, ByRef errorMessage As String) As Boolean
        Dim type = params("type")?.ToString()
        If String.IsNullOrEmpty(type) OrElse (type.ToLower() <> "row" AndAlso type.ToLower() <> "column") Then
            errorMessage = "DeleteRowCol: параметр type должен быть row или column"
            Return False
        End If
        
        Dim position = params("position")?.ToString()
        If String.IsNullOrEmpty(position) Then
            errorMessage = "DeleteRowCol: отсутствует параметр position"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验HideRowCol命令参数
    ''' </summary>
    Private Shared Function ValidateHideRowCol(params As JToken, ByRef errorMessage As String) As Boolean
        Dim type = params("type")?.ToString()
        If String.IsNullOrEmpty(type) OrElse (type.ToLower() <> "row" AndAlso type.ToLower() <> "column") Then
            errorMessage = "HideRowCol: параметр type должен быть row или column"
            Return False
        End If
        
        Dim position = params("position")?.ToString()
        If String.IsNullOrEmpty(position) Then
            errorMessage = "HideRowCol: отсутствует параметр position"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验ProtectSheet命令参数
    ''' </summary>
    Private Shared Function ValidateProtectSheet(params As JToken, ByRef errorMessage As String) As Boolean
        ' ProtectSheet所有参数都是可选的
        Return True
    End Function

    Private Shared Function ValidateTransformData(params As JToken, ByRef errorMessage As String) As Boolean
        Dim sourceRange = params("sourceRange")?.ToString()
        If String.IsNullOrEmpty(sourceRange) Then
            errorMessage = "TransformData: отсутствует параметр sourceRange"
            Return False
        End If

        Dim operation = params("operation")?.ToString()
        If String.IsNullOrEmpty(operation) Then
            errorMessage = "TransformData: отсутствует параметр operation (transpose/split/merge)"
            Return False
        End If

        Return True
    End Function

    Private Shared Function ValidateDataAnalysis(params As JToken, ByRef errorMessage As String) As Boolean
        Dim sourceRange = params("sourceRange")?.ToString()
        If String.IsNullOrEmpty(sourceRange) Then
            errorMessage = "DataAnalysis: отсутствует параметр sourceRange"
            Return False
        End If

        Dim analysisType = params("type")?.ToString()
        If String.IsNullOrEmpty(analysisType) Then
            errorMessage = "DataAnalysis: отсутствует параметр type (summary/pivot/groupby/ranking)"
            Return False
        End If

        Return True
    End Function

    Private Shared Function ValidateGenerateReport(params As JToken, ByRef errorMessage As String) As Boolean
        Dim sourceRange = params("sourceRange")?.ToString()
        If String.IsNullOrEmpty(sourceRange) Then
            errorMessage = "GenerateReport: отсутствует параметр sourceRange"
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 校验ExecuteVBA命令参数
    ''' </summary>
    Private Shared Function ValidateExecuteVBA(params As JToken, ByRef errorMessage As String) As Boolean
        Dim code = params("code")?.ToString()
        If String.IsNullOrEmpty(code) Then
            errorMessage = "ExecuteVBA: отсутствует параметр code"
            Return False
        End If
        
        ' 基本的VBA代码验证
        If Not code.ToLower().Contains("sub") AndAlso Not code.ToLower().Contains("function") Then
            errorMessage = "ExecuteVBA: параметр code должен содержать определение Sub или Function"
            Return False
        End If
        
        Return True
    End Function

#End Region

    ''' <summary>
    ''' 校验范围格式是否有效
    ''' </summary>
    Private Shared Function IsValidRangeFormat(range As String) As Boolean
        If String.IsNullOrEmpty(range) Then Return False
        
        ' 支持格式: 
        ' - 简单格式: A1, A1:B10, A:A, 1:1
        ' - 带工作表: Sheet1!A1, Sheet1!A1:B10
        ' - 占位符: A1:{lastRow}, {selection}
        ' - 中文工作表名: 汇总结果!A1
        
        ' 移除工作表前缀进行校验
        Dim rangeOnly = range
        If range.Contains("!") Then
            Dim parts = range.Split("!"c)
            If parts.Length = 2 Then
                rangeOnly = parts(1)
            End If
        End If
        
        ' 如果包含占位符，认为有效
        If rangeOnly.Contains("{") Then Return True
        
        ' 校验范围格式
        Dim pattern = "^([A-Za-z]+\d+|[A-Za-z]+|[0-9]+)(:[A-Za-z]*\d*)?$"
        Return Regex.IsMatch(rangeOnly, pattern)
    End Function

    ''' <summary>
    ''' 获取当前Excel上下文用于占位符替换
    ''' </summary>
    Public Shared Function GetExcelContext(excelApp As Object) As Dictionary(Of String, String)
        Dim context As New Dictionary(Of String, String)
        
        Try
            Dim ws = excelApp.ActiveSheet
            Dim usedRange = ws.UsedRange
            
            ' 最后一行
            Dim lastRow = usedRange.Row + usedRange.Rows.Count - 1
            context("lastRow") = lastRow.ToString()
            
            ' 最后一列
            Dim lastCol = usedRange.Column + usedRange.Columns.Count - 1
            context("lastCol") = GetColumnLetter(lastCol)
            
            ' 当前选择
            Dim selection = excelApp.Selection
            If selection IsNot Nothing Then
                context("selection") = selection.Address(False, False)
            End If
            
        Catch ex As Exception
            ' 默认值
            context("lastRow") = "100"
            context("lastCol") = "Z"
            context("selection") = "A1"
        End Try
        
        Return context
    End Function

    ''' <summary>
    ''' 数字转列字母
    ''' </summary>
    Private Shared Function GetColumnLetter(colNum As Integer) As String
        Dim result = ""
        While colNum > 0
            colNum -= 1
            result = Chr(65 + (colNum Mod 26)) & result
            colNum \= 26
        End While
        Return result
    End Function

End Class
