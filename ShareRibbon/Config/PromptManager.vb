' ShareRibbon\Config\PromptManager.vb
' 统一提示词管理中心

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json

''' <summary>
''' 统一提示词管理中心 - 单例模式
''' 负责管理所有提示词的加载、组合和获取
''' </summary>
Public Class PromptManager
    Private Shared _instance As PromptManager
    Private _promptConfig As PromptConfiguration

    ''' <summary>
    ''' 获取单例实例
    ''' </summary>
    Public Shared ReadOnly Property Instance As PromptManager
        Get
            If _instance Is Nothing Then
                _instance = New PromptManager()
            End If
            Return _instance
        End Get
    End Property

    ''' <summary>
    ''' 私有构造函数
    ''' </summary>
    Private Sub New()
        LoadPromptConfiguration()
    End Sub

    ''' <summary>
    ''' 加载提示词配置
    ''' </summary>
    Private Sub LoadPromptConfiguration()
        Dim configPath = GetPromptConfigPath()

        Try
            If File.Exists(configPath) Then
                Dim json = File.ReadAllText(configPath, Encoding.UTF8)
                _promptConfig = JsonConvert.DeserializeObject(Of PromptConfiguration)(json)
            Else
                ' 使用默认配置
                _promptConfig = CreateDefaultConfiguration()
                SavePromptConfiguration()
            End If
        Catch ex As Exception
            Debug.WriteLine($"加载提示词配置失败: {ex.Message}")
            _promptConfig = CreateDefaultConfiguration()
        End Try
    End Sub

    ''' <summary>
    ''' 保存提示词配置
    ''' </summary>
    Public Sub SavePromptConfiguration()
        Try
            Dim configPath = GetPromptConfigPath()
            Dim dir = Path.GetDirectoryName(configPath)

            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json = JsonConvert.SerializeObject(_promptConfig, Formatting.Indented)
            File.WriteAllText(configPath, json, Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"保存提示词配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 更新指定应用的JSON Schema约束
    ''' </summary>
    Public Sub UpdateJsonSchemaConstraint(appType As String, constraint As String)
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        If appConfig IsNot Nothing Then
            appConfig.JsonSchemaConstraint = constraint
        End If
    End Sub

    ''' <summary>
    ''' 重置指定应用的JSON Schema约束为默认值
    ''' </summary>
    Public Sub ResetJsonSchemaConstraint(appType As String)
        Dim defaultConfig = CreateDefaultConfiguration()
        Dim defaultAppConfig = defaultConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))

        If defaultAppConfig IsNot Nothing Then
            Dim currentAppConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
            If currentAppConfig IsNot Nothing Then
                currentAppConfig.JsonSchemaConstraint = defaultAppConfig.JsonSchemaConstraint
            End If
        End If
    End Sub

    ''' <summary>
    ''' 获取组合后的提示词（融合模式）
    ''' </summary>
    ''' <param name="context">提示词上下文</param>
    ''' <returns>组合后的完整提示词</returns>
    Public Function GetCombinedPrompt(context As PromptContext) As String
        Dim sb As New StringBuilder()

        ' 判断是否为功能性模式（校对/排版/续写/模板渲染等）
        ' 功能性模式不使用用户配置的提示词，避免污染
        Dim isInFunctionalMode As Boolean = CheckIsFunctionalMode(context.FunctionMode)

        ' 1. 用户配置提示词（仅在非功能性模式下使用）
        If Not isInFunctionalMode AndAlso Not String.IsNullOrEmpty(ConfigSettings.propmtContent) Then
            sb.AppendLine(ConfigSettings.propmtContent)
            sb.AppendLine()
        End If

        ' 2. 意图专用提示词（仅在非功能性模式下使用，置信度>0.2时）
        If Not isInFunctionalMode AndAlso context.IntentResult IsNot Nothing AndAlso context.IntentResult.Confidence > 0.2 Then
            Dim intentPrompt = GetIntentSpecificPrompt(context)
            If Not String.IsNullOrEmpty(intentPrompt) Then
                sb.AppendLine(intentPrompt)
                sb.AppendLine()
            End If
        End If

        ' 3. 功能模式提示词（校对/排版/续写/模板渲染）
        If Not String.IsNullOrEmpty(context.FunctionMode) Then
            Dim modePrompt = GetFunctionModePrompt(context)
            If Not String.IsNullOrEmpty(modePrompt) Then
                sb.AppendLine(modePrompt)
                sb.AppendLine()
            End If
        End If

        ' 4. 输出格式约束（JSON Schema或纯文本）- 仅在非功能性模式下添加
        If Not isInFunctionalMode Then
            Dim formatConstraint = GetOutputFormatConstraint(context)
            If Not String.IsNullOrEmpty(formatConstraint) Then
                sb.AppendLine(formatConstraint)
            End If
        End If

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 判断是否为功能性模式（这些模式使用专用提示词，不受用户配置和历史记录影响）
    ''' </summary>
    Private Function CheckIsFunctionalMode(functionMode As String) As Boolean
        If String.IsNullOrEmpty(functionMode) Then Return False

        Dim functionalModes As String() = {"proofread", "reformat", "continuation", "template_render"}
        Return functionalModes.Contains(functionMode.ToLower())
    End Function

    ''' <summary>
    ''' 获取意图专用提示词
    ''' </summary>
    Private Function GetIntentSpecificPrompt(context As PromptContext) As String
        Dim appType = context.ApplicationType
        Dim intentType = If(context.IntentResult?.OfficeIntent.ToString(), "GENERAL_QUERY")

        ' 从配置中查找对应的意图提示词
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        If appConfig Is Nothing Then Return String.Empty

        Dim intentPrompt = appConfig.IntentPrompts.FirstOrDefault(Function(p) p.IntentType.Equals(intentType, StringComparison.OrdinalIgnoreCase))
        If intentPrompt Is Nothing Then
            ' 如果没有找到特定意图，使用通用提示词
            intentPrompt = appConfig.IntentPrompts.FirstOrDefault(Function(p) p.IntentType.Equals("GENERAL_QUERY", StringComparison.OrdinalIgnoreCase))
        End If

        Return If(intentPrompt?.Content, String.Empty)
    End Function

    ''' <summary>
    ''' 获取功能模式专用提示词
    ''' </summary>
    Private Function GetFunctionModePrompt(context As PromptContext) As String
        Dim appType = context.ApplicationType

        ' 从配置中查找对应的功能模式提示词
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        If appConfig Is Nothing Then Return String.Empty

        Dim modePrompt = appConfig.FunctionModePrompts.FirstOrDefault(Function(p) p.Mode.Equals(context.FunctionMode, StringComparison.OrdinalIgnoreCase))
        Return If(modePrompt?.Content, String.Empty)
    End Function

    ''' <summary>
    ''' 获取输出格式约束
    ''' </summary>
    Private Function GetOutputFormatConstraint(context As PromptContext) As String
        ' 判断是否需要JSON输出
        Dim needsJsonOutput = DetermineIfNeedsJsonOutput(context)

        If needsJsonOutput Then
            ' 返回JSON Schema约束
            Return GetJsonSchemaConstraint(context.ApplicationType)
        Else
            ' 返回纯文本输出约束
            Return GetPlainTextConstraint(context.FunctionMode)
        End If
    End Function

    ''' <summary>
    ''' 判断是否需要JSON输出
    ''' </summary>
    Private Function DetermineIfNeedsJsonOutput(context As PromptContext) As Boolean
        ' 特殊功能模式使用特定格式
        Select Case context.FunctionMode?.ToLower()
            Case "continuation", "template_render"
                ' 续写和模板渲染返回纯文本
                Return False
            Case "proofread", "reformat"
                ' 校对和排版返回JSON数组
                Return True
        End Select

        ' Office应用始终需要JSON Schema约束
        ' 因为用户请求可能涉及命令操作（如翻译、公式、图表等）
        ' JSON Schema中已说明"对于简单问候或问答，直接用中文回复"
        Dim appType = context.ApplicationType?.ToLower()
        If appType = "excel" OrElse appType = "word" OrElse appType = "powerpoint" Then
            Return True
        End If

        ' 意图判断
        If context.IntentResult IsNot Nothing Then
            ' 高置信度（>0.7）时，即使是GENERAL_QUERY也返回JSON（用户需求明确）
            If context.IntentResult.Confidence > 0.7 Then
                Return True
            End If

            ' 中等置信度（>0.2）且不是一般查询，需要JSON
            If context.IntentResult.Confidence > 0.2 AndAlso
               context.IntentResult.OfficeIntent <> OfficeIntentType.GENERAL_QUERY Then
                Return True
            End If
        End If

        Return False
    End Function

    ''' <summary>
    ''' 获取JSON Schema约束（根据Office应用类型）- 供外部调用
    ''' </summary>
    Public Function GetJsonSchemaConstraint(appType As String) As String
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        Dim userConstraint = appConfig?.JsonSchemaConstraint
        
        ' 如果用户配置为空或明显不完整，则使用内置默认值
        If String.IsNullOrEmpty(userConstraint) OrElse
           Not IsValidJsonSchemaConstraint(userConstraint) OrElse
           IsLegacyBuiltInJsonSchemaConstraint(appType, userConstraint) Then
            Return GetDefaultJsonSchemaConstraint(appType)
        End If
        
        Return userConstraint
    End Function

    ''' <summary>
    ''' 识别已经写入用户配置的旧版内置 Prompt。这些 Prompt 声明了不存在的
    ''' executor，不应继续被当作用户自定义配置使用。
    ''' </summary>
    Private Shared Function IsLegacyBuiltInJsonSchemaConstraint(appType As String, constraint As String) As Boolean
        If String.IsNullOrWhiteSpace(constraint) Then Return False
        Select Case If(appType, "").Trim().ToLowerInvariant()
            Case "excel"
                Return constraint.Contains("【Excel支持的22个命令】")
            Case "word"
                Return constraint.Contains("【Word支持的22个命令】")
            Case "powerpoint", "ppt"
                Return constraint.Contains("【PowerPoint支持的22个命令】")
            Case Else
                Return False
        End Select
    End Function
    
    ''' <summary>
    ''' 获取默认的JSON Schema约束（内置硬编码）
    ''' </summary>
    Private Function GetDefaultJsonSchemaConstraint(appType As String) As String
        Select Case appType?.ToLower()
            Case "excel"
                Return GetExcelJsonSchemaConstraintDefault()
            Case "word"
                Return GetWordJsonSchemaConstraintDefault()
            Case "powerpoint"
                Return GetPowerPointJsonSchemaConstraintDefault()
            Case Else
                Return GetExcelJsonSchemaConstraintDefault()
        End Select
    End Function
    
    ''' <summary>
    ''' Excel专用JSON Schema约束（默认值）
    ''' </summary>
    Private Function GetExcelJsonSchemaConstraintDefault() As String
        Return "
【Формат вывода Excel JSON — соблюдать строго】

【Важно】JSON должен возвращаться в формате блока кода Markdown, например:
```json
{""command"": ""ApplyFormula"", ""params"": {...}}
```
Возвращать «голый» JSON без блока кода запрещено!

Ты должен и можешь вернуть только один из двух форматов:

Формат одной команды (обязательно поле command):
```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""A1:B10"", ""formula"": ""=SUM(A1:A10)""}}
```

Формат нескольких команд (обязательно массив commands):
```json
{""commands"": [{""command"": ""WriteData"", ""params"": {""data"": [[""Имя"", ""Возраст""]], ""targetRange"": ""A1""}}, {""command"": ""FormatRange"", ""params"": {""range"": ""A1:B1"", ""style"": ""header""}}]}
```

【24 команд, поддерживаемых Excel】

=== Базовые операции (5) ===
1. ApplyFormula - применить формулу {targetRange:обязательно, formula:обязательно, fillDown:необязательно}. Формула — в англоязычном виде (английские имена функций, запятая как разделитель); для «дата без времени» используй =INT(E2), а не TEXT с кодами DD.MM.YYYY
2. WriteData - записать данные {targetRange:обязательно, data:обязательно(одно значение или двумерный массив)}
3. FormatRange - форматирование {range:обязательно, style:header/total/data, bold/italic/fontSize/backgroundColor/fontColor, borders:true/""all""/""outline""/""none""}
4. CreateChart - создать диаграмму {dataRange:обязательно, type:column/line/pie/bar/scatter/area, title:необязательно, position:необязательно, seriesNames:массив имён серий, categoryAxis:диапазон оси категорий, legendPosition:right/left/top/bottom}
5. CleanData - очистка данных {range:обязательно, operation:removeduplicates/fillempty/trim/replace}

=== Операции с данными (8) ===
6. SortData - сортировка {range:обязательно, sortColumn:номер столбца с 1, order:asc/desc, hasHeader:по умолчанию true}
7. FilterData - фильтр {range:обязательно, column:номер столбца, criteria:условие фильтра, например "">100"", clearFilter:true очищает}
8. RemoveDuplicates - удалить дубликаты {range:обязательно, columns:массив номеров столбцов(необязательно), hasHeader:необязательно}
9. ConditionalFormat - условное форматирование {range:обязательно, rule:highlight/databar/colorscale/iconset, condition:необязательно, color:необязательно}
10. MergeCells - объединить ячейки {range:обязательно, unmerge:true отменяет объединение}
11. AutoFit - автоподбор {range:обязательно, type:columns/rows/both}
12. FindReplace - поиск и замена {range:""all"" или конкретный диапазон, find:обязательно, replace:обязательно, matchCase:необязательно, matchEntireCell:необязательно}
13. CreatePivotTable - сводная таблица {sourceRange:обязательно, targetCell:обязательно, rowFields:массив, valueFields:массив, columnFields:необязательно}

=== Операции с листами (4) ===
14. CreateSheet - создать лист {name:обязательно, position:before/after, referenceSheet:необязательно}
15. DeleteSheet - удалить лист {name:обязательно}
16. RenameSheet - переименовать {oldName:обязательно, newName:обязательно}
17. CopySheet - скопировать лист {sourceName:обязательно, newName:обязательно}

=== Расширенные функции (4) ===
18. InsertRowCol - вставить строки/столбцы {type:row/column, position:номер строки или буква столбца, count:по умолчанию 1}
19. DeleteRowCol - удалить строки/столбцы {type:row/column, position:обязательно, count:по умолчанию 1}
20. HideRowCol - скрыть строки/столбцы {type:row/column, position:обязательно, unhide:true отменяет скрытие}
21. ProtectSheet - защитить лист {sheetName:необязательно, password:необязательно, unprotect:true снимает защиту}

=== Возможности Agent (3) ===
22. TransformData - преобразование данных {sourceRange:обязательно, operation:transpose/split/merge, targetRange/delimiter:необязательно}
23. DataAnalysis - анализ данных {sourceRange:обязательно, type:summary/pivot/groupby/ranking, targetRange/groupBy/valueField/aggregate/topN:по необходимости}
24. GenerateReport - создать отчёт {sourceRange:обязательно, targetSheet/title/includeChart:необязательно}

【Плейсхолдеры динамических диапазонов】
Используй {lastRow} для последней строки, {lastCol} для последнего столбца, {selection} для текущего выделения

【Категорически запрещено】
- Запрещено использовать массивы actions/operations
- Запрещено опускать обёртку params
- Запрещено придумывать другие команды (например, translateText и т. п.)
- Запрещено использовать команды, специфичные для Word/PowerPoint
- Возвращать «голый» JSON без блока кода запрещено

【Неподдерживаемые функции — сообщи пользователю, что нужно использовать кнопки панели】
- Перевод: сообщи пользователю нажать кнопку «AI-перевод» на панели
- Вычитка: сообщи пользователю нажать кнопку «AI-вычитка» на панели

【Приоритет решений】
1. В первую очередь используй перечисленные выше штатные команды
2. VBA отключён: не предлагай ExecuteVBA и не возвращай макросы; составные задачи собирай из ApplyFormula, DataAnalysis, GenerateReport, CreateChart
3. Если запрос неоднозначен, задай уточняющий вопрос на русском языке"
    End Function
    
    ''' <summary>
    ''' Word专用JSON Schema约束（默认值）
    ''' </summary>
    Private Function GetWordJsonSchemaConstraintDefault() As String
        Return "
【Формат вывода Word JSON — соблюдать строго】

【Важно】JSON должен возвращаться в формате блока кода Markdown, например:
```json
{""command"": ""InsertText"", ""params"": {...}}
```
Возвращать «голый» JSON без блока кода запрещено!

Ты должен и можешь вернуть только один из двух форматов:

Формат одной команды:
```json
{""command"": ""InsertText"", ""params"": {""content"": ""вставляемый текст"", ""position"": ""cursor""}}
```

Формат нескольких команд:
```json
{""commands"": [{""command"": ""InsertText"", ""params"": {""content"": ""Заголовок""}}, {""command"": ""FormatText"", ""params"": {""range"": ""selection"", ""bold"": true}}]}
```

【9 команд, поддерживаемых Word】

=== Базовые операции с текстом (4) ===
1. InsertText - вставить текст {content:обязательно, position:cursor/start/end}
2. FormatText - форматировать {range:selection/all, bold/italic/fontSize/fontName/underline/color}
3. ReplaceText - поиск и замена {find:обязательно, replace:обязательно, matchCase:необязательно}
4. DeleteText - удалить текст {range:selection/all}

=== Абзацы и стили (2) ===
5. ApplyStyle - применить стиль {styleName:обязательно, например ""Заголовок 1"", range:selection/paragraph}
6. SetParagraphFormat - формат абзаца {alignment:left/center/right/justify, firstLineIndent/beforeSpacing/afterSpacing}

=== Таблицы (1) ===
7. InsertTable - вставить таблицу {rows:обязательно, cols:обязательно, data:необязательно}

=== Структура документа (1) ===
8. GenerateTOC - создать оглавление {position:start/cursor, levels:1-9}

=== Оформление документа (1) ===
9. BeautifyDocument - оформление {theme:{настройки h1/h2/body}, margins:{top/bottom/left/right}}

【Категорически запрещено】
- Запрещено использовать массивы actions/operations
- Запрещено опускать обёртку params
- Запрещено использовать команды, специфичные для Excel/PowerPoint

【Приоритет решений】
1. Можно использовать только перечисленные выше 9 реализованных команд
2. Для перевода используй кнопку на панели
3. Если запрос неоднозначен, задай уточняющий вопрос на русском языке"
    End Function
    
    ''' <summary>
    ''' PowerPoint专用JSON Schema约束（默认值）
    ''' </summary>
    Private Function GetPowerPointJsonSchemaConstraintDefault() As String
        Return "
【Формат вывода PowerPoint JSON — соблюдать строго】

【Важно】JSON должен возвращаться в формате блока кода Markdown, например:
```json
{""command"": ""InsertSlide"", ""params"": {...}}
```
Возвращать «голый» JSON без блока кода запрещено!

Формат одной команды:
```json
{""command"": ""InsertSlide"", ""params"": {""title"": ""Заголовок"", ""layout"": ""titleAndContent""}}
```

Формат нескольких команд:
```json
{""commands"": [{""command"": ""CreateSlides"", ""params"": {""slides"": [{""title"": ""Первый слайд""}]}}, {""command"": ""AddAnimation"", ""params"": {""effect"": ""fadeIn"", ""scope"": ""all""}}]}
```

【16 команд, поддерживаемых PowerPoint】

=== Операции со слайдами (5) ===
1. InsertSlide - вставить слайд {position:current/end, layout, title, content}
2. DeleteSlide - удалить слайд {slideIndex:обязательно,-1 текущий}
3. DuplicateSlide - дублировать слайд {slideIndex:обязательно}
4. MoveSlide - переместить слайд {fromIndex:обязательно, toIndex:обязательно}
5. CreateSlides - пакетное создание {slides:массив с title/content/layout}

=== Операции с содержимым (3) ===
6. InsertText - вставить текст {content:обязательно, slideIndex:-1 текущий, x/y:необязательно}
7. InsertShape - вставить фигуру {shapeType:обязательно, x:обязательно, y:обязательно}
8. InsertTable - вставить таблицу {rows:обязательно, cols:обязательно, data:необязательно}

=== Стили и анимация (5) ===
9. FormatSlide - форматировать слайд {background, layout}
10. AddAnimation - добавить анимацию {effect:fadeIn/flyIn/zoom/wipe, targetShapes:all/title}
11. ApplyTransition - эффект перехода {transitionType:fade/push/wipe, scope:all/current}
12. BeautifySlides - оформление {scope:all/current, theme:{background/titleFont/bodyFont}}
13. SetSlideLayout - задать макет {layout:title/titleAndContent/blank}

=== Расширенные функции (1) ===
14. AddSpeakerNotes - заметки докладчика {notes:обязательно, slideIndex:необязательно}

=== Тема (1) ===
15. ApplyTheme - применить тему {themeName или themeFile}

=== Резервный VBA (1) ===
16. ExecuteVBA - код VBA {code:обязательно, полный Sub/Function}

【Категорически запрещено】
- Запрещено использовать массивы actions/operations
- Запрещено опускать обёртку params
- Запрещено использовать команды, специфичные для Excel/Word

【Приоритет решений】
1. В первую очередь используй перечисленные выше 16 команд
2. Для сложных задач используй ExecuteVBA
3. Для перевода используй кнопку на панели
4. Если запрос неоднозначен, задай уточняющий вопрос на русском языке"
    End Function
    
    ''' <summary>
    ''' 验证JSON Schema约束是否有效（包含必要的格式要求）
    ''' </summary>
    Private Function IsValidJsonSchemaConstraint(constraint As String) As Boolean
        If String.IsNullOrEmpty(constraint) Then Return False
        
        ' 检查是否包含关键约束词汇
        Dim requiredKeywords() As String = {
            "блок кода Markdown",
            "голый",
            "command",
            "commands",
            "params"
        }
        
        For Each keyword In requiredKeywords
            If Not constraint.Contains(keyword) Then
                Return False
            End If
        Next
        
        Return True
    End Function
    
    ''' <summary>
    ''' 获取纯文本输出约束
    ''' </summary>
    Private Function GetPlainTextConstraint(functionMode As String) As String
        Select Case functionMode?.ToLower()
            Case "continuation"
                Return "
【Важные требования к выводу】
- Выводи только продолжение, без префиксов, суффиксов и пояснений
- Сохраняй язык, стиль и терминологию исходного текста
- Содержание должно быть связным и естественным, не повторяй уже написанное выше"

            Case "template_render"
                Return "
【Важные требования к формату】
- Строго запрещено использовать формат блоков кода Markdown (символы ```)
- Строго запрещено использовать любую Markdown-разметку (#, **, -, > и т. п.)
- Выводи чистый текст, не оборачивай его ни в какие блоки кода
- Не добавляй префиксы, суффиксы, пояснения и служебный текст"

            Case Else
                Return String.Empty
        End Select
    End Function

    ''' <summary>
    ''' 获取配置文件路径
    ''' </summary>
    Private Function GetPromptConfigPath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "prompt_templates.json")
    End Function

    ''' <summary>
    ''' 创建默认配置
    ''' </summary>
    Private Function CreateDefaultConfiguration() As PromptConfiguration
        Dim config As New PromptConfiguration()

        ' Excel应用配置
        config.Applications.Add(CreateExcelConfig())

        ' Word应用配置
        config.Applications.Add(CreateWordConfig())

        ' PowerPoint应用配置
        config.Applications.Add(CreatePowerPointConfig())

        Return config
    End Function

    ''' <summary>
    ''' 创建Excel默认配置
    ''' </summary>
    Private Function CreateExcelConfig() As ApplicationPromptConfig
        Dim excelApp As New ApplicationPromptConfig With {
            .Type = "Excel"
        }

        ' 意图提示词
        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "DATA_ANALYSIS",
            .Content = "Ты помощник по анализу данных Excel. Если запрос понятен, верни фрагмент JSON-кода для выполнения анализа данных. Если запрос неясен, сначала уточни, какой результат анализа нужен. Отвечай только на русском языке."
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "FORMULA_CALC",
            .Content = "Ты помощник по формулам Excel. Если запрос понятен, верни фрагмент JSON-кода для применения формулы. Если запрос неясен, сначала уточни, что именно нужно вычислить."
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "CHART_GEN",
            .Content = "Ты помощник по диаграммам Excel. Если запрос понятен, верни фрагмент JSON-кода для создания диаграммы. Порекомендуй подходящий тип диаграммы по особенностям данных (столбчатая, линейная, круговая и т. д.)."
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "DATA_CLEANING",
            .Content = "Ты помощник по очистке данных Excel. Если запрос понятен, верни фрагмент JSON-кода для очистки данных. Поддерживаются удаление дубликатов, заполнение пустых значений, удаление лишних пробелов и т. п."
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "GENERAL_QUERY",
            .Content = "Ты помощник по Excel. Если запрос понятен и выполним, верни фрагмент JSON-кода; если запрос неясен, сначала уточни; на простые приветствия и вопросы отвечай прямо на русском языке."
        })

        ' JSON Schema约束
        excelApp.JsonSchemaConstraint = "
【Формат вывода Excel JSON — соблюдать строго】

【Важно】JSON должен возвращаться в формате блока кода Markdown, например:
```json
{""command"": ""ApplyFormula"", ""params"": {...}}
```
Возвращать «голый» JSON без блока кода запрещено!

Ты должен и можешь вернуть только один из двух форматов:

Формат одного JSON-кода:
```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1"", ""fillDown"": true}}
```

Формат нескольких JSON-кодов:
```json
{""commands"": [{""command"": ""ApplyFormula"", ""params"": {...}}, {...}]}
```

【Категорически запрещено】
- Запрещено использовать массив actions
- Запрещено использовать массив operations
- Запрещено опускать обёртку params
- Запрещено возвращать типы command, не указанные ниже
- Возвращать «голый» JSON без блока кода запрещено

【Типы Excel command】
1. ApplyFormula - применить формулу (targetRange, formula, fillDown)
2. WriteData - записать данные (targetRange, data)
3. FormatRange - форматировать диапазон (range, style, bold, fontSize, fontColor, bgColor)
4. CreateChart - создать диаграмму (dataRange, chartType, title)
5. CleanData - очистить данные (range, operation: removeDuplicates/fillEmpty/trim)

【Плейсхолдеры динамических диапазонов】
Используй {lastRow} для последней строки, система автоматически подставит фактический номер"

        Return excelApp
    End Function

    ''' <summary>
    ''' 创建Word默认配置
    ''' </summary>
    Private Function CreateWordConfig() As ApplicationPromptConfig
        Dim wordApp As New ApplicationPromptConfig With {
            .Type = "Word"
        }

        ' 意图提示词
        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "DOCUMENT_EDIT",
            .Content = "Ты помощник по редактированию документов Word. Если запрос понятен, верни фрагмент JSON-кода для операции редактирования документа. Поддерживаются вставка, удаление и замена текста и т. п. Отвечай только на русском языке."
        })

        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "TOC_GENERATION",
            .Content = "Ты помощник по созданию оглавления Word. Если пользователь говорит «создай оглавление» или «добавь оглавление», сразу верни команду GenerateTOC. Если нужно уточнение, спроси: оглавление в начале или в текущем месте? сколько уровней заголовков показывать?"
        })

        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "FORMAT_STYLE",
            .Content = "Ты помощник по форматированию и стилям Word. Если нужно оформить документ, верни команду BeautifyDocument. Поддерживаются применение темы, настройка шрифта, интервалов и т. п."
        })

        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "GENERAL_QUERY",
            .Content = "Ты помощник по Word. Если запрос понятен и выполним, верни фрагмент JSON-кода; если запрос неясен, сначала уточни; на простые приветствия и вопросы отвечай прямо на русском языке."
        })

        ' 功能模式提示词
        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "proofread",
            .Content = "Ты помощник по вычитке содержимого Word. Проверь приведённый ниже текст на опечатки, пунктуационные ошибки и неуместные переносы строк и предложи исправления.

【Формат вывода】
Обязательно верни JSON-массив, каждый элемент содержит:
[{""paraIndex"": 0, ""original"": ""фрагмент оригинала"", ""corrected"": ""исправленный текст"", ""reason"": ""краткое обоснование правки""}]

Если исправлять нечего, верни пустой массив []"
        })

        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "reformat",
            .Content = "Ты помощник по вёрстке Word. Я передаю абзацы документа, помоги оптимизировать оформление.

【Правила оформления】
1. Используй единый шрифт Times New Roman
2. Размер основного текста 12pt, заголовки — по уровню (например, 16pt/14pt)
3. Отступ первой строки абзаца 2 знака
4. Междустрочный интервал 1,5

【Формат вывода】
Обязательно верни JSON-массив, формат следующий:
[{""paraIndex"": 0, ""formatting"": {""fontNameCN"": ""Times New Roman"", ""fontNameEN"": ""Times New Roman"", ""fontSize"": 12, ""bold"": false, ""alignment"": ""left"", ""firstLineIndent"": 2, ""lineSpacing"": 1.5}}]"
        })

        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "continuation",
            .Content = "Ты профессиональный помощник по письму. По предоставленному контексту естественно продолжи текст.

Требования:
1. Сохраняй язык, стиль, тон и терминологию оригинала
2. Содержание должно быть связным и естественным, не повторяй написанное выше
3. Выводи только продолжение, без пояснений, префиксов и разметки
4. Если контекста недостаточно, можешь разумно предположить, но осторожно
5. Длина продолжения умеренная, примерно 100–300 знаков, если пользователь не требует иного"
        })

        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "template_render",
            .Content = "Ты профессиональный помощник по генерации содержимого документов. Тебе нужно сгенерировать новое содержимое по структуре шаблона (JSON) и стилю, предоставленным пользователем.

【Описание JSON-структуры шаблона】
- elements: массив элементов документа, каждый содержит type(тип), text(текст), styleName(имя стиля), formatting(детали формата)
- formatting содержит: fontName(шрифт), fontSize(размер), bold(полужирный), italic(курсив), alignment(выравнивание) и т. п.

【Требования к генерации】
1. Строго соблюдай иерархию шаблона (например, соотношение заголовка, подзаголовка и основного текста)
2. Сохраняй тон, терминологию и стиль шаблона
3. Ориентируйся на размеры шрифта в шаблоне для оценки важности (больший размер = заголовок, меньший = основной текст)
4. Содержание должно быть профессиональным, связным и практичным
5. Организуй вывод в порядке элементов шаблона"
        })

        ' JSON Schema约束
        wordApp.JsonSchemaConstraint = "
【Формат вывода Word JSON — соблюдать строго】

【Важно】JSON должен возвращаться в формате блока кода Markdown, например:
```json
{""command"": ""InsertText"", ""params"": {...}}
```
Возвращать «голый» JSON без блока кода запрещено!

Ты должен и можешь вернуть только один из двух форматов:

Формат одного JSON-кода:
```json
{""command"": ""InsertText"", ""params"": {""content"": ""содержимое"", ""position"": ""cursor""}}
```

Формат нескольких JSON-кодов:
```json
{""commands"": [{""command"": ""InsertText"", ""params"": {...}}, {...}]}
```

【Типы Word command】
1. InsertText - вставить текст (content, position: cursor/start/end)
2. FormatText - форматировать текст (range: selection/all, bold, italic, fontSize, fontName)
3. ReplaceText - заменить текст (find, replace, matchCase, matchWholeWord)
4. InsertTable - вставить таблицу (rows, cols, data)
5. ApplyStyle - применить стиль (styleName, range)
6. GenerateTOC - создать оглавление (position: start/cursor, levels: 1-9)
7. BeautifyDocument - оформить документ (theme, margins)

【Категорически запрещено】
- Запрещено использовать команды Excel (WriteData, ApplyFormula и т. п.)
- Запрещено использовать команды PPT (InsertSlide, CreateSlides и т. п.)
- Запрещено возвращать типы command, не указанные выше
- Возвращать «голый» JSON без блока кода запрещено"

        Return wordApp
    End Function

    ''' <summary>
    ''' 创建PowerPoint默认配置
    ''' </summary>
    Private Function CreatePowerPointConfig() As ApplicationPromptConfig
        Dim pptApp As New ApplicationPromptConfig With {
            .Type = "PowerPoint"
        }

        ' 意图提示词
        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "SLIDE_CREATE",
            .Content = "Ты помощник по созданию слайдов PowerPoint. Когда пользователь говорит «создай N слайдов», используй команду CreateSlides для пакетного создания; когда говорит «добавь слайд» — команду InsertSlide для одного слайда. Отвечай только на русском языке."
        })

        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "ANIMATION_EFFECT",
            .Content = "Ты помощник по анимации PowerPoint. Поддерживается добавление анимации входа (fadeIn, flyIn, zoom, wipe и т. д.) и выхода. Анимацию можно применять ко всем фигурам или только к заголовку."
        })

        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "TRANSITION_EFFECT",
            .Content = "Ты помощник по переходам PowerPoint. Поддерживается применение эффекта перехода (fade, push, wipe, split и т. д.) к текущему слайду или ко всем слайдам."
        })

        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "GENERAL_QUERY",
            .Content = "Ты помощник по PowerPoint. Если запрос понятен и выполним, верни фрагмент JSON-кода; если запрос неясен, сначала уточни; на простые приветствия и вопросы отвечай прямо на русском языке."
        })

        ' 功能模式提示词
        pptApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "continuation",
            .Content = "Ты профессиональный помощник по написанию презентаций. По предоставленному контексту слайдов естественно продолжи содержимое.

Требования:
1. Сохраняй стиль и терминологию исходных слайдов
2. Содержание должно быть кратким и выразительным, пригодным для показа
3. Выводи только продолжение, без пояснений
4. Объём каждой страницы — в разумных пределах"
        })

        pptApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "template_render",
            .Content = "Ты профессиональный помощник по генерации содержимого презентаций. Сгенерируй новое содержимое по структуре шаблона PPT.

【Описание структуры шаблона PPT】
- slides: массив слайдов, каждый содержит layout(макет) и elements(список элементов)
- elements содержит: type(тип), text(текст), formatting(формат)

【Требования к генерации】
1. Генерируй содержимое в соответствии с числом и макетом слайдов шаблона
2. Заголовки должны быть краткими и выразительными, тезисы — чёткими
3. Содержание должно подходить для показа, избегай длинных абзацев"
        })

        ' JSON Schema约束
        pptApp.JsonSchemaConstraint = "
【Формат вывода PowerPoint JSON — соблюдать строго】

【Важно】JSON должен возвращаться в формате блока кода Markdown, например:
```json
{""command"": ""InsertSlide"", ""params"": {...}}
```
Возвращать «голый» JSON без блока кода запрещено!

Ты должен и можешь вернуть только один из двух форматов:

Формат одного JSON-кода:
```json
{""command"": ""InsertSlide"", ""params"": {""title"": ""Заголовок"", ""content"": ""Содержимое""}}
```

Формат нескольких JSON-кодов:
```json
{""commands"": [{""command"": ""CreateSlides"", ""params"": {...}}, {...}]}
```

【Типы PowerPoint command】
1. InsertSlide - вставить один слайд (title, content, layout)
2. CreateSlides - пакетно создать несколько слайдов (массив slides, каждый с title/content/layout) 【рекомендуется для нескольких слайдов】
3. InsertText - вставить текст (content, slideIndex)
4. InsertShape - вставить фигуру (shapeType, text)
5. FormatSlide - форматировать слайд (slideIndex, background, theme)
6. InsertTable - вставить таблицу (rows, cols, data)
7. AddAnimation - добавить анимацию (effect: fadeIn/flyIn/zoom/wipe, targetShapes: all/title)
8. ApplyTransition - применить переход (transitionType: fade/push/wipe/split, scope: all/current)
9. BeautifySlides - оформить слайды (theme, colorScheme)

【Пояснение slideIndex】
- -1 или пусто означает текущий слайд
- 0 означает первый слайд

【Категорически запрещено】
- Запрещено использовать команды Excel (WriteData, ApplyFormula и т. п.)
- Запрещено использовать команды Word (Word-версию InsertText, GenerateTOC и т. п.)
- Запрещено возвращать типы command, не указанные выше
- Возвращать «голый» JSON без блока кода запрещено"

        Return pptApp
    End Function
End Class
