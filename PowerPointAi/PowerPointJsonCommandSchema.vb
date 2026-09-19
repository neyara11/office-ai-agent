' PowerPointAi\PowerPointJsonCommandSchema.vb
' PowerPoint JSON命令Schema定义和校验

Imports System.Diagnostics
Imports Newtonsoft.Json.Linq

''' <summary>
''' PowerPoint JSON命令Schema和校验器
''' </summary>
Public Class PowerPointJsonCommandSchema

    ''' <summary>
    ''' 仅列出 ExecutePPTCommand 已实现的命令，避免向 Agent 暴露无执行器的假能力。
    ''' 幻灯片操作: InsertSlide, DeleteSlide, DuplicateSlide, MoveSlide, CreateSlides
    ''' 内容操作: InsertText, InsertShape, InsertTable
    ''' 样式和动画: FormatSlide, AddAnimation, ApplyTransition, BeautifySlides, SetSlideLayout
    ''' 高级功能: AddSpeakerNotes
    ''' 主题: ApplyTheme
    ''' VBA回退: ExecuteVBA
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertSlide",
        "DeleteSlide",
        "DuplicateSlide",
        "MoveSlide",
        "CreateSlides",
        "InsertText",
        "InsertShape",
        "InsertTable",
        "FormatSlide",
        "AddAnimation",
        "ApplyTransition",
        "BeautifySlides",
        "SetSlideLayout",
        "AddSpeakerNotes",
        "ApplyTheme",
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
  ""command"": ""InsertSlide"",
  ""params"": {
    ""position"": ""end"",
    ""layout"": ""title"",
    ""title"": ""Заголовок слайда""
  }
}
```

Формат 2 — несколько команд (пакетная операция):
```json
{
  ""commands"": [
    {""command"": ""InsertSlide"", ""params"": {""title"": ""Первый слайд""}},
    {""command"": ""AddAnimation"", ""params"": {""effect"": ""fadeIn"", ""scope"": ""all""}}
  ]
}
```

【Категорически запрещённые форматы】
- запрещено {""command"": ""xxx"", ""actions"": [...]}
- запрещено {""command"": ""xxx"", ""title"": ""...""} (отсутствует обёртка params)
- запрещено {""operations"": [...]}

【16 команд, поддерживаемых PowerPoint】

=== Операции со слайдами (5) ===
1. InsertSlide — вставить слайд {position:current/end, layout:title/titleAndContent/blank, title:необязательно, content:необязательно}
2. DeleteSlide — удалить слайд {slideIndex:обязательно, -1 означает текущий}
3. DuplicateSlide — дублировать слайд {slideIndex:обязательно, insertAfter:необязательно}
4. MoveSlide — переместить слайд {fromIndex:обязательно, toIndex:обязательно}
5. CreateSlides — пакетное создание {slides:массив, каждый содержит title/content/layout}

=== Операции с содержимым (3) ===
6. InsertText — вставить текст {content:обязательно, slideIndex:-1 означает текущий, x/y:необязательное положение}
7. InsertShape — вставить фигуру {shapeType:rectangle/oval/arrow и т. п., x:обязательно, y:обязательно, width/height:необязательно}
8. InsertTable — вставить таблицу {rows:обязательно, cols:обязательно, data:необязательно, slideIndex:необязательно}

=== Стили и анимация (5) ===
9. FormatSlide — форматировать слайд {slideIndex:необязательно, background:цвет/путь к изображению, layout:необязательно}
10. AddAnimation — добавить анимацию {effect:fadeIn/flyIn/zoom/wipe/appear, slideIndex:необязательно, targetShapes:all/title/content}
11. ApplyTransition — эффект перехода {transitionType:fade/push/wipe/split, scope:all/current, duration:секунды}
12. BeautifySlides — оформить слайды {scope:all/current, theme:{background/titleFont/bodyFont}}
13. SetSlideLayout — задать макет {slideIndex:необязательно, layout:title/titleAndContent/twoContent/blank/comparison}

=== Расширенные функции (1) ===
14. AddSpeakerNotes — заметки докладчика {slideIndex:необязательно, notes:обязательно}

=== Тема (1) ===
15. ApplyTheme — применить тему {themeName:необязательное имя встроенной темы, themeFile:необязательный путь к файлу темы}

=== Резервный VBA (1) ===
16. ExecuteVBA — выполнить код VBA {code:обязательно, полный код Sub или Function}
    Когда перечисленных команд недостаточно, сгенерируй код VBA как резервный вариант

【Пояснение slideIndex】
- -1 или пропуск означает текущий слайд
- 0 означает первый слайд
- положительное число означает конкретный индекс слайда

【Приоритет решений】
1. В первую очередь используй перечисленные выше 16 команд
2. Для сложных задач, которые нельзя решить командой, используй ExecuteVBA для генерации кода VBA
3. Для перевода сообщи пользователю использовать кнопку ""Перевод"" на панели инструментов
4. Если запрос неоднозначен, задай вопрос напрямую"
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
{{""command"": ""InsertSlide"", ""params"": {{""position"": ""end"", ""title"": ""Заголовок""}}}}

Несколько команд:
{{""commands"": [{{""command"": ""InsertSlide"", ""params"": {{""title"": ""Первый слайд""}}}}, {{""command"": ""InsertText"", ""params"": {{""content"": ""содержимое""}}}}]}}

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
    ''' 标准化JSON命令结构 - 将扁平结构自动包装到params中
    ''' </summary>
    Public Shared Function NormalizeCommandStructure(json As JObject) As JObject
        Try
            ' 检查是否已有params字段
            If json("params") IsNot Nothing Then
                Return json
            End If

            ' 检查是否有command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                Return json
            End If

            ' 需要移到params中的字段
            Dim topLevelFields As String() = {"command"}
            Dim paramsFields As New JObject()

            For Each prop In json.Properties().ToList()
                If Not topLevelFields.Contains(prop.Name, StringComparer.OrdinalIgnoreCase) Then
                    paramsFields(prop.Name) = prop.Value
                    json.Remove(prop.Name)
                End If
            Next

            If paramsFields.Count > 0 Then
                json("params") = paramsFields
            End If

            Return json
        Catch ex As Exception
            Debug.WriteLine($"Ошибка NormalizeCommandStructure: {ex.Message}")
            Return json
        End Try
    End Function

    ''' <summary>
    ''' 校验JSON命令是否有效
    ''' </summary>
    Public Shared Function ValidateCommand(json As JObject, ByRef errorMessage As String) As Boolean
        Try
            errorMessage = ""
            
            ' 首先进行结构标准化
            json = NormalizeCommandStructure(json)
            
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                errorMessage = "Отсутствует поле command"
                Return False
            End If
            
            If Not SupportedCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase)) Then
                errorMessage = $"Неподдерживаемая команда: {command}. Поддерживаемые команды: {String.Join(", ", SupportedCommands)}"
                Return False
            End If
            
            Dim params = json("params")
            If params Is Nothing Then
                errorMessage = "Отсутствует поле params"
                Return False
            End If
            
            ' 根据命令类型校验参数
            Select Case command.ToLower()
                ' === 幻灯片操作 ===
                Case "insertslide"
                    Return ValidateInsertSlide(params, errorMessage)
                Case "deleteslide"
                    Return ValidateDeleteSlide(params, errorMessage)
                Case "duplicateslide"
                    Return ValidateDuplicateSlide(params, errorMessage)
                Case "moveslide"
                    Return ValidateMoveSlide(params, errorMessage)
                Case "createslides"
                    Return ValidateCreateSlides(params, errorMessage)
                ' === 内容操作 ===
                Case "inserttext"
                    Return ValidateInsertText(params, errorMessage)
                Case "insertshape"
                    Return ValidateInsertShape(params, errorMessage)
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                ' === 样式和动画 ===
                Case "formatslide"
                    Return ValidateFormatSlide(params, errorMessage)
                Case "addanimation"
                    Return ValidateAddAnimation(params, errorMessage)
                Case "applytransition"
                    Return ValidateApplyTransition(params, errorMessage)
                Case "beautifyslides"
                    Return ValidateBeautifySlides(params, errorMessage)
                Case "setslidelayout"
                    Return ValidateSetSlideLayout(params, errorMessage)
                ' === 高级功能 ===
                Case "addspeakernotes"
                    Return ValidateAddSpeakerNotes(params, errorMessage)
                ' === 母版和主题 ===
                Case "applytheme"
                    Return ValidateApplyTheme(params, errorMessage)
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

    Private Shared Function ValidateInsertSlide(params As JToken, ByRef errorMessage As String) As Boolean
        ' InsertSlide参数都是可选的，基本验证通过
        Return True
    End Function

    Private Shared Function ValidateInsertText(params As JToken, ByRef errorMessage As String) As Boolean
        Dim content = params("content")?.ToString()
        ' 兼容处理：大模型在GENERAL_QUERY模式下可能返回text而不是content
        If String.IsNullOrEmpty(content) Then
            content = params("text")?.ToString()
            ' 如果text存在，将其复制到content字段以便后续执行
            If Not String.IsNullOrEmpty(content) AndAlso params.Type = JTokenType.Object Then
                CType(params, JObject)("content") = content
            End If
        End If
        If String.IsNullOrEmpty(content) Then
            errorMessage = "InsertText: отсутствует параметр content"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertShape(params As JToken, ByRef errorMessage As String) As Boolean
        Dim shapeType = params("shapeType")?.ToString()
        If String.IsNullOrEmpty(shapeType) Then
            errorMessage = "InsertShape: отсутствует параметр shapeType"
            Return False
        End If
        
        If params("x") Is Nothing OrElse params("y") Is Nothing Then
            errorMessage = "InsertShape: отсутствует параметр x или y"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateFormatSlide(params As JToken, ByRef errorMessage As String) As Boolean
        ' FormatSlide至少需要一个格式化属性
        If params("background") Is Nothing AndAlso params("transition") Is Nothing AndAlso
           params("layout") Is Nothing Then
            errorMessage = "FormatSlide: требуется хотя бы один атрибут форматирования (background/transition/layout)"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertTable(params As JToken, ByRef errorMessage As String) As Boolean
        Dim rows = params("rows")
        Dim cols = params("cols")
        
        If rows Is Nothing OrElse cols Is Nothing Then
            errorMessage = "InsertTable: отсутствует параметр rows или cols"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateCreateSlides(params As JToken, ByRef errorMessage As String) As Boolean
        Dim slides = params("slides")
        If slides Is Nothing OrElse slides.Type <> JTokenType.Array Then
            errorMessage = "CreateSlides: отсутствует параметр-массив slides"
            Return False
        End If
        
        Dim slidesArray = CType(slides, JArray)
        If slidesArray.Count = 0 Then
            errorMessage = "CreateSlides: массив slides не может быть пустым"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateAddAnimation(params As JToken, ByRef errorMessage As String) As Boolean
        Dim effect = params("effect")?.ToString()
        If String.IsNullOrEmpty(effect) Then
            errorMessage = "AddAnimation: отсутствует параметр effect"
            Return False
        End If
        
        Dim validEffects = {"fadein", "flyin", "zoom", "wipe", "appear", "float"}
        If Not validEffects.Contains(effect.ToLower()) Then
            errorMessage = $"AddAnimation: недопустимый параметр effect: {effect}. Допустимые значения: fadeIn, flyIn, zoom, wipe, appear, float"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateApplyTransition(params As JToken, ByRef errorMessage As String) As Boolean
        Dim transType = params("transitionType")?.ToString()
        If String.IsNullOrEmpty(transType) Then
            errorMessage = "ApplyTransition: отсутствует параметр transitionType"
            Return False
        End If
        
        Dim validTypes = {"fade", "push", "wipe", "split", "reveal", "random"}
        If Not validTypes.Contains(transType.ToLower()) Then
            errorMessage = $"ApplyTransition: недопустимый параметр transitionType: {transType}. Допустимые значения: fade, push, wipe, split, reveal, random"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateBeautifySlides(params As JToken, ByRef errorMessage As String) As Boolean
        ' BeautifySlides参数都是可选的，但至少需要一个
        If params("scope") Is Nothing AndAlso params("theme") Is Nothing Then
            errorMessage = "BeautifySlides: требуется параметр scope или theme"
            Return False
        End If
        Return True
    End Function

#Region "新增命令验证方法"

    Private Shared Function ValidateDeleteSlide(params As JToken, ByRef errorMessage As String) As Boolean
        If params("slideIndex") Is Nothing Then
            errorMessage = "DeleteSlide: отсутствует параметр slideIndex"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateDuplicateSlide(params As JToken, ByRef errorMessage As String) As Boolean
        If params("slideIndex") Is Nothing Then
            errorMessage = "DuplicateSlide: отсутствует параметр slideIndex"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateMoveSlide(params As JToken, ByRef errorMessage As String) As Boolean
        If params("fromIndex") Is Nothing Then
            errorMessage = "MoveSlide: отсутствует параметр fromIndex"
            Return False
        End If
        If params("toIndex") Is Nothing Then
            errorMessage = "MoveSlide: отсутствует параметр toIndex"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateSetSlideLayout(params As JToken, ByRef errorMessage As String) As Boolean
        Dim layout = params("layout")?.ToString()
        If String.IsNullOrEmpty(layout) Then
            errorMessage = "SetSlideLayout: отсутствует параметр layout"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateAddSpeakerNotes(params As JToken, ByRef errorMessage As String) As Boolean
        Dim notes = params("notes")?.ToString()
        If String.IsNullOrEmpty(notes) Then
            errorMessage = "AddSpeakerNotes: отсутствует параметр notes"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateApplyTheme(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要themeName或themeFile之一
        If params("themeName") Is Nothing AndAlso params("themeFile") Is Nothing Then
            errorMessage = "ApplyTheme: требуется параметр themeName или themeFile"
            Return False
        End If
        Return True
    End Function

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

End Class
