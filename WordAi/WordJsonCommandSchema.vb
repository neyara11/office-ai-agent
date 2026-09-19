' WordAi\WordJsonCommandSchema.vb
' Word JSON命令Schema定义和校验

Imports System.Diagnostics
Imports Newtonsoft.Json.Linq

''' <summary>
''' Word JSON命令Schema和校验器
''' </summary>
Public Class WordJsonCommandSchema

    ''' <summary>
    ''' 仅列出 ExecuteWordCommand 已实现的命令，避免向 Agent 暴露无执行器的假能力。
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertText",
        "FormatText",
        "ReplaceText",
        "DeleteText",
        "ApplyStyle",
        "SetParagraphFormat",
        "InsertTable",
        "GenerateTOC",
        "BeautifyDocument"
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
  ""command"": ""InsertText"",
  ""params"": {
    ""position"": ""cursor"",
    ""content"": ""текст для вставки""
  }
}
```

Формат 2 — несколько команд (пакетная операция):
```json
{
  ""commands"": [
    {""command"": ""InsertText"", ""params"": {""position"": ""cursor"", ""content"": ""Заголовок""}},
    {""command"": ""FormatText"", ""params"": {""range"": ""selection"", ""bold"": true, ""fontSize"": 16}}
  ]
}
```

【Категорически запрещённые форматы】
- запрещено {""command"": ""xxx"", ""actions"": [...]}
- запрещено {""command"": ""xxx"", ""content"": ""...""} (отсутствует обёртка params)
- запрещено {""operations"": [...]}

【9 команд, поддерживаемых Word】

=== Базовые операции с текстом (4) ===
1. InsertText — вставить текст {content:обязательно, position:cursor/start/end}
2. FormatText — форматировать {range:selection/all, bold/italic/fontSize/fontName/underline/color}
3. ReplaceText — поиск и замена {find:обязательно, replace:обязательно, matchCase:необязательно, matchWholeWord:необязательно}
4. DeleteText — удалить текст {range:selection/all}

=== Абзацы и стили (2) ===
5. ApplyStyle — применить стиль {styleName:обязательно, например ""Заголовок 1""/""Обычный"", range:selection/paragraph}
6. SetParagraphFormat — формат абзаца {alignment:left/center/right/justify, firstLineIndent:необязательно, beforeSpacing/afterSpacing:необязательно}

=== Таблицы (1) ===
7. InsertTable — вставить таблицу {rows:обязательно, cols:обязательно, data:необязательно (двумерный массив), style:необязательно}

=== Структура документа (1) ===
8. GenerateTOC — создать оглавление {position:start/cursor, levels:1-9, includePageNumbers:по умолчанию true}

=== Оформление документа (1) ===
9. BeautifyDocument — оформить документ {theme:{настройки шрифтов h1/h2/body}, margins:{top/bottom/left/right}}

【Приоритет решений】
1. Для выполнения запроса можно использовать только перечисленные выше 9 реализованных команд
2. Для перевода сообщи пользователю использовать кнопку ""Перевод"" на панели инструментов
3. Если запрос неоднозначен, задай вопрос напрямую"
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
{{""command"": ""InsertText"", ""params"": {{""position"": ""cursor"", ""content"": ""текст""}}}}

Несколько команд:
{{""commands"": [{{""command"": ""InsertText"", ""params"": {{""content"": ""содержимое 1""}}}}, {{""command"": ""FormatText"", ""params"": {{""range"": ""selection"", ""bold"": true}}}}]}}

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
                ' === 基础文本操作 ===
                Case "inserttext"
                    Return ValidateInsertText(params, errorMessage)
                Case "formattext"
                    Return ValidateFormatText(params, errorMessage)
                Case "replacetext"
                    Return ValidateReplaceText(params, errorMessage)
                Case "deletetext"
                    Return ValidateDeleteText(params, errorMessage)
                ' === 段落和样式 ===
                Case "applystyle"
                    Return ValidateApplyStyle(params, errorMessage)
                Case "setparagraphformat"
                    Return ValidateSetParagraphFormat(params, errorMessage)
                ' === 表格操作 ===
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                ' === 文档结构 ===
                Case "generatetoc"
                    Return ValidateGenerateTOC(params, errorMessage)
                ' === 文档美化 ===
                Case "beautifydocument"
                    Return ValidateBeautifyDocument(params, errorMessage)
                Case Else
                    Return True
            End Select
            
        Catch ex As Exception
            errorMessage = $"Исключение при проверке JSON: {ex.Message}"
            Return False
        End Try
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

    Private Shared Function ValidateFormatText(params As JToken, ByRef errorMessage As String) As Boolean
        ' FormatText至少需要一个格式化属性
        If params("bold") Is Nothing AndAlso params("italic") Is Nothing AndAlso 
           params("fontSize") Is Nothing AndAlso params("fontName") Is Nothing AndAlso
           params("underline") Is Nothing Then
            errorMessage = "FormatText: требуется хотя бы один атрибут форматирования (bold/italic/fontSize/fontName/underline)"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateReplaceText(params As JToken, ByRef errorMessage As String) As Boolean
        Dim find = params("find")?.ToString()
        If String.IsNullOrEmpty(find) Then
            errorMessage = "ReplaceText: отсутствует параметр find"
            Return False
        End If
        
        If params("replace") Is Nothing Then
            errorMessage = "ReplaceText: отсутствует параметр replace"
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

    Private Shared Function ValidateApplyStyle(params As JToken, ByRef errorMessage As String) As Boolean
        Dim styleName = params("styleName")?.ToString()
        If String.IsNullOrEmpty(styleName) Then
            errorMessage = "ApplyStyle: отсутствует параметр styleName"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateGenerateTOC(params As JToken, ByRef errorMessage As String) As Boolean
        ' GenerateTOC参数都是可选的
        Dim levels = params("levels")
        If levels IsNot Nothing Then
            Dim levelValue = levels.Value(Of Integer)()
            If levelValue < 1 OrElse levelValue > 9 Then
                errorMessage = "GenerateTOC: параметр levels должен быть в диапазоне 1-9"
                Return False
            End If
        End If
        Return True
    End Function

    Private Shared Function ValidateBeautifyDocument(params As JToken, ByRef errorMessage As String) As Boolean
        ' BeautifyDocument至少需要theme或margins之一
        If params("theme") Is Nothing AndAlso params("margins") Is Nothing Then
            errorMessage = "BeautifyDocument: требуется параметр theme или margins"
            Return False
        End If
        Return True
    End Function

#Region "新增命令验证方法"

    Private Shared Function ValidateDeleteText(params As JToken, ByRef errorMessage As String) As Boolean
        ' DeleteText可以用range指定范围,默认删除选中内容
        Return True
    End Function

    Private Shared Function ValidateSetParagraphFormat(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要一个段落格式属性
        If params("alignment") Is Nothing AndAlso params("firstLineIndent") Is Nothing AndAlso
           params("beforeSpacing") Is Nothing AndAlso params("afterSpacing") Is Nothing Then
            errorMessage = "SetParagraphFormat: требуется хотя бы один атрибут формата"
            Return False
        End If
        Return True
    End Function

#End Region

End Class
