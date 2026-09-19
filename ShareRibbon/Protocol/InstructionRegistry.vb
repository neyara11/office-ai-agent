' ShareRibbon\Protocol\InstructionRegistry.vb
' 指令注册表 - 维护所有支持的指令及其参数Schema

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' 指令注册表 - 维护所有支持的指令及其参数Schema
''' </summary>
Public Class InstructionRegistry

    Private Shared ReadOnly _registry As New Dictionary(Of String, InstructionDefinition)()
    Private Shared _isInitialized As Boolean = False

    ''' <summary>
    ''' 初始化注册表（首次访问时自动调用）
    ''' </summary>
    Public Shared Sub Initialize()
        If _isInitialized Then Return

        ' ========== 排版指令 ==========

        Register("setParagraphStyle", New InstructionDefinition With {
            .Operation = "setParagraphStyle",
            .Category = "reformat",
            .DisplayName = "Настроить стиль абзаца",
            .Description = "Задать имя стиля, шрифт, выравнивание, межстрочный интервал и другое для абзаца",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.styleName", "params.font", "params.alignment", "params.spacing", "params.indent", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "index", "range", "selection"})},
                {"target.selector", ParamType.StringType()},
                {"target.index", ParamType.NumberType()},
                {"params.styleName", ParamType.StringType()},
                {"params.font.name", ParamType.StringType()},
                {"params.font.size", ParamType.NumberType()},
                {"params.font.bold", ParamType.BooleanType()},
                {"params.font.italic", ParamType.BooleanType()},
                {"params.font.color", ParamType.StringType()},
                {"params.alignment", ParamType.EnumType(New String() {"left", "center", "right", "justify"})},
                {"params.spacing.before", ParamType.NumberType()},
                {"params.spacing.after", ParamType.NumberType()},
                {"params.spacing.line", ParamType.NumberType()},
                {"params.indent.firstLine", ParamType.NumberType()},
                {"params.indent.left", ParamType.NumberType()},
                {"params.indent.right", ParamType.NumberType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"paragraph", "selection"}
        })

        Register("setCharacterFormat", New InstructionDefinition With {
            .Operation = "setCharacterFormat",
            .Category = "reformat",
            .DisplayName = "Настроить формат символов",
            .Description = "Задать формат символов выделенного текста (шрифт, цвет, полужирный и другое)",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.font", "params.color", "params.bold", "params.italic", "params.underline", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "range", "selection"})},
                {"target.selector", ParamType.StringType()},
                {"params.font.name", ParamType.StringType()},
                {"params.font.size", ParamType.NumberType()},
                {"params.color", ParamType.StringType()},
                {"params.bold", ParamType.BooleanType()},
                {"params.italic", ParamType.BooleanType()},
                {"params.underline", ParamType.EnumType(New String() {"none", "single", "double", "wavy"})}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"range", "selection"}
        })

        Register("insertTable", New InstructionDefinition With {
            .Operation = "insertTable",
            .Category = "reformat",
            .DisplayName = "Вставить таблицу",
            .Description = "Вставить таблицу в указанную позицию",
            .RequiredParams = New List(Of String) From {"target", "params.rows", "params.cols"},
            .OptionalParams = New List(Of String) From {"params.style", "params.data", "params.headerRow", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"position", "after", "before"})},
                {"target.position", ParamType.EnumType(New String() {"cursor", "end", "start"})},
                {"params.rows", ParamType.NumberType()},
                {"params.cols", ParamType.NumberType()},
                {"params.style", ParamType.StringType()},
                {"params.headerRow", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"position"}
        })

        Register("formatTable", New InstructionDefinition With {
            .Operation = "formatTable",
            .Category = "reformat",
            .DisplayName = "Форматировать таблицу",
            .Description = "Форматировать стиль, границы и другое у существующей таблицы",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.style", "params.borders", "params.headerRow", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"index", "semantic"})},
                {"target.index", ParamType.NumberType()},
                {"target.selector", ParamType.StringType()},
                {"params.style", ParamType.StringType()},
                {"params.borders", ParamType.BooleanType()},
                {"params.headerRow", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"table"}
        })

        Register("setPageSetup", New InstructionDefinition With {
            .Operation = "setPageSetup",
            .Category = "reformat",
            .DisplayName = "Параметры страницы",
            .Description = "Задать поля страницы, ориентацию и размер бумаги",
            .RequiredParams = New List(Of String) From {},
            .OptionalParams = New List(Of String) From {"params.margins", "params.orientation", "params.paperSize", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.margins.top", ParamType.NumberType()},
                {"params.margins.bottom", ParamType.NumberType()},
                {"params.margins.left", ParamType.NumberType()},
                {"params.margins.right", ParamType.NumberType()},
                {"params.orientation", ParamType.EnumType(New String() {"portrait", "landscape"})},
                {"params.paperSize", ParamType.StringType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"document"}
        })

        Register("insertBreak", New InstructionDefinition With {
            .Operation = "insertBreak",
            .Category = "reformat",
            .DisplayName = "Вставить разделитель",
            .Description = "Вставить разрыв страницы, разрыв раздела или разрыв строки",
            .RequiredParams = New List(Of String) From {"params.type"},
            .OptionalParams = New List(Of String) From {"target", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"position", "after"})},
                {"target.position", ParamType.EnumType(New String() {"cursor", "end"})},
                {"params.type", ParamType.EnumType(New String() {"page", "section", "line"})}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"position", "paragraph"}
        })

        Register("applyListFormat", New InstructionDefinition With {
            .Operation = "applyListFormat",
            .Category = "reformat",
            .DisplayName = "Применить формат списка",
            .Description = "Преобразовать абзацы в список (маркированный или нумерованный)",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.listType", "params.numberFormat", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "index", "range", "selection"})},
                {"params.listType", ParamType.EnumType(New String() {"bullet", "number", "outline"})},
                {"params.numberFormat", ParamType.StringType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"paragraph", "selection"}
        })

        Register("setColumnFormat", New InstructionDefinition With {
            .Operation = "setColumnFormat",
            .Category = "reformat",
            .DisplayName = "Настройка колонок",
            .Description = "Настроить колонки документа",
            .RequiredParams = New List(Of String) From {"params.columnCount"},
            .OptionalParams = New List(Of String) From {"params.columnWidth", "params.separator", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.columnCount", ParamType.NumberType()},
                {"params.columnWidth", ParamType.NumberType()},
                {"params.separator", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"section"}
        })

        Register("insertHeaderFooter", New InstructionDefinition With {
            .Operation = "insertHeaderFooter",
            .Category = "reformat",
            .DisplayName = "Вставить колонтитулы",
            .Description = "Вставить или изменить содержимое верхнего либо нижнего колонтитула",
            .RequiredParams = New List(Of String) From {"params.type", "params.content"},
            .OptionalParams = New List(Of String) From {"params.alignment", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.type", ParamType.EnumType(New String() {"header", "footer"})},
                {"params.content", ParamType.StringType()},
                {"params.alignment", ParamType.EnumType(New String() {"left", "center", "right"})}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"section"}
        })

        Register("generateToc", New InstructionDefinition With {
            .Operation = "generateToc",
            .Category = "reformat",
            .DisplayName = "Создать оглавление",
            .Description = "Автоматически создать оглавление документа",
            .RequiredParams = New List(Of String) From {},
            .OptionalParams = New List(Of String) From {"target.position", "params.levels", "params.includePageNumbers", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.position", ParamType.EnumType(New String() {"start", "cursor"})},
                {"params.levels", ParamType.NumberType()},
                {"params.includePageNumbers", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"document"}
        })

        ' ========== 校对指令 ==========

        Register("suggestCorrection", New InstructionDefinition With {
            .Operation = "suggestCorrection",
            .Category = "proofread",
            .DisplayName = "Предложить исправление",
            .Description = "Предложить заменить исходный текст на новый",
            .RequiredParams = New List(Of String) From {"target", "params.original", "params.suggestion"},
            .OptionalParams = New List(Of String) From {"params.issueType", "params.severity", "params.explanation", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"textMatch", "semantic", "range"})},
                {"target.match", ParamType.StringType()},
                {"params.original", ParamType.StringType()},
                {"params.suggestion", ParamType.StringType()},
                {"params.issueType", ParamType.EnumType(New String() {"spellingError", "wordUsageError", "punctuationError", "grammaticalError", "expressionError", "formatError"})},
                {"params.severity", ParamType.EnumType(New String() {"high", "medium", "low"})},
                {"params.explanation", ParamType.StringType()}
            },
            .IsDestructive = False,
            .RequiresConfirmation = True,
            .SupportedTargetTypes = New List(Of String) From {"text", "range"}
        })

        Register("suggestFormatFix", New InstructionDefinition With {
            .Operation = "suggestFormatFix",
            .Category = "proofread",
            .DisplayName = "Предложить исправление формата",
            .Description = "Предложить исправление проблем форматирования",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.currentFormat", "params.expectedFormat", "params.explanation", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "range", "paragraph"})},
                {"params.currentFormat", ParamType.StringType()},
                {"params.expectedFormat", ParamType.StringType()},
                {"params.explanation", ParamType.StringType()}
            },
            .IsDestructive = False,
            .RequiresConfirmation = True,
            .SupportedTargetTypes = New List(Of String) From {"range", "paragraph"}
        })

        Register("suggestStyleUnify", New InstructionDefinition With {
            .Operation = "suggestStyleUnify",
            .Category = "proofread",
            .DisplayName = "Предложить унификацию стилей",
            .Description = "Предложить унифицировать несогласованные стили",
            .RequiredParams = New List(Of String) From {"params.targetStyle", "params.inconsistentRanges"},
            .OptionalParams = New List(Of String) From {"params.expectedStyle", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.targetStyle", ParamType.StringType()},
                {"params.expectedStyle", ParamType.StringType()},
                {"params.inconsistentRanges", ParamType.ArrayType()}
            },
            .IsDestructive = False,
            .RequiresConfirmation = True,
            .SupportedTargetTypes = New List(Of String) From {"document"}
        })

        Register("markForReview", New InstructionDefinition With {
            .Operation = "markForReview",
            .Category = "proofread",
            .DisplayName = "Пометить для проверки",
            .Description = "Пометить фрагмент содержимого для проверки пользователем",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.note", "params.category", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "range", "paragraph"})},
                {"params.note", ParamType.StringType()},
                {"params.category", ParamType.StringType()}
            },
            .IsDestructive = False,
            .SupportedTargetTypes = New List(Of String) From {"range", "paragraph"}
        })

        _isInitialized = True
    End Sub

    ''' <summary>
    ''' 注册指令定义
    ''' </summary>
    Public Shared Sub Register(operation As String, definition As InstructionDefinition)
        _registry(operation.ToLower()) = definition
    End Sub

    ''' <summary>
    ''' 检查操作是否有效
    ''' </summary>
    Public Shared Function IsValidOperation(operation As String) As Boolean
        EnsureInitialized()
        Return _registry.ContainsKey(operation.ToLower())
    End Function

    ''' <summary>
    ''' 获取指令定义
    ''' </summary>
    Public Shared Function GetDefinition(operation As String) As InstructionDefinition
        EnsureInitialized()
        Dim key = operation.ToLower()
        If _registry.ContainsKey(key) Then
            Return _registry(key)
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 校验指令参数
    ''' </summary>
    Public Shared Function ValidateParameters(operation As String, params As JToken) As ParamValidationResult
        EnsureInitialized()

        Dim key = operation.ToLower()
        If Not _registry.ContainsKey(key) Then
            Return ParamValidationResult.Failure($"Неизвестный тип операции: {operation}")
        End If

        Dim def = _registry(key)

        If params Is Nothing OrElse params.Type <> JTokenType.Object Then
            Return ParamValidationResult.Failure("params должен быть объектом")
        End If

        Dim paramsObj = CType(params, JObject)

        ' 检查必需参数
        For Each required In def.RequiredParams
            Dim token = paramsObj.SelectToken(required)
            If token Is Nothing OrElse token.Type = JTokenType.Null Then
                Return ParamValidationResult.Failure($"Отсутствует обязательный параметр: {required}")
            End If
        Next

        ' 校验参数类型
        For Each kvp In def.ParamSchema
            Dim token = paramsObj.SelectToken(kvp.Key)
            If token IsNot Nothing AndAlso token.Type <> JTokenType.Null Then
                If Not IsTokenTypeMatch(token, kvp.Value) Then
                    Return ParamValidationResult.Failure($"Тип параметра {kvp.Key} не совпадает, ожидается {kvp.Value.BaseType}")
                End If
            End If
        Next

        Return ParamValidationResult.Success()
    End Function

    Private Shared Sub EnsureInitialized()
        If Not _isInitialized Then
            Initialize()
        End If
    End Sub

    Private Shared Function IsTokenTypeMatch(token As JToken, paramType As ParamType) As Boolean
        Select Case paramType.BaseType.ToLower()
            Case "string"
                Return token.Type = JTokenType.String OrElse token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float
            Case "number"
                Return token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float
            Case "boolean"
                Return token.Type = JTokenType.Boolean
            Case "array"
                Return token.Type = JTokenType.Array
            Case "object"
                Return token.Type = JTokenType.Object
            Case "enum"
                If token.Type <> JTokenType.String Then Return False
                If paramType.EnumValues IsNot Nothing AndAlso paramType.EnumValues.Count > 0 Then
                    Return paramType.EnumValues.Contains(token.ToString())
                End If
                Return True
            Case Else
                Return True
        End Select
    End Function

End Class
