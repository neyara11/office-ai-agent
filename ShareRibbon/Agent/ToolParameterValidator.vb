' ShareRibbon\Agent\ToolParameterValidator.vb
' Structural validation of native tool calls before they reach the Office host.

Imports System.Linq
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' One schema violation. Path uses JSON-ish notation, for example slides[2].slideType.
    ''' </summary>
    Public Class ToolSchemaError
        Public Property Path As String = ""
        Public Property Message As String = ""

        Public Overrides Function ToString() As String
            If String.IsNullOrWhiteSpace(Path) Then Return Message
            Return Path & ": " & Message
        End Function
    End Class

    ''' <summary>
    ''' Validates tool parameters against ToolDescriptor metadata and an optional nested
    ''' JSON Schema subset stored in ToolParam.Schema. Unknown parameters are tolerated;
    ''' only explicitly declared constraints are enforced so legacy tool calls keep working.
    ''' </summary>
    Public NotInheritable Class ToolParameterValidator
        Private Const MaxArrayItemsToCheck As Integer = 200

        Private Sub New()
        End Sub

        Public Shared Function Validate(tool As ToolDescriptor, params As JObject) As List(Of ToolSchemaError)
            Dim errors As New List(Of ToolSchemaError)()
            If tool Is Nothing OrElse tool.Parameters Is Nothing OrElse params Is Nothing Then Return errors

            Try
                For Each p In tool.Parameters
                    If p Is Nothing OrElse String.IsNullOrWhiteSpace(p.Name) Then Continue For

                    Dim token = params(p.Name)
                    If token Is Nothing OrElse token.Type = JTokenType.Null OrElse token.Type = JTokenType.Undefined Then
                        If p.Required Then AddError(errors, p.Name, "обязательный параметр не передан")
                        Continue For
                    End If

                    ValidateType(token, p.Type, p.Name, errors)
                    ValidateEnum(token, p.EnumValues, p.Name, errors)
                    If p.Schema IsNot Nothing Then ValidateSchema(token, p.Schema, p.Name, errors)
                Next
            Catch ex As Exception
                AddError(errors, "", "внутренняя ошибка проверки схемы: " & ex.Message)
            End Try

            Return errors
        End Function

        Private Shared Sub ValidateType(token As JToken, expectedType As String, path As String, errors As List(Of ToolSchemaError))
            If token Is Nothing OrElse String.IsNullOrWhiteSpace(expectedType) Then Return
            If TypeMatches(token, expectedType) Then Return
            AddError(errors, path, $"ожидается тип {expectedType}, получено {DescribeToken(token)}")
        End Sub

        Private Shared Function TypeMatches(token As JToken, expectedType As String) As Boolean
            Select Case expectedType.Trim().ToLowerInvariant()
                Case "string"
                    Return token.Type = JTokenType.String OrElse token.Type = JTokenType.Date OrElse token.Type = JTokenType.Guid
                Case "integer"
                    If token.Type = JTokenType.Integer Then Return True
                    If token.Type = JTokenType.Float Then
                        Dim value = token.Value(Of Double)()
                        Return Math.Abs(value - Math.Truncate(value)) < 0.0000001
                    End If
                    Return False
                Case "number"
                    Return token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float
                Case "boolean"
                    Return token.Type = JTokenType.Boolean
                Case "array"
                    Return token.Type = JTokenType.Array
                Case "object"
                    Return token.Type = JTokenType.Object
                Case Else
                    ' Unknown/loose type (for example "any"): do not restrict.
                    Return True
            End Select
        End Function

        Private Shared Function DescribeToken(token As JToken) As String
            If token Is Nothing Then Return "null"
            Select Case token.Type
                Case JTokenType.Integer, JTokenType.Float
                    Return "число"
                Case JTokenType.String, JTokenType.Date, JTokenType.Guid
                    Return "строка"
                Case JTokenType.Boolean
                    Return "boolean"
                Case JTokenType.Array
                    Return "массив"
                Case JTokenType.Object
                    Return "объект"
                Case Else
                    Return token.Type.ToString()
            End Select
        End Function

        Private Shared Sub ValidateEnum(token As JToken, allowed As List(Of String), path As String, errors As List(Of ToolSchemaError))
            If token Is Nothing OrElse allowed Is Nothing OrElse allowed.Count = 0 Then Return
            If token.Type = JTokenType.Array OrElse token.Type = JTokenType.Object Then Return

            Dim value = token.ToString()
            If allowed.Any(Function(item) String.Equals(item, value, StringComparison.OrdinalIgnoreCase)) Then Return
            AddError(errors, path, $"недопустимое значение '{Truncate(value)}'; допустимые: {String.Join(", ", allowed)}")
        End Sub

        Private Shared Sub ValidateSchema(token As JToken, schema As JObject, path As String, errors As List(Of ToolSchemaError))
            If token Is Nothing OrElse schema Is Nothing Then Return

            Dim declaredType = schema("type")
            If declaredType IsNot Nothing Then
                If declaredType.Type = JTokenType.Array Then
                    Dim anyMatch = declaredType.OfType(Of JValue)().
                        Any(Function(item) TypeMatches(token, If(item.Value?.ToString(), "")))
                    If Not anyMatch Then
                        AddError(errors, path, $"ожидается тип {declaredType.ToString(Newtonsoft.Json.Formatting.None)}, получено {DescribeToken(token)}")
                    End If
                Else
                    ValidateType(token, If(declaredType.ToString(), ""), path, errors)
                End If
            End If

            Dim enumNode = schema("enum")
            If enumNode IsNot Nothing AndAlso enumNode.Type = JTokenType.Array Then
                Dim allowed = enumNode.OfType(Of JValue)().
                    Select(Function(item) If(item.Value?.ToString(), "")).
                    Where(Function(item) item.Length > 0).
                    ToList()
                ValidateEnum(token, allowed, path, errors)
            End If

            If token.Type = JTokenType.String Then
                Dim minLength = ReadInteger(schema("minLength"), -1)
                Dim maxLength = ReadInteger(schema("maxLength"), -1)
                Dim text = token.ToString()
                If minLength >= 0 AndAlso text.Length < minLength Then
                    AddError(errors, path, $"минимальная длина строки {minLength}")
                End If
                If maxLength >= 0 AndAlso text.Length > maxLength Then
                    AddError(errors, path, $"максимальная длина строки {maxLength}")
                End If
            End If

            If token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float Then
                Dim value = token.Value(Of Double)()
                Dim minimumNode = schema("minimum")
                Dim maximumNode = schema("maximum")
                If minimumNode IsNot Nothing AndAlso value < minimumNode.Value(Of Double)() Then
                    AddError(errors, path, $"значение меньше минимума {minimumNode.ToString()}")
                End If
                If maximumNode IsNot Nothing AndAlso value > maximumNode.Value(Of Double)() Then
                    AddError(errors, path, $"значение больше максимума {maximumNode.ToString()}")
                End If
            End If

            If token.Type = JTokenType.Object Then
                Dim properties = TryCast(schema("properties"), JObject)
                Dim required = TryCast(schema("required"), JArray)

                If required IsNot Nothing Then
                    For Each requiredName In required.OfType(Of JValue)()
                        Dim name = If(requiredName.Value?.ToString(), "")
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim child = token(name)
                        If child Is Nothing OrElse child.Type = JTokenType.Null OrElse child.Type = JTokenType.Undefined Then
                            AddError(errors, JoinPath(path, name), "обязательное поле не передано")
                        End If
                    Next
                End If

                If properties IsNot Nothing Then
                    For Each propertyDef In properties.Properties()
                        Dim child = token(propertyDef.Name)
                        If child Is Nothing OrElse child.Type = JTokenType.Null Then Continue For
                        Dim childSchema = TryCast(propertyDef.Value, JObject)
                        If childSchema IsNot Nothing Then
                            ValidateSchema(child, childSchema, JoinPath(path, propertyDef.Name), errors)
                        End If
                    Next
                End If

                Dim additional = schema("additionalProperties")
                If additional IsNot Nothing AndAlso additional.Type = JTokenType.Boolean AndAlso
                   Not additional.Value(Of Boolean)() AndAlso properties IsNot Nothing Then
                    For Each propertyItem In token.Children(Of JProperty)()
                        If properties(propertyItem.Name) Is Nothing Then
                            AddError(errors, JoinPath(path, propertyItem.Name), "поле не описано в схеме инструмента")
                        End If
                    Next
                End If
            End If

            If token.Type = JTokenType.Array Then
                Dim minItems = ReadInteger(schema("minItems"), -1)
                Dim maxItems = ReadInteger(schema("maxItems"), -1)
                Dim array = CType(token, JArray)
                If minItems >= 0 AndAlso array.Count < minItems Then
                    AddError(errors, path, $"минимум элементов: {minItems}")
                End If
                If maxItems >= 0 AndAlso array.Count > maxItems Then
                    AddError(errors, path, $"максимум элементов: {maxItems}")
                End If

                Dim itemsSchema = TryCast(schema("items"), JObject)
                If itemsSchema IsNot Nothing Then
                    Dim limit = Math.Min(array.Count, MaxArrayItemsToCheck)
                    For index = 0 To limit - 1
                        ValidateSchema(array(index), itemsSchema, path & "[" & index.ToString() & "]", errors)
                    Next
                End If
            End If
        End Sub

        Private Shared Function ReadInteger(node As JToken, fallback As Integer) As Integer
            If node Is Nothing Then Return fallback
            Try
                Return node.Value(Of Integer)()
            Catch
                Return fallback
            End Try
        End Function

        Private Shared Function JoinPath(parent As String, child As String) As String
            If String.IsNullOrWhiteSpace(parent) Then Return child
            Return parent & "." & child
        End Function

        Private Shared Function Truncate(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If value.Length <= 60 Then Return value
            Return value.Substring(0, 60) & "..."
        End Function

        Private Shared Sub AddError(errors As List(Of ToolSchemaError), path As String, message As String)
            If errors Is Nothing Then Return
            If errors.Count >= 40 Then Return
            Dim safePath = If(path, "")
            Dim safeMessage = If(message, "")
            If errors.Any(Function(item) item.Path = safePath AndAlso item.Message = safeMessage) Then Return
            errors.Add(New ToolSchemaError With {.Path = safePath, .Message = safeMessage})
        End Sub
    End Class

End Namespace
