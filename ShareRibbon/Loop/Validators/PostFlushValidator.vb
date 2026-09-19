' ShareRibbon\Loop\Validators\PostFlushValidator.vb
' Flush后校验器 - 校验AI响应内容的结构正确性、指令可解析性、操作安全性

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Flush后校验器 - 校验AI响应内容
''' </summary>
Public Class PostFlushValidator
    Implements IInstructionValidator

    Public Async Function ValidateAsync(
        aiResponse As String,
        expectedFormat As InstructionFormat,
        context As ExecutionContext) As Task(Of ValidationResult) Implements IInstructionValidator.ValidateAsync

        Dim errors As New List(Of InstructionError)()

        ' 检查1: 响应是否为空或过长
        If String.IsNullOrWhiteSpace(aiResponse) Then
            errors.Add(New InstructionError(ErrorLevel.Critical, "Ответ AI пуст"))
            Return ValidationResult.Failure(errors)
        End If

        ' 检查2: 提取指令内容（支持多种包装格式）
        Dim extractedContent = ExtractInstructionContent(aiResponse, expectedFormat)
        If String.IsNullOrEmpty(extractedContent) Then
            errors.Add(New InstructionError(ErrorLevel.Critical,
                $"Не удалось извлечь содержимое формата {expectedFormat} из ответа AI"))
            Return ValidationResult.Failure(errors, aiResponse)
        End If

        ' 检查3: 根据指令格式进行结构化校验
        Dim parseResult As ParseResult = Nothing
        Select Case expectedFormat
            Case InstructionFormat.DslJson
                parseResult = ValidateDslJson(extractedContent)
            Case InstructionFormat.ProofreadJson
                parseResult = ValidateProofreadJson(extractedContent)
            Case InstructionFormat.LegacyJsonCommand
                parseResult = ValidateLegacyJsonCommand(extractedContent)
            Case Else
                parseResult = New ParseResult With {.IsValid = True}
        End Select

        If Not parseResult.IsValid Then
            errors.AddRange(parseResult.Errors)
            Return ValidationResult.Failure(errors, extractedContent)
        End If

        ' 检查4: 指令语义校验（操作是否安全、是否可执行）
        If parseResult.Instructions IsNot Nothing AndAlso parseResult.Instructions.Count > 0 Then
            Dim semanticCheck = ValidateSemanticSafety(parseResult.Instructions, context)
            If Not semanticCheck.IsSafe Then
                errors.AddRange(semanticCheck.Errors)
            End If

            ' 检查5: 指令一致性校验（排版指令是否冲突）
            Dim consistencyCheck = ValidateConsistency(parseResult.Instructions)
            If Not consistencyCheck.IsConsistent Then
                errors.AddRange(consistencyCheck.Errors)
            End If
        End If

        Dim result As New ValidationResult With {
            .IsValid = errors.Count = 0 OrElse errors.All(Function(e) e.Level <> ErrorLevel.Critical),
            .Errors = errors,
            .ExtractedContent = extractedContent,
            .ParsedInstructions = If(parseResult?.Instructions, New List(Of Instruction)()),
            .CanAutoCorrect = errors.Count > 0 AndAlso errors.All(Function(e) e.IsAutoCorrectable),
            .OriginalResponse = aiResponse
        }

        Return result
    End Function

    ''' <summary>
    ''' 从AI响应中提取指令内容
    ''' </summary>
    Private Function ExtractInstructionContent(aiResponse As String, expectedFormat As InstructionFormat) As String
        Dim content = aiResponse.Trim()

        ' 尝试提取代码块中的JSON
        Dim codeBlockMatch = Regex.Match(content, "```(?:json)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase)
        If codeBlockMatch.Success Then
            Return codeBlockMatch.Groups(1).Value.Trim()
        End If

        ' 尝试直接找到JSON对象
        If content.StartsWith("{") AndAlso content.EndsWith("}") Then
            Return content
        End If

        ' 尝试在文本中找到JSON
        Dim startIdx = content.IndexOf("{")
        Dim endIdx = content.LastIndexOf("}")
        If startIdx >= 0 AndAlso endIdx > startIdx Then
            Return content.Substring(startIdx, endIdx - startIdx + 1)
        End If

        ' 尝试找到JSON数组
        If content.StartsWith("[") AndAlso content.EndsWith("]") Then
            Return content
        End If

        Return String.Empty
    End Function

    ''' <summary>
    ''' 校验DSL JSON格式
    ''' </summary>
    Private Function ValidateDslJson(content As String) As ParseResult
        Try
            ' 清理JSON内容，移除末尾多余的逗号等不规范字符
            Dim cleanedContent = CleanJsonContent(content)
            Dim json = JObject.Parse(cleanedContent)

            ' 必须有version和instructions
            If json("version") Is Nothing Then
                Return ParseResult.Failure(New InstructionError(ErrorLevel.Error, "Отсутствует поле version"))
            End If

            If json("instructions") Is Nothing OrElse json("instructions").Type <> JTokenType.Array Then
                Return ParseResult.Failure(New InstructionError(ErrorLevel.Error, "Отсутствует массив instructions"))
            End If

            Dim instructions As New List(Of Instruction)()
            Dim instructionArray = CType(json("instructions"), JArray)

            For i = 0 To instructionArray.Count - 1
                Dim item = instructionArray(i)
                If item.Type <> JTokenType.Object Then
                    Return ParseResult.Failure(New InstructionError(ErrorLevel.Error, $"instructions[{i}] должен быть объектом"))
                End If

                Dim itemObj = CType(item, JObject)

                If itemObj("op") Is Nothing Then
                    Return ParseResult.Failure(New InstructionError(ErrorLevel.Error, $"instructions[{i}] отсутствует поле op"))
                End If

                Dim op = itemObj("op").ToString()

                ' 校验操作类型是否注册
                If Not InstructionRegistry.IsValidOperation(op) Then
                    Return ParseResult.Failure(New InstructionError(ErrorLevel.Error, $"Неизвестный тип операции: {op}"))
                End If

                ' 校验参数Schema
                Dim params = itemObj("params")
                If params IsNot Nothing Then
                    Dim paramCheck = InstructionRegistry.ValidateParameters(op, params)
                    If Not paramCheck.IsValid Then
                        Return ParseResult.Failure(New InstructionError(ErrorLevel.Error,
                            $"instructions[{i}] ошибка проверки параметров: {paramCheck.ErrorMessage}"))
                    End If
                End If

                Dim instruction = New Instruction(op, If(params, New JObject()), itemObj("id")?.ToString())
                If itemObj("target") IsNot Nothing Then
                    instruction.Target = CType(itemObj("target"), JObject)
                End If
                If itemObj("expected") IsNot Nothing Then
                    instruction.Expected = CType(itemObj("expected"), JObject)
                End If
                If itemObj("rollback") IsNot Nothing Then
                    instruction.Rollback = CType(itemObj("rollback"), JObject)
                End If
                If itemObj("metadata") IsNot Nothing Then
                    instruction.Metadata = CType(itemObj("metadata"), JObject)
                End If

                instructions.Add(instruction)
            Next

            Return ParseResult.Success(instructions)

        Catch ex As JsonException
            Return ParseResult.Failure(New InstructionError(ErrorLevel.Critical, $"Ошибка разбора JSON: {ex.Message}"))
        End Try
    End Function

    ''' <summary>
    ''' 清理JSON内容，移除不规范字符
    ''' </summary>
    Private Function CleanJsonContent(content As String) As String
        If String.IsNullOrWhiteSpace(content) Then
            Return content
        End If

        Dim result = content.Trim()

        ' 找到最外层的{}或[]
        Dim startBrace = result.IndexOf("{"c)
        Dim startBracket = result.IndexOf("["c)
        Dim realStart = -1

        If startBrace >= 0 AndAlso startBracket >= 0 Then
            realStart = Math.Min(startBrace, startBracket)
        ElseIf startBrace >= 0 Then
            realStart = startBrace
        ElseIf startBracket >= 0 Then
            realStart = startBracket
        End If

        If realStart >= 0 Then
            ' 从真正的JSON开始位置截取
            result = result.Substring(realStart)
        End If

        ' 处理结束位置：提取第一个完整JSON根对象/数组，忽略后续逗号或说明文字
        Dim jsonEnd = FindFirstCompleteJsonEnd(result)
        If jsonEnd >= 0 Then
            result = result.Substring(0, jsonEnd + 1)
        End If

        ' 移除末尾多余的逗号（常见问题）
        ' 查找数组或对象末尾可能存在的多余逗号
        result = System.Text.RegularExpressions.Regex.Replace(result, ",\s*([\}\]])", "$1")

        Return result
    End Function

    Private Function FindFirstCompleteJsonEnd(content As String) As Integer
        If String.IsNullOrEmpty(content) Then Return -1
        If Not content.StartsWith("{") AndAlso Not content.StartsWith("[") Then Return -1

        Dim curlyDepth As Integer = 0
        Dim squareDepth As Integer = 0
        Dim inString As Boolean = False
        Dim escaped As Boolean = False

        For i = 0 To content.Length - 1
            Dim ch = content(i)

            If inString Then
                If escaped Then
                    escaped = False
                ElseIf ch = "\"c Then
                    escaped = True
                ElseIf ch = """"c Then
                    inString = False
                End If
                Continue For
            End If

            If ch = """"c Then
                inString = True
            ElseIf ch = "{"c Then
                curlyDepth += 1
            ElseIf ch = "}"c Then
                curlyDepth -= 1
            ElseIf ch = "["c Then
                squareDepth += 1
            ElseIf ch = "]"c Then
                squareDepth -= 1
            End If

            If i > 0 AndAlso curlyDepth = 0 AndAlso squareDepth = 0 Then
                Return i
            End If
        Next

        Return -1
    End Function

    ''' <summary>
    ''' 校验校对JSON格式（支持多种格式）
    ''' </summary>
    Private Function ValidateProofreadJson(content As String) As ParseResult
        Try
            ' 使用新的统一解析器
            Dim proofreadResult = ProofreadJsonParser.Parse(content)

            If Not proofreadResult.Success Then
                Return ParseResult.Failure(New InstructionError(
                    ErrorLevel.Critical, proofreadResult.ErrorMessage))
            End If

            ' 转换成DSL指令
            Dim instructions As New List(Of Instruction)()
            For Each issue In proofreadResult.Issues
                Dim instruction = New Instruction("suggestCorrection", Nothing)

                ' 设置target
                instruction.Target = New JObject From {
                    {"type", "textMatch"},
                    {"match", issue.Original}
                }

                ' 设置params
                instruction.Params = New JObject From {
                    {"original", issue.Original},
                    {"suggestion", issue.Suggestion},
                    {"issueType", issue.IssueType.ToString()},
                    {"severity", issue.Severity.ToString()},
                    {"explanation", issue.Explanation}
                }

                ' 设置expected描述
                Dim originalPreview = If(issue.Original.Length > 20, issue.Original.Substring(0, 20) & "...", issue.Original)
                Dim suggestionPreview = If(issue.Suggestion.Length > 20, issue.Suggestion.Substring(0, 20) & "...", issue.Suggestion)
                instruction.Expected = New JObject From {
                    {"description", $"Рекомендуется исправить '{originalPreview}' на '{suggestionPreview}'"}
                }

                instructions.Add(instruction)
            Next

            Return ParseResult.Success(instructions)

        Catch ex As Exception
            Return ParseResult.Failure(New InstructionError(
                ErrorLevel.Critical, $"Ошибка обработки вычитки: {ex.Message}"))
        End Try
    End Function

    ''' <summary>
    ''' 校验旧版JSON命令格式
    ''' </summary>
    Private Function ValidateLegacyJsonCommand(content As String) As ParseResult
        Try
            Dim json = JObject.Parse(content)
            Dim instructions As New List(Of Instruction)()

            ' 检查commands数组
            If json("commands") IsNot Nothing AndAlso json("commands").Type = JTokenType.Array Then
                Dim commands = CType(json("commands"), JArray)
                For i = 0 To commands.Count - 1
                    Dim cmd = commands(i)
                    If cmd("command") IsNot Nothing Then
                        Dim instruction = ConvertLegacyCommandToDsl(CType(cmd, JObject))
                        If instruction IsNot Nothing Then
                            instructions.Add(instruction)
                        End If
                    End If
                Next
            ElseIf json("command") IsNot Nothing Then
                ' 单命令格式
                Dim instruction = ConvertLegacyCommandToDsl(json)
                If instruction IsNot Nothing Then
                    instructions.Add(instruction)
                End If
            End If

            Return ParseResult.Success(instructions)

        Catch ex As JsonException
            Return ParseResult.Failure(New InstructionError(ErrorLevel.Critical, $"Ошибка разбора устаревшей JSON-команды: {ex.Message}"))
        End Try
    End Function

    ''' <summary>
    ''' 将旧版命令转换为DSL指令
    ''' </summary>
    Private Function ConvertLegacyCommandToDsl(legacyCmd As JObject) As Instruction
        Try
            Dim cmdName = legacyCmd("command")?.ToString()
            Dim params = legacyCmd("params")

            If String.IsNullOrEmpty(cmdName) Then Return Nothing

            ' 简单的命令映射
            Dim opMapping As New Dictionary(Of String, String) From {
                {"FormatText", "setCharacterFormat"},
                {"ApplyStyle", "setParagraphStyle"},
                {"InsertTable", "insertTable"},
                {"FormatTable", "formatTable"},
                {"SetPageMargins", "setPageSetup"},
                {"InsertParagraph", "insertBreak"},
                {"SetLineSpacing", "setParagraphStyle"},
                {"SetIndent", "setParagraphStyle"},
                {"InsertHeader", "insertHeaderFooter"},
                {"InsertFooter", "insertHeaderFooter"},
                {"GenerateTOC", "generateToc"}
            }

            Dim dslOp = If(opMapping.ContainsKey(cmdName), opMapping(cmdName), cmdName)

            Return New Instruction(dslOp, If(params, New JObject()))

        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 语义安全性校验
    ''' </summary>
    Private Function ValidateSemanticSafety(instructions As List(Of Instruction), context As ExecutionContext) As SemanticSafetyResult
        Dim errors As New List(Of InstructionError)()
        Dim isSafe As Boolean = True

        For Each inst In instructions
            Dim def = InstructionRegistry.GetDefinition(inst.Operation)
            If def Is Nothing Then Continue For

            ' 检查破坏性操作
            If def.IsDestructive AndAlso context.OfficeContent IsNot Nothing AndAlso context.OfficeContent.IsReadOnly Then
                errors.Add(New InstructionError(ErrorLevel.Critical,
                    $"Инструкция {inst.Id} ({inst.Operation}) является разрушительной операцией, но документ доступен только для чтения",
                    inst.Id))
                isSafe = False
            End If

            ' 检查需要确认的操作
            If def.RequiresConfirmation Then
                ' 记录但不一定阻止
                errors.Add(New InstructionError(ErrorLevel.Warning,
                    $"Инструкция {inst.Id} ({inst.Operation}) требует подтверждения пользователя",
                    inst.Id))
            End If
        Next

        Return New SemanticSafetyResult With {.IsSafe = isSafe, .Errors = errors}
    End Function

    ''' <summary>
    ''' 指令一致性校验
    ''' </summary>
    Private Function ValidateConsistency(instructions As List(Of Instruction)) As ConsistencyResult
        Dim errors As New List(Of InstructionError)()
        Dim isConsistent As Boolean = True

        ' 检查是否有重复的段落样式设置（冲突）
        Dim styleTargets As New Dictionary(Of String, String)()
        For Each inst In instructions
            If inst.Operation = "setParagraphStyle" OrElse inst.Operation = "setCharacterFormat" Then
                Dim selector = inst.GetTargetSelector()
                If Not String.IsNullOrEmpty(selector) Then
                    If styleTargets.ContainsKey(selector) Then
                        errors.Add(New InstructionError(ErrorLevel.Warning,
                            $"Инструкция {inst.Id} и инструкция {styleTargets(selector)} могут конфликтовать (одинаковая цель)",
                            inst.Id, True, "Объедините в одну инструкцию или скорректируйте диапазон цели"))
                    Else
                        styleTargets(selector) = inst.Id
                    End If
                End If
            End If
        Next

        Return New ConsistencyResult With {.IsConsistent = isConsistent, .Errors = errors}
    End Function

    ''' <summary>
    ''' 语义安全性校验结果
    ''' </summary>
    Private Class SemanticSafetyResult
        Public Property IsSafe As Boolean = True
        Public Property Errors As List(Of InstructionError)

        Public Sub New()
            Errors = New List(Of InstructionError)()
        End Sub
    End Class

    ''' <summary>
    ''' 一致性校验结果
    ''' </summary>
    Private Class ConsistencyResult
        Public Property IsConsistent As Boolean = True
        Public Property Errors As List(Of InstructionError)

        Public Sub New()
            Errors = New List(Of InstructionError)()
        End Sub
    End Class

End Class
