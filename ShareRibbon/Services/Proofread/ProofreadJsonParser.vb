' ShareRibbon/Services/Proofread/ProofreadJsonParser.vb
' 统一的校对JSON解析器 - 支持多种返回格式

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 统一的校对JSON解析器
''' 支持多种格式：
''' 1. 纯数组格式：[{"paragraphIndex":0,...}]
''' 2. 对象包装格式：{"issues": [...], "summary": {...}}
''' 3. 包装在code block：```json [...] ```
''' </summary>
Public Class ProofreadJsonParser

    ''' <summary>
    ''' 解析AI返回的校对结果（支持多种格式）
    ''' </summary>
    Public Shared Function Parse(aiResponse As String) As ProofreadParseResult
        Dim result As New ProofreadParseResult()

        Try
            ' 1. 先清理响应，去除markdown代码块包装
            Dim cleanContent = CleanResponse(aiResponse)

            If String.IsNullOrWhiteSpace(cleanContent) Then
                Return New ProofreadParseResult With {
                    .Success = False,
                    .ErrorMessage = "Ответ пуст"
                }
            End If

            ' 2. 尝试解析
            Dim json As JToken = JToken.Parse(cleanContent)

            Dim issuesArray As JArray = Nothing

            If json.Type = JTokenType.Array Then
                ' 格式A：纯数组
                issuesArray = CType(json, JArray)
                result.FormatDetected = "array"
            ElseIf json.Type = JTokenType.Object Then
                ' 格式B：对象包装
                Dim obj = CType(json, JObject)

                ' 尝试找到问题数组
                issuesArray = FindIssuesArrayFromObject(obj)

                ' 提取摘要（如果有）
                If obj("summary") IsNot Nothing Then
                    result.Summary = obj("summary").ToString()
                End If
            End If

            ' 3. 如果没找到issues数组但确实是对象，把对象当单个issue
            If issuesArray Is Nothing AndAlso json.Type = JTokenType.Object Then
                issuesArray = New JArray()
                issuesArray.Add(json)
                result.FormatDetected = "single-object-as-array"
            End If

            If issuesArray Is Nothing Then
                Return New ProofreadParseResult With {
                    .Success = False,
                    .ErrorMessage = "Не удалось найти массив issues"
                }
            End If

            ' 4. 解析issue数组
            Dim issues As New List(Of ProofreadIssue)()
            For Each item In issuesArray
                Try
                    Dim issue = ParseSingleIssue(item)
                    If issue IsNot Nothing Then
                        issues.Add(issue)
                    End If
                Catch ex As Exception
                    ' 跳过单个解析失败的issue
                    Debug.WriteLine($"解析单个issue失败: {ex.Message}")
                End Try
            Next

            result.Success = True
            result.Issues = issues

        Catch ex As JsonException
            Return New ProofreadParseResult With {
                .Success = False,
                .ErrorMessage = $"Сбой разбора JSON: {ex.Message}"
            }
        Catch ex As Exception
            Return New ProofreadParseResult With {
                .Success = False,
                .ErrorMessage = $"Сбой разбора: {ex.Message}"
            }
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 从对象中寻找issues数组
    ''' </summary>
    Private Shared Function FindIssuesArrayFromObject(obj As JObject) As JArray
        ' 尝试各种可能的字段名
        Dim possibleFieldNames As String() = {
            "issues", "corrections", "results", "data",
            "proofread", "problems", "errors", "suggestions"
        }

        For Each fieldName In possibleFieldNames
            If obj(fieldName) IsNot Nothing AndAlso obj(fieldName).Type = JTokenType.Array Then
                Return CType(obj(fieldName), JArray)
            End If
        Next

        ' 尝试找第一个数组属性
        For Each prop In obj.Properties()
            If prop.Value.Type = JTokenType.Array Then
                Return CType(prop.Value, JArray)
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' 清理响应内容
    ''' </summary>
    Private Shared Function CleanResponse(response As String) As String
        If String.IsNullOrEmpty(response) Then Return response

        Dim content = response.Trim()

        ' 去除markdown代码块
        ' 格式1: ```json ... ```
        Dim codeBlockMatch = Regex.Match(content, "```(?:json)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase)
        If codeBlockMatch.Success Then
            content = codeBlockMatch.Groups(1).Value.Trim()
        End If

        ' 格式2: ` ... `
        Dim inlineMatch = Regex.Match(content, "`\s*([\s\S]*?)\s*`")
        If inlineMatch.Success Then
            content = inlineMatch.Groups(1).Value.Trim()
        End If

        ' 尝试找到第一个 [ 或 {
        Dim firstBracket = Math.Max(content.IndexOf("["c), content.IndexOf("{"c))
        Dim lastBracket = Math.Max(content.LastIndexOf("]"c), content.LastIndexOf("}"c))

        If firstBracket >= 0 AndAlso lastBracket > firstBracket Then
            content = content.Substring(firstBracket, lastBracket - firstBracket + 1)
        End If

        ' 移除末尾多余的逗号（常见问题）
        ' 查找数组或对象末尾可能存在的多余逗号
        content = Regex.Replace(content, ",\s*([\}\]])", "$1")

        Return content
    End Function

    ''' <summary>
    ''' 解析单个issue
    ''' </summary>
    Private Shared Function ParseSingleIssue(jsonItem As JToken) As ProofreadIssue
        Dim issue As New ProofreadIssue()

        ' 兼容多种字段命名
        issue.Id = Guid.NewGuid().ToString()

        ' paragraphIndex / paraIndex / index
        issue.ParagraphIndex = GetIntValue(jsonItem, "paragraphIndex", "paraIndex", "index")

        ' original / originalText / text
        issue.Original = GetStringValue(jsonItem, "original", "originalText", "text")

        ' suggestion / correction / corrected / replacement
        issue.Suggestion = GetStringValue(jsonItem, "suggestion", "correction", "corrected", "replacement")

        ' 如果suggestion为空，尝试用original+correction的组合
        If String.IsNullOrEmpty(issue.Suggestion) Then
            issue.Suggestion = GetStringValue(jsonItem, "correctedText", "newText")
        End If

        ' issueType / type / category
        issue.IssueType = ParseIssueType(GetStringValue(jsonItem, "issueType", "type", "category"))

        ' severity / priority / level
        issue.Severity = ParseSeverity(GetStringValue(jsonItem, "severity", "priority", "level"))

        ' explanation / reason / description / note
        issue.Explanation = GetStringValue(jsonItem, "explanation", "reason", "description", "note")

        ' 验证必需字段 - 至少需要original
        If String.IsNullOrEmpty(issue.Original) Then
            ' 尝试从其他字段推断
            Dim anyText = GetStringValue(jsonItem, "content", "value")
            If Not String.IsNullOrEmpty(anyText) Then
                issue.Original = anyText
            Else
                Return Nothing
            End If
        End If

        ' 如果没有suggestion但有original，留空（标记为仅提示）
        If String.IsNullOrEmpty(issue.Suggestion) Then
            issue.Suggestion = issue.Original
        End If

        Return issue
    End Function

    Private Shared Function GetStringValue(obj As JToken, ParamArray fieldNames As String()) As String
        For Each name In fieldNames
            If obj(name) IsNot Nothing Then
                Return obj(name).ToString()
            End If
        Next
        Return ""
    End Function

    Private Shared Function GetIntValue(obj As JToken, ParamArray fieldNames As String()) As Integer
        For Each name In fieldNames
            If obj(name) IsNot Nothing Then
                Dim val = obj(name)
                If val.Type = JTokenType.Integer Then
                    Return val.Value(Of Integer)()
                Else
                    Dim intVal As Integer
                    If Integer.TryParse(val.ToString(), intVal) Then
                        Return intVal
                    End If
                End If
            End If
        Next
        Return 0
    End Function

    Private Shared Function ParseIssueType(typeStr As String) As IssueType
        If String.IsNullOrEmpty(typeStr) Then Return IssueType.ExpressionError

        Select Case typeStr.ToLower()
            Case "spelling", "spellingerror", "spell", "拼写", "错别字"
                Return IssueType.SpellingError
            Case "wordusage", "wordusageerror", "word", "usage", "用词", "的地得"
                Return IssueType.WordUsageError
            Case "punctuation", "punctuationerror", "punct", "标点", "标点符号"
                Return IssueType.PunctuationError
            Case "grammar", "grammaticalerror", "grammatical", "语法"
                Return IssueType.GrammaticalError
            Case "expression", "expressionerror", "style", "表达", "表述"
                Return IssueType.ExpressionError
            Case "format", "formaterror", "格式"
                Return IssueType.FormatError
            Case Else
                Return IssueType.ExpressionError
        End Select
    End Function

    Private Shared Function ParseSeverity(severityStr As String) As IssueSeverity
        If String.IsNullOrEmpty(severityStr) Then Return IssueSeverity.Medium

        Select Case severityStr.ToLower()
            Case "high", "critical", "error", "must", "必须", "严重"
                Return IssueSeverity.High
            Case "medium", "warning", "should", "suggest", "建议", "中等"
                Return IssueSeverity.Medium
            Case "low", "info", "optional", "可选", "轻微"
                Return IssueSeverity.Low
            Case Else
                Return IssueSeverity.Medium
        End Select
    End Function
End Class

''' <summary>
''' 校对解析结果
''' </summary>
Public Class ProofreadParseResult
    Public Property Success As Boolean
    Public Property ErrorMessage As String
    Public Property Issues As List(Of ProofreadIssue)
    Public Property Summary As String
    Public Property FormatDetected As String  ' 检测到的格式
End Class

''' <summary>
''' 校对分析状态，用于区分真正无问题和解析失败。
''' </summary>
Public Enum ProofreadAnalysisStatus
    HasIssues
    NoIssues
    ParseFailed
    ModelFailed
End Enum

''' <summary>
''' 校对分析结果。不要再用空列表同时表达“无问题”和“解析失败”。
''' </summary>
Public Class ProofreadAnalysisResult
    Public Property Status As ProofreadAnalysisStatus = ProofreadAnalysisStatus.ModelFailed
    Public Property Issues As New List(Of ProofreadIssue)()
    Public Property ErrorMessage As String
    Public Property RawResponsePreview As String
    Public Property Summary As String
    Public Property FormatDetected As String

    Public ReadOnly Property HasIssues As Boolean
        Get
            Return Status = ProofreadAnalysisStatus.HasIssues AndAlso Issues IsNot Nothing AndAlso Issues.Count > 0
        End Get
    End Property
End Class
