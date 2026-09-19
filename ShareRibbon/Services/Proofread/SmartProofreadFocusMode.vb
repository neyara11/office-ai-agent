' ShareRibbon\Services\Proofread\SmartProofreadFocusMode.vb
' 校对专注模式 - WPS风格体验：波浪线标注 + Hover Tooltip + 问题列表

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 校对专注模式 - 进入后提供类WPS的校对体验
''' 特点：全文校对 + 内联标注 + Hover提示 + 问题列表
''' </summary>
Public Class SmartProofreadFocusMode

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _applyCorrection As Func(Of String, String, Task(Of Boolean))
    Private ReadOnly _state As ProofreadFocusState
    
    Private _annotationMap As New Dictionary(Of String, AnnotationInfo)()

    ''' <summary>
    ''' 标注信息
    ''' </summary>
    Private Class AnnotationInfo
        Public Property Issue As ProofreadIssue
        Public Property RangeStart As Integer
        Public Property RangeEnd As Integer
        Public Property IsIgnored As Boolean = False
    End Class

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(
        executeScript As Func(Of String, Task),
        applyCorrection As Func(Of String, String, Task(Of Boolean)))
        
        _executeScript = executeScript
        _applyCorrection = applyCorrection
        _state = New ProofreadFocusState()
    End Sub

    ''' <summary>
    ''' 是否处于校对专注模式
    ''' </summary>
    Public ReadOnly Property IsActive As Boolean
        Get
            Return _state.IsActive
        End Get
    End Property

    ''' <summary>
    ''' 当前校对问题列表
    ''' </summary>
    Public ReadOnly Property CurrentIssues As List(Of ProofreadIssue)
        Get
            Return _state.CurrentIssues
        End Get
    End Property

    ''' <summary>
    ''' 进入校对专注模式
    ''' </summary>
    Public Async Function EnterAsync() As Task
        If _state.IsActive Then Return
        _state.IsActive = True
        
        Debug.WriteLine("[SmartProofreadFocusMode] 进入校对专注模式")
        
        ' 显示校对模式UI
        Await ShowProofreadModeUIAsync()
    End Function

    ''' <summary>
    ''' 执行校对分析
    ''' </summary>
    ''' <param name="aiResponse">AI返回的校对JSON</param>
    ''' <param name="paragraphs">段落文本列表</param>
    ''' <param name="wordApp">Word应用对象</param>
    ''' <param name="selectionStartOffset">选区在文档中的起始位置偏移（用于精确定位波浪线）</param>
    Public Async Function AnalyzeAsync(
        aiResponse As String,
        paragraphs As List(Of String),
        wordApp As Object,
        Optional selectionStartOffset As Integer = 0) As Task
        
        Try
            ' 1. 解析AI返回的校对问题，区分无问题和解析失败
            Dim analysis = ProofreadPromptBuilder.AnalyzeProofreadResponse(aiResponse, paragraphs)
            _state.ProcessedParagraphs = paragraphs

            If analysis Is Nothing Then
                Await ShowParseErrorMessageAsync("Результат анализа корректуры пуст", "")
                Await _executeScript("window.proofreadIssueCount = 0;")
                Return
            End If

            If analysis.Status = ProofreadAnalysisStatus.ParseFailed OrElse
               analysis.Status = ProofreadAnalysisStatus.ModelFailed Then
                Await ShowParseErrorMessageAsync(analysis.ErrorMessage, analysis.RawResponsePreview)
                Await _executeScript("window.proofreadIssueCount = 0;")
                Return
            End If

            _state.CurrentIssues = If(analysis.Issues, New List(Of ProofreadIssue)())

            If analysis.Status = ProofreadAnalysisStatus.NoIssues OrElse _state.CurrentIssues.Count = 0 Then
                Await ShowNoIssuesMessageAsync()
                Await _executeScript("window.proofreadIssueCount = 0;")
                Return
            End If

            ' 将问题数量传递到JS端，供校对模式下聊天时注入上下文
            Await _executeScript($"window.proofreadIssueCount = {_state.CurrentIssues.Count};")
            
            ' 2. 计算位置信息（传入选区起始偏移）
            CalculatePositions(_state.CurrentIssues, paragraphs, selectionStartOffset)
            
            ' 3. 在Word中创建内联标注（在当前线程执行，避免COM跨线程）
            CreateInlineAnnotations(_state.CurrentIssues, wordApp)
            
            ' 4. 显示问题列表面板
            Await ShowProofreadListPanelAsync(_state.CurrentIssues)
            
            ' 5. 显示校对摘要
            Await ShowProofreadSummaryAsync(_state.CurrentIssues)
            
        Catch ex As Exception
            Debug.WriteLine($"[SmartProofreadFocusMode] 校对分析失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 退出校对专注模式
    ''' </summary>
    Public Async Function ExitAsync() As Task
        If Not _state.IsActive Then Return
        _state.IsActive = False
        
        ' 1. 清除所有内联标注
        ClearAllAnnotations()
        
        ' 2. 隐藏校对UI
        Await HideProofreadModeUIAsync()
        
        ' 3. 清空状态
        _state.Reset()
        
        Debug.WriteLine("[SmartProofreadFocusMode] 退出校对专注模式")
    End Function

    ''' <summary>
    ''' 接受修正建议
    ''' </summary>
    Public Async Function AcceptCorrectionAsync(issueId As String) As Task
        Dim issue = _state.CurrentIssues?.FirstOrDefault(Function(i) i.Id = issueId)
        If issue Is Nothing Then Return
        
        ' 应用修正
        Dim success = Await _applyCorrection(issue.Original, issue.Suggestion)
        If success Then
            issue.IsCorrected = True
            
            ' 移除标注
            RemoveAnnotation(issueId)
            
            ' 更新列表
            Await UpdateProofreadListAsync()
        End If
    End Function

    ''' <summary>
    ''' 忽略问题
    ''' </summary>
    Public Async Function IgnoreIssueAsync(issueId As String) As Task
        Dim issue = _state.CurrentIssues?.FirstOrDefault(Function(i) i.Id = issueId)
        If issue IsNot Nothing Then
            issue.IsIgnored = True
        End If
        
        MarkAsIgnored(issueId)
        Await UpdateProofreadListAsync()
    End Function

    ''' <summary>
    ''' 接受所有修正
    ''' </summary>
    Public Async Function AcceptAllAsync() As Task
        Dim issues = If(_state.CurrentIssues, New List(Of ProofreadIssue)())
        For Each issue In issues.Where(Function(i) Not i.IsIgnored AndAlso Not i.IsCorrected)
            Await _applyCorrection(issue.Original, issue.Suggestion)
            issue.IsCorrected = True
        Next

        ' 清除所有标注
        ClearAllAnnotations()
        _state.CurrentIssues.Clear()

        ' 更新UI
        Await ShowAllCorrectedMessageAsync()
    End Function

    ''' <summary>
    ''' 显示校对模式UI
    ''' </summary>
    Private Async Function ShowProofreadModeUIAsync() As Task
        ' 1. 显示吸顶提示
        Await _executeScript("showProofreadModeIndicator();")
        
        ' 2. 显示侧边校对面板
        Await _executeScript("showProofreadSidePanel();")
    End Function

    ''' <summary>
    ''' 隐藏校对模式UI
    ''' </summary>
    Private Async Function HideProofreadModeUIAsync() As Task
        Await _executeScript("hideProofreadModeIndicator();")
        Await _executeScript("hideProofreadSidePanel();")
    End Function

    ''' <summary>
    ''' 计算问题在文档中的位置
    ''' </summary>
    ''' <param name="issues">问题列表</param>
    ''' <param name="paragraphs">段落文本列表</param>
    ''' <param name="selectionStartOffset">选区在文档中的起始位置（绝对偏移）</param>
    Private Sub CalculatePositions(issues As List(Of ProofreadIssue), paragraphs As List(Of String), selectionStartOffset As Integer)
        Dim offset As Integer = selectionStartOffset
        
        For i = 0 To paragraphs.Count - 1
            Dim para = paragraphs(i)
            Dim paragraphIndex = i
            For Each issue In issues.Where(Function(it) it.ParagraphIndex = paragraphIndex)
                Dim pos = para.IndexOf(issue.Original, StringComparison.Ordinal)
                If pos >= 0 Then
                    issue.StartPosition = offset + pos
                    issue.EndPosition = issue.StartPosition + issue.Original.Length
                End If
            Next
            offset += para.Length + 1 ' +1 for paragraph break
        Next
    End Sub

    ''' <summary>
    ''' 在Word中创建内联标注（必须在UI线程执行，避免COM跨线程异常）
    ''' </summary>
    Private Sub CreateInlineAnnotations(
        issues As List(Of ProofreadIssue),
        wordApp As Object)
        
        For Each issue In issues
            CreateWavyUnderline(issue, wordApp)
        Next
    End Sub

    ''' <summary>
    ''' 创建波浪线标注
    ''' </summary>
    Private Sub CreateWavyUnderline(issue As ProofreadIssue, wordApp As Object)
        If wordApp Is Nothing OrElse issue Is Nothing Then Return
        
        Try
            Dim doc = wordApp.ActiveDocument
            If doc Is Nothing Then Return
            
            ' 创建Range
            Dim range = doc.Range(issue.StartPosition, issue.EndPosition)
            If range Is Nothing Then Return

            ' 设置波浪线格式（wdUnderlineWave = 20）
            With range.Font
                .Underline = 20  ' wdUnderlineWave
                .UnderlineColor = GetUnderlineColor(issue.Severity)
            End With
            
            ' 保存标注信息
            _annotationMap(issue.Id) = New AnnotationInfo With {
                .Issue = issue,
                .RangeStart = issue.StartPosition,
                .RangeEnd = issue.EndPosition
            }
            
            Debug.WriteLine($"[SmartProofreadFocusMode] 创建标注: {issue.Id}")
            
        Catch ex As Exception
            Debug.WriteLine($"[SmartProofreadFocusMode] 创建标注失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取波浪线颜色（OLE颜色值，格式 0x00BBGGRR）
    ''' </summary>
    Private Function GetUnderlineColor(severity As IssueSeverity) As Integer
        Select Case severity
            Case IssueSeverity.High
                Return &HFF        ' 红色 (B=0, G=0, R=255)
            Case IssueSeverity.Medium
                Return &HFF0000    ' 蓝色 (B=255, G=0, R=0)
            Case IssueSeverity.Low
                Return &HFF00      ' 绿色 (B=0, G=255, R=0)
            Case Else
                Return &HFF        ' 红色
        End Select
    End Function

    ''' <summary>
    ''' 移除标注
    ''' </summary>
    Private Sub RemoveAnnotation(issueId As String)
        If Not _annotationMap.ContainsKey(issueId) Then Return
        
        Dim info = _annotationMap(issueId)
        
        Try
            _annotationMap(issueId).IsIgnored = True
            ' 标注信息保留用于状态跟踪
        Catch ex As Exception
            Debug.WriteLine($"[SmartProofreadFocusMode] 移除标注失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 标记为已忽略
    ''' </summary>
    Private Sub MarkAsIgnored(issueId As String)
        If _annotationMap.ContainsKey(issueId) Then
            _annotationMap(issueId).IsIgnored = True
        End If
    End Sub

    ''' <summary>
    ''' 清除所有标注
    ''' </summary>
    Private Sub ClearAllAnnotations()
        _annotationMap.Clear()
    End Sub

    ''' <summary>
    ''' 显示问题列表面板
    ''' </summary>
    Private Async Function ShowProofreadListPanelAsync(issues As List(Of ProofreadIssue)) As Task
        Dim html = GenerateProofreadListHtml(issues)
        Await _executeScript($"showProofreadList({JsonConvert.SerializeObject(html)});")
    End Function

    ''' <summary>
    ''' 生成校对列表HTML
    ''' </summary>
    Private Function GenerateProofreadListHtml(issues As List(Of ProofreadIssue)) As String
        Dim sb As New StringBuilder()
        
        sb.AppendLine("<div class=""proofread-list"">")
        sb.AppendLine("  <div class=""proofread-list-header"">")
        sb.AppendLine("    <span class=""proofread-list-icon"">🔍</span>")
        sb.AppendLine($"    <span class=""proofread-list-title"">Результат корректуры ({issues.Count} проблем)</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""proofread-summary"" id=""proofread-summary""></div>")
        
        ' 按严重程度分组
        Dim highIssues = issues.Where(Function(i) i.Severity = IssueSeverity.High AndAlso Not i.IsIgnored AndAlso Not i.IsCorrected).ToList()
        Dim mediumIssues = issues.Where(Function(i) i.Severity = IssueSeverity.Medium AndAlso Not i.IsIgnored AndAlso Not i.IsCorrected).ToList()
        Dim lowIssues = issues.Where(Function(i) i.Severity = IssueSeverity.Low AndAlso Not i.IsIgnored AndAlso Not i.IsCorrected).ToList()
        
        ' 高严重程度问题
        If highIssues.Count > 0 Then
            sb.AppendLine("  <div class=""proofread-severity-group"">")
            sb.AppendLine($"    <div class=""severity-header high"">⚠️ Обязательно исправить ({highIssues.Count})</div>")
            For Each issue In highIssues
                sb.AppendLine(GenerateIssueItemHtml(issue))
            Next
            sb.AppendLine("  </div>")
        End If
        
        ' 中等严重程度问题
        If mediumIssues.Count > 0 Then
            sb.AppendLine("  <div class=""proofread-severity-group"">")
            sb.AppendLine($"    <div class=""severity-header medium"">💡 Рекомендуется исправить ({mediumIssues.Count})</div>")
            For Each issue In mediumIssues
                sb.AppendLine(GenerateIssueItemHtml(issue))
            Next
            sb.AppendLine("  </div>")
        End If
        
        ' 低严重程度问题
        If lowIssues.Count > 0 Then
            sb.AppendLine("  <div class=""proofread-severity-group"">")
            sb.AppendLine($"    <div class=""severity-header low"">ℹ️ Необязательные улучшения ({lowIssues.Count})</div>")
            For Each issue In lowIssues.Take(5)
                sb.AppendLine(GenerateIssueItemHtml(issue))
            Next
            If lowIssues.Count > 5 Then
                sb.AppendLine($"    <div class=""proofread-more"">Ещё {lowIssues.Count - 5}...</div>")
            End If
            sb.AppendLine("  </div>")
        End If
        
        ' 批量操作按钮
        sb.AppendLine("  <div class=""proofread-list-actions"">")
        sb.AppendLine("    <button class=""proofread-btn primary"" onclick=""proofreadAcceptAll()"">Принять все</button>")
        sb.AppendLine("    <button class=""proofread-btn secondary"" onclick=""proofreadExit()"">Завершить корректуру</button>")
        sb.AppendLine("  </div>")
        
        sb.AppendLine("</div>")
        
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成单个问题项HTML
    ''' </summary>
    Private Function GenerateIssueItemHtml(issue As ProofreadIssue) As String
        Dim severityClass = issue.Severity.ToString().ToLower()
        Dim issueTypeName = GetIssueTypeName(issue.IssueType)
        
        Return $"
        <div class=""proofread-issue-item {severityClass}"" data-issue-id=""{issue.Id}"">
            <div class=""issue-header"">
                <span class=""issue-location"">Абзац {issue.ParagraphIndex + 1}</span>
                <span class=""issue-type"">{issueTypeName}</span>
            </div>
            <div class=""issue-content"">
                <div class=""issue-original"">
                    <span class=""label"">Оригинал:</span>
                    <span class=""text"">{System.Web.HttpUtility.HtmlEncode(TruncateText(issue.Original, 50))}</span>
                </div>
                <div class=""issue-suggestion"">
                    <span class=""label"">Правка:</span>
                    <span class=""text"">{System.Web.HttpUtility.HtmlEncode(TruncateText(issue.Suggestion, 50))}</span>
                </div>
            </div>
            <div class=""issue-explanation"">{System.Web.HttpUtility.HtmlEncode(TruncateText(issue.Explanation, 100))}</div>
            <div class=""issue-actions"">
                <button class=""issue-btn accept"" data-issue-id=""{issue.Id}"">Принять</button>
                <button class=""issue-btn ignore"" data-issue-id=""{issue.Id}"">Пропустить</button>
            </div>
        </div>"
    End Function

    ''' <summary>
    ''' 获取问题类型名称
    ''' </summary>
    Private Function GetIssueTypeName(issueType As IssueType) As String
        Select Case issueType
            Case IssueType.SpellingError : Return "Орфография"
            Case IssueType.WordUsageError : Return "Словоупотребление"
            Case IssueType.PunctuationError : Return "Пунктуация"
            Case IssueType.GrammaticalError : Return "Грамматика"
            Case IssueType.ExpressionError : Return "Формулировка"
            Case IssueType.FormatError : Return "Формат"
            Case Else : Return "Другое"
        End Select
    End Function

    ''' <summary>
    ''' 截断文本
    ''' </summary>
    Private Function TruncateText(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "..."
    End Function

    ''' <summary>
    ''' 更新校对列表
    ''' </summary>
    Private Async Function UpdateProofreadListAsync() As Task
        If _state.CurrentIssues Is Nothing Then Return
        Dim remaining = _state.CurrentIssues.Where(Function(i) Not i.IsIgnored AndAlso Not i.IsCorrected).ToList()
        Await ShowProofreadListPanelAsync(remaining)
    End Function

    ''' <summary>
    ''' 显示无问题消息
    ''' </summary>
    Private Async Function ShowNoIssuesMessageAsync() As Task
        Await _executeScript("showProofreadNoIssues();")
    End Function

    ''' <summary>
    ''' 显示校对解析失败消息。
    ''' </summary>
    Private Async Function ShowParseErrorMessageAsync(errorMessage As String, rawPreview As String) As Task
        Dim payload As New JObject From {
            {"errorMessage", If(errorMessage, "AI вернул неверный формат, не удалось построить список правок.")},
            {"rawPreview", If(rawPreview, "")}
        }
        Await _executeScript($"showProofreadParseError({payload.ToString(Formatting.None)});")
    End Function

    ''' <summary>
    ''' 显示全部修正完成消息
    ''' </summary>
    Private Async Function ShowAllCorrectedMessageAsync() As Task
        Await _executeScript("showProofreadAllCorrected();")
    End Function

    ''' <summary>
    ''' 显示校对摘要
    ''' </summary>
    Private Async Function ShowProofreadSummaryAsync(issues As List(Of ProofreadIssue)) As Task
        Dim highCount As Integer = 0
        Dim mediumCount As Integer = 0
        Dim lowCount As Integer = 0
        For Each issue In issues
            Select Case issue.Severity
                Case IssueSeverity.High
                    highCount += 1
                Case IssueSeverity.Medium
                    mediumCount += 1
                Case IssueSeverity.Low
                    lowCount += 1
            End Select
        Next
        Dim summary = New ProofreadSummary With {
            .TotalCount = issues.Count,
            .HighCount = highCount,
            .MediumCount = mediumCount,
            .LowCount = lowCount
        }
        Await _executeScript($"updateProofreadSummary({summary.TotalCount}, {summary.HighCount}, {summary.MediumCount}, {summary.LowCount});")
    End Function

End Class
