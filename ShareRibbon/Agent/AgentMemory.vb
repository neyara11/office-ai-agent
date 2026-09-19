Imports System.IO
Imports System.Linq
Imports Newtonsoft.Json

Namespace Agent

    ''' <summary>
    ''' 会话消息
    ''' </summary>
    Public Class SessionMessage
        Public Property Role As String
        Public Property Content As String
        Public Property Timestamp As DateTime = DateTime.Now
        Public Property IsSummary As Boolean = False
    End Class

    Friend Class AgentLongTermMemoryData
        Public Property TaskHistory As New List(Of AgentTaskRecord)()
        Public Property LongTermMemory As New Dictionary(Of String, String)()
    End Class

    Friend Class AgentTaskRecord
        Public Property Id As String = Guid.NewGuid().ToString()
        Public Property Timestamp As DateTime = DateTime.Now
        Public Property UserInput As String
        Public Property Intent As String
        Public Property Plan As String
        Public Property Result As String
        Public Property Success As Boolean
        Public Property ApplicationType As String
    End Class

    ''' <summary>
    ''' Agent 三层 Memory 系统
    ''' Working(步骤级) -> Short-term(会话级) -> Long-term(跨会话)
    ''' </summary>
    Public Class AgentMemory

        ' === Working Memory (步骤级，内存) ===
        Private ReadOnly _workingContext As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        ' === Short-term Memory (会话级，内存) ===
        Private ReadOnly _sessionHistory As New List(Of SessionMessage)()
        Private Const CompactThreshold As Integer = 20
        Private Const CompactBatchSize As Integer = 10

        ' === Long-term Memory (跨会话，JSON文件) ===
        Private _longTermData As AgentLongTermMemoryData
        Private ReadOnly _memoryFilePath As String
        Private ReadOnly _lock As New Object()

        ' AI 请求委托（用于摘要压缩）
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))

        Public Sub New()
            _memoryFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "OfficeAiAgent",
                "agent_memory.json"
            )
            LoadLongTerm()
        End Sub

#Region "Working Memory"

        Public Sub SetWorking(key As String, value As Object)
            SyncLock _lock
                _workingContext(key) = value
            End SyncLock
        End Sub

        Public Function GetWorking(key As String) As Object
            SyncLock _lock
                If _workingContext.ContainsKey(key) Then
                    Return _workingContext(key)
                End If
                Return Nothing
            End SyncLock
        End Function

        Public Function GetWorkingString(key As String) As String
            Dim val = GetWorking(key)
            If val Is Nothing Then Return ""
            Return val.ToString()
        End Function

        Public Sub ClearWorking()
            SyncLock _lock
                _workingContext.Clear()
            End SyncLock
        End Sub

#End Region

#Region "Short-term Memory"

        Public Sub AddSessionMessage(role As String, content As String)
            SyncLock _lock
                _sessionHistory.Add(New SessionMessage With {
                    .Role = role,
                    .Content = content,
                    .Timestamp = DateTime.Now
                })

                ' 超过阈值时压缩旧消息
                If _sessionHistory.Count > CompactThreshold Then
                    Dim unused = CompactOldMessagesAsync()
                End If
            End SyncLock
        End Sub

        Public Function GetRecentMessages(count As Integer) As List(Of HistoryMessage)
            SyncLock _lock
                Dim result As New List(Of HistoryMessage)()
                Dim startIdx = Math.Max(0, _sessionHistory.Count - count)
                For i = startIdx To _sessionHistory.Count - 1
                    Dim msg = _sessionHistory(i)
                    result.Add(New HistoryMessage With {
                        .role = msg.Role,
                        .content = msg.Content
                    })
                Next
                Return result
            End SyncLock
        End Function

        Private Async Function CompactOldMessagesAsync() As Task
            If SendAIRequest Is Nothing Then Return

            Dim oldMessages As List(Of SessionMessage)
            SyncLock _lock
                If _sessionHistory.Count <= CompactBatchSize Then Return
                oldMessages = _sessionHistory.Take(CompactBatchSize).ToList()
            End SyncLock

            Dim sb As New Text.StringBuilder()
            For Each msg In oldMessages
                sb.AppendLine($"[{msg.Role}] {msg.Content}")
            Next

            Dim summaryPrompt = $"请将以下对话历史压缩为一段简洁的摘要，保留关键决策和上下文：

{sb.ToString()}"

            Try
                Dim summary = Await SendAIRequest(summaryPrompt, "Ты специалист по краткому изложению диалогов. Отвечай только на русском языке, кратко и по делу.", Nothing)
                If Not String.IsNullOrWhiteSpace(summary) Then
                    SyncLock _lock
                        If _sessionHistory.Count > CompactBatchSize Then
                            _sessionHistory.RemoveRange(0, CompactBatchSize)
                            _sessionHistory.Insert(0, New SessionMessage With {
                                .Role = "system",
                                .Content = $"[历史摘要] {summary.Trim()}",
                                .IsSummary = True
                            })
                        End If
                    End SyncLock
                End If
            Catch ex As Exception
                Debug.WriteLine($"[AgentMemory] 摘要压缩失败: {ex.Message}")
            End Try
        End Function

#End Region

#Region "Long-term Memory"

        Private Sub LoadLongTerm()
            Try
                SyncLock _lock
                    If File.Exists(_memoryFilePath) Then
                        Dim json = File.ReadAllText(_memoryFilePath)
                        _longTermData = JsonConvert.DeserializeObject(Of AgentLongTermMemoryData)(json)
                    End If
                End SyncLock
            Catch ex As Exception
                Debug.WriteLine($"[AgentMemory] 加载长期记忆失败: {ex.Message}")
            End Try

            If _longTermData Is Nothing Then
                _longTermData = New AgentLongTermMemoryData()
            End If

            ' 确保集合初始化
            If _longTermData.TaskHistory Is Nothing Then _longTermData.TaskHistory = New List(Of AgentTaskRecord)()
            If _longTermData.LongTermMemory Is Nothing Then _longTermData.LongTermMemory = New Dictionary(Of String, String)()
        End Sub

        Public Sub SaveLongTerm()
            Try
                SyncLock _lock
                    Dim dir = Path.GetDirectoryName(_memoryFilePath)
                    If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                    Dim json = JsonConvert.SerializeObject(_longTermData, Formatting.Indented)
                    File.WriteAllText(_memoryFilePath, json)
                End SyncLock
            Catch ex As Exception
                Debug.WriteLine($"[AgentMemory] 保存长期记忆失败: {ex.Message}")
            End Try
        End Sub

        Public Sub AddTaskRecord(result As AgentResult)
            SyncLock _lock
                _longTermData.TaskHistory.Add(New AgentTaskRecord With {
                    .UserInput = result.SessionId,
                    .Intent = "agent_task",
                    .Plan = $"迭代次数: {result.IterationsCompleted}",
                    .Result = result.Message,
                    .Success = result.Success,
                    .ApplicationType = "agent"
                })

                ' 保持历史不超过100条
                If _longTermData.TaskHistory.Count > 100 Then
                    _longTermData.TaskHistory.RemoveAt(0)
                End If

                SaveLongTerm()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 混合检索：关键词 + 简单语义匹配
        ''' </summary>
        Public Function Search(query As String, topK As Integer) As List(Of String)
            Dim results As New List(Of String)()
            If String.IsNullOrWhiteSpace(query) Then Return results

            SyncLock _lock
                ' 1. 关键词匹配
                Dim keywords = query.ToLower().Split({" "c, "，"c, ","c}, StringSplitOptions.RemoveEmptyEntries) _
                    .Where(Function(k) k.Length > 1).ToList()

                ' 搜索长期记忆
                For Each kvp In _longTermData.LongTermMemory
                    Dim score = 0
                    For Each kw In keywords
                        If kvp.Key.ToLower().Contains(kw) OrElse kvp.Value.ToLower().Contains(kw) Then
                            score += 1
                        End If
                    Next
                    If score > 0 Then
                        results.Add($"{kvp.Key}: {kvp.Value}")
                    End If
                Next

                ' 搜索任务历史
                For Each task In _longTermData.TaskHistory.Take(20)
                    Dim score = 0
                    For Each kw In keywords
                        If (task.UserInput > "" AndAlso task.UserInput.ToLower().Contains(kw)) OrElse
                           (task.Plan > "" AndAlso task.Plan.ToLower().Contains(kw)) Then
                            score += 1
                        End If
                    Next
                    If score > 0 Then
                        results.Add($"[历史任务 {task.Timestamp:MM-dd}] {task.UserInput}: {If(task.Success, "成功", "失败")}")
                    End If
                Next
            End SyncLock

            Return results.Take(topK).ToList()
        End Function

#End Region

    End Class

End Namespace
