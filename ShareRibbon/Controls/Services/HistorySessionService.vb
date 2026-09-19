Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 历史会话服务：会话列表、加载和新建会话。
''' </summary>
Public Class HistorySessionService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _chatStateService As ChatStateService
    Private ReadOnly _getAppType As Func(Of String)
    Private ReadOnly _invokeOnUiThread As Action(Of Action)

    Public Sub New(
        executeScript As Func(Of String, Task),
        chatStateService As ChatStateService,
        getAppType As Func(Of String),
        invokeOnUiThread As Action(Of Action))

        _executeScript = executeScript
        _chatStateService = chatStateService
        _getAppType = getAppType
        _invokeOnUiThread = invokeOnUiThread
    End Sub

    ''' <summary>
    ''' 获取近期会话列表（来自 session_summary），供历史侧边栏展示
    ''' </summary>
    Public Async Sub HandleGetSessionList()
        Dim errorScript As String = Nothing
        Try
            Dim limit As Integer = 50
            Dim summaries = MemoryRepository.GetRecentSessionSummaries(limit)
            Dim list As New List(Of Object)()
            For Each s In summaries
                list.Add(New With {
                    .sessionId = s.SessionId,
                    .title = If(String.IsNullOrEmpty(s.Title), "Сессия", s.Title),
                    .snippet = If(String.IsNullOrEmpty(s.Snippet), "", s.Snippet),
                    .createdAt = s.CreatedAt,
                    .fileName = s.Title,
                    .fullPath = s.SessionId,
                    .lastModified = s.CreatedAt
                })
            Next
            Dim jsonResult As String = JsonConvert.SerializeObject(list)
            Await _executeScript($"setHistoryFilesList({jsonResult});")
        Catch ex As Exception
            Debug.WriteLine("HandleGetSessionList 失败: " & ex.Message)
            errorScript = "setHistoryFilesList([]);"
        End Try
        If errorScript IsNot Nothing Then
            Await _executeScript(errorScript)
        End If
    End Sub

    ''' <summary>
    ''' 加载指定会话到当前 Chat 并渲染消息
    ''' </summary>
    Public Async Sub HandleLoadSession(jsonDoc As JObject)
        Try
            Dim sessionId As String = jsonDoc("sessionId")?.ToString()
            If String.IsNullOrEmpty(sessionId) Then Return
            _chatStateService.SwitchToSession(sessionId)
            Dim messages As New List(Of Object)()
            For Each m In _chatStateService.HistoryMessages
                If m.role = "user" OrElse m.role = "assistant" Then
                    messages.Add(New With {.role = m.role, .content = m.content, .createTime = m.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")})
                End If
            Next
            Dim jsonResult As String = JsonConvert.SerializeObject(messages)
            Await _executeScript($"setChatMessages({jsonResult});")
        Catch ex As Exception
            Debug.WriteLine("HandleLoadSession 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("Не удалось загрузить сессию")
        End Try
    End Sub

    ''' <summary>
    ''' 新建会话：清空状态并清空聊天区域
    ''' </summary>
    Public Async Sub HandleNewSession()
        Try
            _chatStateService.StartNewSession()
            Await _executeScript("if(typeof clearChatContent==='function')clearChatContent();")
            GlobalStatusStrip.ShowInfo("Создана новая сессия")
        Catch ex As Exception
            Debug.WriteLine("HandleNewSession 失败: " & ex.Message)
        End Try
    End Sub

End Class
