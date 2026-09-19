' ShareRibbon\Controls\HostTaskPaneRegistry.vb
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Tools

''' <summary>
''' 按 Office 文档窗口维护 CustomTaskPane 的生命周期。
''' Word/Excel/PowerPoint 使用不同的窗口类型，因此窗口对象与句柄解析由宿主提供委托。
''' 每个窗口拥有独立面板和控件，避免多窗口时所有窗口共用同一个（第一个）面板。
''' </summary>
Public Class HostTaskPaneRegistry

    Public Class Entry
        Public Property Handle As Integer
        Public Property Control As UserControl
        Public Property Pane As CustomTaskPane
        Public Property HtmlLoaded As Boolean
        ''' <summary>本次是否为新建（用于只在创建时挂一次事件）。</summary>
        Public Property IsNew As Boolean
    End Class

    Private ReadOnly _panes As New Dictionary(Of Integer, Entry)()
    Private ReadOnly _collection As CustomTaskPaneCollection
    Private ReadOnly _getWindowHandle As Func(Of Object, Integer)

    Public Sub New(collection As CustomTaskPaneCollection, getWindowHandle As Func(Of Object, Integer))
        If collection Is Nothing Then Throw New ArgumentNullException(NameOf(collection))
        If getWindowHandle Is Nothing Then Throw New ArgumentNullException(NameOf(getWindowHandle))
        _collection = collection
        _getWindowHandle = getWindowHandle
    End Sub

    ''' <summary>当前注册的所有面板（快照）。</summary>
    Public ReadOnly Property Entries As IEnumerable(Of Entry)
        Get
            Return New List(Of Entry)(_panes.Values)
        End Get
    End Property

    ''' <summary>返回指定窗口句柄的既有面板，不存在返回 Nothing。</summary>
    Public Function GetExisting(handle As Integer) As Entry
        Dim entry As Entry = Nothing
        _panes.TryGetValue(handle, entry)
        Return entry
    End Function

    ''' <summary>
    ''' 返回指定窗口的既有面板；不存在时用 controlFactory 创建控件并绑定到该窗口。
    ''' </summary>
    Public Function GetOrCreate(window As Object, title As String, controlFactory As Func(Of UserControl)) As Entry
        If window Is Nothing OrElse controlFactory Is Nothing Then Return Nothing
        ' 句柄为 0 时（个别环境取不到 HWND）仍按 0 作为单一回退键，保持旧行为
        Dim handle As Integer = _getWindowHandle(window)

        Dim existing As Entry = Nothing
        If _panes.TryGetValue(handle, existing) Then
            If existing.Pane IsNot Nothing Then
                existing.IsNew = False
                EnsureStatusStripAttached(existing.Control)
                Return existing
            End If
            _panes.Remove(handle)
        End If

        Dim control = controlFactory()
        If control Is Nothing Then Return Nothing

        Try
            Dim pane = _collection.Add(control, title, window)
            Dim entry = New Entry With {
                .Handle = handle,
                .Control = control,
                .Pane = pane,
                .IsNew = True
            }
            _panes(handle) = entry
            EnsureStatusStripAttached(control)
            Return entry
        Catch
            Try
                control.Dispose()
            Catch ex As Exception
                Debug.WriteLine($"[HostTaskPaneRegistry] 释放控件失败: {ex.Message}")
            End Try
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 移除对应窗口/文稿已关闭的孤儿面板。
    ''' liveHandlesProvider 返回 Nothing 时表示无法判断，跳过清理。
    ''' </summary>
    Public Sub PruneClosed(liveHandlesProvider As Func(Of HashSet(Of Integer)))
        If liveHandlesProvider Is Nothing Then Return

        Dim liveHandles As HashSet(Of Integer) = Nothing
        Try
            liveHandles = liveHandlesProvider()
        Catch ex As Exception
            Debug.WriteLine($"[HostTaskPaneRegistry] 枚举活动窗口失败: {ex.Message}")
            Return
        End Try
        If liveHandles Is Nothing Then Return

        Dim dead As New List(Of Integer)
        For Each kv In _panes
            Dim alive As Boolean = False
            Try
                Dim paneWindow = kv.Value.Pane.Window
                If paneWindow IsNot Nothing Then
                    alive = liveHandles.Contains(_getWindowHandle(paneWindow))
                End If
            Catch
                alive = False
            End Try
            If Not alive Then dead.Add(kv.Key)
        Next

        For Each handle In dead
            RemoveByHandle(handle)
        Next
    End Sub

    ''' <summary>移除并释放指定窗口句柄的面板。</summary>
    Public Function RemoveByHandle(handle As Integer) As Boolean
        Dim entry As Entry = Nothing
        If Not _panes.TryGetValue(handle, entry) Then Return False
        _panes.Remove(handle)
        Detach(entry)
        Return True
    End Function

    ''' <summary>把共享的全局状态栏挂到当前活动控件的底部（多窗口时跟随活动面板）。</summary>
    Private Shared Sub EnsureStatusStripAttached(control As UserControl)
        Try
            If control Is Nothing Then Return
            If GlobalStatusStrip.StatusStrip.Parent Is control Then Return
            control.Controls.Add(GlobalStatusStrip.StatusStrip)
        Catch ex As Exception
            Debug.WriteLine($"[HostTaskPaneRegistry] 挂载状态栏失败: {ex.Message}")
        End Try
    End Sub

    Private Sub Detach(entry As Entry)
        Try
            ' 先从控件上摘下共享的全局状态栏，避免 Remove 连带释放它
            If entry.Control IsNot Nothing AndAlso GlobalStatusStrip.StatusStrip.Parent Is entry.Control Then
                entry.Control.Controls.Remove(GlobalStatusStrip.StatusStrip)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[HostTaskPaneRegistry] 摘除状态栏失败: {ex.Message}")
        End Try
        Try
            If entry.Pane IsNot Nothing Then _collection.Remove(entry.Pane)
        Catch ex As Exception
            Debug.WriteLine($"[HostTaskPaneRegistry] 移除面板失败: {ex.Message}")
        End Try
    End Sub
End Class
