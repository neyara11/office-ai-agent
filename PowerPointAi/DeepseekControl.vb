Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Vbe.Interop
Imports ShareRibbon

Public Class DeepseekControl
    Inherits BaseDeepseekChat


    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        'InitializeWebView2Script()
    End Sub

    Protected Overrides Sub SendChatMessage(message As String)
        Debug.WriteLine("DeepseekControl SendChatMessage: " & message)
    End Sub

    Protected Overrides Sub GetSelectionContent(target As Object)
        ' Deepseek 网页控件不维护本地引用卡片；发送时由 AppendCurrentSelectedContent 读取当前选区。
    End Sub

    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            If Globals.ThisAddIn.Application.ActivePresentation IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActivePresentation.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取 PowerPoint 工作目录失败: {ex.Message}")
        End Try
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Dim selection As Microsoft.Office.Interop.PowerPoint.Selection = Nothing
        Try
            selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then Return message

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("--- 当前 PowerPoint 选区 ---")

            Select Case selection.Type
                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText
                    sb.AppendLine(selection.TextRange.Text)
                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes
                    For i As Integer = 1 To selection.ShapeRange.Count
                        Dim shp As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                        Try
                            shp = selection.ShapeRange(i)
                            If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                               shp.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                sb.AppendLine(shp.TextFrame.TextRange.Text)
                            End If
                        Finally
                            ComObjectHelper.ReleaseComObject(shp)
                        End Try
                    Next
                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides
                    For Each slide As Microsoft.Office.Interop.PowerPoint.Slide In selection.SlideRange
                        Try
                        sb.AppendLine($"幻灯片 {slide.SlideIndex}")
                        For Each shp As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                            Try
                                If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                                   shp.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    sb.AppendLine(shp.TextFrame.TextRange.Text)
                                End If
                            Finally
                                ComObjectHelper.ReleaseComObject(shp)
                            End Try
                        Next
                        Finally
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next
                Case Else
                    Return message
            End Select

            Dim selectedText = sb.ToString().TrimEnd()
            If String.IsNullOrWhiteSpace(selectedText) OrElse selectedText = "--- 当前 PowerPoint 选区 ---" Then Return message
            Return message & vbCrLf & vbCrLf & selectedText
        Catch ex As Exception
            Debug.WriteLine($"附加 PowerPoint 选区失败: {ex.Message}")
            Return message
        Finally
            ComObjectHelper.ReleaseComObject(selection)
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    Protected Overrides Function GetVBProject() As VBProject
        Try
            Dim project = Globals.ThisAddIn.Application.VBE.ActiveVBProject
            Return project
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return Nothing
        End Try
    End Function

    Protected Overrides Function RunCode(code As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(code)
            Return True
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return False
        Catch ex As Exception
            MessageBox.Show("Ошибка при выполнении кода: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 执行前预览代码
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean)
        Return True
    End Function

    ' 提供Excel应用程序对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function


    ' PowerPoint 不支持公式评估，此功能仅适用于 Excel
    Protected Overrides Function EvaluateFormula(formulaCode As String, preview As Boolean) As Boolean
        GlobalStatusStrip.ShowWarning("Оценка формул не поддерживается в PowerPoint")
        Return False
    End Function
End Class
