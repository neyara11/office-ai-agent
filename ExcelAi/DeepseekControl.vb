Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Vbe.Interop
Imports ShareRibbon

Public Class DeepseekControl
    Inherits BaseDeepseekChat


    Public Sub New()
        ' 此调用是设计师所必需的。
        Try
            InitializeComponent()
            Debug.WriteLine("DeepseekControl 初始化完成")
        Catch ex As Exception
            SimpleLogger.LogError("DeepseekControl 构造异常", ex)
            MessageBox.Show("Ошибка загрузки DeepseekControl: " & ex.Message)
        End Try

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
            If Globals.ThisAddIn.Application.ActiveWorkbook IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取 Excel 工作目录失败: {ex.Message}")
        End Try
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection Is Nothing OrElse Not TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                Return message
            End If

            Dim range = DirectCast(selection, Microsoft.Office.Interop.Excel.Range)
            Dim text As String = ConvertRangeToText(range)
            If String.IsNullOrWhiteSpace(text) Then Return message

            Return message & vbCrLf & vbCrLf &
                   "--- 当前 Excel 选区 ---" & vbCrLf &
                   $"工作表: {range.Worksheet.Name}" & vbCrLf &
                   $"范围: {range.Address(False, False)}" & vbCrLf &
                   text
        Catch ex As Exception
            Debug.WriteLine($"附加 Excel 选区失败: {ex.Message}")
            Return message
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function

    Private Function ConvertRangeToText(range As Microsoft.Office.Interop.Excel.Range) As String
        Dim sb As New System.Text.StringBuilder()
        Try
            Dim rowCount As Integer = Math.Min(range.Rows.Count, 50)
            Dim colCount As Integer = Math.Min(range.Columns.Count, 20)

            For r As Integer = 1 To rowCount
                Dim rowValues As New List(Of String)()
                For c As Integer = 1 To colCount
                    Dim value = range.Cells(r, c).Value
                    rowValues.Add(If(value Is Nothing, "", value.ToString()))
                Next
                sb.AppendLine(String.Join(vbTab, rowValues))
            Next

            If range.Rows.Count > rowCount OrElse range.Columns.Count > colCount Then
                sb.AppendLine($"[选区过大，仅包含前 {rowCount} 行、{colCount} 列]")
            End If
        Catch ex As Exception
            Debug.WriteLine($"读取 Excel 选区失败: {ex.Message}")
        End Try
        Return sb.ToString().TrimEnd()
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


    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean)
        ' 如果需要预览
        Dim previewTool As New EnhancedPreviewAndConfirm()
        ' 允许用户预览代码变更
        If previewTool.PreviewAndConfirmVbaExecution(vbaCode) Then
            Debug.Print("预览结束，用户同意执行代码: " & vbaCode)
            Return True
        Else
            ' 用户取消或拒绝
            Return False
        End If
    End Function

    ' 提供Excel应用程序对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function


    ' 实现Excel公式评估' 执行Excel公式或函数 - 增强版支持赋值和预览
    Protected Overrides Function EvaluateFormula(formulaCode As String, preview As Boolean) As Boolean
        Try
            ' 检查是否是赋值语句 (例如 C1=A1+B1)
            Dim isAssignment As Boolean = Regex.IsMatch(formulaCode, "^[A-Za-z]+[0-9]+\s*=")

            If isAssignment Then
                ' 解析赋值语句
                Dim parts As String() = formulaCode.Split(New Char() {"="c}, 2)
                Dim targetCell As String = parts(0).Trim()
                Dim formula As String = parts(1).Trim()

                ' 如果公式以=开头，则移除
                If formula.StartsWith("=") Then
                    formula = formula.Substring(1)
                End If

                ' 如果需要预览，显示预览对话框
                If preview Then
                    Dim excel As Object = Globals.ThisAddIn.Application
                    Dim currentValue As Object = Nothing
                    Try
                        currentValue = excel.Range(targetCell).Value
                    Catch ex As Exception
                        ' 单元格可能不存在值
                    End Try

                    ' 计算新值
                    Dim newValue As Object = excel.Evaluate(formula)

                    ' 创建预览对话框
                    Dim previewMsg As String = $"В ячейку {targetCell} будет применена формула:" & vbCrLf & vbCrLf &
                                          $"={formula}" & vbCrLf & vbCrLf &
                                          $"Текущее значение: {If(currentValue Is Nothing, "(пусто)", currentValue)}" & vbCrLf &
                                          $"Новое значение: {If(newValue Is Nothing, "(пусто)", newValue)}"

                    Dim result As DialogResult = MessageBox.Show(previewMsg, "Предпросмотр формулы Excel",
                                                          MessageBoxButtons.OKCancel,
                                                          MessageBoxIcon.Information)

                    If result <> DialogResult.OK Then
                        Return False
                    End If
                End If

                ' 执行赋值
                Dim range As Object = Globals.ThisAddIn.Application.Range(targetCell)
                range.Formula = "=" & formula

                GlobalStatusStrip.ShowInfo($"Формула '={formula}' применена к ячейке {targetCell}")
                Return True
            Else
                ' 普通公式计算 (不包含赋值)
                ' 去除可能的等号前缀
                If formulaCode.StartsWith("=") Then
                    formulaCode = formulaCode.Substring(1)
                End If

                ' 计算公式结果
                Dim result As Object = Globals.ThisAddIn.Application.Evaluate(formulaCode)

                ' 如果需要预览，显示计算结果
                If preview Then
                    Dim previewMsg As String = $"Результат вычисления формулы:" & vbCrLf & vbCrLf &
                                         $"={formulaCode}" & vbCrLf & vbCrLf &
                                         $"Результат: {If(result Is Nothing, "(пусто)", result)}"

                    MessageBox.Show(previewMsg, "Результат формулы Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' 显示结果
                    GlobalStatusStrip.ShowInfo($"Результат формулы '={formulaCode}': {result}")
                End If

                Return True
            End If
        Catch ex As Exception
            MessageBox.Show("Ошибка при выполнении формулы Excel: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
End Class
