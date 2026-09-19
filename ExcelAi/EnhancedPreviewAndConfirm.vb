Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop
Imports ShareRibbon
Imports Button = System.Windows.Forms.Button
Imports Font = System.Drawing.Font
Imports Point = System.Drawing.Point
Imports ScrollBars = System.Windows.Forms.ScrollBars
Imports TextBox = System.Windows.Forms.TextBox

Public Class EnhancedPreviewAndConfirm
    ' ���ڱ��湤����״̬��Ϣ����
    Private Class WorksheetState
        Public Name As String
        Public Cells As Dictionary(Of String, Object)
        Public UsedRangeAddress As String
        Public SheetExists As Boolean

        Public Sub New(name As String)
            Me.Name = name
            Me.Cells = New Dictionary(Of String, Object)
            Me.SheetExists = True
        End Sub
    End Class

    ' ���ڱ�ʾ��Ԫ��������
    Private Class CellDifference
        Public Address As String
        Public SheetName As String
        Public OldValue As Object
        Public NewValue As Object
        Public ChangeType As String ' "����", "�޸�", "ɾ��"

        Public Sub New(sheetName As String, address As String, oldValue As Object, newValue As Object, changeType As String)
            Me.SheetName = sheetName
            Me.Address = address
            Me.OldValue = oldValue
            Me.NewValue = newValue
            Me.ChangeType = changeType
        End Sub
    End Class

    ' ���ڱ�ʾ�������������
    Private Class SheetDifference
        Public SheetName As String
        Public ChangeType As String ' "����", "ɾ��", "�޸�"

        Public Sub New(sheetName As String, changeType As String)
            Me.SheetName = sheetName
            Me.ChangeType = changeType
        End Sub
    End Class


    ' ʹ���첽��ʽ������������濨��
    Public Async Function PreviewAndConfirmVbaExecutionAsync(vbaCode As String) As Task(Of Boolean)
        Dim application As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
        Dim originalWorkbook As Workbook = application.ActiveWorkbook

        If originalWorkbook Is Nothing Then
            MessageBox.Show("Нет открытой книги, невозможно показать предпросмотр изменений", "Предпросмотр изменений", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        ' ����1: ����ǰ������״̬
        Dim beforeState = Await Task.Run(Function() CaptureWorkbookState(originalWorkbook))

        ' ����2: ������ʱ������������ִ�д���
        Dim tempWorkbookPath As String = Nothing
        Dim tempWorkbook As Workbook = Nothing
        Dim tempFileName As String = IO.Path.GetTempFileName()

        Try
            ' ʹ��SaveCopyAs����SaveAs����������ı�ԭʼ��������·��
            tempWorkbookPath = IO.Path.ChangeExtension(tempFileName, ".xlsx")
            application.DisplayAlerts = False
            originalWorkbook.SaveCopyAs(tempWorkbookPath)
            application.DisplayAlerts = True

            ' �򿪸ոմ����ĸ���
            tempWorkbook = application.Workbooks.Open(tempWorkbookPath)
            tempWorkbook.Activate() ' ȷ����������ʱ��������ִ��

            ' �첽ִ��VBA
            Dim executionResult = Await Task.Run(Function() ExecuteCodeInTemporaryModule(tempWorkbook, vbaCode))
            If Not executionResult Then Return False

            ' ����4: ����ִ�к��״̬
            Dim afterState = Await Task.Run(Function() CaptureWorkbookState(tempWorkbook))

            ' ����5: �Ƚ�״̬
            Dim cellDifferences As New List(Of CellDifference)()
            Dim sheetDifferences As New List(Of SheetDifference)()
            CompareWorkbookStates(beforeState, afterState, cellDifferences, sheetDifferences)

            ' ����6: ��ʾ�Ż����Ԥ������
            Dim userConfirmed = ShowDifferencePreview(vbaCode, cellDifferences, sheetDifferences)
            Return userConfirmed

        Catch ex As Exception
            MessageBox.Show("Ошибка при выполнении предпросмотра изменений: " & ex.Message, "Предпросмотр изменений", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            application.DisplayAlerts = False
            ' �ر���ʱ������
            If tempWorkbook IsNot Nothing Then
                Try
                    tempWorkbook.Close(SaveChanges:=False)
                    Marshal.ReleaseComObject(tempWorkbook)
                Catch
                End Try
            End If

            ' ���¼���ԭʼ������
            If originalWorkbook IsNot Nothing Then
                Try
                    originalWorkbook.Activate()
                Catch
                    ' ���Լ������
                End Try
            End If
            application.DisplayAlerts = True

            ' ɾ����ʱ�ļ�
            Try
                If tempWorkbookPath IsNot Nothing AndAlso IO.File.Exists(tempWorkbookPath) Then
                    IO.File.Delete(tempWorkbookPath)
                End If
                If IO.File.Exists(tempFileName) Then
                    IO.File.Delete(tempFileName)
                End If
            Catch
                ' ����ɾ����ʱ�ļ��Ĵ���
            End Try
        End Try
    End Function

    ' ��ԭ��ͬ��������Ϊ�����첽���������⿨UI
    Public Function PreviewAndConfirmVbaExecution(vbaCode As String) As Boolean
        Dim ok = ShareRibbon.SyncOverAsync.Run(Function() PreviewAndConfirmVbaExecutionAsync(vbaCode), 300000)
        Return ok
    End Function
    ' �Ż�Ԥ����������
    Private Function ShowDifferencePreview(code As String,
                                          cellDifferences As List(Of CellDifference),
                                          sheetDifferences As List(Of SheetDifference)) As Boolean
        Dim previewForm As New Form() With {
            .Text = "Предварительный просмотр выполнения VBA",
            .Size = New Size(950, 650),
            .StartPosition = FormStartPosition.CenterScreen,
            .MinimizeBox = False,
            .MaximizeBox = True,
            .FormBorderStyle = FormBorderStyle.Sizable
        }

        ' ��TabControl����
        Dim tabControl As New TabControl() With {
            .Dock = DockStyle.Fill
        }

        ' ������
        Dim codeTab As New TabPage("Код VBA")
        Dim codeTextBox As New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .Text = code,
            .Font = New Font("Consolas", 10),
            .Dock = DockStyle.Fill,
            .WordWrap = False
        }
        codeTab.Controls.Add(codeTextBox)

        ' ���������
        Dim sheetTab As New TabPage("Изменения листов")
        Dim sheetListView As New ListView() With {
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .Dock = DockStyle.Fill
        }
        sheetListView.Columns.Add("Имя листа", 150)
        sheetListView.Columns.Add("Тип изменения", 100)

        For Each diff In sheetDifferences
            Dim item As New ListViewItem(diff.SheetName)
            item.SubItems.Add(diff.ChangeType)
            Select Case diff.ChangeType
                Case "����"
                    item.BackColor = Color.LightGreen
                Case "ɾ��"
                    item.BackColor = Color.LightPink
            End Select
            sheetListView.Items.Add(item)
        Next
        If sheetDifferences.Count = 0 Then
            sheetListView.Items.Add(New ListViewItem("Нет изменений листов"))
        End If
        sheetTab.Controls.Add(sheetListView)

        ' ��Ԫ����
        Dim cellTab As New TabPage("Изменения ячеек")
        Dim cellListView As New ListView() With {
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .Dock = DockStyle.Fill
        }
        cellListView.Columns.Add("Лист", 80)
        cellListView.Columns.Add("Ячейка", 80)
        cellListView.Columns.Add("Тип изменения", 80)
        cellListView.Columns.Add("Прежнее значение", 150)
        cellListView.Columns.Add("Новое значение", 150)

        For Each diff In cellDifferences
            Dim item As New ListViewItem(diff.SheetName)
            item.SubItems.Add(diff.Address)
            item.SubItems.Add(diff.ChangeType)
            item.SubItems.Add(If(diff.OldValue Is Nothing, "(пусто)", diff.OldValue.ToString()))
            item.SubItems.Add(If(diff.NewValue Is Nothing, "(пусто)", diff.NewValue.ToString()))
            Select Case diff.ChangeType
                Case "����"
                    item.BackColor = Color.LightGreen
                Case "ɾ��"
                    item.BackColor = Color.LightPink
                Case "�޸�"
                    item.BackColor = Color.LightYellow
            End Select
            cellListView.Items.Add(item)
        Next
        If cellDifferences.Count = 0 Then
            cellListView.Items.Add(New ListViewItem("Нет изменений ячеек"))
        End If
        cellTab.Controls.Add(cellListView)

        ' ժҪ
        Dim summaryTab As New TabPage("Сводка изменений")
        Dim summaryTextBox As New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill,
            .Font = New Font("΢���ź�", 10)
        }
        summaryTextBox.Text = GenerateSummary(sheetDifferences, cellDifferences)
        summaryTab.Controls.Add(summaryTextBox)

        tabControl.TabPages.Add(summaryTab)
        tabControl.TabPages.Add(cellTab)
        tabControl.TabPages.Add(sheetTab)
        tabControl.TabPages.Add(codeTab)

        Dim buttonPanel As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 50
        }

        ' �� buttonPanel ��һ�� FlowLayoutPanel���򻯰�ť����
        Dim flowLayout As New FlowLayoutPanel() With {
            .FlowDirection = FlowDirection.RightToLeft,
            .Dock = DockStyle.Fill
        }
        buttonPanel.Controls.Add(flowLayout)

        Dim acceptButton As New Button() With {
            .Text = "Применить изменения",
            .DialogResult = DialogResult.Yes,
            .AutoSize = True
        }
        Dim cancelButton As New Button() With {
            .Text = "Отмена",
            .DialogResult = DialogResult.No,
            .AutoSize = True
        }

        ' ��ʽ�����´�����������
        flowLayout.Controls.Add(cancelButton)
        flowLayout.Controls.Add(acceptButton)

        ' ���ν� panel��tabControl �ŵ� form
        previewForm.Controls.Add(buttonPanel)
        previewForm.Controls.Add(tabControl)


        ' �Զ���λ
        'acceptButton.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        'cancelButton.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        'acceptButton.Location = New Point(previewForm.ClientSize.Width - 240, 10)
        'cancelButton.Location = New Point(previewForm.ClientSize.Width - 120, 10)

        'buttonPanel.Controls.Add(acceptButton)
        'buttonPanel.Controls.Add(cancelButton)
        'previewForm.Controls.Add(buttonPanel)
        'previewForm.Controls.Add(tabControl)

        ' Ĭ��ȷ�ϰ�ť
        previewForm.AcceptButton = cancelButton
        ' ��ʾ�Ի�������û��㡰Ӧ�ñ�����򷵻� True
        Return (previewForm.ShowDialog() = DialogResult.Yes)
    End Function

    Private Function GenerateSummary(sheetDiffs As List(Of SheetDifference),
                                    cellDiffs As List(Of CellDifference)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("# Сводка изменений")
        sb.AppendLine()

        If sheetDiffs.Count > 0 Then
            sb.AppendLine("## Изменения листов")
            For Each diff In sheetDiffs
                sb.AppendLine($"- {diff.SheetName}: {diff.ChangeType}")
            Next
            sb.AppendLine()
        End If

        If cellDiffs.Count > 0 Then
            Dim grouped = cellDiffs.GroupBy(Function(d) d.SheetName)
            sb.AppendLine("## Изменения ячеек")
            For Each group In grouped
                sb.AppendLine($"### Лист: {group.Key}")
                Dim addCount = group.Count(Function(d) d.ChangeType = "����")
                Dim modifyCount = group.Count(Function(d) d.ChangeType = "�޸�")
                Dim deleteCount = group.Count(Function(d) d.ChangeType = "ɾ��")

                If addCount > 0 Then
                    sb.AppendLine($"- Добавлено: {addCount} ячеек")
                End If
                If modifyCount > 0 Then
                    sb.AppendLine($"- Изменено: {modifyCount} ячеек")
                End If
                If deleteCount > 0 Then
                    sb.AppendLine($"- Удалено: {deleteCount} ячеек")
                End If
                sb.AppendLine()
            Next
        End If

        If sheetDiffs.Count = 0 AndAlso cellDiffs.Count = 0 Then
            sb.AppendLine("После выполнения этого кода изменения данных отсутствуют.")
        End If
        Return sb.ToString()
    End Function


    ' ����ģ���еĵ�һ��������
    Private Function FindFirstProcedureName(comp As VBComponent) As String
        Try
            Dim codeModule As CodeModule = comp.CodeModule
            Dim lineCount As Integer = codeModule.CountOfLines
            Dim line As Integer = 1

            While line <= lineCount
                Dim procName As String = codeModule.ProcOfLine(line, vbext_ProcKind.vbext_pk_Proc)
                If Not String.IsNullOrEmpty(procName) Then
                    Return procName
                End If
                line = codeModule.ProcStartLine(procName, vbext_ProcKind.vbext_pk_Proc) + codeModule.ProcCountLines(procName, vbext_ProcKind.vbext_pk_Proc)
            End While

            Return String.Empty
        Catch
            ' �������������ʹ���������ʽ�Ӵ�������ȡ
            Dim code As String = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            Dim match As Match = Regex.Match(code, "^\s*(Sub|Function)\s+(\w+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)

            If match.Success AndAlso match.Groups.Count > 2 Then
                Return match.Groups(2).Value
            End If

            Return String.Empty
        End Try
    End Function


    ' �������Ƿ������������
    Private Function ContainsProcedureDeclaration(code As String) As Boolean
        ' ʹ�ü򵥵��������ʽ����Ƿ���� Sub �� Function ����
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function

    ' ִ��ǰ�˴����� VBA ����Ƭ��
    Private Function ExecuteCodeInTemporaryModule(workbook As Workbook, vbaCode As String)
        ' ��ȡ VBA ��Ŀ
        Dim vbProj As VBProject = workbook.VBProject

        ' ���ӿ�ֵ���
        If vbProj Is Nothing Then
            Return Nothing
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempPreviewMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' ������ʱģ��
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' �������Ƿ��Ѱ��� Sub/Function ����
            If ContainsProcedureDeclaration(vbaCode) Then
                ' �����Ѱ�������������ֱ������
                vbComp.CodeModule.AddFromString(vbaCode)

                ' ���ҵ�һ����������ִ��
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    workbook.Application.Run(tempModuleName & "." & procName)
                Else
                    'MessageBox.Show("�޷��ڴ������ҵ���ִ�еĹ���")
                    GlobalStatusStrip.ShowWarning("Не удалось найти исполняемую процедуру в коде")
                End If
            Else
                ' ���벻�������������������װ�� Auto_Run ������
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' ִ�� Auto_Run ����
                workbook.Application.Run(tempModuleName & ".Auto_Run")
            End If

        Catch ex As Exception
            MessageBox.Show("Ошибка при выполнении временного кода VBA: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ���۳ɹ�����ʧ�ܣ���ɾ����ʱģ��
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' ������������
            End Try
        End Try
        Return True
    End Function

    ' ��������״̬
    Private Function CaptureWorkbookState(workbook As Workbook) As Dictionary(Of String, WorksheetState)
        Dim state As New Dictionary(Of String, WorksheetState)

        For Each worksheet As Worksheet In workbook.Worksheets
            Dim sheetState As New WorksheetState(worksheet.Name)

            ' ��ȡʹ�÷�Χ
            Dim usedRange As Range = worksheet.UsedRange
            sheetState.UsedRangeAddress = usedRange.Address

            ' �������е�Ԫ���ֵ
            For Each cell As Range In usedRange
                Dim address As String = cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                sheetState.Cells(address) = cell.Value2
            Next

            state(worksheet.Name) = sheetState
        Next

        Return state
    End Function

    ' �ȽϹ�����״̬
    Private Sub CompareWorkbookStates(
        beforeState As Dictionary(Of String, WorksheetState),
        afterState As Dictionary(Of String, WorksheetState),
        cellDifferences As List(Of CellDifference),
        sheetDifferences As List(Of SheetDifference))

        ' ��鹤��������ĸ��ģ�����/ɾ����������
        For Each beforeSheet In beforeState.Values
            If Not afterState.ContainsKey(beforeSheet.Name) Then
                ' ��������ɾ��
                sheetDifferences.Add(New SheetDifference(beforeSheet.Name, "ɾ��"))
            End If
        Next

        For Each afterSheet In afterState.Values
            If Not beforeState.ContainsKey(afterSheet.Name) Then
                ' ������������
                sheetDifferences.Add(New SheetDifference(afterSheet.Name, "����"))

                ' ���������¹������ĵ�Ԫ����Ϊ"����"
                For Each cell In afterSheet.Cells
                    cellDifferences.Add(New CellDifference(
                        afterSheet.Name, cell.Key, Nothing, cell.Value, "����"))
                Next
            Else
                ' ����������������״̬�У��Ƚϵ�Ԫ��
                Dim beforeSheet = beforeState(afterSheet.Name)

                ' ��鵥Ԫ�����
                For Each afterCell In afterSheet.Cells
                    Dim address As String = afterCell.Key
                    Dim newValue As Object = afterCell.Value

                    If beforeSheet.Cells.ContainsKey(address) Then
                        Dim oldValue As Object = beforeSheet.Cells(address)

                        ' �Ƚ�ֵ�Ƿ����
                        If Not AreValuesEqual(oldValue, newValue) Then
                            cellDifferences.Add(New CellDifference(
                                afterSheet.Name, address, oldValue, newValue, "�޸�"))
                        End If
                    Else
                        ' �����ӵĵ�Ԫ��
                        cellDifferences.Add(New CellDifference(
                            afterSheet.Name, address, Nothing, newValue, "����"))
                    End If
                Next

                ' ���ɾ���ĵ�Ԫ��
                For Each beforeCell In beforeSheet.Cells
                    Dim address As String = beforeCell.Key
                    If Not afterSheet.Cells.ContainsKey(address) Then
                        cellDifferences.Add(New CellDifference(
                            beforeSheet.Name, address, beforeCell.Value, Nothing, "ɾ��"))
                    End If
                Next
            End If
        Next
    End Sub

    ' �Ƚ�����ֵ�Ƿ����
    Private Function AreValuesEqual(value1 As Object, value2 As Object) As Boolean
        If value1 Is Nothing AndAlso value2 Is Nothing Then
            Return True
        ElseIf value1 Is Nothing OrElse value2 Is Nothing Then
            Return False
        Else
            Return value1.ToString() = value2.ToString()
        End If
    End Function

End Class