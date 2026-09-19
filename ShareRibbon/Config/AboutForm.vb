' ShareRibbon\Config\AboutForm.vb
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' 关于对话框 - 显示插件信息和开源地址
''' </summary>
Public Class AboutForm
    Inherits Form

    Private lblTitle As Label
    Private lblDescription As Label
    Private lblDataPath As Label
    Private btnClose As Button

    Public Sub New()
        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        Me.Text = "О программе Office MOSS"
        Me.Size = New Size(470, 340)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White

        ' 标题
        lblTitle = New Label()
        lblTitle.Text = "Помощник Office MOSS"
        lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(74, 111, 165)
        lblTitle.Location = New Point(20, 20)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        ' 描述
        lblDescription = New Label()
        lblDescription.Text = "AI-ассистент для Word, Excel и PowerPoint." & vbCrLf & vbCrLf &
                             "Понимает команды на русском языке и помогает с анализом данных," & vbCrLf &
                             "форматированием, вычиткой и переводом документов." & vbCrLf & vbCrLf &
                             "Сборка предназначена для работы во внутреннем контуре и не" & vbCrLf &
                             "обращается к внешним сервисам и сайтам."
        lblDescription.Font = New Font("Segoe UI", 9)
        lblDescription.ForeColor = Color.FromArgb(80, 80, 80)
        lblDescription.Location = New Point(20, 55)
        lblDescription.Size = New Size(420, 150)
        Me.Controls.Add(lblDescription)

        ' 数据路径
        lblDataPath = New Label()
        lblDataPath.Text = "Каталог данных: Документы\" & ConfigSettings.OfficeAiAppDataFolder
        lblDataPath.Font = New Font("Segoe UI", 9)
        lblDataPath.ForeColor = Color.Gray
        lblDataPath.Location = New Point(20, 215)
        lblDataPath.AutoSize = True
        Me.Controls.Add(lblDataPath)

        ' 关闭按钮
        btnClose = New Button()
        btnClose.Text = "Закрыть"
        btnClose.Size = New Size(80, 30)
        btnClose.Location = New Point(370, 250)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.BackColor = Color.FromArgb(74, 111, 165)
        btnClose.ForeColor = Color.White
        btnClose.Font = New Font("Segoe UI", 9)
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        Me.Controls.Add(btnClose)
        Me.AcceptButton = btnClose
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class
