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
    Private lblBili As LinkLabel
    Private lblGitee As LinkLabel
    Private lblGithub As LinkLabel
    Private btnClose As Button

    Public Sub New()
        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        Me.Text = "О программе Office MOSS"
        Me.Size = New Size(450, 420)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White

        ' 标题
        lblTitle = New Label()
        lblTitle.Text = "Помощник Office MOSS"
        lblTitle.Font = New Font("微软雅黑", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(74, 111, 165)
        lblTitle.Location = New Point(20, 20)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        ' 描述
        lblDescription = New Label()
        lblDescription.Text = "Привет! Я Цзюньгэ с Bilibili, мой канал — «君哥聊编程»." & vbCrLf & vbCrLf &
                             "Идея плагина пришла от подписчика на Bilibili, который работает" & vbCrLf &
                             "банковским аудитором и постоянно имеет дело с таблицами." & vbCrLf &
                             "Часто данные в таблицах невозможно вычислить по фиксированным" & vbCrLf &
                             "формулам, хотя для человека они имеют одинаковый смысл." & vbCrLf &
                             "Так появился Office MOSS." & vbCrLf & vbCrLf &
                             "Плагин продолжает развиваться — оставляйте комментарии и предложения, чтобы мы вместе делали его лучше."
        lblDescription.Font = New Font("微软雅黑", 9)
        lblDescription.ForeColor = Color.FromArgb(80, 80, 80)
        lblDescription.Location = New Point(20, 55)
        lblDescription.Size = New Size(400, 130)
        Me.Controls.Add(lblDescription)

        ' 数据路径
        lblDataPath = New Label()
        lblDataPath.Text = "Каталог данных: Документы\" & ConfigSettings.OfficeAiAppDataFolder
        lblDataPath.Font = New Font("微软雅黑", 9)
        lblDataPath.ForeColor = Color.Gray
        lblDataPath.Location = New Point(20, 190)
        lblDataPath.AutoSize = True
        Me.Controls.Add(lblDataPath)

        ' 开源地址标题
        Dim lblOpenSource As New Label()
        lblOpenSource.Text = "Открытый исходный код:"
        lblOpenSource.Font = New Font("微软雅黑", 9, FontStyle.Bold)
        lblOpenSource.Location = New Point(20, 225)
        lblOpenSource.AutoSize = True
        Me.Controls.Add(lblOpenSource)

        ' B站链接
        lblBili = New LinkLabel()
        lblBili.Text = "bilibili: https://www.bilibili.com/video/BV17vNRz1ELn"
        lblBili.Font = New Font("微软雅黑", 9)
        lblBili.Location = New Point(20, 250)
        lblBili.AutoSize = True
        lblBili.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblBili.LinkClicked, AddressOf Bilibili_LinkClicked
        Me.Controls.Add(lblBili)

        ' Gitee链接
        lblGitee = New LinkLabel()
        lblGitee.Text = "Gitee: https://gitee.com/it235/office-ai-agent"
        lblGitee.Font = New Font("微软雅黑", 9)
        lblGitee.Location = New Point(20, 275)
        lblGitee.AutoSize = True
        lblGitee.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblGitee.LinkClicked, AddressOf Gitee_LinkClicked
        Me.Controls.Add(lblGitee)

        ' Github链接
        lblGithub = New LinkLabel()
        lblGithub.Text = "Github: https://github.com/it235/office-ai-agent"
        lblGithub.Font = New Font("微软雅黑", 9)
        lblGithub.Location = New Point(20, 300)
        lblGithub.AutoSize = True
        lblGithub.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblGithub.LinkClicked, AddressOf Github_LinkClicked
        Me.Controls.Add(lblGithub)

        ' 关闭按钮
        btnClose = New Button()
        btnClose.Text = "Закрыть"
        btnClose.Size = New Size(80, 30)
        btnClose.Location = New Point(350, 330)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.BackColor = Color.FromArgb(74, 111, 165)
        btnClose.ForeColor = Color.White
        btnClose.Font = New Font("微软雅黑", 9)
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        Me.Controls.Add(btnClose)
        Me.AcceptButton = btnClose
    End Sub

    Private Sub Bilibili_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://www.bilibili.com/video/BV17vNRz1ELn")
        Catch ex As Exception
            MessageBox.Show("Не удалось открыть ссылку: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Gitee_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://gitee.com/it235/office-ai-agent")
        Catch ex As Exception
            MessageBox.Show("Не удалось открыть ссылку: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Github_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://github.com/it235/office-ai-agent")
        Catch ex As Exception
            MessageBox.Show("Не удалось открыть ссылку: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class
