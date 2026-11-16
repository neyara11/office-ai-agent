' ShareRibbon\Config\AboutForm.vb
Imports System.Diagnostics
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' О программе: версия, конфигурация и путь к загруженной сборке.
''' Раскладка на FlowLayoutPanel, чтобы текст не наезжал друг на друга
''' при любом DPI и длине пути.
''' </summary>
Public Class AboutForm
    Inherits Form

    Public Sub New()
        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        Me.Text = "О программе Office MOSS"
        Me.ClientSize = New Size(620, 470)
        Me.MinimumSize = New Size(520, 420)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.Font = New Font("Segoe UI", 9)

        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 2
        root.Padding = New Padding(22, 18, 22, 16)
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Me.Controls.Add(root)

        Dim content As New FlowLayoutPanel()
        content.Dock = DockStyle.Fill
        content.FlowDirection = FlowDirection.TopDown
        content.WrapContents = False
        content.AutoScroll = True
        content.Padding = New Padding(0)
        root.Controls.Add(content, 0, 0)

        Dim hostAssembly = FindHostAssembly()
        Dim hostName = If(hostAssembly IsNot Nothing, hostAssembly.GetName().Name, "")
        Dim hostVersionText = GetAssemblyVersionText(hostAssembly)
        Dim coreVersionText = GetAssemblyVersionText(GetType(AboutForm).Assembly)
        Dim configuration = If(IsDebugBuild(), "Debug", "Release")

        Dim title As New Label()
        title.Text = "Помощник Office MOSS"
        title.Font = New Font("Segoe UI", 18, FontStyle.Bold)
        title.ForeColor = Color.FromArgb(74, 111, 165)
        title.AutoSize = True
        title.Margin = New Padding(0, 0, 0, 2)
        content.Controls.Add(title)

        Dim subtitle As New Label()
        Dim displayVersion = If(hostVersionText.Length > 0, hostVersionText, coreVersionText)
        subtitle.Text = $"Версия {displayVersion} · {configuration}"
        subtitle.Font = New Font("Segoe UI", 11.5F, FontStyle.Bold)
        subtitle.ForeColor = Color.FromArgb(60, 60, 60)
        subtitle.AutoSize = True
        subtitle.Margin = New Padding(0, 0, 0, 10)
        content.Controls.Add(subtitle)

        Dim details As New Label()
        Dim detailText As New StringBuilder()
        If hostName.Length > 0 Then
            detailText.AppendLine($"Надстройка: {hostName} {hostVersionText}")
        Else
            detailText.AppendLine("Надстройка: не определена (запустите из Word/Excel/PowerPoint)")
        End If
        detailText.AppendLine($"Ядро: ShareRibbon {coreVersionText} ({configuration})")
        detailText.AppendLine($"Сборка: {BuildStamp()}")
        details.Text = detailText.ToString().TrimEnd()
        details.Font = New Font("Consolas", 8.5F)
        details.ForeColor = Color.FromArgb(100, 100, 100)
        details.AutoSize = True
        details.MaximumSize = New Size(560, 0)
        details.Margin = New Padding(0, 0, 0, 12)
        content.Controls.Add(details)

        Dim description As New Label()
        description.Text = "AI-ассистент для Word, Excel и PowerPoint." & vbCrLf & vbCrLf &
                           "Понимает команды на русском языке и помогает с анализом данных," & vbCrLf &
                           "форматированием, вычиткой и переводом документов." & vbCrLf & vbCrLf &
                           "Сборка предназначена для работы во внутреннем контуре и не" & vbCrLf &
                           "обращается к внешним сервисам и сайтам."
        description.Font = New Font("Segoe UI", 9)
        description.ForeColor = Color.FromArgb(80, 80, 80)
        description.AutoSize = True
        description.MaximumSize = New Size(560, 0)
        description.Margin = New Padding(0, 0, 0, 10)
        content.Controls.Add(description)

        Dim dataPath As New Label()
        dataPath.Text = "Каталог данных: Документы\" & ConfigSettings.OfficeAiAppDataFolder
        dataPath.Font = New Font("Segoe UI", 9)
        dataPath.ForeColor = Color.Gray
        dataPath.AutoSize = True
        dataPath.Margin = New Padding(0, 0, 0, 0)
        content.Controls.Add(dataPath)

        Dim footer As New FlowLayoutPanel()
        footer.Dock = DockStyle.Fill
        footer.FlowDirection = FlowDirection.RightToLeft
        footer.WrapContents = False
        footer.AutoSize = True
        footer.Padding = New Padding(0, 10, 0, 0)
        root.Controls.Add(footer, 0, 1)

        Dim btnClose As New Button()
        btnClose.Text = "Закрыть"
        btnClose.Size = New Size(96, 30)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.FlatAppearance.BorderColor = Color.FromArgb(74, 111, 165)
        btnClose.BackColor = Color.FromArgb(74, 111, 165)
        btnClose.ForeColor = Color.White
        btnClose.Font = New Font("Segoe UI", 9)
        AddHandler btnClose.Click, Sub(sender As Object, e As EventArgs) Me.Close()
        footer.Controls.Add(btnClose)
        Me.AcceptButton = btnClose
        Me.CancelButton = btnClose
    End Sub

    ''' <summary>
    ''' AssemblyInfo хранит версию сборки 2.0.0.0, а реальная версия продукта — FileVersion.
    ''' </summary>
    Private Shared Function GetAssemblyVersionText(assemblyItem As Assembly) As String
        Try
            If assemblyItem IsNot Nothing AndAlso Not String.IsNullOrEmpty(assemblyItem.Location) Then
                Dim info = FileVersionInfo.GetVersionInfo(assemblyItem.Location)
                If Not String.IsNullOrWhiteSpace(info.FileVersion) Then Return info.FileVersion.Trim()
            End If
            If assemblyItem IsNot Nothing Then Return assemblyItem.GetName().Version.ToString()
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function BuildStamp() As String
        Try
            Dim assemblyItem = GetType(AboutForm).Assembly
            Dim location = assemblyItem.Location
            If String.IsNullOrEmpty(location) Then Return "путь недоступен"
            Dim builtAt = If(IO.File.Exists(location), IO.File.GetLastWriteTime(location).ToString("dd.MM.yyyy HH:mm"), "")
            If builtAt.Length = 0 Then Return location
            Return $"{builtAt} · {location}"
        Catch
            Return "неизвестно"
        End Try
    End Function

    Private Shared Function IsDebugBuild() As Boolean
#If DEBUG Then
        Return True
#Else
        Return False
#End If
    End Function

    Private Shared Function FindHostAssembly() As Assembly
        Try
            For Each assemblyItem In AppDomain.CurrentDomain.GetAssemblies()
                Dim assemblyName = assemblyItem.GetName()
                If assemblyName.Name = "PowerPointAi" OrElse
                   assemblyName.Name = "WordAi" OrElse
                   assemblyName.Name = "ExcelAi" Then
                    Return assemblyItem
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function
End Class
