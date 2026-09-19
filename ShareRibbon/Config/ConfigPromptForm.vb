Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json

' 大模型提示词配置 - 现代化UI
Public Class ConfigPromptForm
    Inherits Form
    Private ReadOnly _applicationInfo As ApplicationInfo

    Public Shared Property ConfigPromptData As List(Of PromptConfigItem)

    ' UI控件
    Private tabControl As TabControl
    Private tabBasic As TabPage
    Private tabAgentProfile As TabPage
    Private tabAdvanced As TabPage
    Private tabQuickQuestions As TabPage

    ' 基础配置控件
    Private promptListBox As ListBox
    Private promptNameTextBox As TextBox
    Private promptContentTextBox As TextBox
    Private btnAdd As Button
    Private btnDelete As Button
    Private btnUse As Button

    ' 高级配置控件
    Private jsonSchemaTextBox As TextBox
    Private btnSaveSchema As Button
    Private btnResetSchema As Button

    ' 快捷问题控件
    Private quickQuestionsListBox As ListBox
    Private quickQuestionTextBox As TextBox
    Private btnAddQuestion As Button
    Private btnDeleteQuestion As Button
    Private btnSaveQuestions As Button
    Private btnResetQuestions As Button

    ' 快捷问题数据
    Private Shared _quickQuestions As List(Of String)

    ' 默认快捷问题（与前端predefinedPrompts保持一致）
    Private Shared ReadOnly DEFAULT_QUICK_QUESTIONS As String() = {
        "Помоги записать в столбец C сумму значений столбцов A и B",
        "Объедини таблицы Sheet1 и Sheet2 по именам",
        "Раздели данные Sheet1 на несколько файлов xlsx по именам",
        "Приведи в порядок форматирование выделенного содержимого Word",
        "Создай презентацию PPT с еженедельным отчётом на 3 страницы",
        "Нет нужного вопроса? Нажмите здесь, чтобы настроить список"
    }

    Private Const MAX_QUICK_QUESTIONS As Integer = 6

    ' 属性
    Public Property propmtName As String
    Public Property propmtContent As String

    ' 默认提示词（静态常量，供 LoadConfigStatic 和实例方法共用）
    Private Shared ReadOnly DEFAULT_PROMPTS As New Dictionary(Of String, String) From {
        {"Excel", "Ты эксперт по Excel, специализируешься на анализе данных и вычислении формул. Если запрос пользователя ясен, верни JSON-команду для выполнения; если запрос неясен — сначала задай уточняющий вопрос."},
        {"Word", "Ты эксперт по документам Word, специализируешься на редактировании документов, форматировании и создании содержимого. Если запрос пользователя ясен, верни JSON-команду для выполнения; если запрос неясен — сначала задай уточняющий вопрос."},
        {"PowerPoint", "Ты эксперт по презентациям PowerPoint, специализируешься на дизайне слайдов, анимации и создании содержимого. Если запрос пользователя ясен, верни JSON-команду для выполнения; если запрос неясен — сначала задай уточняющий вопрос."}
    }

    Public Sub New(applicationInfo As ApplicationInfo)
        _applicationInfo = applicationInfo
        LoadConfig()
        AddHandler Me.FormClosing, AddressOf HandleFormClosing
        InitializeUI()
    End Sub

    Private Sub HandleFormClosing(sender As Object, e As FormClosingEventArgs)
        If Me.Controls.Contains(GlobalStatusStrip.StatusStrip) Then
            Me.Controls.Remove(GlobalStatusStrip.StatusStrip)
        End If
    End Sub

    Private Sub InitializeUI()
        ' 窗体设置
        Me.Text = $"Настройка промптов - {_applicationInfo.Type}"
        Me.Size = New Size(600, 520)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 创建TabControl
        tabControl = New TabControl() With {
            .Location = New Point(10, 10),
            .Size = New Size(565, 420),
            .Font = New Font("Microsoft YaHei UI", 9)
        }

        ' 基础配置页
        tabBasic = New TabPage("Промпты чата")
        InitializeBasicTab()
        tabControl.TabPages.Add(tabBasic)

        ' Agent 提示词层配置说明
        tabAgentProfile = New TabPage("Слои промптов Agent")
        InitializeAgentProfileTab()
        tabControl.TabPages.Add(tabAgentProfile)

        ' 高级配置页
        tabAdvanced = New TabPage("Ограничения формата JSON")
        InitializeAdvancedTab()
        tabControl.TabPages.Add(tabAdvanced)

        ' 快捷问题配置页
        tabQuickQuestions = New TabPage("Быстрые вопросы")
        InitializeQuickQuestionsTab()
        tabControl.TabPages.Add(tabQuickQuestions)

        Me.Controls.Add(tabControl)

        ' 底部关闭按钮
        Dim btnClose As New Button() With {
            .Text = "Закрыть",
            .Location = New Point(490, 440),
            .Size = New Size(80, 30)
        }
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub InitializeBasicTab()
        ' 说明标签
        Dim lblDesc As New Label() With {
            .Text = "Промпт задаёт роль ИИ и делает ответы профессиональнее. Выберите промпт и нажмите «Использовать».",
            .Location = New Point(10, 10),
            .Size = New Size(530, 20),
            .ForeColor = Color.Gray
        }
        tabBasic.Controls.Add(lblDesc)

        ' 左侧：提示词列表
        Dim lblList As New Label() With {
            .Text = "Сохранённые промпты:",
            .Location = New Point(10, 35),
            .AutoSize = True
        }
        tabBasic.Controls.Add(lblList)

        promptListBox = New ListBox() With {
            .Location = New Point(10, 55),
            .Size = New Size(180, 200),
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        AddHandler promptListBox.SelectedIndexChanged, AddressOf PromptListBox_SelectedIndexChanged
        tabBasic.Controls.Add(promptListBox)

        ' 列表操作按钮
        btnUse = New Button() With {
            .Text = "Использовать выбранный",
            .Location = New Point(10, 260),
            .Size = New Size(85, 28),
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnUse.Click, AddressOf BtnUse_Click
        tabBasic.Controls.Add(btnUse)

        btnDelete = New Button() With {
            .Text = "Удалить",
            .Location = New Point(105, 260),
            .Size = New Size(85, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDelete.Click, AddressOf BtnDelete_Click
        tabBasic.Controls.Add(btnDelete)

        ' 右侧：编辑区域
        Dim lblName As New Label() With {
            .Text = "Имя промпта:",
            .Location = New Point(210, 35),
            .AutoSize = True
        }
        tabBasic.Controls.Add(lblName)

        promptNameTextBox = New TextBox() With {
            .Location = New Point(210, 55),
            .Size = New Size(330, 25)
        }
        tabBasic.Controls.Add(promptNameTextBox)

        Dim lblContent As New Label() With {
            .Text = "Содержимое промпта:",
            .Location = New Point(210, 85),
            .AutoSize = True
        }
        tabBasic.Controls.Add(lblContent)

        promptContentTextBox = New TextBox() With {
            .Location = New Point(210, 105),
            .Size = New Size(330, 150),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        tabBasic.Controls.Add(promptContentTextBox)

        ' 编辑操作按钮
        btnAdd = New Button() With {
            .Text = "Добавить/Сохранить",
            .Location = New Point(210, 260),
            .Size = New Size(100, 28),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAdd.Click, AddressOf BtnAdd_Click
        tabBasic.Controls.Add(btnAdd)

        Dim btnClear As New Button() With {
            .Text = "Очистить поля",
            .Location = New Point(320, 260),
            .Size = New Size(80, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnClear.Click, Sub(s, e)
                                       promptNameTextBox.Clear()
                                       promptContentTextBox.Clear()
                                       promptListBox.ClearSelected()
                                   End Sub
        tabBasic.Controls.Add(btnClear)

        ' 当前使用的提示词显示
        Dim lblCurrent As New Label() With {
            .Text = "Текущий:",
            .Location = New Point(10, 300),
            .AutoSize = True,
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        }
        tabBasic.Controls.Add(lblCurrent)

        Dim lblCurrentValue As New Label() With {
            .Name = "lblCurrentValue",
            .Text = If(String.IsNullOrEmpty(ConfigSettings.propmtName), "(не задан)", ConfigSettings.propmtName),
            .Location = New Point(80, 300),
            .Size = New Size(460, 20),
            .ForeColor = Color.FromArgb(70, 130, 180)
        }
        tabBasic.Controls.Add(lblCurrentValue)

        ' 加载数据到列表
        RefreshPromptList()
    End Sub

    Private Sub InitializeAgentProfileTab()
        Dim externalPromptDir = Agent.PromptProfileService.GetExternalPromptDirectory()
        Dim appScenario = GetCurrentPromptScenario()

        Dim lblDesc As New Label() With {
            .Text = "Agent собирает промпт по фиксированным слоям: системный протокол -> контекст Office -> инструменты/Skill -> предпочтения пользователя -> память. Предпочтения пользователя не могут переопределять протокол инструментов.",
            .Location = New Point(10, 10),
            .Size = New Size(530, 36),
            .ForeColor = Color.DimGray
        }
        tabAgentProfile.Controls.Add(lblDesc)

        Dim lblSelected As New Label() With {
            .Text = "Текущий промпт чата:",
            .Location = New Point(10, 55),
            .AutoSize = True,
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        }
        tabAgentProfile.Controls.Add(lblSelected)

        Dim lblSelectedValue As New Label() With {
            .Text = If(String.IsNullOrEmpty(ConfigSettings.propmtName), "(не задан)", ConfigSettings.propmtName),
            .Location = New Point(120, 55),
            .Size = New Size(420, 20),
            .ForeColor = Color.FromArgb(70, 130, 180)
        }
        tabAgentProfile.Controls.Add(lblSelectedValue)

        Dim lblFolder As New Label() With {
            .Text = "Каталог внешних промптов:",
            .Location = New Point(10, 90),
            .AutoSize = True,
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        }
        tabAgentProfile.Controls.Add(lblFolder)

        Dim txtFolder As New TextBox() With {
            .Location = New Point(10, 112),
            .Size = New Size(430, 25),
            .Text = externalPromptDir,
            .ReadOnly = True
        }
        tabAgentProfile.Controls.Add(txtFolder)

        Dim btnOpenFolder As New Button() With {
            .Text = "Открыть каталог",
            .Location = New Point(450, 110),
            .Size = New Size(90, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnOpenFolder.Click, Sub(s, e) OpenExternalPromptDirectory(externalPromptDir)
        tabAgentProfile.Controls.Add(btnOpenFolder)

        Dim lblFiles As New Label() With {
            .Text = "Поддерживаемые файлы:" & Environment.NewLine &
                    $"common.md / common.txt / common.json: общие для всех приложений Office" & Environment.NewLine &
                    $"{appScenario}.md / {appScenario}.txt / {appScenario}.json: только для текущего приложения" & Environment.NewLine &
                    $"common\*.md|*.txt|*.json и {appScenario}\*.md|*.txt|*.json: можно разбить на несколько тематических файлов",
            .Location = New Point(10, 150),
            .Size = New Size(530, 78),
            .ForeColor = Color.DimGray
        }
        tabAgentProfile.Controls.Add(lblFiles)

        Dim lblJson As New Label() With {
            .Text = "Необязательные поля JSON-файла: enabled, application/appType, content/prompt. Пример:" & Environment.NewLine &
                    "{""enabled"":true,""application"":""" & appScenario & """,""content"":""Отвечай кратко, по делу, задавай минимум уточняющих вопросов.""}",
            .Location = New Point(10, 235),
            .Size = New Size(530, 48),
            .ForeColor = Color.DimGray
        }
        tabAgentProfile.Controls.Add(lblJson)

        Dim lblPriority As New Label() With {
            .Text = "Приоритет: личный стиль на этой странице, внешние промпты и профиль пользователя влияют только на стиль изложения, бизнес-предпочтения и предметный контекст; они не переопределяют Harness, Agent Loop, схему инструментов, границы приложения и протокол выполнения.",
            .Location = New Point(10, 295),
            .Size = New Size(530, 48),
            .ForeColor = Color.FromArgb(120, 80, 20),
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Italic)
        }
        tabAgentProfile.Controls.Add(lblPriority)

        Dim btnCreateExample As New Button() With {
            .Text = "Создать пример common.md",
            .Location = New Point(10, 350),
            .Size = New Size(140, 30),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnCreateExample.Click, Sub(s, e) CreateExternalPromptExample(externalPromptDir)
        tabAgentProfile.Controls.Add(btnCreateExample)
    End Sub

    Private Function GetCurrentPromptScenario() As String
        Select Case _applicationInfo.Type
            Case OfficeApplicationType.Word
                Return "word"
            Case OfficeApplicationType.PowerPoint
                Return "ppt"
            Case Else
                Return "excel"
        End Select
    End Function

    Private Sub OpenExternalPromptDirectory(folderPath As String)
        Try
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If
            System.Diagnostics.Process.Start(folderPath)
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"Не удалось открыть каталог: {ex.Message}")
        End Try
    End Sub

    Private Sub CreateExternalPromptExample(folderPath As String)
        Try
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            Dim examplePath = Path.Combine(folderPath, "common.md")
            If File.Exists(examplePath) Then
                If MessageBox.Show("Файл common.md уже существует. Перезаписать содержимое примера?", "Подтверждение перезаписи", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                    Return
                End If
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine("# Пример личных предпочтений Office Agent")
            sb.AppendLine()
            sb.AppendLine("- В ответах сначала давай практический вывод, меньше общих объяснений.")
            sb.AppendLine("- Если пользователь явно просит операцию Office, сначала формируй план, предпросмотр и выполнение, не переспрашивай наблюдаемую информацию.")
            sb.AppendLine("- Тон общения должен оставаться профессиональным и кратким; при необходимости поясняй сделанные допущения.")
            sb.AppendLine("- Эти предпочтения не могут переопределять схему инструментов, границы приложения, Agent Loop и ограничения безопасности.")

            File.WriteAllText(examplePath, sb.ToString(), Encoding.UTF8)
            GlobalStatusStrip.ShowInfo("Пример внешнего промпта common.md создан")
            OpenExternalPromptDirectory(folderPath)
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"Не удалось создать пример: {ex.Message}")
        End Try
    End Sub

    Private Sub InitializeAdvancedTab()
        ' 说明标签
        Dim lblDesc As New Label() With {
            .Text = $"Ограничения формата JSON задают формат команд, возвращаемых ИИ, для их корректного разбора и выполнения. Текущее приложение: {_applicationInfo.Type}",
            .Location = New Point(10, 10),
            .Size = New Size(530, 20),
            .ForeColor = Color.Gray
        }
        tabAdvanced.Controls.Add(lblDesc)

        Dim lblWarning As New Label() With {
            .Text = "⚠ Изменение этого содержимого может привести к сбою выполнения команд. Будьте осторожны!",
            .Location = New Point(10, 32),
            .Size = New Size(530, 20),
            .ForeColor = Color.OrangeRed,
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        }
        tabAdvanced.Controls.Add(lblWarning)

        ' JSON Schema 编辑框
        jsonSchemaTextBox = New TextBox() With {
            .Location = New Point(10, 55),
            .Size = New Size(530, 270),
            .Multiline = True,
            .ScrollBars = ScrollBars.Both,
            .Font = New Font("Consolas", 9),
            .WordWrap = False
        }
        tabAdvanced.Controls.Add(jsonSchemaTextBox)

        ' 加载当前的 JSON Schema
        LoadJsonSchema()

        ' 操作按钮
        btnSaveSchema = New Button() With {
            .Text = "Сохранить изменения",
            .Location = New Point(10, 335),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSaveSchema.Click, AddressOf BtnSaveSchema_Click
        tabAdvanced.Controls.Add(btnSaveSchema)

        btnResetSchema = New Button() With {
            .Text = "Сбросить к значениям по умолчанию",
            .Location = New Point(120, 335),
            .Size = New Size(100, 30),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnResetSchema.Click, AddressOf BtnResetSchema_Click
        tabAdvanced.Controls.Add(btnResetSchema)
    End Sub

    Private Sub LoadJsonSchema()
        Try
            Dim schema = PromptManager.Instance.GetJsonSchemaConstraint(_applicationInfo.Type.ToString())
            jsonSchemaTextBox.Text = If(String.IsNullOrEmpty(schema), "(не настроено)", schema)
        Catch ex As Exception
            jsonSchemaTextBox.Text = $"(Ошибка загрузки: {ex.Message})"
        End Try
    End Sub

    Private Sub BtnSaveSchema_Click(sender As Object, e As EventArgs)
        Try
            ' 保存到 PromptManager
            PromptManager.Instance.UpdateJsonSchemaConstraint(_applicationInfo.Type.ToString(), jsonSchemaTextBox.Text)
            PromptManager.Instance.SavePromptConfiguration()
            GlobalStatusStrip.ShowInfo("Ограничения формата JSON сохранены!")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"Не удалось сохранить: {ex.Message}")
        End Try
    End Sub

    Private Sub BtnResetSchema_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("Сбросить ограничения формата JSON к значениям по умолчанию?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                PromptManager.Instance.ResetJsonSchemaConstraint(_applicationInfo.Type.ToString())
                PromptManager.Instance.SavePromptConfiguration()
                LoadJsonSchema()
                GlobalStatusStrip.ShowInfo("Настройки по умолчанию восстановлены!")
            Catch ex As Exception
                GlobalStatusStrip.ShowWarning($"Не удалось восстановить: {ex.Message}")
            End Try
        End If
    End Sub

    ' ============ 快捷问题Tab初始化 ============
    Private Sub InitializeQuickQuestionsTab()
        ' 加载快捷问题数据
        LoadQuickQuestions()

        ' 说明标签
        Dim lblDesc As New Label() With {
            .Text = "Быстрые вопросы появляются при нажатии # в поле ввода для быстрого выбора частых запросов. Можно хранить не более 6.",
            .Location = New Point(10, 10),
            .Size = New Size(530, 20),
            .ForeColor = Color.Gray
        }
        tabQuickQuestions.Controls.Add(lblDesc)

        ' 左侧：快捷问题列表
        Dim lblList As New Label() With {
            .Text = "Сохранённые быстрые вопросы:",
            .Location = New Point(10, 35),
            .AutoSize = True
        }
        tabQuickQuestions.Controls.Add(lblList)

        quickQuestionsListBox = New ListBox() With {
            .Location = New Point(10, 55),
            .Size = New Size(530, 150),
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        AddHandler quickQuestionsListBox.SelectedIndexChanged, AddressOf QuickQuestionsListBox_SelectedIndexChanged
        tabQuickQuestions.Controls.Add(quickQuestionsListBox)

        ' 编辑区域
        Dim lblEdit As New Label() With {
            .Text = "Редактировать текст вопроса:",
            .Location = New Point(10, 215),
            .AutoSize = True
        }
        tabQuickQuestions.Controls.Add(lblEdit)

        quickQuestionTextBox = New TextBox() With {
            .Location = New Point(10, 235),
            .Size = New Size(530, 25),
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        tabQuickQuestions.Controls.Add(quickQuestionTextBox)

        ' 操作按钮行
        btnAddQuestion = New Button() With {
            .Text = "Добавить/Обновить",
            .Location = New Point(10, 270),
            .Size = New Size(90, 28),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAddQuestion.Click, AddressOf BtnAddQuestion_Click
        tabQuickQuestions.Controls.Add(btnAddQuestion)

        btnDeleteQuestion = New Button() With {
            .Text = "Удалить выбранное",
            .Location = New Point(110, 270),
            .Size = New Size(90, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDeleteQuestion.Click, AddressOf BtnDeleteQuestion_Click
        tabQuickQuestions.Controls.Add(btnDeleteQuestion)

        btnSaveQuestions = New Button() With {
            .Text = "Сохранить конфигурацию",
            .Location = New Point(350, 270),
            .Size = New Size(90, 28),
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSaveQuestions.Click, AddressOf BtnSaveQuestions_Click
        tabQuickQuestions.Controls.Add(btnSaveQuestions)

        btnResetQuestions = New Button() With {
            .Text = "Сбросить к значениям по умолчанию",
            .Location = New Point(450, 270),
            .Size = New Size(90, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnResetQuestions.Click, AddressOf BtnResetQuestions_Click
        tabQuickQuestions.Controls.Add(btnResetQuestions)

        ' 提示信息
        Dim lblTip As New Label() With {
            .Text = "💡 Подсказка: после сохранения нажмите # в поле ввода чата, чтобы увидеть актуальный список быстрых вопросов.",
            .Location = New Point(10, 310),
            .Size = New Size(530, 20),
            .ForeColor = Color.FromArgb(70, 130, 180),
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Italic)
        }
        tabQuickQuestions.Controls.Add(lblTip)

        ' 刷新列表
        RefreshQuickQuestionsList()
    End Sub

    Private Sub QuickQuestionsListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If quickQuestionsListBox.SelectedItem IsNot Nothing Then
            quickQuestionTextBox.Text = quickQuestionsListBox.SelectedItem.ToString()
        End If
    End Sub

    Private Sub BtnAddQuestion_Click(sender As Object, e As EventArgs)
        Dim question = quickQuestionTextBox.Text.Trim()
        If String.IsNullOrEmpty(question) Then
            GlobalStatusStrip.ShowWarning("Введите текст быстрого вопроса!")
            Return
        End If

        If quickQuestionsListBox.SelectedIndex >= 0 Then
            ' 更新选中项
            _quickQuestions(quickQuestionsListBox.SelectedIndex) = question
            GlobalStatusStrip.ShowInfo("Быстрый вопрос обновлён!")
        Else
            ' 新增
            If _quickQuestions.Count >= MAX_QUICK_QUESTIONS Then
                GlobalStatusStrip.ShowWarning($"Можно хранить не более {MAX_QUICK_QUESTIONS} быстрых вопросов!")
                Return
            End If
            _quickQuestions.Add(question)
            GlobalStatusStrip.ShowInfo("Быстрый вопрос добавлен!")
        End If

        RefreshQuickQuestionsList()
        quickQuestionTextBox.Clear()
        quickQuestionsListBox.ClearSelected()
    End Sub

    Private Sub BtnDeleteQuestion_Click(sender As Object, e As EventArgs)
        If quickQuestionsListBox.SelectedIndex < 0 Then
            GlobalStatusStrip.ShowWarning("Сначала выберите быстрый вопрос для удаления!")
            Return
        End If

        Dim selectedIndex = quickQuestionsListBox.SelectedIndex
        If MessageBox.Show($"Удалить «{_quickQuestions(selectedIndex)}»?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            _quickQuestions.RemoveAt(selectedIndex)
            RefreshQuickQuestionsList()
            quickQuestionTextBox.Clear()
            GlobalStatusStrip.ShowInfo("Удалено!")
        End If
    End Sub

    Private Sub BtnSaveQuestions_Click(sender As Object, e As EventArgs)
        Try
            SaveQuickQuestions()
            GlobalStatusStrip.ShowInfo("Конфигурация быстрых вопросов сохранена! Изменения вступят в силу после повторного открытия панели чата.")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"Не удалось сохранить: {ex.Message}")
        End Try
    End Sub

    Private Sub BtnResetQuestions_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("Восстановить быстрые вопросы по умолчанию?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            _quickQuestions = DEFAULT_QUICK_QUESTIONS.ToList()
            RefreshQuickQuestionsList()
            SaveQuickQuestions()
            GlobalStatusStrip.ShowInfo("Быстрые вопросы восстановлены по умолчанию!")
        End If
    End Sub

    Private Sub RefreshQuickQuestionsList()
        quickQuestionsListBox.Items.Clear()
        For Each q In _quickQuestions
            quickQuestionsListBox.Items.Add(q)
        Next
    End Sub

    ' ============ 快捷问题数据持久化 ============
    Private Sub LoadQuickQuestions()
        _quickQuestions = New List(Of String)()
        Dim filePath = GetQuickQuestionsFilePath()

        If File.Exists(filePath) Then
            Try
                Dim json = File.ReadAllText(filePath)
                _quickQuestions = JsonConvert.DeserializeObject(Of List(Of String))(json)
            Catch ex As Exception
                Debug.WriteLine($"加载快捷问题失败: {ex.Message}")
            End Try
        End If

        ' 如果为空，使用默认值
        If _quickQuestions Is Nothing OrElse _quickQuestions.Count = 0 Then
            _quickQuestions = DEFAULT_QUICK_QUESTIONS.ToList()
        End If
    End Sub

    Private Sub SaveQuickQuestions()
        Dim filePath = GetQuickQuestionsFilePath()
        Dim dir = Path.GetDirectoryName(filePath)
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If

        Dim json = JsonConvert.SerializeObject(_quickQuestions, Formatting.Indented)
        File.WriteAllText(filePath, json)
    End Sub

    Private Function GetQuickQuestionsFilePath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "OfficeAiAppData",
            "quick_questions_config.json")
    End Function

    ''' <summary>
    ''' 获取当前快捷问题列表（供HTML页面调用）
    ''' </summary>
    Public Shared Function GetQuickQuestionsList() As List(Of String)
        If _quickQuestions IsNot Nothing AndAlso _quickQuestions.Count > 0 Then
            Return _quickQuestions
        End If

        ' 尝试从文件加载
        Dim filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "OfficeAiAppData",
            "quick_questions_config.json")

        If File.Exists(filePath) Then
            Try
                Dim json = File.ReadAllText(filePath)
                Dim questions = JsonConvert.DeserializeObject(Of List(Of String))(json)
                If questions IsNot Nothing AndAlso questions.Count > 0 Then
                    Return questions
                End If
            Catch ex As Exception
                Debug.WriteLine($"读取快捷问题失败: {ex.Message}")
            End Try
        End If

        ' 返回默认值
        Return DEFAULT_QUICK_QUESTIONS.ToList()
    End Function

    Private Sub RefreshPromptList()
        promptListBox.Items.Clear()
        For Each item In ConfigPromptData
            promptListBox.Items.Add(item)
        Next

        ' 选中当前使用的
        For i As Integer = 0 To promptListBox.Items.Count - 1
            If CType(promptListBox.Items(i), PromptConfigItem).selected Then
                promptListBox.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub PromptListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If promptListBox.SelectedItem IsNot Nothing Then
            Dim item = CType(promptListBox.SelectedItem, PromptConfigItem)
            promptNameTextBox.Text = item.name
            promptContentTextBox.Text = item.content
        End If
    End Sub

    Private Sub BtnUse_Click(sender As Object, e As EventArgs)
        If promptListBox.SelectedItem Is Nothing Then
            GlobalStatusStrip.ShowWarning("Сначала выберите промпт!")
            Return
        End If

        Dim selectedItem = CType(promptListBox.SelectedItem, PromptConfigItem)

        ' 更新选中状态
        For Each item In ConfigPromptData
            item.selected = (item.name = selectedItem.name)
        Next

        ' 保存并更新全局配置
        SaveConfig()
        ConfigSettings.propmtName = selectedItem.name
        ConfigSettings.propmtContent = selectedItem.content

        ' 更新显示
        Dim lblCurrentValue = tabBasic.Controls.Find("lblCurrentValue", False).FirstOrDefault()
        If lblCurrentValue IsNot Nothing Then
            lblCurrentValue.Text = selectedItem.name
        End If

        GlobalStatusStrip.ShowInfo($"Промпт включён: {selectedItem.name}")
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        Dim name = promptNameTextBox.Text.Trim()
        Dim content = promptContentTextBox.Text.Trim()

        If String.IsNullOrEmpty(name) Then
            GlobalStatusStrip.ShowWarning("Введите имя промпта!")
            Return
        End If

        If String.IsNullOrEmpty(content) Then
            GlobalStatusStrip.ShowWarning("Введите содержимое промпта!")
            Return
        End If

        ' 检查是否存在
        Dim existingItem = ConfigPromptData.FirstOrDefault(Function(item) item.name = name)
        If existingItem IsNot Nothing Then
            ' 更新
            existingItem.content = content
            GlobalStatusStrip.ShowInfo($"Промпт обновлён: {name}")
        Else
            ' 新增
            ConfigPromptData.Add(New PromptConfigItem() With {
                .name = name,
                .content = content,
                .selected = False
            })
            GlobalStatusStrip.ShowInfo($"Промпт добавлен: {name}")
        End If

        SaveConfig()
        RefreshPromptList()
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs)
        If promptListBox.SelectedItem Is Nothing Then
            GlobalStatusStrip.ShowWarning("Сначала выберите промпт для удаления!")
            Return
        End If

        Dim selectedItem = CType(promptListBox.SelectedItem, PromptConfigItem)

        If selectedItem.selected Then
            GlobalStatusStrip.ShowWarning("Нельзя удалить используемый промпт!")
            Return
        End If

        If MessageBox.Show($"Удалить «{selectedItem.name}»?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ConfigPromptData.Remove(selectedItem)
            SaveConfig()
            RefreshPromptList()
            promptNameTextBox.Clear()
            promptContentTextBox.Clear()
            GlobalStatusStrip.ShowInfo("Удалено!")
        End If
    End Sub

    ''' <summary>
    ''' 共享配置路径（用于静态加载）
    ''' </summary>
    Private Shared Function GetConfigFilePath(appType As String) As String
        Dim appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), ConfigSettings.OfficeAiAppDataFolder)
        Return Path.Combine(appDataFolder, $"prompt_config_{appType.ToLower()}.json")
    End Function

    ''' <summary>
    ''' 加载提示词配置（静态版本，不创建UI，用于启动时后台加载）
    ''' </summary>
    Public Shared Sub LoadConfigStatic(appType As String)
        Dim configFilePath = GetConfigFilePath(appType)
        ConfigPromptData = New List(Of PromptConfigItem)()

        If File.Exists(configFilePath) Then
            Try
                Dim json As String = File.ReadAllText(configFilePath)
                ConfigPromptData = JsonConvert.DeserializeObject(Of List(Of PromptConfigItem))(json)
            Catch ex As Exception
                Debug.WriteLine($"[ConfigPromptForm] 静态加载提示词配置失败: {ex.Message}")
            End Try
        End If

        ' 如果为空，添加默认配置
        If ConfigPromptData Is Nothing OrElse ConfigPromptData.Count = 0 Then
            ConfigPromptData = New List(Of PromptConfigItem)()
            Dim defaultPrompt = GetDefaultPromptStatic(appType)
            ConfigPromptData.Add(defaultPrompt)
            Try
                Dim dir = Path.GetDirectoryName(configFilePath)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                Dim json As String = JsonConvert.SerializeObject(ConfigPromptData, Formatting.Indented)
                File.WriteAllText(configFilePath, json)
            Catch ex2 As Exception
                Debug.WriteLine($"[ConfigPromptForm] 保存默认提示词配置失败: {ex2.Message}")
            End Try
        End If

        ' 初始化全局配置
        For Each item In ConfigPromptData
            If item.selected Then
                ConfigSettings.propmtName = item.name
                ConfigSettings.propmtContent = item.content
                Exit For
            End If
        Next
    End Sub

    ''' <summary>
    ''' 获取默认提示词（静态版本）
    ''' </summary>
    Private Shared Function GetDefaultPromptStatic(appType As String) As PromptConfigItem
        Dim content = If(DEFAULT_PROMPTS.ContainsKey(appType), DEFAULT_PROMPTS(appType), "Ты помощник по работе с Office.")
        Return New PromptConfigItem() With {
            .name = $"{appType} - помощник",
            .content = content,
            .selected = True
        }
    End Function

    ''' <summary>
    ''' 加载提示词配置（实例版本，创建UI后调用）
    ''' </summary>
    Public Sub LoadConfig()
        ConfigPromptData = New List(Of PromptConfigItem)()

        If File.Exists(configFilePath) Then
            Try
                Dim json As String = File.ReadAllText(configFilePath)
                ConfigPromptData = JsonConvert.DeserializeObject(Of List(Of PromptConfigItem))(json)
            Catch ex As Exception
                Debug.WriteLine($"加载提示词配置失败: {ex.Message}")
            End Try
        End If

        ' 如果为空，添加默认配置
        If ConfigPromptData Is Nothing OrElse ConfigPromptData.Count = 0 Then
            ConfigPromptData = New List(Of PromptConfigItem)()
            Dim defaultPrompt = GetDefaultPrompt()
            ConfigPromptData.Add(defaultPrompt)
            SaveConfig()
        End If

        ' 初始化全局配置
        For Each item In ConfigPromptData
            If item.selected Then
                ConfigSettings.propmtName = item.name
                ConfigSettings.propmtContent = item.content
                Exit For
            End If
        Next
    End Sub

    Private Function GetDefaultPrompt() As PromptConfigItem
        Dim appType = _applicationInfo.Type.ToString()
        Dim content = If(DEFAULT_PROMPTS.ContainsKey(appType), DEFAULT_PROMPTS(appType), "Ты помощник по работе с Office.")

        Return New PromptConfigItem() With {
            .name = $"{appType} - помощник",
            .content = content,
            .selected = True
        }
    End Function

    Public Sub SaveConfig()
        Try
            Dim dir = Path.GetDirectoryName(configFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json As String = JsonConvert.SerializeObject(ConfigPromptData, Formatting.Indented)
            File.WriteAllText(configFilePath, json)
        Catch ex As Exception
            Debug.WriteLine($"保存提示词配置失败: {ex.Message}")
        End Try
    End Sub

    Private ReadOnly Property configFilePath As String
        Get
            Return _applicationInfo.GetPromptConfigFilePath()
        End Get
    End Property



    ' 提示词配置项
    Public Class PromptConfigItem
        Public Property name As String
        Public Property content As String
        Public Property selected As Boolean
        Public Overrides Function ToString() As String
            Return name
        End Function
    End Class
End Class
