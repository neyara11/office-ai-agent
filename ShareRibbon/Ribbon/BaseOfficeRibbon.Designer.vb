' ShareRibbon\Ribbon\BaseOfficeRibbon.Designer.vb
Partial Class BaseOfficeRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Sub New(ByVal factory As Microsoft.Office.Tools.Ribbon.RibbonFactory)
        MyBase.New(factory)
        InitializeComponent()
    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Protected Sub InitializeComponent()
        Me.TabAI = Me.Factory.CreateRibbonTab

        ' Group 1: 免费强化版 - Deepseek/Doubao
        Me.GroupDeepseek = Me.Factory.CreateRibbonGroup
        Me.DeepseekButton = Me.Factory.CreateRibbonButton()
        Me.DoubaoButton = Me.Factory.CreateRibbonButton()

        ' Group 2: 大模型配置 - 配置API/提示词
        Me.GroupConfig = Me.Factory.CreateRibbonGroup
        Me.ConfigApiButton = Me.Factory.CreateRibbonButton
        Me.PromptConfigButton = Me.Factory.CreateRibbonButton

        ' Group 3: AI对话 - Chat AI/AI翻译
        Me.GroupChat = Me.Factory.CreateRibbonGroup
        Me.ChatButton = Me.Factory.CreateRibbonButton
        Me.TranslateButton = Me.Factory.CreateRibbonButton

        ' Group 4: AI内容提效 - 续写/校对/排版/模板排版/接受补全 (Word/PPT专用)
        Me.GroupAIContent = Me.Factory.CreateRibbonGroup
        Me.ContinuationButton = Me.Factory.CreateRibbonButton
        Me.ProofreadButton = Me.Factory.CreateRibbonButton
        Me.ReformatButton = Me.Factory.CreateRibbonButton
        Me.TemplateFormatButton = Me.Factory.CreateRibbonButton

        ' Group 5: MCP连接
        Me.GroupMCP = Me.Factory.CreateRibbonGroup
        Me.MCPButton = Me.Factory.CreateRibbonButton()

        ' Group 6: 关于与设置
        Me.GroupAbout = Me.Factory.CreateRibbonGroup
        Me.AboutButton = Me.Factory.CreateRibbonButton
        Me.ClearCacheButton = Me.Factory.CreateRibbonButton

        ' Group 7: 帮助与学习
        Me.GroupHelp = Me.Factory.CreateRibbonGroup
        Me.StudyButton = Me.Factory.CreateRibbonButton

        ' Group: 工具箱 (Excel专用)
        Me.GroupTools = Me.Factory.CreateRibbonGroup
        Me.DataAnalysisButton = Me.Factory.CreateRibbonButton
        Me.WebCaptureButton = Me.Factory.CreateRibbonButton
        Me.SpotlightButton = Me.Factory.CreateRibbonButton
        Me.BatchDataGenButton = Me.Factory.CreateRibbonButton()

        Me.TabAI.SuspendLayout()
        Me.SuspendLayout()

        ' ========== TabAI 布局 ==========
        Me.TabAI.Groups.Add(Me.GroupDeepseek)   ' 1. 免费强化版
        Me.TabAI.Groups.Add(Me.GroupConfig)     ' 2. 大模型配置
        Me.TabAI.Groups.Add(Me.GroupChat)       ' 3. AI对话
        Me.TabAI.Groups.Add(Me.GroupAIContent)  ' 4. AI内容提效
        Me.TabAI.Groups.Add(Me.GroupTools)      ' 5. 工具箱
        Me.TabAI.Groups.Add(Me.GroupMCP)        ' 6. MCP连接
        Me.TabAI.Groups.Add(Me.GroupAbout)      ' 7. 关于与设置
        Me.TabAI.Groups.Add(Me.GroupHelp)       ' 8. 帮助与学习

        Me.TabAI.Label = "AI-ассистент"
        Me.TabAI.Name = "TabAI"

        ' ========== Group 1: 免费强化版 ==========
        Me.GroupDeepseek.Items.Add(Me.DeepseekButton)
        Me.GroupDeepseek.Items.Add(Me.DoubaoButton)
        Me.GroupDeepseek.Label = "Бесплатные сервисы"
        Me.GroupDeepseek.Name = "GroupDeepseek"

        Me.DeepseekButton.Label = "Deepseek"
        Me.DeepseekButton.Name = "DeepseekButton"
        Me.DeepseekButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DeepseekButton.ShowImage = True
        Me.DeepseekButton.ScreenTip = "Бесплатная расширенная версия"
        Me.DeepseekButton.SuperTip = "Добавляет возможности агента к обычному диалогу"

        Me.DoubaoButton.Label = "Doubao"
        Me.DoubaoButton.Name = "DoubaoButton"
        Me.DoubaoButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DoubaoButton.ShowImage = True
        Me.DoubaoButton.ScreenTip = "Умный помощник Doubao"
        Me.DoubaoButton.SuperTip = "Диалоговый помощник на базе Doubao с поддержкой выполнения кода"

        ' ========== Group 2: 大模型配置 ==========
        Me.GroupConfig.Items.Add(Me.ConfigApiButton)
        Me.GroupConfig.Items.Add(Me.PromptConfigButton)
        Me.GroupConfig.Label = "Настройка моделей"
        Me.GroupConfig.Name = "GroupConfig"

        Me.ConfigApiButton.Label = "Настройка API"
        Me.ConfigApiButton.Name = "ConfigApiButton"
        Me.ConfigApiButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ConfigApiButton.ShowImage = True
        Me.ConfigApiButton.ScreenTip = "Настроить API языковой модели"
        Me.ConfigApiButton.SuperTip = "Перед использованием AI-функций укажите apiKey"

        Me.PromptConfigButton.Label = "Промпты"
        Me.PromptConfigButton.Name = "PromptConfigButton"
        Me.PromptConfigButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.PromptConfigButton.ShowImage = True
        Me.PromptConfigButton.ScreenTip = "Настроить промпты"
        Me.PromptConfigButton.SuperTip = "Управление системными промптами AI-диалога"

        ' ========== Group 3: AI对话 ==========
        Me.GroupChat.Items.Add(Me.ChatButton)
        Me.GroupChat.Items.Add(Me.TranslateButton)
        Me.GroupChat.Label = "AI-диалог"
        Me.GroupChat.Name = "GroupChat"

        Me.ChatButton.Label = "AI-чат"
        Me.ChatButton.Name = "ChatButton"
        Me.ChatButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ChatButton.ShowImage = True
        Me.ChatButton.ScreenTip = "Помощник AI-диалога"
        Me.ChatButton.SuperTip = "Открыть панель AI-диалога: многоходовые беседы и выполнение кода"

        Me.TranslateButton.Label = "AI-перевод"
        Me.TranslateButton.Name = "TranslateButton"
        Me.TranslateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TranslateButton.ShowImage = True
        Me.TranslateButton.ScreenTip = "Перевести содержимое документа одним щелчком"
        Me.TranslateButton.SuperTip = "Полный текст, выделение, иммерсивный режим и другие режимы перевода"

        ' ========== Group 4: AI内容提效 ==========
        Me.GroupAIContent.Items.Add(Me.ContinuationButton)
        Me.GroupAIContent.Items.Add(Me.ProofreadButton)
        Me.GroupAIContent.Items.Add(Me.ReformatButton)
        Me.GroupAIContent.Items.Add(Me.TemplateFormatButton)
        Me.GroupAIContent.Label = "AI-обработка текста"
        Me.GroupAIContent.Name = "GroupAIContent"

        Me.ContinuationButton.Label = "Продолжить"
        Me.ContinuationButton.Name = "ContinuationButton"
        Me.ContinuationButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ContinuationButton.ShowImage = True
        Me.ContinuationButton.ScreenTip = "Умное продолжение текста"
        Me.ContinuationButton.SuperTip = "Продолжает текст с учётом контекста позиции курсора"

        Me.ProofreadButton.Label = "Вычитка"
        Me.ProofreadButton.Name = "ProofreadButton"
        Me.ProofreadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ProofreadButton.ShowImage = True
        Me.ProofreadButton.ScreenTip = "Языковая проверка выделения или всего текста"
        Me.ProofreadButton.SuperTip = "Исправляет грамматику и орфографию, возвращает разбираемый JSON правок"

        Me.ReformatButton.Label = "Форматирование"
        Me.ReformatButton.Name = "ReformatButton"
        Me.ReformatButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ReformatButton.ShowImage = True
        Me.ReformatButton.ScreenTip = "Умное форматирование: определяет тип документа и применяет стандарт"
        Me.ReformatButton.SuperTip = "Автоматически определяет тип документа (служебный/научный/отчёт и т. п.), подбирает стандарт (например, ГОСТ) и форматирует одним щелчком. Поддерживает правки в диалоге и клонирование по образцу."

        Me.TemplateFormatButton.Label = "Формат по шаблону"
        Me.TemplateFormatButton.Name = "TemplateFormatButton"
        Me.TemplateFormatButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TemplateFormatButton.ShowImage = True
        Me.TemplateFormatButton.Visible = False
        Me.TemplateFormatButton.ScreenTip = "Форматирование по шаблону"
        Me.TemplateFormatButton.SuperTip = "Выберите шаблон оформления: AI учтёт шрифт, размер и абзацы шаблона (объединено с умным форматированием, кнопка скрыта)"


        ' ========== Group 5: 工具箱 (Excel专用) ==========
        Me.GroupTools.Items.Add(Me.DataAnalysisButton)
        Me.GroupTools.Items.Add(Me.WebCaptureButton)
        Me.GroupTools.Items.Add(Me.SpotlightButton)
        Me.GroupTools.Items.Add(Me.BatchDataGenButton)
        Me.GroupTools.Label = "Инструменты"
        Me.GroupTools.Name = "GroupTools"

        Me.DataAnalysisButton.Label = "Анализ данных"
        Me.DataAnalysisButton.Name = "DataAnalysisButton"
        Me.DataAnalysisButton.ShowImage = True
        Me.DataAnalysisButton.ScreenTip = "Умный анализ данных"
        Me.DataAnalysisButton.SuperTip = "AI-помощь в анализе данных Excel"

        Me.WebCaptureButton.Label = "Захват страниц"
        Me.WebCaptureButton.Name = "WebCaptureButton"
        Me.WebCaptureButton.ShowImage = True
        Me.WebCaptureButton.SuperTip = "Открыть инструмент захвата веб-страниц"

        Me.SpotlightButton.Label = "Подсветка"
        Me.SpotlightButton.Name = "SpotlightButton"
        Me.SpotlightButton.ShowImage = True
        Me.SpotlightButton.SuperTip = "Подсвечивает строку и столбец выбранной ячейки"

        Me.BatchDataGenButton.Label = "Пакетная генерация"
        Me.BatchDataGenButton.Name = "BatchDataGenButton"
        Me.BatchDataGenButton.ShowImage = True
        Me.BatchDataGenButton.ScreenTip = "Настройка и генерация пакетных данных"
        Me.BatchDataGenButton.SuperTip = "Настройте поля и связи столбцов и сгенерируйте данные в книгу"

        ' ========== Group 6: MCP连接 ==========
        Me.GroupMCP.Items.Add(Me.MCPButton)
        Me.GroupMCP.Label = "Подключение MCP"
        Me.GroupMCP.Name = "GroupMCP"

        Me.MCPButton.Label = "MCP"
        Me.MCPButton.Name = "MCPButton"
        Me.MCPButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.MCPButton.ShowImage = True
        Me.MCPButton.ScreenTip = "Настройка MCP-серверов"
        Me.MCPButton.SuperTip = "Настройте MCP-сервер и вызывайте модель как клиент"

        ' ========== Group 7: 关于与设置 ==========
        Me.GroupAbout.Items.Add(Me.AboutButton)
        Me.GroupAbout.Items.Add(Me.ClearCacheButton)
        Me.GroupAbout.Label = "О программе и настройки"
        Me.GroupAbout.Name = "GroupAbout"

        Me.AboutButton.Label = "О программе"
        Me.AboutButton.Name = "AboutButton"
        Me.AboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.AboutButton.ShowImage = True
        Me.AboutButton.ScreenTip = "О плагине"
        Me.AboutButton.SuperTip = "Информация о плагине и ссылка на исходный код"

        Me.ClearCacheButton.Label = "Очистить кэш"
        Me.ClearCacheButton.Name = "ClearCacheButton"
        Me.ClearCacheButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ClearCacheButton.ShowImage = True
        Me.ClearCacheButton.ScreenTip = "Очистить кэш конфигурации"
        Me.ClearCacheButton.SuperTip = "Удаляет все настройки и историю"

        ' ========== Group 8: 帮助与学习 ==========
        Me.GroupHelp.Items.Add(Me.StudyButton)
        Me.GroupHelp.Label = "Справка и обучение"
        Me.GroupHelp.Name = "GroupHelp"

        Me.StudyButton.Label = "Документация"
        Me.StudyButton.Name = "StudyButton"
        Me.StudyButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.StudyButton.ShowImage = True
        Me.StudyButton.ScreenTip = "Открыть документацию"
        Me.StudyButton.SuperTip = "Открывает справку по всем функциям"

        ' BaseOfficeRibbon
        Me.Name = "BaseOfficeRibbon"
        Me.Tabs.Add(Me.TabAI)

        Me.TabAI.ResumeLayout(False)
        Me.TabAI.PerformLayout()
        Me.ResumeLayout(False)
    End Sub

    ' Tab
    Protected WithEvents TabAI As Microsoft.Office.Tools.Ribbon.RibbonTab

    ' Group 1: 免费强化版
    Protected WithEvents GroupDeepseek As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents DeepseekButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents DoubaoButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 2: 大模型配置
    Protected WithEvents GroupConfig As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ConfigApiButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents PromptConfigButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 3: AI对话
    Protected WithEvents GroupChat As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ChatButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents TranslateButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 4: AI内容提效
    Protected WithEvents GroupAIContent As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents ContinuationButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ProofreadButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ReformatButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents TemplateFormatButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 5: 工具箱
    Protected WithEvents GroupTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents DataAnalysisButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents WebCaptureButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents SpotlightButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents BatchDataGenButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 6: MCP连接
    Protected WithEvents GroupMCP As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents MCPButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 7: 关于与设置
    Protected WithEvents GroupAbout As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents AboutButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents ClearCacheButton As Microsoft.Office.Tools.Ribbon.RibbonButton

    ' Group 8: 帮助与学习
    Protected WithEvents GroupHelp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Protected WithEvents StudyButton As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class
