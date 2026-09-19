' ShareRibbon\Ribbon\BaseOfficeRibbon.vb
Imports System.IO
Imports System.Net
Imports System.Threading.Tasks
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json.Linq


Public MustInherit Class BaseOfficeRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    Private Async Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        ' GetApplication() 涉及 COM 对象访问，必须在 UI 线程（STA）调用，提前捕获
        Dim appInfo = GetApplication()

        ' 不在 Ribbon 加载时主动预加载程序集。
        ' AssemblyResolve 已在 ThisAddIn.Startup 注册，缺依赖时会按需解析。

        ' 将文件 I/O 密集的配置加载推迟到后台线程：
        '   ConfigManager.LoadConfig()          读取 API 配置 JSON
        '   ConfigPromptForm.LoadConfigStatic() 读取提示词模板 JSON（静态方法，不创建UI）
        ' 两者均为纯文件操作，不涉及 COM/UI，在后台线程执行安全，可避免阻塞 Ribbon 渲染
        Await Task.Run(Sub()
            ' 加载 API 配置
            Dim apiConfig As New ConfigManager()
            apiConfig.LoadConfig()
            ' 加载提示词配置（静态方法，不创建UI控件，避免阻塞启动）
            ConfigPromptForm.LoadConfigStatic(appInfo.Type.ToString())
        End Sub)

        ' Phase 2 不在 Office 启动期自动预热。
        ' WebView2/SQLite/ResourceExtractor 都比较重，即使放到后台也会和 Office 启动抢 CPU/磁盘。
        ' 保持 OnDemand：首次打开 AI 面板时由 EnsureCoreServicesLoaded/InitializeWebView2 触发。

        InitializeBaseRibbon()
    End Sub

    Protected Overridable Sub InitializeBaseRibbon()
        ' 基类初始化方法，子类可以重写
    End Sub

    ' 关于我按钮点击事件 - 显示带git链接的对话框
    Private Sub AboutButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles AboutButton.Click
        Using aboutForm As New AboutForm()
            aboutForm.ShowDialog()
        End Using
    End Sub

    ' 清理缓存配置按钮点击事件
    ' 使用递归删除子目录（包含 SQLite 数据库、日志、chat 历史等子目录），并对失败项做明细反馈
    Private Sub ClearCacheConfig_Click_1(sender As Object, e As RibbonControlEventArgs) Handles ClearCacheButton.Click
        ' 弹出确认框
        Dim result = MessageBox.Show("Будут безвозвратно удалены все настройки и история чатов в каталоге «Документы\" & ConfigSettings.OfficeAiAppDataFolder & "». Продолжить?", "Подтверждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result <> DialogResult.OK Then
            Return
        End If

        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & ConfigSettings.OfficeAiAppDataFolder
        If Not System.IO.Directory.Exists(appDataPath) Then
            MsgBox("Каталог кэша не существует!")
            Return
        End If

        Dim failedItems As New List(Of String)()
        Try
            ' 1. 删除根目录下的文件
            For Each filePath As String In System.IO.Directory.GetFiles(appDataPath)
                Try
                    System.IO.File.SetAttributes(filePath, IO.FileAttributes.Normal)
                    System.IO.File.Delete(filePath)
                Catch ex As Exception
                    failedItems.Add(System.IO.Path.GetFileName(filePath) & " (" & ex.Message & ")")
                End Try
            Next
            ' 2. 递归删除子目录（覆盖 SQLite/日志/chatHistory 等）
            For Each dirPath As String In System.IO.Directory.GetDirectories(appDataPath)
                Try
                    System.IO.Directory.Delete(dirPath, recursive:=True)
                Catch ex As Exception
                    failedItems.Add(System.IO.Path.GetFileName(dirPath) & "\ (" & ex.Message & ")")
                End Try
            Next

            If failedItems.Count = 0 Then
                MsgBox("Кэш конфигурации очищен. Перезапустите приложения Office!")
            Else
                Dim msg As String = "Очистка кэша завершена, но следующие элементы удалить не удалось (возможно, они заняты процессом; закройте все приложения Office и повторите):" & Environment.NewLine & String.Join(Environment.NewLine, failedItems)
                MsgBox(msg, vbExclamation)
            End If
        Catch ex As Exception
            MsgBox("Ошибка при очистке кэша конфигурации: " & ex.Message, vbCritical)
        End Try
    End Sub

    ' 点击Ribbon区的配置API按钮后触发
    Private Sub ConfigApiButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfigApiButton.Click
        ' 创建并显示配置 API 的对话框（Using 确保 Form 资源释放）
        Using configForm As New ConfigApiForm()
            configForm.ShowDialog()
        End Using
    End Sub
    Private Sub PromptConfigButton_Click(sender As Object, e As RibbonControlEventArgs) Handles PromptConfigButton.Click
        ' 创建并显示配置 API 的对话框（Using 确保 Form 资源释放）
        Using configForm As New ConfigPromptForm(GetApplication())
            configForm.ShowDialog()
        End Using
    End Sub

    ' Кнопка справки — открывает локальную HTML-справку, без внешних ссылок
    Private Sub StudyButton_Click(sender As Object, e As RibbonControlEventArgs) Handles StudyButton.Click
        Dim appInfo = GetApplication()
        Dim page As String

        Select Case appInfo.Type
            Case OfficeApplicationType.Word
                page = "word"
            Case OfficeApplicationType.Excel
                page = "excel"
            Case OfficeApplicationType.PowerPoint
                page = "ppt"
            Case Else
                page = "index"
        End Select

        Try
            Dim wwwRoot As String = ResourceExtractor.ExtractResources()
            If String.IsNullOrEmpty(wwwRoot) Then
                MessageBox.Show("Не удалось подготовить локальную справку: " & ResourceExtractor.LastError, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim helpPath As String = System.IO.Path.Combine(wwwRoot, "help", "ru", page & ".html")
            If System.IO.File.Exists(helpPath) Then
                System.Diagnostics.Process.Start(helpPath)
            Else
                MessageBox.Show("Файл справки не найден: " & helpPath, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Не удалось открыть справку: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ' AI聊天实现
    Protected MustOverride Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ChatButton.Click

    ' web爬虫实现
    Protected MustOverride Sub WebCaptureButton_Click(sender As Object, e As RibbonControlEventArgs) Handles WebCaptureButton.Click

    ' 聚光灯实现（跟随鼠标选中整行和整列并高亮）
    Protected MustOverride Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs) Handles SpotlightButton.Click

    ' 数据魔法分析实现
    Protected MustOverride Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DataAnalysisButton.Click
    Protected MustOverride Function GetApplication() As ApplicationInfo


    ' 新增：校对与排版按钮的抽象事件（由子类实现具体流程）
    Protected MustOverride Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ProofreadButton.Click
    Protected MustOverride Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ReformatButton.Click

    ' Deepseek按钮点击事件
    Protected MustOverride Sub DeepseekButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DeepseekButton.Click

    ' Doubao按钮点击事件
    Protected MustOverride Sub DoubaoButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DoubaoButton.Click

    ' 批量数据生成按钮点击事件
    Protected MustOverride Sub BatchDataGenButton_Click(sender As Object, e As RibbonControlEventArgs) Handles BatchDataGenButton.Click

    ' MCP按钮点击事件 - 三端共用同一对话框，提供基类默认实现，子类可按需重写
    Protected Overridable Sub MCPButton_Click(sender As Object, e As RibbonControlEventArgs) Handles MCPButton.Click
        Using mcpConfigForm As New MCPConfigForm()
            mcpConfigForm.ShowDialog()
        End Using
    End Sub

    ' 一键翻译按钮点击事件（抽象方法，由子类实现）
    Protected MustOverride Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs) Handles TranslateButton.Click

    ' AI续写按钮点击事件（抽象方法，由子类实现）
    Protected MustOverride Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ContinuationButton.Click

    ' 模板排版按钮点击事件（抽象方法，由子类实现）
    Protected MustOverride Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles TemplateFormatButton.Click
End Class
