Imports System.IO
Imports System.Diagnostics

Public Class ResourceExtractor
    Private Shared _lastError As String = String.Empty

    ''' <summary>
    ''' 资源版本号 — 更新此值可强制刷新所有前端资源文件
    ''' </summary>
    Private Shared _resourceVersion As String = "2026.09.20.2"

    ''' <summary>
    ''' Манифест извлечённых файлов (имя|размер). Позволяет подхватывать правки JS/CSS
    ''' без ручного поднятия _resourceVersion: если размер файла разошёлся — ресурсы переизвлекаются.
    ''' </summary>
    Private Const ManifestFileName As String = ".manifest"

    ''' <summary>
    ''' 获取最后一次错误信息
    ''' </summary>
    Public Shared ReadOnly Property LastError As String
        Get
            Return _lastError
        End Get
    End Property

    Public Shared Function ExtractResources() As String
        _lastError = String.Empty

#If DEBUG Then
        DebugListResources()
#End If

        Try
            ' 获取用户本地应用程序数据目录
            Dim appDataPath As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAI",
                "www"
            )

            ' 检查版本标记文件，匹配则跳过提取
            Dim versionFile = Path.Combine(appDataPath, ".version")
            Dim manifestFile = Path.Combine(appDataPath, ManifestFileName)
            If File.Exists(versionFile) AndAlso File.Exists(manifestFile) Then
                Try
                    Dim savedVersion = File.ReadAllText(versionFile).Trim()
                    If savedVersion = _resourceVersion AndAlso ManifestMatches(appDataPath, manifestFile) Then
                        Debug.WriteLine("[ResourceExtractor] 资源版本匹配，跳过提取")
                        Return appDataPath
                    End If
                Catch
                    ' 读取失败，继续执行提取
                End Try
            End If

            ' 确保目录存在
            Directory.CreateDirectory(appDataPath)
            Directory.CreateDirectory(Path.Combine(appDataPath, "css"))
            Directory.CreateDirectory(Path.Combine(appDataPath, "js"))

            ' 获取资源管理器 - 使用当前类型的程序集，确保在子插件中也能正确获取 ShareRibbon 资源
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", GetType(ResourceExtractor).Assembly)

            ' 第三方库资源文件映射
            Dim libraryResources As New Dictionary(Of String, String) From {
                {"marked_min", "marked.min.js"},
                {"highlight_min", "highlight.min.js"},
                {"vbscript_min", "vbscript.min.js"},
                {"github_min", "github.min.css"}
            }

            ' 释放第三方库资源
            Dim extractErrors As New List(Of String)()
            For Each kvp In libraryResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, If(kvp.Value.EndsWith(".js"), "js", "css")), rm:=rm)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next

            ' 自定义CSS资源文件映射
            Dim cssResources As New Dictionary(Of String, String) From {
                {"styles", "styles.css"}
            }

            ' 释放CSS资源
            For Each kvp In cssResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "css"), rm:=rm)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next

            ' 自定义JS资源文件映射
            Dim jsResources As New Dictionary(Of String, String) From {
                {"office_ai_bridge", "office-ai-bridge.js"},
                {"utils", "utils.js"},
                {"core", "core.js"},
                {"markdown_renderer", "markdown-renderer.js"},
                {"chat_manager", "chat-manager.js"},
                {"message_sender", "message-sender.js"},
                {"code_handler", "code-handler.js"},
                {"settings_manager", "settings-manager.js"},
                {"mcp_manager", "mcp-manager.js"},
                {"revision_manager", "revision-manager.js"},
                {"history_manager", "history-manager.js"},
                {"autocomplete", "autocomplete.js"},
                {"agent_protocol", "agent-protocol.js"},
                {"agent_card", "agent-card.js"},
                {"model_switcher", "model-switcher.js"},
                {"reformat_template", "reformat-template.js"},
                {"reformat_chat", "reformat-chat.js"},
                {"proofread_ui", "proofread-ui.js"}
            }

            ' 释放JS资源
            For Each kvp In jsResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "js"), rm:=rm)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next
            
            ' Локальная справка (RU) — работает без интернета
            Directory.CreateDirectory(Path.Combine(appDataPath, "help", "ru"))
            Dim helpResources As New Dictionary(Of String, String) From {
                {"help_ru_index", "index.html"},
                {"help_ru_word", "word.html"},
                {"help_ru_excel", "excel.html"},
                {"help_ru_ppt", "ppt.html"},
                {"help_ru_faq", "faq.html"}
            }
            For Each kvp In helpResources
                Dim errMsg As String = ExtractResourceToFileFromManager(kvp.Key, kvp.Value, targetDir:=Path.Combine(appDataPath, "help", "ru"), rm:=rm)
                If Not String.IsNullOrEmpty(errMsg) Then
                    extractErrors.Add(errMsg)
                End If
            Next

            ' 如果有提取错误，记录但仍然返回路径（部分资源可能已成功）
            If extractErrors.Count > 0 Then
                _lastError = String.Join(Environment.NewLine, extractErrors)
                Debug.WriteLine($"资源提取部分失败: {_lastError}")
            End If

            ' 写入版本标记文件，下次启动时跳过提取
            Try
                Directory.CreateDirectory(appDataPath)
                File.WriteAllText(Path.Combine(appDataPath, ".version"), _resourceVersion)
                WriteManifest(appDataPath, manifestFile)
            Catch
                ' 写入失败不影响功能
            End Try

            Return appDataPath
        Catch ex As Exception
            _lastError = $"释放资源失败: {ex.Message}"
            Debug.WriteLine(_lastError)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Проверяет, что все файлы из манифеста на месте и их размеры не изменились.
    ''' </summary>
    Private Shared Function ManifestMatches(appDataPath As String, manifestPath As String) As Boolean
        Try
            For Each line In File.ReadAllLines(manifestPath)
                If String.IsNullOrWhiteSpace(line) Then Continue For

                Dim parts = line.Split("|"c)
                If parts.Length <> 2 Then Return False

                Dim target = Path.Combine(appDataPath, parts(0))
                If Not File.Exists(target) Then Return False

                Dim expected As Long = 0
                If Not Long.TryParse(parts(1), expected) Then Return False
                If New FileInfo(target).Length <> expected Then Return False
            Next

            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Пишет манифест извлечённых файлов (относительный путь|размер).
    ''' </summary>
    Private Shared Sub WriteManifest(appDataPath As String, manifestPath As String)
        Try
            Dim files = Directory.GetFiles(appDataPath, "*.*", SearchOption.AllDirectories)
            Array.Sort(files, StringComparer.OrdinalIgnoreCase)

            Dim sb As New System.Text.StringBuilder()
            For Each file In files
                Dim name = Path.GetFileName(file)
                If name = ".version" OrElse name = ManifestFileName Then Continue For

                Dim relative = file.Substring(appDataPath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                sb.Append(relative).Append("|"c).Append(New FileInfo(file).Length).AppendLine()
            Next

            File.WriteAllText(manifestPath, sb.ToString())
        Catch ex As Exception
            ' Манифест не критичен: без него ресурсы просто переизвлекутся
            Debug.WriteLine($"Не удалось записать манифест ресурсов: {ex.Message}")
        End Try
    End Sub

    Private Shared Function ExtractResourceToFileFromManager(resourceName As String, targetFileName As String, targetDir As String, rm As Resources.ResourceManager) As String
        Try
            ' 从资源管理器中获取资源
            Dim resourceObj = rm.GetObject(resourceName)

            If resourceObj IsNot Nothing Then
                Dim targetPath = Path.Combine(targetDir, targetFileName)

                ' 根据资源类型处理
                If TypeOf resourceObj Is Byte() Then
                    ' 如果是字节数组，直接写入
                    File.WriteAllBytes(targetPath, DirectCast(resourceObj, Byte()))
                ElseIf TypeOf resourceObj Is String Then
                    ' 如果是字符串，转换为字节后写入
                    File.WriteAllText(targetPath, DirectCast(resourceObj, String))
                Else
                    Dim errMsg = $"Unsupported resource type for {resourceName}: {resourceObj.GetType().Name}"
                    Debug.WriteLine(errMsg)
                    Return errMsg
                End If

                Debug.WriteLine($"Successfully extracted {resourceName} to {targetPath}")
                Return String.Empty
            Else
                Dim errMsg = $"Resource not found: {resourceName}"
                Debug.WriteLine(errMsg)

                ' 列出所有可用的资源，以便调试
                Try
                    Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
                    For Each entry As DictionaryEntry In resourceSet
                        Debug.WriteLine($"Available resource: {entry.Key}")
                    Next
                Catch
                End Try
                
                Return errMsg
            End If
        Catch ex As Exception
            Dim errMsg = $"提取资源 {resourceName} 失败: {ex.Message}"
            Debug.WriteLine(errMsg)
            Return errMsg
        End Try
    End Function

    ' 调试方法 - 列出所有资源
    Private Shared Sub DebugListResources()
        Try
            Dim assembly = GetType(ResourceExtractor).Assembly
            Debug.WriteLine("=== 所有嵌入资源 ===")
            For Each resourceName In assembly.GetManifestResourceNames()
                Debug.WriteLine($"Found resource: {resourceName}")
            Next

            Debug.WriteLine("=== Resources 中的资源 ===")
            Dim rm As New Resources.ResourceManager("ShareRibbon.Resources", assembly)
            Dim resourceSet = rm.GetResourceSet(Globalization.CultureInfo.CurrentUICulture, True, True)
            For Each entry As DictionaryEntry In resourceSet
                Debug.WriteLine($"Resource: {entry.Key} (Type: {If(entry.Value IsNot Nothing, entry.Value.GetType().ToString(), "null")})")
            Next
        Catch ex As Exception
            Debug.WriteLine($"列出资源时出错: {ex.Message}")
        End Try
    End Sub
End Class
