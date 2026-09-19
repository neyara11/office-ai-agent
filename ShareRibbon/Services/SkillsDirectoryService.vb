' ShareRibbon\Services\SkillsDirectoryService.vb
' Skills目录管理服务：从文件系统读取Claude规范的Skills

Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports System.Diagnostics
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Skills文件定义（Claude规范格式，支持Front Matter）
''' </summary>
Public Class SkillFileDefinition
    ' 必需字段
    Public Property Name As String
    Public Property Description As String
    Public Property Content As String

    ' 可选字段（Front Matter元数据）
    Public Property License As String
    Public Property Compatibility As String
    Public Property AllowedTools As List(Of String)
    Public Property Metadata As Dictionary(Of String, Object)
    Public Property ArgumentHint As String
    Public Property DisableModelInvocation As Boolean = False
    Public Property UserInvocable As Boolean = True
    Public Property Model As String
    Public Property Context As String
    Public Property Agent As String
    Public Property Hooks As Dictionary(Of String, String)

    ' 扩展字段
    Public Property FilePath As String

    ''' <summary>是否已加载完整 Skill 内容。目录扫描阶段为 False，命中后再二次加载。</summary>
    Public Property IsContentLoaded As Boolean = True

    ''' <summary>标签列表（从 front matter tags 字段解析）</summary>
    Public Property Tags As List(Of String) = New List(Of String)()

    ''' <summary>使用次数（运行时统计，不持久化到文件）</summary>
    Public Property UsageCount As Integer = 0

    ''' <summary>最后使用时间</summary>
    Public Property LastUsedAt As DateTime? = Nothing

    ''' <summary>成功率</summary>
    Public Property SuccessRate As Double? = Nothing

    ''' <summary>应用类型（用于分类）</summary>
    Public Property Application As String = ""

    ''' <summary>脚本列表（从 scripts/ 目录解析）</summary>
    Public Property Scripts As List(Of SkillScript) = New List(Of SkillScript)()

    Public ReadOnly Property AllowedToolsText As String
        Get
            If AllowedTools Is Nothing OrElse AllowedTools.Count = 0 Then Return ""
            Return String.Join(", ", AllowedTools)
        End Get
    End Property

    Public ReadOnly Property Author As String
        Get
            If Metadata IsNot Nothing AndAlso Metadata.ContainsKey("author") Then
                Return Metadata("author")?.ToString()
            End If
            Return ""
        End Get
    End Property

    Public ReadOnly Property Version As String
        Get
            If Metadata IsNot Nothing AndAlso Metadata.ContainsKey("version") Then
                Return Metadata("version")?.ToString()
            End If
            Return "1.0"
        End Get
    End Property

End Class

''' <summary>
''' Skill脚本定义
''' </summary>
Public Class SkillScript
    ''' <summary>脚本文件名</summary>
    Public Property FileName As String

    ''' <summary>脚本完整路径</summary>
    Public Property FilePath As String

    ''' <summary>脚本类型：python / powershell / shell / batch</summary>
    Public Property ScriptType As String

    ''' <summary>脚本描述</summary>
    Public Property Description As String

    ''' <summary>是否可执行</summary>
    Public Property Executable As Boolean = True

    ''' <summary>脚本参数说明</summary>
    Public Property ArgsHint As String = ""

    ''' <summary>工作目录（相对于Skill目录）</summary>
    Public Property WorkingDirectory As String = ""
End Class

''' <summary>
''' Skills目录服务
''' 管理Skills文件的读取、解析和缓存
''' </summary>
Public Class SkillsDirectoryService
    Private Shared _skillsDirectory As String = ""
    Private Shared _cachedSkills As New List(Of SkillFileDefinition)()
    Private Shared _cachedSkillCatalog As New List(Of SkillFileDefinition)()
    Private Shared _lastRefreshTime As DateTime = DateTime.MinValue
    Private Shared _lastCatalogRefreshTime As DateTime = DateTime.MinValue
    Private Shared ReadOnly _cacheDuration As TimeSpan = TimeSpan.FromMinutes(5)

    ''' <summary>
    ''' 获取Skills目录路径
    ''' </summary>
    Public Shared Function GetSkillsDirectory() As String
        If Not String.IsNullOrEmpty(_skillsDirectory) Then
            Return _skillsDirectory
        End If

        ' 默认路径：Documents\OfficeAiAppData\Skills
        Dim appDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "Skills")
        _skillsDirectory = appDataPath
        Return _skillsDirectory
    End Function

    ''' <summary>
    ''' 获取所有 Skill 搜索根目录。用户目录用于安装个人 Skill，程序集 Skills 目录用于内置 Skill。
    ''' </summary>
    Public Shared Function GetSkillsDirectories() As List(Of String)
        Dim dirs As New List(Of String)()
        Dim userDir = GetSkillsDirectory()
        AddSkillsDirectory(dirs, userDir)

        Try
            Dim roots As New List(Of String) From {
                Path.GetDirectoryName(GetType(SkillsDirectoryService).Assembly.Location),
                AppDomain.CurrentDomain.BaseDirectory
            }

            ' VSTO 项目引用不会稳定地把 ShareRibbon 的 Content 传递复制到宿主输出目录。
            ' 从程序集/宿主目录向上查找，兼容开发目录、安装目录和直接运行 ShareRibbon。
            For Each root In roots.Where(Function(r) Not String.IsNullOrWhiteSpace(r))
                Dim current = root
                While Not String.IsNullOrWhiteSpace(current)
                    AddSkillsDirectory(dirs, Path.Combine(current, "Skills"))
                    AddSkillsDirectory(dirs, Path.Combine(current, "ShareRibbon", "Skills"))
                    current = Path.GetDirectoryName(current)
                End While
            Next
        Catch ex As Exception
            Debug.WriteLine($"[SkillsDirectoryService] 获取内置 Skills 目录失败: {ex.Message}")
        End Try

        Return dirs
    End Function

    Private Shared Sub AddSkillsDirectory(dirs As List(Of String), path As String)
        If dirs Is Nothing OrElse String.IsNullOrWhiteSpace(path) Then Return
        If dirs.Any(Function(d) SamePath(d, path)) Then Return
        dirs.Add(path)
    End Sub

    ''' <summary>
    ''' 确保Skills目录存在
    ''' </summary>
    Public Shared Sub EnsureDirectoryExists()
        Dim dir = GetSkillsDirectory()
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        EnsureBuiltInSkillsExtracted()
    End Sub

    ''' <summary>
    ''' Разворачивает встроенные Skill-каталоги в пользовательский каталог Skills.
    ''' Существующие файлы не перезаписываются, чтобы не затирать правки пользователя.
    ''' Нужно, потому что установленная раскладка MSI не содержит каталог Skills.
    ''' </summary>
    Public Shared Sub EnsureBuiltInSkillsExtracted()
        Try
            Dim userDir = GetSkillsDirectory()
            If String.IsNullOrWhiteSpace(userDir) Then Return
            Directory.CreateDirectory(userDir)

            Dim asm = GetType(SkillsDirectoryService).Assembly
            Dim extracted As Integer = 0
            For Each resName In asm.GetManifestResourceNames()
                If String.IsNullOrWhiteSpace(resName) Then Continue For
                If Not resName.StartsWith("Skills", StringComparison.OrdinalIgnoreCase) Then Continue For
                If resName.EndsWith(".resources", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim rel = resName.Substring("Skills".Length).TrimStart("."c, "/"c, "\"c)
                If String.IsNullOrWhiteSpace(rel) Then Continue For
                rel = rel.Replace("/"c, Path.DirectorySeparatorChar).Replace("\"c, Path.DirectorySeparatorChar)

                Dim targetPath = Path.Combine(userDir, rel)
                Dim targetDir = Path.GetDirectoryName(targetPath)
                If Not String.IsNullOrWhiteSpace(targetDir) AndAlso Not Directory.Exists(targetDir) Then
                    Directory.CreateDirectory(targetDir)
                End If

                Using stream = asm.GetManifestResourceStream(resName)
                    If stream Is Nothing Then Continue For
                    Using memory As New IO.MemoryStream()
                        stream.CopyTo(memory)
                        Dim content = memory.ToArray()

                        ' Обновляем встроенный файл, если он изменился; иначе не трогаем.
                        If IO.File.Exists(targetPath) Then
                            Dim existing = IO.File.ReadAllBytes(targetPath)
                            If existing.Length = content.Length AndAlso existing.SequenceEqual(content) Then
                                Continue For
                            End If
                        End If

                        IO.File.WriteAllBytes(targetPath, content)
                        extracted += 1
                    End Using
                End Using
            Next

            If extracted > 0 Then
                Debug.WriteLine($"[SkillsDirectoryService] встроенных Skills извлечено: {extracted} -> {userDir}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SkillsDirectoryService] не удалось извлечь встроенные Skills: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取所有Skills（带缓存）
    ''' </summary>
    Public Shared Function GetAllSkills(Optional forceRefresh As Boolean = False) As List(Of SkillFileDefinition)
        If Not forceRefresh AndAlso _cachedSkills.Count > 0 AndAlso (DateTime.Now - _lastRefreshTime) < _cacheDuration Then
            Return _cachedSkills.ToList()
        End If

        RefreshSkills()
        Return _cachedSkills.ToList()
    End Function

    ''' <summary>
    ''' 获取 Skills 元数据目录。不会读取 reference/examples 全文，也不会把 Skill 内容注入上下文。
    ''' </summary>
    Public Shared Function GetSkillsCatalog(Optional forceRefresh As Boolean = False) As List(Of SkillFileDefinition)
        If Not forceRefresh AndAlso _cachedSkillCatalog.Count > 0 AndAlso (DateTime.Now - _lastCatalogRefreshTime) < _cacheDuration Then
            Return _cachedSkillCatalog.ToList()
        End If

        RefreshSkillCatalog()
        Return _cachedSkillCatalog.ToList()
    End Function

    ''' <summary>
    ''' 刷新Skills缓存
    ''' </summary>
    Public Shared Sub RefreshSkills()
        _cachedSkills.Clear()

        EnsureBuiltInSkillsExtracted()

        For Each skillsRoot In GetSkillsDirectories()
            If Not Directory.Exists(skillsRoot) Then Continue For
            LoadSkillsFromRoot(skillsRoot, includeDetails:=True, target:=_cachedSkills)
        Next

        _lastRefreshTime = DateTime.Now
    End Sub

    ''' <summary>
    ''' 只刷新元数据目录，用于请求前第一阶段扫描。
    ''' </summary>
    Public Shared Sub RefreshSkillCatalog()
        _cachedSkillCatalog.Clear()

        EnsureBuiltInSkillsExtracted()

        For Each skillsRoot In GetSkillsDirectories()
            If Not Directory.Exists(skillsRoot) Then Continue For
            LoadSkillsFromRoot(skillsRoot, includeDetails:=False, target:=_cachedSkillCatalog)
        Next

        _lastCatalogRefreshTime = DateTime.Now
        Dim existingRoots = GetSkillsDirectories().Where(Function(d) Directory.Exists(d)).ToList()
        Debug.WriteLine($"[SkillsDirectoryService] Skill catalog={_cachedSkillCatalog.Count}, roots={String.Join("; ", existingRoots)}")
    End Sub

    Private Shared Sub LoadSkillsFromRoot(rootDir As String, includeDetails As Boolean, target As List(Of SkillFileDefinition))
        For Each skillDir In Directory.GetDirectories(rootDir)
            Try
                Dim skill = ParseSkillDirectory(skillDir, includeDetails)
                AddOrReplaceSkill(target, skill)
            Catch ex As Exception
                Debug.WriteLine($"[SkillsDirectoryService] 解析目录失败: {skillDir}, 错误: {ex.Message}")
            End Try
        Next

        For Each file In Directory.GetFiles(rootDir, "*.json", SearchOption.TopDirectoryOnly)
            Try
                Dim skill = ParseSkillJsonFile(file, includeDetails)
                AddOrReplaceSkill(target, skill)
            Catch ex As Exception
                Debug.WriteLine($"[SkillsDirectoryService] 解析JSON文件失败: {file}, 错误: {ex.Message}")
            End Try
        Next
    End Sub

    Private Shared Sub AddOrReplaceSkill(target As List(Of SkillFileDefinition), skill As SkillFileDefinition)
        If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then Return

        Dim existingIndex = target.FindIndex(Function(s) String.Equals(s.Name, skill.Name, StringComparison.OrdinalIgnoreCase))
        If existingIndex >= 0 Then
            ' 用户目录先被扫描，内置目录不覆盖同名用户 Skill。
            Return
        End If

        target.Add(skill)
    End Sub

    ''' <summary>
    ''' 二次加载完整 Skill。只在召回命中后调用。
    ''' </summary>
    Public Shared Function LoadSkillDetail(skill As SkillFileDefinition) As SkillFileDefinition
        If skill Is Nothing Then Return Nothing
        If skill.IsContentLoaded Then Return skill

        Try
            If Directory.Exists(skill.FilePath) Then
                Return ParseSkillDirectory(skill.FilePath, includeDetails:=True)
            End If

            If File.Exists(skill.FilePath) Then
                Return ParseSkillJsonFile(skill.FilePath, includeDetails:=True)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SkillsDirectoryService] 二次加载 Skill 失败: {skill.FilePath}, 错误: {ex.Message}")
        End Try

        Return skill
    End Function

    Public Shared Function GetSkillByNameOrPath(skillName As String, filePath As String) As SkillFileDefinition
        Dim catalog = GetSkillsCatalog()
        Dim matched = catalog.FirstOrDefault(Function(s) String.Equals(s.Name, skillName, StringComparison.OrdinalIgnoreCase) OrElse
                                                        SamePath(s.FilePath, filePath))
        If matched Is Nothing Then Return Nothing
        Return LoadSkillDetail(matched)
    End Function

    ''' <summary>
    ''' 解析Skill目录（Claude规范）
    ''' 目录结构：
    ''' my-skill/
    '''   ├── SKILL.md (required)
    '''   ├── reference.md (optional)
    '''   ├── examples.md (optional)
    '''   ├── scripts/ (optional)
    '''   └── templates/ (optional)
    ''' </summary>
    Private Shared Function ParseSkillDirectory(dirPath As String, Optional includeDetails As Boolean = True) As SkillFileDefinition
        Dim skillMdPath = Path.Combine(dirPath, "SKILL.md")
        If Not File.Exists(skillMdPath) Then
            Return Nothing
        End If

        Dim skillName = Path.GetFileName(dirPath)
        Dim fileContent = File.ReadAllText(skillMdPath)

        Dim skill As New SkillFileDefinition()
        skill.Name = skillName
        skill.FilePath = dirPath
        skill.IsContentLoaded = includeDetails

        ' 解析Front Matter和内容
        ParseFrontMatterAndContent(fileContent, skill, includeDetails)

        If includeDetails Then
            ' 尝试读取reference.md
            Dim refPath = Path.Combine(dirPath, "reference.md")
            If File.Exists(refPath) Then
                skill.Content &= vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf & File.ReadAllText(refPath)
            End If

            ' 支持官方 Skills 风格的 references/ 目录，命中后才二次加载。
            Dim referencesDir = Path.Combine(dirPath, "references")
            If Directory.Exists(referencesDir) Then
                For Each refFile In Directory.GetFiles(referencesDir, "*.md", SearchOption.TopDirectoryOnly).OrderBy(Function(p) p)
                    skill.Content &= vbCrLf & vbCrLf & "---" & vbCrLf & $"# references/{Path.GetFileName(refFile)}" & vbCrLf & vbCrLf & File.ReadAllText(refFile)
                Next
            End If

            ' 尝试读取examples.md
            Dim examplesPath = Path.Combine(dirPath, "examples.md")
            If File.Exists(examplesPath) Then
                skill.Content &= vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf & File.ReadAllText(examplesPath)
            End If
        End If

        ' 发现并解析 scripts/ 目录下的脚本
        DiscoverSkillScripts(dirPath, skill)

        Return skill
    End Function

    ''' <summary>
    ''' 发现 Skill 目录下的脚本文件
    ''' </summary>
    Private Shared Sub DiscoverSkillScripts(skillDir As String, skill As SkillFileDefinition)
        Dim scriptsPath = Path.Combine(skillDir, "scripts")
        If Not Directory.Exists(scriptsPath) Then Return

        Try
            Dim scriptFiles = Directory.GetFiles(scriptsPath, "*.*", SearchOption.TopDirectoryOnly)
            For Each scriptPath In scriptFiles
                Dim fileName = Path.GetFileName(scriptPath)
                Dim ext = Path.GetExtension(fileName).ToLowerInvariant()

                ' 只处理可执行脚本类型
                Select Case ext
                    Case ".py", ".ps1", ".sh", ".bat", ".cmd"
                        Dim scriptType = GetScriptType(ext)
                        If scriptType <> "" Then
                            Dim script = New SkillScript() With {
                                .FileName = fileName,
                                .FilePath = scriptPath,
                                .ScriptType = scriptType,
                                .Description = $"Запуск скрипта {fileName}",
                                .Executable = True
                            }

                            ' 尝试从同名的 .md 文件读取脚本说明
                            Dim descPath = Path.Combine(scriptsPath, fileName & ".md")
                            If File.Exists(descPath) Then
                                Dim descContent = File.ReadAllText(descPath)
                                ' 解析脚本说明（支持 frontmatter 格式）
                                Dim descLines = descContent.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                                Dim inFrontMatter = False
                                For Each line In descLines
                                    If line.Trim() = "---" Then
                                        inFrontMatter = Not inFrontMatter
                                        Continue For
                                    End If
                                    If Not inFrontMatter AndAlso line.StartsWith("#") Then
                                        script.Description = line.TrimStart("#"c).Trim()
                                        Exit For
                                    End If
                                Next
                                If script.Description = $"Запуск скрипта {fileName}" Then
                                    ' 如果没有找到标题，使用第一行非空非frontmatter行
                                    For Each line In descLines
                                        If Not line.Trim().StartsWith("---") AndAlso Not line.Trim().StartsWith("#") AndAlso Not String.IsNullOrWhiteSpace(line.Trim()) Then
                                            script.Description = line.Trim()
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If

                            skill.Scripts.Add(script)
                        End If
                End Select
            Next
        Catch ex As Exception
            Debug.WriteLine($"[SkillsDirectoryService] 发现脚本失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 根据扩展名获取脚本类型
    ''' </summary>
    Private Shared Function GetScriptType(ext As String) As String
        Select Case ext.ToLower()
            Case ".py" : Return "python"
            Case ".ps1" : Return "powershell"
            Case ".sh" : Return "shell"
            Case ".bat", ".cmd" : Return "batch"
            Case Else : Return ""
        End Select
    End Function

    ''' <summary>
    ''' 解析Front Matter和内容
    ''' </summary>
    Private Shared Sub ParseFrontMatterAndContent(fileContent As String, skill As SkillFileDefinition, Optional includeContent As Boolean = True)
        ' 查找Front Matter分隔符
        Dim lines = fileContent.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        Dim frontMatterLines As New List(Of String)()
        Dim contentLines As New List(Of String)()
        Dim inFrontMatter = False
        Dim frontMatterEnded = False
        Dim frontMatterStartIndex = -1
        Dim frontMatterEndIndex = -1

        For i = 0 To lines.Length - 1
            Dim line = lines(i).Trim()

            If line = "---" Then
                If Not frontMatterEnded Then
                    If frontMatterStartIndex = -1 Then
                        frontMatterStartIndex = i
                        inFrontMatter = True
                    Else
                        frontMatterEndIndex = i
                        inFrontMatter = False
                        frontMatterEnded = True
                    End If
                End If
            ElseIf inFrontMatter Then
                frontMatterLines.Add(lines(i))
            ElseIf frontMatterEnded OrElse frontMatterStartIndex = -1 Then
                contentLines.Add(lines(i))
            End If
        Next

        ' 解析Front Matter
        If frontMatterLines.Count > 0 Then
            ParseFrontMatterLines(frontMatterLines, skill)
        End If

        ' 剩余部分作为内容。目录扫描阶段不保留全文，避免第一阶段加载过重。
        If includeContent Then
            skill.Content = String.Join(vbCrLf, contentLines).Trim()
        Else
            skill.Content = ""
        End If

        ' 如果没有从Front Matter获取到name，从第一个标题获取
        If String.IsNullOrWhiteSpace(skill.Name) OrElse skill.Name = Path.GetFileName(skill.FilePath) Then
            For Each line In contentLines
                If line.StartsWith("# ") Then
                    skill.Name = line.Substring(2).Trim()
                    Exit For
                End If
            Next
        End If

        If String.IsNullOrWhiteSpace(skill.Description) Then
            skill.Description = FirstNonEmptyContentLine(contentLines)
        End If
    End Sub

    ''' <summary>
    ''' 解析Front Matter行
    ''' </summary>
    Private Shared Sub ParseFrontMatterLines(lines As List(Of String), skill As SkillFileDefinition)
        skill.Metadata = New Dictionary(Of String, Object)()
        Dim inMetadata = False

        For Each line In lines
            Dim trimmedLine = line.Trim()

            If String.IsNullOrWhiteSpace(trimmedLine) Then
                Continue For
            End If

            If trimmedLine.StartsWith("metadata:") Then
                inMetadata = True
                Continue For
            End If

            If inMetadata AndAlso trimmedLine.StartsWith("  ") Then
                ' metadata下的子项
                Dim colonIndex = trimmedLine.IndexOf(":")
                If colonIndex > 0 Then
                    Dim key = trimmedLine.Substring(0, colonIndex).Trim()
                    Dim value = trimmedLine.Substring(colonIndex + 1).Trim()
                    ' 移除引号
                    If value.StartsWith("""") AndAlso value.EndsWith("""") Then
                        value = value.Substring(1, value.Length - 2)
                    ElseIf value.StartsWith("'") AndAlso value.EndsWith("'") Then
                        value = value.Substring(1, value.Length - 2)
                    End If
                    skill.Metadata(key) = value
                End If
                Continue For
            ElseIf inMetadata Then
                inMetadata = False
            End If

            ' 标准字段
            Dim colonIndex2 = trimmedLine.IndexOf(":")
            If colonIndex2 > 0 Then
                Dim key = trimmedLine.Substring(0, colonIndex2).Trim()
                Dim value = trimmedLine.Substring(colonIndex2 + 1).Trim()
                ' 移除引号
                If value.StartsWith("""") AndAlso value.EndsWith("""") Then
                    value = value.Substring(1, value.Length - 2)
                ElseIf value.StartsWith("'") AndAlso value.EndsWith("'") Then
                    value = value.Substring(1, value.Length - 2)
                End If
                ' 移除#注释
                Dim hashIndex = value.IndexOf("#")
                If hashIndex > 0 Then
                    value = value.Substring(0, hashIndex).Trim()
                End If

                Select Case key.ToLowerInvariant()
                    Case "name"
                        skill.Name = value
                    Case "description"
                        skill.Description = value
                    Case "license"
                        skill.License = value
                    Case "compatibility"
                        skill.Compatibility = value
                    Case "allowed-tools", "allowed_tools", "tools"
                        skill.AllowedTools = ParseListValue(value)
                    Case "argument-hint"
                        skill.ArgumentHint = value
                    Case "disable-model-invocation"
                        Dim boolVal As Boolean
                        If Boolean.TryParse(value, boolVal) Then
                            skill.DisableModelInvocation = boolVal
                        End If
                    Case "user-invocable"
                        Dim boolVal As Boolean
                        If Boolean.TryParse(value, boolVal) Then
                            skill.UserInvocable = boolVal
                        End If
                    Case "model"
                        skill.Model = value
                    Case "context"
                        skill.Context = value
                    Case "agent"
                        skill.Agent = value
                    Case "tags"
                        skill.Tags = ParseListValue(value)
                    Case "application", "app", "app-scope", "app_scope"
                        skill.Application = value
                    Case "intent", "intents", "intent-type", "intent_type", "intent_types"
                        If skill.Metadata Is Nothing Then skill.Metadata = New Dictionary(Of String, Object)()
                        skill.Metadata("intent_types") = value
                    Case Else
                        ' 保留扩展 front matter，供运行时选择策略使用，例如
                        ' default_for_application。未知元数据不应因解析器升级滞后而丢失。
                        If skill.Metadata Is Nothing Then skill.Metadata = New Dictionary(Of String, Object)()
                        skill.Metadata(key.ToLowerInvariant().Replace("-", "_")) = value
                End Select
            End If
        Next
    End Sub

    Private Shared Function ParseListValue(value As String) As List(Of String)
        Dim result As New List(Of String)()
        If String.IsNullOrWhiteSpace(value) Then Return result

        Dim text = value.Trim()
        If text.StartsWith("[") AndAlso text.EndsWith("]") Then
            text = text.Substring(1, text.Length - 2)
        End If

        For Each part In text.Split({","c, "|"c, ";"c, "；"c, "，"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim item = part.Trim().Trim(""""c).Trim("'"c)
            If Not String.IsNullOrWhiteSpace(item) Then result.Add(item)
        Next

        Return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    ''' <summary>
    ''' 解析单个Skill JSON文件（兼容旧格式）
    ''' </summary>
    Private Shared Function ParseSkillJsonFile(filePath As String, Optional includeDetails As Boolean = True) As SkillFileDefinition
        Dim content = File.ReadAllText(filePath)
        Dim jo = JObject.Parse(content)

        Dim skill As New SkillFileDefinition()
        skill.FilePath = filePath
        skill.IsContentLoaded = includeDetails

        ' 读取基本信息
        skill.Name = If(jo("name")?.ToString(), Path.GetFileNameWithoutExtension(filePath))
        skill.Description = If(jo("description")?.ToString(), "")
        skill.Application = If(jo("application")?.ToString(), If(jo("appType")?.ToString(), ""))

        Dim allowedTools = TryCast(jo("allowedTools"), JArray)
        If allowedTools Is Nothing Then allowedTools = TryCast(jo("requiredTools"), JArray)
        If allowedTools IsNot Nothing Then
            skill.AllowedTools = allowedTools.Select(Function(t) t.ToString()).Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
        End If

        Dim tags = TryCast(jo("tags"), JArray)
        If tags IsNot Nothing Then
            skill.Tags = tags.Select(Function(t) t.ToString()).Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
        End If

        Dim triggerPatterns = TryCast(jo("triggerPatterns"), JArray)
        If triggerPatterns IsNot Nothing Then
            For Each token In triggerPatterns
                Dim tag = token.ToString()
                If Not String.IsNullOrWhiteSpace(tag) AndAlso Not skill.Tags.Contains(tag, StringComparer.OrdinalIgnoreCase) Then
                    skill.Tags.Add(tag)
                End If
            Next
        End If

        If skill.Metadata Is Nothing Then skill.Metadata = New Dictionary(Of String, Object)()
        If jo("intent_types") IsNot Nothing Then skill.Metadata("intent_types") = jo("intent_types").ToString()
        If triggerPatterns IsNot Nothing Then skill.Metadata("triggers") = String.Join(",", skill.Tags)

        ' 读取keywords兼容
        Dim keywordsToken = jo("keywords")
        If keywordsToken IsNot Nothing AndAlso TypeOf keywordsToken Is JArray Then
            ' 不做任何处理，保持兼容性
        ElseIf keywordsToken IsNot Nothing AndAlso TypeOf keywordsToken Is JValue Then
            ' 不做任何处理，保持兼容性
        End If

        If includeDetails Then
            ' 读取Skill内容
            Dim contentToken = jo("content")
            If contentToken IsNot Nothing Then
                skill.Content = contentToken.ToString()
            Else
                Dim promptToken = jo("prompt")
                If promptToken IsNot Nothing Then
                    skill.Content = promptToken.ToString()
                Else
                    Dim promptTemplateToken = jo("promptTemplate")
                    If promptTemplateToken IsNot Nothing Then
                        skill.Content = promptTemplateToken.ToString()
                    End If
                End If
            End If
        Else
            skill.Content = ""
        End If

        Return skill
    End Function

    Private Shared Function FirstNonEmptyContentLine(lines As List(Of String)) As String
        If lines Is Nothing Then Return ""
        For Each line In lines
            Dim text = If(line, "").Trim()
            If text.Length > 0 AndAlso Not text.StartsWith("#") Then Return text
        Next
        Return ""
    End Function

    Private Shared Function SamePath(leftPath As String, rightPath As String) As Boolean
        If String.IsNullOrWhiteSpace(leftPath) OrElse String.IsNullOrWhiteSpace(rightPath) Then Return False
        Try
            Return String.Equals(Path.GetFullPath(leftPath), Path.GetFullPath(rightPath), StringComparison.OrdinalIgnoreCase)
        Catch
            Return String.Equals(leftPath.Trim(), rightPath.Trim(), StringComparison.OrdinalIgnoreCase)
        End Try
    End Function

    ''' <summary>
    ''' 打开Skills目录
    ''' </summary>
    Public Shared Sub OpenSkillsDirectory()
        EnsureDirectoryExists()
        Dim dir = GetSkillsDirectory()
        Process.Start("explorer.exe", dir)
    End Sub

    ''' <summary>
    ''' 打开指定Skill的目录
    ''' </summary>
    Public Shared Sub OpenSkillDirectory(skill As SkillFileDefinition)
        If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.FilePath) Then
            OpenSkillsDirectory()
            Return
        End If

        If Directory.Exists(skill.FilePath) Then
            Process.Start("explorer.exe", skill.FilePath)
        ElseIf File.Exists(skill.FilePath) Then
            Process.Start("explorer.exe", Path.GetDirectoryName(skill.FilePath))
        Else
            OpenSkillsDirectory()
        End If
    End Sub

End Class
