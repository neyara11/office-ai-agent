' ShareRibbon\Services\SkillsService.vb
' Skills服务：实现Claude Skills规范和渐进式披露
' 支持从文件系统目录读取Skills（类似Trae/Cursor模式）

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Skills匹配结果
''' </summary>
Public Class SkillMatchResult
    Public Property Skill As SkillFileDefinition
    Public Property MatchScore As Double
    Public Property MatchedKeywords As List(Of String)
End Class

''' <summary>
''' Skills使用统计（持久化）
''' </summary>
Public Class SkillUsageStats
    Public Property SkillName As String
    Public Property UsageCount As Integer
    Public Property LastUsedAt As DateTime?
    Public Property SuccessCount As Integer
    Public Property TotalTokens As Long
End Class

''' <summary>
''' Skills使用统计存储
''' </summary>
Public Class SkillsUsageStorage
    Public Property Skills As New Dictionary(Of String, SkillUsageStats)()
    Public Property LastUpdated As DateTime = DateTime.Now
End Class

''' <summary>
''' Skills服务：实现渐进式披露和智能匹配
''' 从Skills目录读取Claude规范的Skills文件
''' </summary>
Public Class SkillsService

    ' 同义词词库（用于语义匹配）
    Private Shared ReadOnly Synonyms As New Dictionary(Of String, List(Of String))() From {
        {"excel", New List(Of String) From {"电子表格", "spreadsheet", "xlsx", "xls"}},
        {"word", New List(Of String) From {"文档", "docx", "doc", "文字处理"}},
        {"powerpoint", New List(Of String) From {"ppt", "pptx", "演示", "幻灯片"}},
        {"数据", New List(Of String) From {"data", "dataset", "数据库"}},
        {"分析", New List(Of String) From {"analyze", "analysis", "统计"}},
        {"图表", New List(Of String) From {"chart", "graph", "可视化", "visualization"}},
        {"公式", New List(Of String) From {"function", "formula", "函数", "计算"}},
        {"格式", New List(Of String) From {"format", "样式", "style", "排版"}},
        {"表格", New List(Of String) From {"table", "range", "区域"}},
        {"单元格", New List(Of String) From {"cell", "单元格"}},
        {"脚本", New List(Of String) From {"script", "vba", "宏", "macro"}},
        {"模板", New List(Of String) From {"template", "模板"}},
        {"报告", New List(Of String) From {"report", "报表", "summary"}},
        {"清理", New List(Of String) From {"clean", "清洗", "整理"}},
        {"转换", New List(Of String) From {"convert", "transform", "转换"}},
        {"批量", New List(Of String) From {"batch", "批量", "mass"}},
        {"智能", New List(Of String) From {"ai", "智能", "smart"}},
        {"自动", New List(Of String) From {"auto", "自动", "automation"}},
        {"助手", New List(Of String) From {"assistant", "helper", "助手"}},
        {"专家", New List(Of String) From {"expert", "specialist", "专家"}},
        {"顾问", New List(Of String) From {"advisor", "consultant", "顾问"}}
    }

    ' Служебные слова не несут смысла для выбора Skill, но как подстроки дают ложные
    ' срабатывания (например, "на" внутри "навыков"), поэтому исключаются из запроса.
    Private Shared ReadOnly StopWords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "на", "в", "во", "и", "с", "со", "по", "за", "из", "к", "ко", "о", "об", "от", "до",
        "для", "при", "про", "над", "под", "у", "не", "ни", "же", "бы", "ли", "это", "этот",
        "мне", "мой", "моя", "мы", "вы", "ты", "он", "она", "они", "как", "что", "чтобы",
        "the", "a", "an", "of", "to", "in", "on", "for", "and", "or", "with", "is", "are", "be", "at", "by"
    }

    ' 使用统计文件路径
    Private Shared ReadOnly UsageStatsPath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        "skills_usage.json"
    )

    ' 缓存的使用统计
    Private Shared _usageStorage As SkillsUsageStorage = Nothing
    Private Shared _usageStorageLock As New Object()

    ''' <summary>
    ''' 获取Skills目录（用于渐进式披露的第一步）
    ''' 只返回Skill的元数据，不返回详细内容
    ''' </summary>
    Public Shared Function GetSkillsCatalog() As List(Of SkillFileDefinition)
        Dim skills = SkillsDirectoryService.GetSkillsCatalog()

        ' 加载使用统计并合并到Skills
        LoadUsageStats()
        For Each skill In skills
            If _usageStorage.Skills.ContainsKey(skill.Name.ToLowerInvariant()) Then
                Dim stats = _usageStorage.Skills(skill.Name.ToLowerInvariant())
                skill.UsageCount = stats.UsageCount
                skill.LastUsedAt = stats.LastUsedAt
            End If
        Next
        ApplyRegistryUsageStats(skills)

        Return skills
    End Function

    Private Shared Sub ApplyRegistryUsageStats(skills As List(Of SkillFileDefinition))
        If skills Is Nothing OrElse skills.Count = 0 Then Return

        For Each skill In skills
            Try
                If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then Continue For
                Dim record = AgentMemoryRepository.GetSkillRegistryByName(skill.Name)
                If record Is Nothing Then Continue For

                If record.UsageCount > 0 Then skill.UsageCount = record.UsageCount
                If record.SuccessCount > 0 AndAlso record.UsageCount > 0 Then
                    skill.SuccessRate = CDbl(record.SuccessCount) / record.UsageCount
                End If

                Dim lastUsed As DateTime
                If Not String.IsNullOrWhiteSpace(record.LastIndexedAt) AndAlso DateTime.TryParse(record.LastIndexedAt, lastUsed) Then
                    skill.LastUsedAt = lastUsed
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SkillsService] ApplyRegistryUsageStats failed: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 智能匹配Skills（基于用户查询，支持语义匹配）
    ''' </summary>
    Public Shared Function MatchSkills(userQuery As String, Optional topN As Integer = 5) As List(Of SkillMatchResult)
        Dim results As New List(Of SkillMatchResult)()
        Dim allSkills = GetSkillsCatalog()

        If allSkills.Count = 0 Then
            Return results
        End If

        Dim queryLower = userQuery.ToLowerInvariant()
        Dim queryWords = TokenizeQuery(queryLower)

        For Each skill In allSkills
            Dim matchResult = CalculateMatchScore(queryLower, queryWords, skill)
            If matchResult.MatchScore > 0 Then
                results.Add(matchResult)
            End If
        Next

        ' 按匹配分数排序，分数相同则按使用次数排序
        Return results.OrderByDescending(Function(r) r.MatchScore) _
                      .ThenByDescending(Function(r) r.Skill.UsageCount) _
                      .Take(topN).ToList()
    End Function

    ''' <summary>
    ''' 将查询分词
    ''' </summary>
    Private Shared Function TokenizeQuery(query As String) As List(Of String)
        Dim words As New List(Of String)()

        ' 简单分词：按常见分隔符分割
        Dim tokens = query.Split({" "c, ","c, "，"c, "。"c, "."c, "、"c, "/"c, "\"c,
                                 "("c, ")"c, "["c, "]"c, "{"c, "}"c, "："c, ":"c,
                                 "!"c, "！"c, "?"c, "？"c, ";"c, "；"c},
                                 StringSplitOptions.RemoveEmptyEntries)

        For Each token In tokens
            Dim t = token.Trim()
            If t.Length > 0 AndAlso Not StopWords.Contains(t) Then
                words.Add(t)
                ' 添加同义词
                For Each kvp In Synonyms
                    If kvp.Value.Contains(t) OrElse kvp.Key = t Then
                        If Not words.Contains(kvp.Key) Then
                            words.Add(kvp.Key)
                        End If
                        For Each syn In kvp.Value
                            If Not words.Contains(syn) Then
                                words.Add(syn)
                            End If
                        Next
                    End If
                Next
            End If
        Next

        Return words.Distinct().ToList()
    End Function

    ''' <summary>
    ''' 计算Skill匹配分数（增强版，支持语义匹配）
    ''' </summary>
    Private Shared Function CalculateMatchScore(queryLower As String, queryWords As List(Of String), skill As SkillFileDefinition) As SkillMatchResult
        Dim score As Double = 0
        Dim matchedKeywords As New List(Of String)()

        ' === 1. 精确匹配名称（最高权重） ===
        If Not String.IsNullOrWhiteSpace(skill.Name) Then
            Dim nameLower = skill.Name.ToLowerInvariant()
            If queryLower.Contains(nameLower) Then
                score += 30  ' 提高名称匹配权重
                matchedKeywords.Add(skill.Name)
            End If

            ' 查询词完全包含在Skill名称中
            For Each word In queryWords
                If word.Length >= 3 AndAlso nameLower.Contains(word) Then
                    score += 10
                    If Not matchedKeywords.Contains(word) Then
                        matchedKeywords.Add(word)
                    End If
                End If
            Next
        End If

        ' === 2. 匹配描述词（增强版） ===
        If Not String.IsNullOrWhiteSpace(skill.Description) Then
            Dim descLower = skill.Description.ToLowerInvariant()

            ' 查询词匹配
            For Each word In queryWords
                If word.Length >= 3 AndAlso descLower.Contains(word) Then
                    score += 5
                    If Not matchedKeywords.Contains(word) Then
                        matchedKeywords.Add(word)
                    End If
                End If
            Next
        End If

        ' === 3. 匹配 tags（增强版） ===
        If skill.Tags IsNot Nothing Then
            For Each tag In skill.Tags
                If Not String.IsNullOrWhiteSpace(tag) Then
                    Dim tagLower = tag.ToLowerInvariant()

                    ' 直接匹配
                    If queryLower.Contains(tagLower) Then
                        score += 8
                        If Not matchedKeywords.Contains(tag) Then
                            matchedKeywords.Add(tag)
                        End If
                    End If

                    ' 同义词匹配
                    For Each word In queryWords
                        If Synonyms.ContainsKey(word) AndAlso Synonyms(word).Contains(tagLower) Then
                            score += 6
                            If Not matchedKeywords.Contains(tag) Then
                                matchedKeywords.Add(tag)
                            End If
                            Exit For
                        End If
                        If Synonyms.ContainsKey(tagLower) AndAlso Synonyms(tagLower).Contains(word) Then
                            score += 6
                            If Not matchedKeywords.Contains(tag) Then
                                matchedKeywords.Add(tag)
                            End If
                            Exit For
                        End If
                    Next
                End If
            Next
        End If

        ' 使用频率和最近使用只能用于相关结果之间的排序，不能凭空制造语义命中。
        ' 否则任何近期使用过的 Skill 都会污染后续不相关请求。
        ' Specialized Skills can publish fresh routing vocabulary without adding
        ' request-specific branches to the runtime selector.
        If skill.Metadata IsNot Nothing Then
            For Each metadataKey In New String() {"keywords", "triggers", "trigger_keywords"}
                Dim rawTerms As Object = Nothing
                If Not skill.Metadata.TryGetValue(metadataKey, rawTerms) OrElse rawTerms Is Nothing Then Continue For
                For Each term In rawTerms.ToString().Split({","c, ";"c, "|"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim normalizedTerm = term.Trim().ToLowerInvariant()
                    If normalizedTerm.Length <= 1 OrElse Not queryLower.Contains(normalizedTerm) Then Continue For
                    score += 10
                    If Not matchedKeywords.Contains(normalizedTerm, StringComparer.OrdinalIgnoreCase) Then
                        matchedKeywords.Add(normalizedTerm)
                    End If
                Next
            Next
        End If

        If score <= 0 Then
            Return New SkillMatchResult With {
                .Skill = skill,
                .MatchScore = 0,
                .MatchedKeywords = matchedKeywords
            }
        End If

        ' === 4. 使用频率加权（热门 Skill 加分，最多 +5） ===
        If skill.UsageCount > 0 Then
            score += Math.Min(5.0, skill.UsageCount * 0.8)
        End If

        ' === 5. 最近使用加分（活跃Skill） ===
        If skill.LastUsedAt.HasValue Then
            Dim daysSinceUse = (DateTime.Now - skill.LastUsedAt.Value).TotalDays
            If daysSinceUse < 1 Then
                score += 5  ' 24小时内用过
            ElseIf daysSinceUse < 7 Then
                score += 3  ' 一周内用过
            ElseIf daysSinceUse < 30 Then
                score += 1  ' 一个月内用过
            End If
        End If

        Return New SkillMatchResult With {
            .Skill = skill,
            .MatchScore = score,
            .MatchedKeywords = matchedKeywords
        }
    End Function

    ' 批量保存计数器
    Private Shared _unsavedChanges As Integer = 0
    Private Const SAVE_BATCH_SIZE As Integer = 10

    ''' <summary>
    ''' 记录 Skill 使用情况（持久化，批量写入）
    ''' </summary>
    Public Shared Sub RecordSkillUsage(skillName As String, Optional success As Boolean = True, Optional tokensUsed As Long = 0)
        Try
            LoadUsageStats()

            Dim key = skillName.ToLowerInvariant()
            If Not _usageStorage.Skills.ContainsKey(key) Then
                _usageStorage.Skills(key) = New SkillUsageStats With {
                    .SkillName = skillName,
                    .UsageCount = 0,
                    .SuccessCount = 0,
                    .TotalTokens = 0
                }
            End If

            Dim stats = _usageStorage.Skills(key)
            stats.UsageCount += 1
            stats.LastUsedAt = DateTime.Now
            If success Then
                stats.SuccessCount += 1
            End If
            stats.TotalTokens += tokensUsed

            Try
                AgentMemoryRepository.RecordSkillRegistryUsage(skillName, success)
                MemoryRepository.RecordSkillUsage(skillName, success, tokensUsed)
                SkillsIndexService.KickoffIndex()
            Catch ex As Exception
                Debug.WriteLine($"[SkillsService] RecordSkillUsage registry sync failed: {ex.Message}")
            End Try

            ' 批量保存：每N次或程序退出时才写文件
            _unsavedChanges += 1
            If _unsavedChanges >= SAVE_BATCH_SIZE Then
                SaveUsageStats()
                _unsavedChanges = 0
            End If

            ' 更新内存中的Skill对象（使用缓存，避免刷新开销）
            Dim skill = SkillsDirectoryService.GetSkillsCatalog().FirstOrDefault(Function(s) String.Equals(s.Name, skillName, StringComparison.OrdinalIgnoreCase))
            If skill IsNot Nothing Then
                skill.UsageCount = stats.UsageCount
                skill.LastUsedAt = stats.LastUsedAt
            End If

            Debug.WriteLine($"[SkillsService] 记录 Skill 使用: {skillName}, 累计: {stats.UsageCount}")
        Catch ex As Exception
            Debug.WriteLine($"[SkillsService] RecordSkillUsage 失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 记录 Skill 的真实执行结果，并把失败原因/使用场景沉淀为可检索记忆。
    ''' </summary>
    Public Shared Sub RecordSkillExecution(skillName As String,
                                           success As Boolean,
                                           Optional failureReason As String = "",
                                           Optional scenario As String = "",
                                           Optional tokensUsed As Long = 0)
        If String.IsNullOrWhiteSpace(skillName) Then Return

        RecordSkillUsage(skillName, success, tokensUsed)

        Try
            Dim scene = If(String.IsNullOrWhiteSpace(scenario), "未记录场景", scenario.Trim())
            Dim outcome = If(success, "成功", "失败")
            Dim content = $"Skill '{skillName}' 执行{outcome}。场景: {scene}"
            If Not success AndAlso Not String.IsNullOrWhiteSpace(failureReason) Then
                content &= $"。失败原因: {failureReason.Trim()}"
            End If

            MemoryRepository.InsertMemory(
                content,
                Nothing,
                Nothing,
                Nothing,
                If(success, "short_term", "long_term"),
                importance:=If(success, 0.45, 0.7),
                sourceType:="skill_execution"
            )
        Catch ex As Exception
            Debug.WriteLine($"[SkillsService] RecordSkillExecution memory failed: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 加载使用统计
    ''' </summary>
    Private Shared Sub LoadUsageStats()
        SyncLock _usageStorageLock
            If _usageStorage IsNot Nothing Then
                Return
            End If

            Try
                If File.Exists(UsageStatsPath) Then
                    Dim json = File.ReadAllText(UsageStatsPath)
                    _usageStorage = JsonConvert.DeserializeObject(Of SkillsUsageStorage)(json)
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SkillsService] LoadUsageStats 失败: {ex.Message}")
            End Try

            If _usageStorage Is Nothing Then
                _usageStorage = New SkillsUsageStorage()
            End If
        End SyncLock
    End Sub

    ''' <summary>
    ''' 保存使用统计
    ''' </summary>
    Private Shared Sub SaveUsageStats()
        SyncLock _usageStorageLock
            Try
                Dim dir = Path.GetDirectoryName(UsageStatsPath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                _usageStorage.LastUpdated = DateTime.Now
                Dim json = JsonConvert.SerializeObject(_usageStorage, Formatting.Indented)
                File.WriteAllText(UsageStatsPath, json)
            Catch ex As Exception
                Debug.WriteLine($"[SkillsService] SaveUsageStats 失败: {ex.Message}")
            End Try
        End SyncLock
    End Sub

    ''' <summary>
    ''' 构建渐进式披露的第一步：Skills目录（增强版）
    ''' </summary>
    Public Shared Function BuildSkillsCatalogMessage(skills As List(Of SkillFileDefinition), Optional maxItems As Integer = 20) As String
        If skills Is Nothing OrElse skills.Count = 0 Then
            Return ""
        End If

        ' 按使用次数排序
        Dim sortedSkills = skills.OrderByDescending(Function(s) s.UsageCount) _
                                 .ThenByDescending(Function(s) s.LastUsedAt.GetValueOrDefault()) _
                                 .ToList()
        Dim displayedSkills = sortedSkills.Take(Math.Max(1, maxItems)).ToList()

        Dim sb As New StringBuilder()
        sb.AppendLine("## Доступные Skills (каталог)")
        sb.AppendLine()
        sb.AppendLine($"Установлено Skills: {skills.Count}. Ниже — краткие метаданные не более {displayedSkills.Count} навыков; полное содержимое загружается только для выбранного.")
        sb.AppendLine()

        ' 分类显示：热门、最近使用、其他
        Dim hotSkills = displayedSkills.Where(Function(s) s.UsageCount >= 3).ToList()
        Dim recentSkills = displayedSkills.Where(Function(s) s.UsageCount < 3 AndAlso s.LastUsedAt.HasValue AndAlso (DateTime.Now - s.LastUsedAt.Value).TotalDays < 7).ToList()
        Dim otherSkills = displayedSkills.Where(Function(s) Not hotSkills.Contains(s) AndAlso Not recentSkills.Contains(s)).ToList()

        If hotSkills.Count > 0 Then
            sb.AppendLine("### 🔥 Популярные навыки")
            sb.AppendLine()
            For Each skill In hotSkills
                sb.AppendLine($"- **{skill.Name}**")
                If Not String.IsNullOrWhiteSpace(skill.Description) Then
                    sb.AppendLine($"  {skill.Description}")
                End If
                If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                    sb.AppendLine($"  *теги: {String.Join(", ", skill.Tags)}*")
                End If
                sb.AppendLine($"  *использований: {skill.UsageCount}*")
                sb.AppendLine()
            Next
        End If

        If recentSkills.Count > 0 Then
            sb.AppendLine("### ⏰ Недавно использованные")
            sb.AppendLine()
            For Each skill In recentSkills
                sb.AppendLine($"- **{skill.Name}**")
                If Not String.IsNullOrWhiteSpace(skill.Description) Then
                    sb.AppendLine($"  {skill.Description}")
                End If
                sb.AppendLine()
            Next
        End If

        If otherSkills.Count > 0 Then
            sb.AppendLine("### 📚 Все навыки")
            sb.AppendLine()
            For Each skill In otherSkills
                sb.AppendLine($"- **{skill.Name}**")
                If Not String.IsNullOrWhiteSpace(skill.Description) Then
                    sb.AppendLine($"  {skill.Description}")
                End If
                If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                    sb.AppendLine($"  *теги: {String.Join(", ", skill.Tags)}*")
                End If
                sb.AppendLine()
            Next
        End If

        sb.AppendLine("---")
        sb.AppendLine("**Как использовать**:")
        sb.AppendLine("1. По запросу пользователя выбери наиболее подходящий навык из списка выше")
        sb.AppendLine("2. Если нужно подробное содержимое навыка, явно укажи, какой навык требуется")
        sb.AppendLine("3. Можно использовать несколько навыков одновременно")
        If skills.Count > displayedSkills.Count Then
            sb.AppendLine($"4. Ещё {skills.Count - displayedSkills.Count} Skills не раскрыты; при необходимости вызови их по индексу.")
        End If

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建渐进式披露的第二步：披露选中的Skill详细内容
    ''' </summary>
    Public Shared Function BuildSkillDetailMessage(skill As SkillFileDefinition) As String
        If skill Is Nothing Then
            Return ""
        End If

        Dim detailSkill = SkillsDirectoryService.LoadSkillDetail(skill)
        If detailSkill IsNot Nothing Then skill = detailSkill

        Dim sb As New StringBuilder()
        sb.AppendLine($"## Skill: {skill.Name}")
        sb.AppendLine()

        If Not String.IsNullOrWhiteSpace(skill.Description) Then
            sb.AppendLine($"**Описание**: {skill.Description}")
            sb.AppendLine()
        End If

        If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
            sb.AppendLine($"**Теги**: {String.Join(", ", skill.Tags)}")
            sb.AppendLine()
        End If

        If skill.UsageCount > 0 Then
            sb.AppendLine($"**Статистика использования**: {skill.UsageCount}")
            If skill.LastUsedAt.HasValue Then
                sb.AppendLine($"**Последнее использование**: {skill.LastUsedAt.Value.ToString("yyyy-MM-dd HH:mm")}")
            End If
            sb.AppendLine()
        End If

        ' Skill详细内容
        sb.AppendLine("**Содержимое Skill**:")
        sb.AppendLine("```")
        sb.AppendLine(skill.Content)
        sb.AppendLine("```")

        Return sb.ToString()
    End Function

End Class
