' ShareRibbon\Services\Reformat\ScenarioManager.vb
' 场景化提示词管理器 - 从 JSON 文件加载排版场景

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Reformat

    ''' <summary>
    ''' 排版场景定义
    ''' </summary>
    Public Class FormattingScenario
        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property TriggerKeywords As List(Of String)
        Public Property IdentificationPatterns As JObject
        Public Property FormatSpecification As JObject
        Public Property StructureGuidance As StructureGuidance
        Public Property Examples As List(Of ExampleAnnotation)
    End Class

    ''' <summary>
    ''' 结构指导
    ''' </summary>
    Public Class StructureGuidance
        Public Property Title As String
        Public Property Order As List(Of String)
        Public Property Rules As List(Of String)
    End Class

    ''' <summary>
    ''' 示例标注
    ''' </summary>
    Public Class ExampleAnnotation
        Public Property ParaIndex As Integer
        Public Property Text As String
        Public Property Tag As String
        Public Property Reason As String
    End Class

    ''' <summary>
    ''' 场景管理器 - 加载和匹配排版场景
    ''' </summary>
    Public Class ScenarioManager

        Private ReadOnly _scenarios As New Dictionary(Of String, FormattingScenario)()
        Private ReadOnly _scenarioDir As String

        Public Sub New()
            ' 场景目录：ShareRibbon/Prompts/Scenarios/
            Dim baseDir = Path.GetDirectoryName(GetType(ScenarioManager).Assembly.Location)
            _scenarioDir = Path.Combine(baseDir, "Prompts", "Scenarios")
            LoadScenarios()
        End Sub

        ''' <summary>
        ''' 从目录加载所有场景
        ''' </summary>
        Private Sub LoadScenarios()
            If Not Directory.Exists(_scenarioDir) Then
                Debug.WriteLine($"[ScenarioManager] 场景目录不存在: {_scenarioDir}")
                Return
            End If

            For Each file In Directory.GetFiles(_scenarioDir, "*.json")
                Try
                    Dim json = IO.File.ReadAllText(file, Text.Encoding.UTF8)
                    Dim scenario = JsonConvert.DeserializeObject(Of FormattingScenario)(json)
                    If scenario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(scenario.Id) Then
                        _scenarios(scenario.Id) = scenario
                        Debug.WriteLine($"[ScenarioManager] 加载场景: {scenario.Name} ({scenario.Id})")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ScenarioManager] 加载场景失败 {file}: {ex.Message}")
                End Try
            Next

            Debug.WriteLine($"[ScenarioManager] 共加载 {_scenarios.Count} 个场景")
        End Sub

        ''' <summary>
        ''' 根据文档类型上下文匹配场景
        ''' </summary>
        ''' <param name="documentTypeContext">文档类型描述，如"公文"、"论文"</param>
        Public Function MatchScenario(documentTypeContext As String) As FormattingScenario
            If String.IsNullOrWhiteSpace(documentTypeContext) Then
                Return Nothing
            End If

            Dim contextLower = documentTypeContext.ToLower()

            ' 遍历所有场景，匹配关键词
            For Each scenario In _scenarios.Values
                If scenario.TriggerKeywords IsNot Nothing Then
                    For Each keyword In scenario.TriggerKeywords
                        If contextLower.Contains(keyword.ToLower()) Then
                            Debug.WriteLine($"[ScenarioManager] 匹配场景: {scenario.Name} (关键词: {keyword})")
                            Return scenario
                        End If
                    Next
                End If
            Next

            Debug.WriteLine($"[ScenarioManager] 未匹配到场景: {documentTypeContext}")
            Return Nothing
        End Function

        ''' <summary>
        ''' 构建场景的结构指导文本
        ''' </summary>
        Public Shared Function BuildStructureGuidanceText(scenario As FormattingScenario) As String
            If scenario Is Nothing OrElse scenario.StructureGuidance Is Nothing Then
                Return String.Empty
            End If

            Dim sb As New Text.StringBuilder()

            ' 标题
            If Not String.IsNullOrEmpty(scenario.StructureGuidance.Title) Then
                sb.AppendLine(scenario.StructureGuidance.Title)
            End If

            ' 结构顺序
            If scenario.StructureGuidance.Order IsNot Nothing AndAlso scenario.StructureGuidance.Order.Count > 0 Then
                sb.AppendLine($"{scenario.Name}: фиксированный порядок структуры, распознавай в этом порядке:")
                sb.AppendLine(String.Join(" → ", scenario.StructureGuidance.Order))
                sb.AppendLine()
            End If

            ' 规则
            If scenario.StructureGuidance.Rules IsNot Nothing AndAlso scenario.StructureGuidance.Rules.Count > 0 Then
                sb.AppendLine("Внимание:")
                For Each rule In scenario.StructureGuidance.Rules
                    sb.AppendLine($"- {rule}")
                Next
                sb.AppendLine()
            End If

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 构建场景的示例文本
        ''' </summary>
        Public Shared Function BuildExamplesText(scenario As FormattingScenario) As String
            If scenario Is Nothing OrElse scenario.Examples Is Nothing OrElse scenario.Examples.Count = 0 Then
                Return String.Empty
            End If

            Dim sb As New Text.StringBuilder()
            sb.AppendLine($"【Пример разметки: {scenario.Name}】")

            For Each example In scenario.Examples
                sb.AppendLine($"段落{example.ParaIndex}：「{example.Text}」")
                sb.AppendLine($"  → 标签：{example.Tag}  原因：{example.Reason}")
            Next

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 构建场景的识别模式文本（详细的模式匹配规则）
        ''' </summary>
        Public Shared Function BuildIdentificationPatternsText(scenario As FormattingScenario) As String
            If scenario Is Nothing OrElse scenario.IdentificationPatterns Is Nothing Then
                Return String.Empty
            End If

            Dim sb As New Text.StringBuilder()
            sb.AppendLine($"【{scenario.Name}识别模式】")
            sb.AppendLine("Ниже — подробные признаки распознавания для каждого типа абзацев; используй их в совокупности:")
            sb.AppendLine()

            For Each kvp In scenario.IdentificationPatterns
                Dim elementName As String = kvp.Key
                Dim pattern As JObject = CType(kvp.Value, JObject)

                sb.AppendLine($"**{elementName}**")

                ' 位置
                If pattern("position") IsNot Nothing Then
                    sb.AppendLine($"  位置：{pattern("position")}")
                End If

                ' 模式
                If pattern("patterns") IsNot Nothing Then
                    sb.AppendLine($"  匹配模式：")
                    For Each p In CType(pattern("patterns"), JArray)
                        sb.AppendLine($"    - {p}")
                    Next
                End If

                ' 特征
                If pattern("features") IsNot Nothing Then
                    sb.AppendLine($"  识别特征：")
                    For Each f In CType(pattern("features"), JArray)
                        sb.AppendLine($"    ✓ {f}")
                    Next
                End If

                ' 示例
                If pattern("examples") IsNot Nothing Then
                    sb.Append($"  示例：")
                    Dim examples As New List(Of String)
                    For Each e In CType(pattern("examples"), JArray)
                        examples.Add($"「{e}」")
                    Next
                    sb.AppendLine(String.Join("、", examples))
                End If

                sb.AppendLine()
            Next

            Return sb.ToString()
        End Function

    End Class

End Namespace
