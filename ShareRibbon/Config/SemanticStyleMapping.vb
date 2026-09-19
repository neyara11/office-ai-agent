' ShareRibbon\Config\SemanticStyleMapping.vb
' 语义排版映射数据模型 - 统一中间格式

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

''' <summary>
''' 语义排版映射（核心中间格式）
''' 模板排版和规范排版最终都转换为此格式
''' </summary>
Public Class SemanticStyleMapping
    ''' <summary>映射ID（唯一标识）</summary>
    Public Property Id As String = Guid.NewGuid().ToString()

    ''' <summary>映射名称</summary>
    Public Property Name As String = ""

    ''' <summary>来源类型</summary>
    <JsonConverter(GetType(StringEnumConverter))>
    Public Property SourceType As SemanticMappingSourceType = SemanticMappingSourceType.FromLegacy

    ''' <summary>来源ID（关联的模板或规范ID）</summary>
    Public Property SourceId As String = ""

    ''' <summary>原始文件路径（FromDocxTemplate时为拷贝到数据目录的.docx文件路径）</summary>
    Public Property SourceFilePath As String = ""

    ''' <summary>语义标签集合</summary>
    Public Property SemanticTags As List(Of SemanticTag)

    ''' <summary>页面设置（复用现有PageConfig）</summary>
    Public Property PageConfig As PageConfig

    ''' <summary>版式骨架（复用现有LayoutConfig结构）</summary>
    Public Property LayoutSkeleton As LayoutConfig

    ''' <summary>创建时间</summary>
    Public Property CreatedAt As DateTime = DateTime.Now

    ''' <summary>最后修改时间</summary>
    Public Property LastModified As DateTime = DateTime.Now

    Public Sub New()
        SemanticTags = New List(Of SemanticTag)()
        PageConfig = New PageConfig()
        LayoutSkeleton = New LayoutConfig()
    End Sub

    ''' <summary>
    ''' 根据TagId查找语义标签，找不到则回退到父级标签
    ''' </summary>
    Public Function FindTag(tagId As String) As SemanticTag
        ' 精确匹配
        Dim tag = SemanticTags.FirstOrDefault(Function(t) t.TagId = tagId)
        If tag IsNot Nothing Then Return tag

        ' 回退到父级
        Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
        If Not String.IsNullOrEmpty(parentId) Then
            Return SemanticTags.FirstOrDefault(Function(t) t.TagId = parentId)
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' 获取所有可用的TagId列表
    ''' </summary>
    Public Function GetAvailableTagIds() As List(Of String)
        Return SemanticTags.Select(Function(t) t.TagId).ToList()
    End Function

    ''' <summary>
    ''' 补齐 AI Native 排版流程必须可用的基础语义标签。
    ''' 用户导入模板或通用文档没有完整 mapping 时，仍要让 AI 标注和渲染有稳定落点。
    ''' </summary>
    Public Sub EnsureBaselineTags()
        If SemanticTags Is Nothing Then SemanticTags = New List(Of SemanticTag)()

        AddMissingTag(CreateDefaultBodyNormalTag())
        AddMissingTag(CreateDefaultTitleTag(SemanticTagRegistry.TAG_TITLE_1, "Заголовок 1 уровня", 16, True))
        AddMissingTag(CreateDefaultTitleTag(SemanticTagRegistry.TAG_TITLE_2, "Заголовок 2 уровня", 14, True))
        AddMissingTag(CreateDefaultTitleTag(SemanticTagRegistry.TAG_TITLE_3, "Заголовок 3 уровня", 12, True))
        AddMissingTag(CreateDefaultHeadingTag("heading.1", "Заголовок раздела 1 уровня", 16, True))
        AddMissingTag(CreateDefaultHeadingTag("heading.2", "Заголовок раздела 2 уровня", 14, True))
        AddMissingTag(CreateDefaultHeadingTag("heading.3", "Заголовок раздела 3 уровня", 12, True))
        AddMissingTag(CreateDefaultListTag(SemanticTagRegistry.TAG_LIST_ORDERED, "Нумерованный список"))
        AddMissingTag(CreateDefaultListTag(SemanticTagRegistry.TAG_LIST_UNORDERED, "Маркированный список"))
    End Sub

    Private Sub AddMissingTag(tag As SemanticTag)
        If tag Is Nothing OrElse String.IsNullOrWhiteSpace(tag.TagId) Then Return
        If SemanticTags.Any(Function(t) String.Equals(t.TagId, tag.TagId, StringComparison.OrdinalIgnoreCase)) Then Return
        SemanticTags.Add(tag)
    End Sub

    Private Shared Function CreateDefaultBodyNormalTag() As SemanticTag
        Dim tag As New SemanticTag(SemanticTagRegistry.TAG_BODY_NORMAL, "Основной текст", SemanticTagRegistry.TAG_BODY, 2, "Обычный абзац основного текста")
        tag.Font = New FontConfig("宋体", "Times New Roman", 12)
        tag.Paragraph = New ParagraphConfig("justify", 2, 1.5)
        tag.Color = New ColorConfig("#000000")
        Return tag
    End Function

    Private Shared Function CreateDefaultTitleTag(tagId As String, displayName As String, fontSize As Double, bold As Boolean) As SemanticTag
        Dim tag As New SemanticTag(tagId, displayName, SemanticTagRegistry.TAG_TITLE, 2, displayName & ": используется для заголовка документа или нумерованного заголовка")
        tag.Font = New FontConfig("黑体", "Times New Roman", fontSize, bold)
        tag.Paragraph = New ParagraphConfig("left", 0, 1.5)
        tag.Paragraph.SpaceBefore = 0.5
        tag.Paragraph.SpaceAfter = 0.25
        tag.Paragraph.KeepWithNext = True
        tag.Color = New ColorConfig("#000000")
        Return tag
    End Function

    Private Shared Function CreateDefaultHeadingTag(tagId As String, displayName As String, fontSize As Double, bold As Boolean) As SemanticTag
        Dim tag As New SemanticTag(tagId, displayName, SemanticTagRegistry.TAG_HEADING, 2, displayName & ": используется для структуры разделов документа")
        tag.Font = New FontConfig("黑体", "Times New Roman", fontSize, bold)
        tag.Paragraph = New ParagraphConfig("left", 0, 1.5)
        tag.Paragraph.SpaceBefore = 0.5
        tag.Paragraph.SpaceAfter = 0.25
        tag.Paragraph.KeepWithNext = True
        tag.Color = New ColorConfig("#000000")
        Return tag
    End Function

    Private Shared Function CreateDefaultListTag(tagId As String, displayName As String) As SemanticTag
        Dim tag As New SemanticTag(tagId, displayName, SemanticTagRegistry.TAG_LIST, 2, displayName & ": используется для элементов списка с нумерацией или маркерами")
        tag.Font = New FontConfig("宋体", "Times New Roman", 12)
        tag.Paragraph = New ParagraphConfig("left", 0, 1.5)
        tag.Paragraph.LeftIndent = 0.75
        tag.Color = New ColorConfig("#000000")
        Return tag
    End Function
End Class

''' <summary>
''' 语义标签 - 定义一种语义类型的完整格式规则
''' </summary>
Public Class SemanticTag
    ''' <summary>标签ID（如 "title.1", "body.normal"）</summary>
    Public Property TagId As String = ""

    ''' <summary>显示名称（如 "一级标题", "正文"）</summary>
    Public Property DisplayName As String = ""

    ''' <summary>父级标签ID（如 "title", "body"）</summary>
    Public Property ParentTagId As String = ""

    ''' <summary>标签层级（1=固定语义层, 2=模板细分层）</summary>
    Public Property Level As Integer = 2

    ''' <summary>给AI的匹配提示（如 "包含'第X章'开头"）</summary>
    Public Property MatchHint As String = ""

    ''' <summary>字体配置</summary>
    Public Property Font As FontConfig

    ''' <summary>段落配置</summary>
    Public Property Paragraph As ParagraphConfig

    ''' <summary>颜色配置</summary>
    Public Property Color As ColorConfig

    Public Sub New()
        Font = New FontConfig()
        Paragraph = New ParagraphConfig()
        Color = New ColorConfig()
    End Sub

    Public Sub New(tagId As String, displayName As String, parentTagId As String,
                   Optional level As Integer = 2, Optional matchHint As String = "")
        Me.TagId = tagId
        Me.DisplayName = displayName
        Me.ParentTagId = parentTagId
        Me.Level = level
        Me.MatchHint = matchHint
        Font = New FontConfig()
        Paragraph = New ParagraphConfig()
        Color = New ColorConfig()
    End Sub

    Public Overrides Function ToString() As String
        Return $"[{TagId}] {DisplayName}"
    End Function
End Class

''' <summary>
''' 语义映射来源类型
''' </summary>
Public Enum SemanticMappingSourceType
    ''' <summary>从.docx模板提取</summary>
    FromDocxTemplate = 0
    ''' <summary>从文本规范转换</summary>
    FromStyleGuide = 1
    ''' <summary>从旧格式模板转换</summary>
    FromLegacy = 2
End Enum
