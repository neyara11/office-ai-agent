' ShareRibbon\Services\Reformat\LegacyTemplateConverter.vb
' 旧模板 → SemanticStyleMapping 转换器

''' <summary>
''' 旧模板转换器 - 将现有ReformatTemplate转换为SemanticStyleMapping
''' 提供向后兼容能力
''' </summary>
Public Class LegacyTemplateConverter

    ''' <summary>
    ''' 将ReformatTemplate转换为SemanticStyleMapping
    ''' </summary>
    Public Shared Function Convert(template As ReformatTemplate) As SemanticStyleMapping
        If template Is Nothing Then Return Nothing

        ' 先检查缓存
        Dim cached = SemanticMappingManager.Instance.GetMappingBySourceId(template.Id)
        If cached IsNot Nothing Then Return cached

        Dim mapping As New SemanticStyleMapping()
        mapping.Name = template.Name
        mapping.SourceType = SemanticMappingSourceType.FromLegacy
        mapping.SourceId = template.Id

        ' 转换正文样式规则
        ConvertBodyStyles(template.BodyStyles, mapping)

        ' 转换版式骨架
        ConvertLayout(template.Layout, mapping)

        ' 转换页面设置（直接复用）
        If template.PageSettings IsNot Nothing Then
            mapping.PageConfig = template.PageSettings
        End If

        ' 确保基础标签
        EnsureBasicTags(mapping)

        ' 公文场景：添加公文特有的语义标签
        If template.Category = "公文" OrElse template.Name.Contains("公文") OrElse
           template.Name.IndexOf("приказ", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           template.Name.IndexOf("официальн", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           template.Name.IndexOf("официальный документ", StringComparison.OrdinalIgnoreCase) >= 0 Then
            AddOfficialDocumentTags(mapping)
        End If

        ' 缓存转换结果
        SemanticMappingManager.Instance.AddMapping(mapping)

        Return mapping
    End Function

    ''' <summary>转换正文样式规则为语义标签</summary>
    Private Shared Sub ConvertBodyStyles(styles As List(Of StyleRule), mapping As SemanticStyleMapping)
        If styles Is Nothing Then Return

        For Each rule In styles
            Dim tag = MapRuleToTag(rule)
            If tag IsNot Nothing Then
                mapping.SemanticTags.Add(tag)
            End If
        Next
    End Sub

    ''' <summary>将StyleRule映射到SemanticTag</summary>
    Private Shared Function MapRuleToTag(rule As StyleRule) As SemanticTag
        If rule Is Nothing Then Return Nothing

        Dim tagId As String = ""
        Dim displayName As String = rule.RuleName
        Dim parentId As String = ""
        Dim matchHint As String = rule.MatchCondition

        Dim name = If(rule.RuleName, "").ToLower()

        Select Case True
            Case name.Contains("一级") OrElse name.Contains("大标题") OrElse name.Contains("章标题") OrElse
                 name.Contains("заголовок 1") OrElse name.Contains("глава") OrElse name.Contains("уровень 1")
                tagId = SemanticTagRegistry.TAG_TITLE_1
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("二级") OrElse name.Contains("节标题") OrElse
                 name.Contains("заголовок 2") OrElse name.Contains("раздел") OrElse name.Contains("уровень 2")
                tagId = SemanticTagRegistry.TAG_TITLE_2
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("三级") OrElse name.Contains("小标题") OrElse
                 name.Contains("заголовок 3") OrElse name.Contains("подзаголовок") OrElse name.Contains("уровень 3")
                tagId = SemanticTagRegistry.TAG_TITLE_3
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("正文") OrElse name.Contains("body") OrElse
                 name.Contains("основной текст") OrElse name.Contains("обычный") OrElse name.Contains("текст")
                tagId = SemanticTagRegistry.TAG_BODY_NORMAL
                parentId = SemanticTagRegistry.TAG_BODY

            Case name.Contains("强调") OrElse name.Contains("emphasis") OrElse name.Contains("акцент")
                tagId = SemanticTagRegistry.TAG_BODY_EMPHASIS
                parentId = SemanticTagRegistry.TAG_BODY

            Case name.Contains("列表") OrElse name.Contains("list") OrElse name.Contains("список")
                tagId = SemanticTagRegistry.TAG_LIST_ORDERED
                parentId = SemanticTagRegistry.TAG_LIST

            Case name.Contains("引用") OrElse name.Contains("quote") OrElse name.Contains("цитата")
                tagId = SemanticTagRegistry.TAG_QUOTE
                parentId = ""

            Case name.Contains("题注") OrElse name.Contains("caption") OrElse
                 name.Contains("название") OrElse name.Contains("подпись")
                tagId = SemanticTagRegistry.TAG_CAPTION
                parentId = ""

            Case Else
                ' 无法识别的规则作为正文处理
                tagId = SemanticTagRegistry.TAG_BODY_NORMAL
                parentId = SemanticTagRegistry.TAG_BODY
                matchHint = $"Из старого правила: {rule.RuleName}"
        End Select

        ' 避免重复
        Dim tag As New SemanticTag(tagId, displayName, parentId, SemanticTagRegistry.GetTagLevel(tagId), matchHint)

        ' 复制格式配置
        If rule.Font IsNot Nothing Then tag.Font = rule.Font
        If rule.Paragraph IsNot Nothing Then tag.Paragraph = rule.Paragraph
        If rule.Color IsNot Nothing Then tag.Color = rule.Color

        Return tag
    End Function

    ''' <summary>转换版式骨架</summary>
    Private Shared Sub ConvertLayout(layout As LayoutConfig, mapping As SemanticStyleMapping)
        If layout Is Nothing OrElse layout.Elements Is Nothing Then Return
        mapping.LayoutSkeleton = layout
    End Sub

    ''' <summary>确保基础标签存在</summary>
    Private Shared Sub EnsureBasicTags(mapping As SemanticStyleMapping)
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_BODY_NORMAL) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_BODY_NORMAL, "Основной текст",
                SemanticTagRegistry.TAG_BODY, 2, "Обычный абзац основного текста"))
        End If

        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_TITLE_1) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_TITLE_1, "Заголовок 1",
                SemanticTagRegistry.TAG_TITLE, 2, "Заголовок основного раздела"))
        End If
    End Sub

    ''' <summary>添加公文特有的语义标签</summary>
    Private Shared Sub AddOfficialDocumentTags(mapping As SemanticStyleMapping)
        ' 发文机关标志（红色大字）
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "header.org") Then
            Dim tag = New SemanticTag("header.org", "Бланк органа власти", "header", 2,
                "Полное наименование органа и слово 'документ', например 'Правительство РФ', обычно по центру красным крупным шрифтом")
            tag.Font = New FontConfig("方正小标宋简体", "Arial", 22, True)
            tag.Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceBefore = 2}
            tag.Color = New ColorConfig("#C00000")
            mapping.SemanticTags.Add(tag)
        End If

        ' 发文字号
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "header.refno") Then
            Dim tag = New SemanticTag("header.refno", "Номер исходящего документа", "header", 2,
                "Регистрационный номер документа, например '№ 15 от 2024 г.', обычно под красной разделительной линией")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("center", 0, 1.0)
            mapping.SemanticTags.Add(tag)
        End If

        ' 签发人（上行文）
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "header.signer") Then
            Dim tag = New SemanticTag("header.signer", "Подписывающий", "header", 2,
                "Сведения о подписывающем лице, только для исходящих документов, формат 'Подписал: ...', обычно по правому краю")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("right", 0, 1.0) With {.SpaceBefore = 0.5}
            mapping.SemanticTags.Add(tag)
        End If

        ' 文件标题
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "title.main") Then
            Dim tag = New SemanticTag("title.main", "Заголовок документа", "title", 2,
                "Основной заголовок документа, например 'О мерах по обеспечению безопасности', обычно по центру")
            tag.Font = New FontConfig("方正小标宋简体", "Arial", 22, True)
            tag.Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceBefore = 1.5, .SpaceAfter = 1.5}
            mapping.SemanticTags.Add(tag)
        End If

        ' 主送机关
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "title.recipient") Then
            Dim tag = New SemanticTag("title.recipient", "Адресат", "title", 2,
                "Основной адресат документа, например 'Руководителям региональных органов:', обычно без отступа, по левому краю")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("left", 0, 1.0)
            mapping.SemanticTags.Add(tag)
        End If

        ' 附件说明
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "body.attachment") Then
            Dim tag = New SemanticTag("body.attachment", "Отметка о приложении", "body", 2,
                "Абзац с перечнем приложений, формат 'Приложение: 1. ...', обычно с отступом первой строки")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("left", 2, 1.875) With {.SpaceBefore = 1}
            mapping.SemanticTags.Add(tag)
        End If

        ' 发文机关署名
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "footer.signature") Then
            Dim tag = New SemanticTag("footer.signature", "Подпись органа", "footer", 2,
                "Наименование органа в конце документа, например 'Правительство РФ', обычно по правому краю")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("right", 2, 1.0) With {.SpaceBefore = 2}
            mapping.SemanticTags.Add(tag)
        End If

        ' 成文日期
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "footer.date") Then
            Dim tag = New SemanticTag("footer.date", "Дата документа", "footer", 2,
                "Дата подписания документа, формат '15.01.2024', обычно по правому краю")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16)
            tag.Paragraph = New ParagraphConfig("right", 0, 1.0)
            mapping.SemanticTags.Add(tag)
        End If

        ' 抄送
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = "footer.cc") Then
            Dim tag = New SemanticTag("footer.cc", "Копия (рассылка)", "footer", 2,
                "Сведения о рассылке копий, формат 'Копия: ...'")
            tag.Font = New FontConfig("仿宋_GB2312", "Times New Roman", 14)
            tag.Paragraph = New ParagraphConfig("left", 2, 1.0)
            mapping.SemanticTags.Add(tag)
        End If
    End Sub
End Class
