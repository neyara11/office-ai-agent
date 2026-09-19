' ShareRibbon\Services\Reformat\FormattingStandardData.vb
' 内置排版标准数据

Imports System.Collections.Generic

''' <summary>
''' 内置排版标准数据提供器 - 创建各排版标准的 SemanticStyleMapping 数据
''' </summary>
Public Module FormattingStandardData

    ''' <summary>
    ''' 获取所有内置排版标准
    ''' </summary>
    Public Function GetAllBuiltInStandards() As List(Of FormattingStandard)
        Return New List(Of FormattingStandard) From {
            CreateGbt9704Standard(),
            CreateGbt7714Standard(),
            CreateAcademicStandard(),
            CreateBusinessStandard()
        }
    End Function

#Region "GB/T 9704-2012 党政机关公文格式"

    ''' <summary>
    ''' 创建 GB/T 9704-2012 党政机关公文格式标准
    ''' </summary>
    Private Function CreateGbt9704Standard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "gbt-9704-2012"
        standard.Name = "GB/T 9704-2012"
        standard.Description = "Национальный стандарт оформления официальных документов государственных органов, применяется к документам органов всех уровней. Содержит полные правила оформления страницы, обозначения органа, номера документа, заголовка, основного текста, приложений, подписи, даты документа, номера страницы."
        standard.ApplicableDocumentTypes.Add(DocumentType.OfficialDocument.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateGbt9704Mapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建 GB/T 9704-2012 的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateGbt9704Mapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "gbt-9704-2012"
        mapping.Name = "GB/T 9704-2012 Оформление официальных документов"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' ============ 页面设置 ============
        mapping.PageConfig.Margins = New MarginsConfig(3.7, 3.5, 2.8, 2.6)

        ' ============ 语义标签 ============
        Dim tags = mapping.SemanticTags

        ' 1. 发文机关标志
        Dim headerOrg As New SemanticTag("header.org", "Обозначение органа", "header", 2, "Наименование органа или его официальное сокращение")
        headerOrg.Font = New FontConfig("方正小标宋简体", "", 22, True)
        headerOrg.Paragraph = New ParagraphConfig("center")
        headerOrg.Color = New ColorConfig("#C00000")
        tags.Add(headerOrg)

        ' 2. 红色分隔线
        Dim headerSep As New SemanticTag("header.separator", "Красная разделительная линия", "header", 2, "Красная горизонтальная разделительная линия под номером документа")
        headerSep.Color = New ColorConfig("#C00000")
        tags.Add(headerSep)

        ' 3. 发文字号
        Dim headerRefno As New SemanticTag("header.refno", "Номер документа", "header", 2, "Номер документа, включающий код органа, год и порядковый номер, например ""×发〔2026〕×号""")
        headerRefno.Font = New FontConfig("仿宋_GB2312", "", 16)
        headerRefno.Paragraph = New ParagraphConfig("center")
        tags.Add(headerRefno)

        ' 4. 文件标题
        Dim titleMain As New SemanticTag("title.main", "Заголовок документа", "title", 2, "Главный заголовок документа, выравнивание по центру, при нескольких строках — ромбовидная раскладка")
        titleMain.Font = New FontConfig("方正小标宋简体", "", 22, True)
        titleMain.Paragraph = New ParagraphConfig("center")
        titleMain.Paragraph.SpaceAfter = 1
        tags.Add(titleMain)

        ' 5. 主送机关
        Dim titleRecipient As New SemanticTag("title.recipient", "Адресат", "title", 2, "Наименование адресата, выравнивание по левому краю без отступа, несколько адресатов разделяются запятыми")
        titleRecipient.Font = New FontConfig("仿宋_GB2312", "", 16)
        titleRecipient.Paragraph = New ParagraphConfig("left")
        tags.Add(titleRecipient)

        ' 6. 一级标题（黑体）
        Dim title1 As New SemanticTag("title.1", "Заголовок первого уровня", "title", 2, "Заголовок первого уровня основного текста документа, например ""一、""""二、"" и т. д.")
        title1.Font = New FontConfig("黑体", "", 16, False)
        title1.Paragraph = New ParagraphConfig("left", 0, 1.75)
        tags.Add(title1)

        ' 7. 二级标题（楷体）
        Dim title2 As New SemanticTag("title.2", "Заголовок второго уровня", "title", 2, "Заголовок второго уровня основного текста документа, например ""（一）""""（二）"" и т. д.")
        title2.Font = New FontConfig("楷体_GB2312", "", 16, False)
        title2.Paragraph = New ParagraphConfig("left", 0, 1.75)
        tags.Add(title2)

        ' 8. 三级标题（仿宋加粗）
        Dim title3 As New SemanticTag("title.3", "Заголовок третьего уровня", "title", 2, "Заголовок третьего уровня основного текста документа, например ""1.""""2."" и т. д.")
        title3.Font = New FontConfig("仿宋_GB2312", "", 16, True)
        title3.Paragraph = New ParagraphConfig("left", 0, 1.75)
        tags.Add(title3)

        ' 9. 正文
        Dim bodyNormal As New SemanticTag("body.normal", "Основной текст", "body", 2, "Абзацы основного текста документа, 22 строки на странице, 28 знаков в строке")
        bodyNormal.Font = New FontConfig("仿宋_GB2312", "", 16)
        bodyNormal.Paragraph = New ParagraphConfig("justify", 2, 1.75)
        tags.Add(bodyNormal)

        ' 10. 附件说明
        Dim bodyAttachment As New SemanticTag("body.attachment", "Отметка о приложении", "body", 2, "Отметка о приложении, например ""附件：1. ×××""")
        bodyAttachment.Font = New FontConfig("仿宋_GB2312", "", 16)
        bodyAttachment.Paragraph = New ParagraphConfig("left")
        bodyAttachment.Paragraph.SpaceBefore = 1
        tags.Add(bodyAttachment)

        ' 11. 发文机关署名
        Dim footerSignature As New SemanticTag("footer.signature", "Подпись органа", "footer", 2, "Подпись органа, располагается над датой документа")
        footerSignature.Font = New FontConfig("仿宋_GB2312", "", 16)
        footerSignature.Paragraph = New ParagraphConfig("right")
        tags.Add(footerSignature)

        ' 12. 成文日期
        Dim footerDate As New SemanticTag("footer.date", "Дата документа", "footer", 2, "Дата документа, арабскими цифрами, без ведущих нулей")
        footerDate.Font = New FontConfig("仿宋_GB2312", "", 16)
        footerDate.Paragraph = New ParagraphConfig("right")
        tags.Add(footerDate)

        ' 13. 签发人
        Dim headerSigner As New SemanticTag("header.signer", "Подписант", "header", 2, "Сведения о подписанте исходящего документа, например ""签发人：×××""")
        headerSigner.Font = New FontConfig("仿宋_GB2312", "", 16)
        headerSigner.Paragraph = New ParagraphConfig("right")
        tags.Add(headerSigner)

        ' 14. 附注
        Dim footerNote As New SemanticTag("footer.note", "Примечание", "footer", 2, "Примечание к документу, например контактное лицо и телефон, слева с отступом в два знака в круглых скобках")
        footerNote.Font = New FontConfig("仿宋_GB2312", "", 16)
        footerNote.Paragraph = New ParagraphConfig("left", 2, 1.75)
        tags.Add(footerNote)

        ' 15. 抄送机关
        Dim footerCc As New SemanticTag("footer.cc", "Рассылка копий", "footer", 2, "Сведения о рассылке копий в колонтитуле, начинаются с ""抄送：""")
        footerCc.Font = New FontConfig("仿宋_GB2312", "", 16)
        footerCc.Paragraph = New ParagraphConfig("left", 0, 1.0)
        tags.Add(footerCc)

        ' 16. 页码
        Dim footerPage As New SemanticTag("footer.page", "Номер страницы", "footer", 2, "Номер страницы документа, полуширинный шрифт SimSun, размер 4")
        footerPage.Font = New FontConfig("宋体", "", 14)
        footerPage.Paragraph = New ParagraphConfig("center")
        tags.Add(footerPage)

        ' ============ 版式骨架 ============
        Dim layout As New LayoutConfig()

        Dim el1 As New LayoutElement()
        el1.Name = "Обозначение органа"
        el1.ElementType = "text"
        el1.SortOrder = 1
        el1.Required = True
        el1.PlaceholderContent = "{{发文机关}}文件"
        layout.Elements.Add(el1)

        Dim el2 As New LayoutElement()
        el2.Name = "Красная линия"
        el2.ElementType = "redLine"
        el2.SortOrder = 2
        el2.Required = True
        el2.Color = New ColorConfig("#C00000")
        el2.SpecialProps("lineWidth") = "2pt"
        layout.Elements.Add(el2)

        Dim el3 As New LayoutElement()
        el3.Name = "Номер документа"
        el3.ElementType = "text"
        el3.SortOrder = 3
        el3.Required = True
        el3.PlaceholderContent = "×发〔2026〕×号"
        layout.Elements.Add(el3)

        Dim el4 As New LayoutElement()
        el4.Name = "Заголовок документа"
        el4.ElementType = "text"
        el4.SortOrder = 4
        el4.Required = True
        layout.Elements.Add(el4)

        Dim el5 As New LayoutElement()
        el5.Name = "Адресат"
        el5.ElementType = "text"
        el5.SortOrder = 5
        el5.Required = True
        el5.PlaceholderContent = "各有关单位："
        layout.Elements.Add(el5)

        mapping.LayoutSkeleton = layout

        Return mapping
    End Function

#End Region

#Region "GB/T 7714-2015 参考文献著录规则"

    ''' <summary>
    ''' 创建 GB/T 7714-2015 参考文献著录规则标准
    ''' </summary>
    Private Function CreateGbt7714Standard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "gbt-7714-2015"
        standard.Name = "GB/T 7714-2015"
        standard.Description = "Национальный стандарт описания библиографических ссылок в области информации и документации; устанавливает состав, порядок, знаки, язык и формат описания библиографических ссылок в научных работах и изданиях."
        standard.ApplicableDocumentTypes.Add(DocumentType.AcademicPaper.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateGbt7714Mapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建 GB/T 7714-2015 的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateGbt7714Mapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "gbt-7714-2015"
        mapping.Name = "GB/T 7714-2015 Правила описания библиографических ссылок"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' 语义标签
        Dim tags = mapping.SemanticTags

        ' 参考文献标题
        Dim titleRef As New SemanticTag("title.references", "Заголовок списка литературы", "title", 2, "Заголовок раздела списка литературы, например ""参考文献""")
        titleRef.Font = New FontConfig("黑体", "", 14, True)
        titleRef.Paragraph = New ParagraphConfig("left")
        titleRef.Paragraph.SpaceBefore = 1
        titleRef.Paragraph.SpaceAfter = 0.5
        tags.Add(titleRef)

        ' 参考文献条目
        Dim bodyRef As New SemanticTag("body.reference", "Запись списка литературы", "body", 2, "Отдельная библиографическая запись в списке литературы")
        bodyRef.Font = New FontConfig("宋体", "Times New Roman", 9)
        bodyRef.Paragraph = New ParagraphConfig("justify")
        bodyRef.Paragraph.LineSpacing = 1.25
        tags.Add(bodyRef)

        Return mapping
    End Function

#End Region

#Region "学术论文通用格式"

    ''' <summary>
    ''' 创建学术论文通用格式标准（基于 GB/T 7713.1）
    ''' </summary>
    Private Function CreateAcademicStandard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "academic-general"
        standard.Name = "Общий формат научной работы"
        standard.Description = "Общие правила оформления научных статей и диссертаций. Содержит стандартное оформление заголовка, аннотации, ключевых слов, заголовков разделов, основного текста, списка литературы и других элементов."
        standard.ApplicableDocumentTypes.Add(DocumentType.AcademicPaper.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateAcademicMapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建学术论文通用格式的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateAcademicMapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "academic-general"
        mapping.Name = "Общий формат научной работы"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' 页面设置
        mapping.PageConfig.Margins = New MarginsConfig(2.54, 2.54, 3.18, 3.18)

        ' 语义标签
        Dim tags = mapping.SemanticTags

        ' 论文标题
        Dim titleMain As New SemanticTag("title.main", "Заголовок научной работы", "title", 2, "Главный заголовок научной работы")
        titleMain.Font = New FontConfig("黑体", "Times New Roman", 18, True)
        titleMain.Paragraph = New ParagraphConfig("center", 0, 1.5)
        titleMain.Paragraph.SpaceBefore = 2
        titleMain.Paragraph.SpaceAfter = 1
        tags.Add(titleMain)

        ' 摘要标题
        Dim titleAbstract As New SemanticTag("title.abstract", "Заголовок аннотации", "title", 2, """摘要"" строка заголовка")
        titleAbstract.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        titleAbstract.Paragraph = New ParagraphConfig("left")
        titleAbstract.Paragraph.SpaceBefore = 1
        titleAbstract.Paragraph.SpaceAfter = 0.5
        tags.Add(titleAbstract)

        ' 摘要正文
        Dim bodyAbstract As New SemanticTag("body.abstract", "Текст аннотации", "body", 2, "Абзац текста аннотации")
        bodyAbstract.Font = New FontConfig("宋体", "Times New Roman", 12)
        bodyAbstract.Paragraph = New ParagraphConfig("justify", 2, 1.5)
        tags.Add(bodyAbstract)

        ' 关键词标题
        Dim titleKeywords As New SemanticTag("title.keywords", "Заголовок ключевых слов", "title", 2, """Ключевые слова"" строка-маркер")
        titleKeywords.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        titleKeywords.Paragraph = New ParagraphConfig("left")
        titleKeywords.Paragraph.SpaceBefore = 0.5
        tags.Add(titleKeywords)

        ' 关键词
        Dim bodyKeywords As New SemanticTag("body.keywords", "Ключевые слова", "body", 2, "Список ключевых слов, разделённых точкой с запятой")
        bodyKeywords.Font = New FontConfig("宋体", "Times New Roman", 12)
        bodyKeywords.Paragraph = New ParagraphConfig("left")
        bodyKeywords.Paragraph.SpaceAfter = 1
        tags.Add(bodyKeywords)

        ' 一级标题（章）
        Dim heading1 As New SemanticTag("heading.1", "Заголовок первого уровня", "", 2, "Заголовок главы работы, например ""第1章 引言""")
        heading1.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        heading1.Paragraph = New ParagraphConfig("left")
        heading1.Paragraph.SpaceBefore = 1
        heading1.Paragraph.SpaceAfter = 0.5
        tags.Add(heading1)

        ' 二级标题（节）
        Dim heading2 As New SemanticTag("heading.2", "Заголовок второго уровня", "", 2, "Заголовок параграфа работы, например ""1.1 研究背景""")
        heading2.Font = New FontConfig("黑体", "Times New Roman", 12, True)
        heading2.Paragraph = New ParagraphConfig("left")
        heading2.Paragraph.SpaceBefore = 0.5
        heading2.Paragraph.SpaceAfter = 0.25
        tags.Add(heading2)

        ' 正文
        Dim bodyNormal As New SemanticTag("body.normal", "Основной текст", "body", 2, "Абзацы основного текста научной работы")
        bodyNormal.Font = New FontConfig("宋体", "Times New Roman", 12)
        bodyNormal.Paragraph = New ParagraphConfig("justify", 2, 1.5)
        tags.Add(bodyNormal)

        ' 参考文献标题
        Dim titleRef As New SemanticTag("title.references", "Заголовок списка литературы", "title", 2, """参考文献"" заголовок раздела")
        titleRef.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        titleRef.Paragraph = New ParagraphConfig("left")
        titleRef.Paragraph.SpaceBefore = 1
        titleRef.Paragraph.SpaceAfter = 0.5
        tags.Add(titleRef)

        ' 参考文献条目
        Dim bodyRef As New SemanticTag("body.reference", "Запись списка литературы", "body", 2, "Отдельная запись в списке литературы")
        bodyRef.Font = New FontConfig("宋体", "Times New Roman", 10)
        bodyRef.Paragraph = New ParagraphConfig("justify")
        bodyRef.Paragraph.LineSpacing = 1.25
        tags.Add(bodyRef)

        ' 页码
        Dim footerPage As New SemanticTag("footer.page", "Номер страницы", "footer", 2, "Номер страницы научной работы")
        footerPage.Font = New FontConfig("宋体", "Times New Roman", 10)
        footerPage.Paragraph = New ParagraphConfig("center")
        tags.Add(footerPage)

        Return mapping
    End Function

#End Region

#Region "商务报告通用规范"

    ''' <summary>
    ''' 创建商务报告通用规范标准
    ''' </summary>
    Private Function CreateBusinessStandard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "business-report"
        standard.Name = "Общие правила оформления бизнес-отчётов"
        standard.Description = "Общие правила оформления деловых документов: бизнес-отчётов, бизнес-планов, отчётов о работе. Используется современный деловой стиль вёрстки."
        standard.ApplicableDocumentTypes.Add(DocumentType.BusinessReport.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateBusinessMapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建商务报告通用规范的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateBusinessMapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "business-report"
        mapping.Name = "Общие правила оформления бизнес-отчётов"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' 页面设置（商务报告通常使用稍宽的页边距）
        mapping.PageConfig.Margins = New MarginsConfig(2.54, 2.54, 3.18, 3.18)

        ' 语义标签
        Dim tags = mapping.SemanticTags

        ' 报告标题
        Dim titleMain As New SemanticTag("title.main", "Заголовок отчёта", "title", 2, "Главный заголовок бизнес-отчёта")
        titleMain.Font = New FontConfig("微软雅黑", "Arial", 20, True)
        titleMain.Paragraph = New ParagraphConfig("center", 0, 1.5)
        titleMain.Paragraph.SpaceBefore = 3
        titleMain.Paragraph.SpaceAfter = 1
        tags.Add(titleMain)

        ' 一级标题
        Dim heading1 As New SemanticTag("heading.1", "Заголовок первого уровня", "", 2, "Заголовок раздела отчёта")
        heading1.Font = New FontConfig("微软雅黑", "Arial", 16, True)
        heading1.Paragraph = New ParagraphConfig("left")
        heading1.Paragraph.SpaceBefore = 1.5
        heading1.Paragraph.SpaceAfter = 0.5
        tags.Add(heading1)

        ' 二级标题
        Dim heading2 As New SemanticTag("heading.2", "Заголовок второго уровня", "", 2, "Заголовок подраздела отчёта")
        heading2.Font = New FontConfig("微软雅黑", "Arial", 14, True)
        heading2.Paragraph = New ParagraphConfig("left")
        heading2.Paragraph.SpaceBefore = 1
        heading2.Paragraph.SpaceAfter = 0.25
        tags.Add(heading2)

        ' 正文
        Dim bodyNormal As New SemanticTag("body.normal", "Основной текст", "body", 2, "Абзацы основного текста бизнес-отчёта")
        bodyNormal.Font = New FontConfig("微软雅黑", "Arial", 11)
        bodyNormal.Paragraph = New ParagraphConfig("justify", 0, 1.25)
        tags.Add(bodyNormal)

        ' 摘要/总结段落
        Dim bodySummary As New SemanticTag("body.summary", "Аннотация", "body", 2, "Абзац аннотации или краткого содержания отчёта")
        bodySummary.Font = New FontConfig("微软雅黑", "Arial", 11)
        bodySummary.Paragraph = New ParagraphConfig("justify", 0, 1.25)
        bodySummary.Paragraph.SpaceBefore = 1
        bodySummary.Paragraph.SpaceAfter = 1
        tags.Add(bodySummary)

        ' 页码
        Dim footerPage As New SemanticTag("footer.page", "Номер страницы", "footer", 2, "Номер страницы отчёта")
        footerPage.Font = New FontConfig("微软雅黑", "Arial", 9)
        footerPage.Paragraph = New ParagraphConfig("center")
        tags.Add(footerPage)

        Return mapping
    End Function

#End Region

End Module
