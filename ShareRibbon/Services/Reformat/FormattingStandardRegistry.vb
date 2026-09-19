Imports System.Linq

Public Enum FormattingStandardSourceType
    BuiltIn
    Template
    StyleGuide
    DocxMapping
End Enum

Public Class FormattingStandardCandidate
    Public Property Standard As FormattingStandard
    Public Property SourceType As FormattingStandardSourceType = FormattingStandardSourceType.BuiltIn
    Public Property SourceId As String = ""
    Public Property SourceName As String = ""
    Public Property Confidence As Double = 0.0
    Public Property Reason As String = ""
End Class

''' <summary>
''' Реестр стандартов форматирования.
''' Сводит встроенные стандарты, пользовательские шаблоны, стилевые руководства и docx-сопоставления в единый выбор FormattingStandard.
''' </summary>
Public Class FormattingStandardRegistry
    Private ReadOnly _knowledgeEngine As FormattingKnowledgeEngine

    Public Sub New(Optional knowledgeEngine As FormattingKnowledgeEngine = Nothing)
        _knowledgeEngine = If(knowledgeEngine, New FormattingKnowledgeEngine())
    End Sub

    Public Function GetAllCandidates() As List(Of FormattingStandardCandidate)
        Dim result As New List(Of FormattingStandardCandidate)()
        AddBuiltInCandidates(result)
        AddTemplateCandidates(result)
        AddStyleGuideCandidates(result)
        AddDocxMappingCandidates(result)
        Return Deduplicate(result)
    End Function

    Public Function FindStandardByName(name As String) As FormattingStandard
        If String.IsNullOrWhiteSpace(name) Then Return Nothing

        Dim candidate = GetAllCandidates().
            FirstOrDefault(Function(c) c.Standard IsNot Nothing AndAlso
                                       String.Equals(c.Standard.Name, name, StringComparison.OrdinalIgnoreCase))
        If candidate IsNot Nothing Then Return candidate.Standard

        candidate = GetAllCandidates().
            FirstOrDefault(Function(c) c.Standard IsNot Nothing AndAlso
                                       c.Standard.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
        Return If(candidate IsNot Nothing, candidate.Standard, Nothing)
    End Function

    Public Function GetStandardForDocumentType(docType As DocumentType) As FormattingStandard
        If docType = DocumentType.Unknown Then Return Nothing
        Dim typeName = docType.ToString()
        Dim candidate = GetAllCandidates().
            OrderByDescending(Function(c) c.Confidence).
            FirstOrDefault(Function(c) c.Standard IsNot Nothing AndAlso
                                       c.Standard.IsActive AndAlso
                                       c.Standard.ApplicableDocumentTypes IsNot Nothing AndAlso
                                       c.Standard.ApplicableDocumentTypes.Contains(typeName))
        Return If(candidate IsNot Nothing, candidate.Standard, Nothing)
    End Function

    Public Function SelectBest(intent As FormatIntent, analysis As DocumentAnalysisResult) As FormattingStandard
        If intent IsNot Nothing Then
            If Not String.IsNullOrWhiteSpace(intent.TargetStandardName) Then
                Dim byName = FindStandardByName(intent.TargetStandardName)
                If byName IsNot Nothing Then Return byName
            End If

            If intent.TargetDocumentType <> DocumentType.Unknown Then
                Dim byIntentType = GetStandardForDocumentType(intent.TargetDocumentType)
                If byIntentType IsNot Nothing Then Return byIntentType
            End If
        End If

        If analysis IsNot Nothing AndAlso analysis.DocumentType <> DocumentType.Unknown Then
            Dim byAnalysisType = GetStandardForDocumentType(analysis.DocumentType)
            If byAnalysisType IsNot Nothing Then Return byAnalysisType
        End If

        Dim general = _knowledgeEngine.GetStandardForDocumentType(DocumentType.GeneralDocument)
        If general IsNot Nothing Then Return general

        Return CreateGeneralFallbackStandard()
    End Function

    Private Sub AddBuiltInCandidates(result As List(Of FormattingStandardCandidate))
        For Each standard In _knowledgeEngine.GetActiveStandards()
            result.Add(New FormattingStandardCandidate With {
                .Standard = standard,
                .SourceType = FormattingStandardSourceType.BuiltIn,
                .SourceId = standard.Id,
                .SourceName = standard.Name,
                .Confidence = 0.8,
                .Reason = "Встроенный стандарт вёрстки"
            })
        Next
    End Sub

    Private Sub AddTemplateCandidates(result As List(Of FormattingStandardCandidate))
        Try
            For Each template In ReformatTemplateManager.Instance.Templates
                If template Is Nothing Then Continue For
                If Not IsWordTarget(template.TargetApp) Then Continue For

                Dim mapping = LegacyTemplateConverter.Convert(template)
                If mapping Is Nothing OrElse mapping.SemanticTags Is Nothing OrElse mapping.SemanticTags.Count = 0 Then Continue For

                result.Add(New FormattingStandardCandidate With {
                    .Standard = BuildStandard(template.Name, template.Description, template.Category, mapping, template.IsPreset),
                    .SourceType = FormattingStandardSourceType.Template,
                    .SourceId = template.Id,
                    .SourceName = template.Name,
                    .Confidence = If(template.IsPreset, 0.7, 0.9),
                    .Reason = "Пользовательский или предустановленный шаблон"
                })
            Next
        Catch ex As Exception
            Debug.WriteLine("Не удалось загрузить шаблонные стандарты: " & ex.Message)
        End Try
    End Sub

    Private Sub AddStyleGuideCandidates(result As List(Of FormattingStandardCandidate))
        Try
            For Each guide In StyleGuideManager.Instance.StyleGuides
                If guide Is Nothing Then Continue For
                If Not IsWordTarget(guide.TargetApp) Then Continue For

                Dim mapping = SemanticMappingManager.Instance.GetMappingBySourceId(guide.Id)
                If mapping Is Nothing OrElse mapping.SemanticTags Is Nothing OrElse mapping.SemanticTags.Count = 0 Then Continue For

                result.Add(New FormattingStandardCandidate With {
                    .Standard = BuildStandard(guide.Name, guide.Description, guide.Category, mapping, guide.IsPreset),
                    .SourceType = FormattingStandardSourceType.StyleGuide,
                    .SourceId = guide.Id,
                    .SourceName = guide.Name,
                    .Confidence = If(guide.IsPreset, 0.75, 0.88),
                    .Reason = "Преобразованный стандарт вёрстки"
                })
            Next
        Catch ex As Exception
            Debug.WriteLine("Не удалось загрузить стилевые стандарты: " & ex.Message)
        End Try
    End Sub

    Private Sub AddDocxMappingCandidates(result As List(Of FormattingStandardCandidate))
        Try
            For Each mapping In SemanticMappingManager.Instance.Mappings
                If mapping Is Nothing Then Continue For
                If mapping.SourceType <> SemanticMappingSourceType.FromDocxTemplate Then Continue For
                If mapping.SemanticTags Is Nothing OrElse mapping.SemanticTags.Count = 0 Then Continue For

                result.Add(New FormattingStandardCandidate With {
                    .Standard = BuildStandard(mapping.Name, "Стандарт вёрстки, извлечённый из шаблона Word", "", mapping, False),
                    .SourceType = FormattingStandardSourceType.DocxMapping,
                    .SourceId = mapping.Id,
                    .SourceName = mapping.Name,
                    .Confidence = 0.92,
                    .Reason = "Семантическое сопоставление docx"
                })
            Next
        Catch ex As Exception
            Debug.WriteLine("Не удалось загрузить docx-сопоставления: " & ex.Message)
        End Try
    End Sub

    Private Shared Function BuildStandard(name As String,
                                          description As String,
                                          category As String,
                                          mapping As SemanticStyleMapping,
                                          isBuiltIn As Boolean) As FormattingStandard
        Dim standard As New FormattingStandard(name, description)
        standard.Id = If(Not String.IsNullOrWhiteSpace(mapping.SourceId), mapping.SourceId, mapping.Id)
        standard.SemanticMapping = mapping
        standard.SemanticMapping.EnsureBaselineTags()
        standard.IsBuiltIn = isBuiltIn
        standard.IsActive = True
        For Each docType In InferDocumentTypes(name, category)
            If Not standard.ApplicableDocumentTypes.Contains(docType.ToString()) Then
                standard.ApplicableDocumentTypes.Add(docType.ToString())
            End If
        Next
        If standard.ApplicableDocumentTypes.Count = 0 Then
            standard.ApplicableDocumentTypes.Add(DocumentType.GeneralDocument.ToString())
        End If
        Return standard
    End Function

    Private Shared Function CreateGeneralFallbackStandard() As FormattingStandard
        Dim standard As New FormattingStandard("Универсальная вёрстка документа", "Базовое оформление заголовков, основного текста и нумерованных списков для документов без явно определённого типа.")
        standard.Id = "general-document-ai-native"
        standard.IsBuiltIn = True
        standard.IsActive = True
        standard.ApplicableDocumentTypes.Add(DocumentType.GeneralDocument.ToString())
        standard.SemanticMapping = New SemanticStyleMapping()
        standard.SemanticMapping.Id = standard.Id
        standard.SemanticMapping.Name = standard.Name
        standard.SemanticMapping.SourceType = SemanticMappingSourceType.FromStyleGuide
        standard.SemanticMapping.SourceId = standard.Id
        standard.SemanticMapping.EnsureBaselineTags()
        Return standard
    End Function

    Private Shared Function InferDocumentTypes(name As String, category As String) As List(Of DocumentType)
        Dim text = (If(name, "") & " " & If(category, "")).ToLowerInvariant()
        Dim result As New List(Of DocumentType)()
        If ContainsAny(text, {"公文", "行政", "gb/t", "9704", "гост", "приказ", "распоряжение", "официальн"}) Then result.Add(DocumentType.OfficialDocument)
        If ContainsAny(text, {"论文", "学术", "paper", "academic", "диплом", "курсов", "диссертац", "реферат", "научн"}) Then result.Add(DocumentType.AcademicPaper)
        If ContainsAny(text, {"报告", "商务", "商业", "business", "report", "отчёт", "отчет", "доклад", "бизнес"}) Then result.Add(DocumentType.BusinessReport)
        If ContainsAny(text, {"合同", "协议", "contract", "договор", "контракт", "соглашени"}) Then result.Add(DocumentType.Contract)
        If ContainsAny(text, {"简历", "resume", "резюме"}) Then result.Add(DocumentType.[Resume])
        If result.Count = 0 Then result.Add(DocumentType.GeneralDocument)
        Return result
    End Function

    Private Shared Function ContainsAny(text As String, values As IEnumerable(Of String)) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each value In values
            If text.Contains(value.ToLowerInvariant()) Then Return True
        Next
        Return False
    End Function

    Private Shared Function IsWordTarget(targetApp As String) As Boolean
        If String.IsNullOrWhiteSpace(targetApp) Then Return True
        Return targetApp.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Shared Function Deduplicate(candidates As List(Of FormattingStandardCandidate)) As List(Of FormattingStandardCandidate)
        Dim result As New List(Of FormattingStandardCandidate)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each candidate In candidates.OrderByDescending(Function(c) c.Confidence)
            If candidate.Standard Is Nothing Then Continue For
            Dim key = candidate.SourceType.ToString() & ":" & If(candidate.SourceId, candidate.Standard.Name)
            If seen.Add(key) Then result.Add(candidate)
        Next
        Return result
    End Function
End Class
