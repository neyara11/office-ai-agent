' ShareRibbon/Agent/DeckTemplates/TemplateCatalog.vb
' Каталог макетов: чтение, поиск, компактные записи для модели и UI.

Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.DeckTemplates

    Public NotInheritable Class TemplateCatalog
        Private Sub New()
        End Sub

        Public Shared Function ListAll() As List(Of SlideTemplateManifest)
            Return TemplateStorage.LoadAll()
        End Function

        Public Shared Function FindById(id As String) As SlideTemplateManifest
            Return TemplateStorage.Load(id)
        End Function

        Public Shared Function FindByArchetype(archetype As String) As List(Of SlideTemplateManifest)
            Dim normalized = TemplateArchetypes.Normalize(archetype)
            Return TemplateStorage.LoadAll().
                Where(Function(item) String.Equals(item.Archetype, normalized, StringComparison.OrdinalIgnoreCase)).
                ToList()
        End Function

        Public Shared Function Search(query As String,
                                      Optional archetype As String = "",
                                      Optional limit As Integer = 20) As List(Of SlideTemplateManifest)
            Dim items = TemplateStorage.LoadAll().AsEnumerable()
            Dim normalizedArchetype = TemplateArchetypes.Normalize(archetype)
            If Not String.IsNullOrWhiteSpace(archetype) Then
                items = items.Where(Function(item) String.Equals(item.Archetype, normalizedArchetype, StringComparison.OrdinalIgnoreCase))
            End If

            Dim term = If(query, "").Trim()
            If term.Length > 0 Then
                items = items.Where(Function(item) Matches(item, term))
            End If

            Return items.Take(Math.Max(1, limit)).ToList()
        End Function

        Public Shared Function ListSummaries(Optional limit As Integer = 20,
                                             Optional archetype As String = "",
                                             Optional query As String = "") As List(Of SlideTemplateSummary)
            If limit <= 0 Then limit = 20
            If limit > 50 Then limit = 50
            Return Search(query, archetype, limit).Select(Function(item) ToSummary(item)).ToList()
        End Function

        Public Shared Function ToSummary(manifest As SlideTemplateManifest) As SlideTemplateSummary
            If manifest Is Nothing Then Return New SlideTemplateSummary()
            Dim roles = If(manifest.Slots Is Nothing,
                           New List(Of String)(),
                           manifest.Slots.Where(Function(slot) slot IsNot Nothing).
                               Select(Function(slot) slot.Role).
                               Where(Function(role) Not String.IsNullOrWhiteSpace(role)).
                               Distinct(StringComparer.OrdinalIgnoreCase).
                               ToList())
            Dim hasPreview = False
            Try
                hasPreview = Not String.IsNullOrWhiteSpace(manifest.PreviewFullPath) AndAlso IO.File.Exists(manifest.PreviewFullPath)
            Catch
            End Try
            Return New SlideTemplateSummary With {
                .Id = manifest.Id,
                .Name = manifest.Name,
                .SetId = manifest.SetId,
                .Archetype = manifest.Archetype,
                .SlotCount = If(manifest.Slots Is Nothing, 0, manifest.Slots.Count),
                .SlotRoles = roles,
                .HasPreview = hasPreview,
                .Tags = If(manifest.Tags, New List(Of String)()),
                .CreatedAt = manifest.CreatedAt
            }
        End Function

        Private Shared Function Matches(manifest As SlideTemplateManifest, term As String) As Boolean
            If manifest Is Nothing Then Return False
            If Contains(manifest.Name, term) OrElse Contains(manifest.Id, term) OrElse Contains(manifest.SetId, term) Then Return True
            If manifest.Tags IsNot Nothing AndAlso manifest.Tags.Any(Function(tag) Contains(tag, term)) Then Return True
            Return False
        End Function

        Private Shared Function Contains(value As String, term As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(value) AndAlso
                   value.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

    End Class

End Namespace
