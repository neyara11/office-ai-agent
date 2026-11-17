' ShareRibbon/Agent/DeckTemplates/SlideTemplateSelector.vb
' Детерминированный подбор макета под контент слайда.

Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.DeckTemplates

    Public Class SlideTemplateMatch
        Public Property Template As SlideTemplateManifest
        Public Property Score As Integer
        Public Property MatchedBy As New List(Of String)()
        Public Property Missing As New List(Of String)()

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return Template Is Nothing
            End Get
        End Property
    End Class

    Public NotInheritable Class SlideTemplateSelector
        Private Sub New()
        End Sub

        ''' <summary>Подбирает лучший макет из явного каталога.</summary>
        Public Shared Function SelectFromCatalog(spec As SlideSpec, Optional limit As Integer = 50) As SlideTemplateMatch
            Dim templates = TemplateCatalog.ListAll().Take(Math.Max(1, limit)).ToList()
            Return SelectBest(spec, templates)
        End Function

        ''' <summary>Подбирает лучший макет из переданного списка (каталог или слайды колоды).</summary>
        Public Shared Function SelectBest(spec As SlideSpec, templates As IEnumerable(Of SlideTemplateManifest)) As SlideTemplateMatch
            If templates Is Nothing Then Return New SlideTemplateMatch()
            If spec Is Nothing Then spec = New SlideSpec()
            spec.Normalize()

            Dim matches As New List(Of SlideTemplateMatch)()
            For Each template In templates
                If template Is Nothing Then Continue For
                matches.Add(Score(template, spec))
            Next
            If matches.Count = 0 Then Return New SlideTemplateMatch()

            Return matches.
                OrderByDescending(Function(item) item.Score).
                ThenBy(Function(item) If(item.Template.Name, ""), StringComparer.OrdinalIgnoreCase).
                ThenBy(Function(item) If(item.Template.Id, ""), StringComparer.OrdinalIgnoreCase).
                First()
        End Function

        ''' <summary>
        ''' Оценка макета под spec. Положительные факторы — покрытие блоков, вместимость,
        ''' совпадение архетипа; отрицательные — отсутствие нужного слота или переполнение.
        ''' </summary>
        Public Shared Function Score(template As SlideTemplateManifest, spec As SlideSpec) As SlideTemplateMatch
            Dim match As New SlideTemplateMatch With {.Template = template}
            If template Is Nothing Then Return match
            If spec Is Nothing Then spec = New SlideSpec()
            If template.Slots Is Nothing Then template.Slots = New List(Of TemplateSlot)()

            Dim slots = template.Slots.Where(Function(slot) slot IsNot Nothing).ToList()
            Dim byRole = Function(role As String) slots.FirstOrDefault(Function(slot) String.Equals(slot.Role, role, StringComparison.OrdinalIgnoreCase))

            Dim textContent = BuildTextContent(spec)
            Dim archetype = If(spec.Archetype, "").Trim().ToLowerInvariant()
            If archetype.Length > 0 Then
                If String.Equals(template.Archetype, archetype, StringComparison.OrdinalIgnoreCase) Then
                    match.Score += 30
                    match.MatchedBy.Add("archetype:" & archetype)
                Else
                    match.Score -= 5
                End If
            End If

            If Not String.IsNullOrWhiteSpace(spec.Title) Then
                Dim titleSlot = byRole(TemplateSlotRoles.Title)
                If titleSlot Is Nothing Then
                    match.Score -= 25
                    match.Missing.Add(TemplateSlotRoles.Title)
                Else
                    match.Score += 10
                    match.MatchedBy.Add("slot:title")
                    If titleSlot.CapacityChars > 0 AndAlso spec.Title.Length > titleSlot.CapacityChars Then
                        match.Score -= 4
                        match.Missing.Add("title-capacity")
                    End If
                End If
            End If

            If Not String.IsNullOrWhiteSpace(spec.Subtitle) Then
                If byRole(TemplateSlotRoles.Subtitle) IsNot Nothing Then
                    match.Score += 6
                    match.MatchedBy.Add("slot:subtitle")
                Else
                    match.Score -= 6
                    match.Missing.Add(TemplateSlotRoles.Subtitle)
                End If
            End If

            If textContent.Length > 0 Then
                Dim bulletsSlot = byRole(TemplateSlotRoles.Bullets)
                Dim bodySlot = byRole(TemplateSlotRoles.Body)
                Dim textSlot = If(bulletsSlot, bodySlot)
                If textSlot Is Nothing Then
                    match.Score -= 12
                    match.Missing.Add(TemplateSlotRoles.Body)
                Else
                    match.Score += If(bulletsSlot IsNot Nothing, 12, 8)
                    match.MatchedBy.Add(If(bulletsSlot IsNot Nothing, "slot:bullets", "slot:body"))
                    If textSlot.CapacityChars > 0 AndAlso textContent.Length > textSlot.CapacityChars Then
                        match.Score -= 4
                        match.Missing.Add("body-capacity")
                    End If
                    If spec.Bullets.Count > 0 AndAlso textSlot.ItemCount > 0 AndAlso spec.Bullets.Count > textSlot.ItemCount Then
                        match.Score -= 4
                        match.Missing.Add("bullets-capacity")
                    End If
                End If
            End If

            If spec.Table IsNot Nothing Then
                Dim tableSlot = byRole(TemplateSlotRoles.Table)
                If tableSlot Is Nothing Then
                    match.Score -= 20
                    match.Missing.Add(TemplateSlotRoles.Table)
                Else
                    match.Score += 18
                    match.MatchedBy.Add("slot:table")
                    If tableSlot.ItemCount > 0 AndAlso spec.Table.RowCount > tableSlot.ItemCount Then
                        match.Score -= 8
                        match.Missing.Add("table-rows")
                    End If
                End If
            End If

            If spec.Chart IsNot Nothing Then
                Dim chartSlot = byRole(TemplateSlotRoles.Chart)
                If chartSlot Is Nothing Then
                    match.Score -= 20
                    match.Missing.Add(TemplateSlotRoles.Chart)
                Else
                    match.Score += 18
                    match.MatchedBy.Add("slot:chart")
                    If chartSlot.ItemCount > 0 AndAlso spec.Chart.PointCount > chartSlot.ItemCount Then
                        match.Score -= 6
                        match.Missing.Add("chart-points")
                    End If
                End If
            End If

            If spec.Image IsNot Nothing Then
                If byRole(TemplateSlotRoles.Image) IsNot Nothing Then
                    match.Score += 8
                    match.MatchedBy.Add("slot:image")
                Else
                    match.Score -= 8
                    match.Missing.Add(TemplateSlotRoles.Image)
                End If
            End If

            If spec.Kpis IsNot Nothing AndAlso spec.Kpis.Count > 0 Then
                If String.Equals(template.Archetype, TemplateArchetypes.Kpi, StringComparison.OrdinalIgnoreCase) Then
                    match.Score += 10
                    match.MatchedBy.Add("archetype:kpi")
                Else
                    Dim contentSlots = slots.Where(Function(slot) slot.Role = TemplateSlotRoles.Body OrElse
                                                               slot.Role = TemplateSlotRoles.Kpi).ToList()
                    If contentSlots.Count >= spec.Kpis.Count Then
                        match.Score += 6
                        match.MatchedBy.Add("kpi-slots")
                    Else
                        match.Score -= 6
                        match.Missing.Add("kpi-slots")
                    End If
                End If
            End If

            If Not String.IsNullOrWhiteSpace(spec.Caption) Then
                If byRole(TemplateSlotRoles.Caption) IsNot Nothing Then match.Score += 2
            End If
            If Not String.IsNullOrWhiteSpace(spec.Footer) Then
                If byRole(TemplateSlotRoles.Footer) IsNot Nothing Then match.Score += 2
            End If

            Return match
        End Function

        Private Shared Function BuildTextContent(spec As SlideSpec) As String
            Dim parts As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(spec.Body) Then parts.Add(spec.Body)
            If spec.Bullets IsNot Nothing Then parts.AddRange(spec.Bullets.Where(Function(item) Not String.IsNullOrWhiteSpace(item)))
            Return String.Join(vbLf, parts)
        End Function

    End Class

End Namespace
