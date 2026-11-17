' ShareRibbon/Agent/DeckTemplates/DeckOutline.vb
' План новой колоды: титул и последовательность контентных слайдов.

Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.DeckTemplates

    Public Class DeckOutline
        Public Property Title As String = ""
        Public Property Subtitle As String = ""
        Public Property SetId As String = ""
        Public Property Slides As New List(Of SlideSpec)()

        Public Sub Normalize()
            Title = If(Title, "").Trim()
            Subtitle = If(Subtitle, "").Trim()
            SetId = If(SetId, "").Trim()
            If Slides Is Nothing Then Slides = New List(Of SlideSpec)()
            Slides = Slides.Where(Function(slide) slide IsNot Nothing).ToList()
            For Each slide In Slides
                slide.Normalize()
            Next
        End Sub
    End Class

    Public Class DeckPlanStep
        Public Property Index As Integer
        Public Property Spec As SlideSpec
        Public Property Match As SlideTemplateMatch
    End Class

    ''' <summary>
    ''' Строит детерминированный план колоды: для каждого SlideSpec выбирает макет
    ''' из набора (явный setId) или из переданных кандидатов.
    ''' </summary>
    Public NotInheritable Class DeckBuildPlanner
        Private Sub New()
        End Sub

        Public Shared Function Plan(outline As DeckOutline,
                                    templates As IEnumerable(Of SlideTemplateManifest),
                                    Optional maxSlides As Integer = 20) As List(Of DeckPlanStep)
            Dim steps As New List(Of DeckPlanStep)()
            If outline Is Nothing Then Return steps
            outline.Normalize()
            If maxSlides <= 0 Then maxSlides = 20
            If maxSlides > 40 Then maxSlides = 40

            Dim candidates = If(templates, Enumerable.Empty(Of SlideTemplateManifest)()).ToList()
            If Not String.IsNullOrWhiteSpace(outline.SetId) Then
                Dim setCandidates = candidates.
                    Where(Function(item) String.Equals(item.SetId, outline.SetId, StringComparison.OrdinalIgnoreCase)).
                    ToList()
                If setCandidates.Count > 0 Then candidates = setCandidates
            End If
            If candidates.Count = 0 Then Return steps

            Dim index = 0
            For Each slideSpec In outline.Slides
                If index >= maxSlides Then Exit For
                index += 1
                Dim effectiveSpec = slideSpec
                If index = 1 AndAlso Not String.IsNullOrWhiteSpace(outline.Title) AndAlso
                   String.IsNullOrWhiteSpace(effectiveSpec.Title) Then
                    effectiveSpec = CloneSpec(effectiveSpec)
                    effectiveSpec.Title = outline.Title
                    If String.IsNullOrWhiteSpace(effectiveSpec.Subtitle) Then effectiveSpec.Subtitle = outline.Subtitle
                End If

                steps.Add(New DeckPlanStep With {
                    .Index = index,
                    .Spec = effectiveSpec,
                    .Match = SlideTemplateSelector.SelectBest(effectiveSpec, candidates)
                })
            Next
            Return steps
        End Function

        Private Shared Function CloneSpec(source As SlideSpec) As SlideSpec
            Dim copy = source.Clone()
            Return copy
        End Function

    End Class

End Namespace
