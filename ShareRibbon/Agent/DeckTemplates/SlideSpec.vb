' ShareRibbon/Agent/DeckTemplates/SlideSpec.vb
' Запрос контента от модели: что именно нужно разместить на новом слайде.

Imports System.Collections.Generic
Imports System.Linq

Namespace Agent.DeckTemplates

    ''' <summary>
    ''' Контент слайда, который нужно заполнить в макете. Поля необязательны:
    ''' селектор подбирает макет по тем блокам, которые реально переданы.
    ''' </summary>
    Public Class SlideSpec
        Public Property Archetype As String = ""
        Public Property Title As String = ""
        Public Property Subtitle As String = ""
        Public Property Body As String = ""
        Public Property Bullets As New List(Of String)()
        Public Property Table As SlideTableSpec = Nothing
        Public Property Chart As SlideChartSpec = Nothing
        Public Property Image As SlideImageSpec = Nothing
        Public Property Kpis As New List(Of SlideKpiSpec)()
        Public Property Caption As String = ""
        Public Property Footer As String = ""
        Public Property Notes As String = ""
        Public Property Blocks As New List(Of SlideBlockSpec)()

        Public Sub Normalize()
            Archetype = TemplateArchetypes.Normalize(Archetype)
            Title = TrimText(Title)
            Subtitle = TrimText(Subtitle)
            Body = TrimText(Body)
            Caption = TrimText(Caption)
            Footer = TrimText(Footer)
            Notes = TrimText(Notes)
            If Bullets Is Nothing Then Bullets = New List(Of String)()
            Bullets = Bullets.Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                Select(Function(item) item.Trim()).ToList()
            If Kpis Is Nothing Then Kpis = New List(Of SlideKpiSpec)()
            Kpis = Kpis.Where(Function(item) item IsNot Nothing AndAlso
                                          (Not String.IsNullOrWhiteSpace(item.Label) OrElse
                                           Not String.IsNullOrWhiteSpace(item.Value))).ToList()
            If Blocks Is Nothing Then Blocks = New List(Of SlideBlockSpec)()
            Blocks = Blocks.Where(Function(item) item IsNot Nothing).ToList()
            For Each block In Blocks
                block.Normalize()
            Next
            If Table IsNot Nothing Then Table.Normalize()
            If Chart IsNot Nothing Then Chart.Normalize()
        End Sub

        Public Function HasAnyContent() As Boolean
            Return Not String.IsNullOrWhiteSpace(Title) OrElse
                   Not String.IsNullOrWhiteSpace(Subtitle) OrElse
                   Not String.IsNullOrWhiteSpace(Body) OrElse
                   Bullets.Count > 0 OrElse
                   Table IsNot Nothing OrElse
                   Chart IsNot Nothing OrElse
                   Image IsNot Nothing OrElse
                   Kpis.Count > 0 OrElse
                   Blocks.Count > 0
        End Function

        Public Function Clone() As SlideSpec
            Dim copy = DirectCast(Me.MemberwiseClone(), SlideSpec)
            copy.Bullets = New List(Of String)(Bullets)
            copy.Kpis = New List(Of SlideKpiSpec)(Kpis)
            copy.Blocks = New List(Of SlideBlockSpec)(Blocks)
            Return copy
        End Function

        Private Shared Function TrimText(value As String) As String
            Return If(value, "").Trim()
        End Function
    End Class

    ''' <summary>Один смысловой блок слайда (для макетов из нескольких зон).</summary>
    Public Class SlideBlockSpec
        ''' <summary>text | bullets | table | chart | image | kpi</summary>
        Public Property Kind As String = "text"
        Public Property Text As String = ""
        Public Property Items As New List(Of String)()
        Public Property Table As SlideTableSpec = Nothing
        Public Property Chart As SlideChartSpec = Nothing
        Public Property Kpis As New List(Of SlideKpiSpec)()

        Public Sub Normalize()
            Kind = If(Kind, "text").Trim().ToLowerInvariant()
            Text = If(Text, "").Trim()
            If Items Is Nothing Then Items = New List(Of String)()
            Items = Items.Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                Select(Function(item) item.Trim()).ToList()
            If Kpis Is Nothing Then Kpis = New List(Of SlideKpiSpec)()
            If Table IsNot Nothing Then Table.Normalize()
            If Chart IsNot Nothing Then Chart.Normalize()
        End Sub
    End Class

    Public Class SlideTableSpec
        Public Property Headers As New List(Of String)()
        Public Property Rows As New List(Of List(Of String))()

        Public Sub Normalize()
            If Headers Is Nothing Then Headers = New List(Of String)()
            Headers = Headers.Select(Function(item) If(item, "").Trim()).ToList()
            If Rows Is Nothing Then Rows = New List(Of List(Of String))()
            Dim normalizedRows As New List(Of List(Of String))()
            For Each row In Rows
                If row Is Nothing Then Continue For
                normalizedRows.Add(row.Select(Function(cell) If(cell, "").Trim()).ToList())
            Next
            Rows = normalizedRows
        End Sub

        Public ReadOnly Property RowCount As Integer
            Get
                Return If(Rows Is Nothing, 0, Rows.Count)
            End Get
        End Property

        Public ReadOnly Property ColumnCount As Integer
            Get
                Dim headerCount = If(Headers Is Nothing, 0, Headers.Count)
                Dim rowMax = 0
                If Rows IsNot Nothing Then
                    For Each row In Rows
                        If row IsNot Nothing AndAlso row.Count > rowMax Then rowMax = row.Count
                    Next
                End If
                Return Math.Max(headerCount, rowMax)
            End Get
        End Property
    End Class

    Public Class SlideChartSpec
        Public Property ChartType As String = "doughnut"
        Public Property Categories As New List(Of String)()
        Public Property Values As New List(Of Double)()
        Public Property SeriesName As String = ""

        Public Sub Normalize()
            ChartType = If(ChartType, "doughnut").Trim().ToLowerInvariant()
            If Categories Is Nothing Then Categories = New List(Of String)()
            Categories = Categories.Select(Function(item) If(item, "").Trim()).ToList()
            If Values Is Nothing Then Values = New List(Of Double)()
            SeriesName = If(SeriesName, "").Trim()
        End Sub

        Public ReadOnly Property PointCount As Integer
            Get
                Dim categoryCount = If(Categories Is Nothing, 0, Categories.Count)
                Dim valueCount = If(Values Is Nothing, 0, Values.Count)
                Return Math.Max(categoryCount, valueCount)
            End Get
        End Property
    End Class

    Public Class SlideImageSpec
        Public Property Path As String = ""
        Public Property Description As String = ""
    End Class

    Public Class SlideKpiSpec
        Public Property Label As String = ""
        Public Property Value As String = ""
    End Class

End Namespace
