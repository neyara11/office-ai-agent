' ShareRibbon/Agent/DeckTemplates/SlideTemplateManifest.vb
' Контракт макета слайда: метаданные, слоты, источник. Хранится рядом с template.pptx.

Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json

Namespace Agent.DeckTemplates

    Public NotInheritable Class TemplateSchema
        Public Const CurrentVersion As Integer = 1
        Public Const ManifestFileName As String = "manifest.json"
        Public Const TemplateFileName As String = "template.pptx"
        Public Const PreviewFileName As String = "preview.png"

        Private Sub New()
        End Sub
    End Class

    ''' <summary>
    ''' Канонические архетипы макета. Нормализация не запрещает свои значения,
    ''' но селектор и UI опираются на этот список.
    ''' </summary>
    Public NotInheritable Class TemplateArchetypes
        Public Const Cover As String = "cover"
        Public Const Section As String = "section"
        Public Const Bullets As String = "bullets"
        Public Const TwoColumn As String = "two-column"
        Public Const Chart As String = "chart"
        Public Const Table As String = "table"
        Public Const Kpi As String = "kpi"
        Public Const ImageText As String = "image-text"
        Public Const Quote As String = "quote"
        Public Const Closing As String = "closing"
        Public Const Custom As String = "custom"

        Private Sub New()
        End Sub

        Public Shared Function Normalize(value As String) As String
            Dim normalized = If(value, "").Trim().ToLowerInvariant().Replace("_"c, "-"c).Replace(" "c, "-"c)
            If normalized.Length = 0 Then Return Custom
            If normalized.Length > 40 Then normalized = normalized.Substring(0, 40)
            Return normalized
        End Function
    End Class

    ''' <summary>
    ''' Роли слотов, которые умеет заполнять renderer.
    ''' </summary>
    Public NotInheritable Class TemplateSlotRoles
        Public Const Title As String = "title"
        Public Const Subtitle As String = "subtitle"
        Public Const Body As String = "body"
        Public Const Bullets As String = "bullets"
        Public Const Table As String = "table"
        Public Const Chart As String = "chart"
        Public Const Image As String = "image"
        Public Const Caption As String = "caption"
        Public Const Footer As String = "footer"
        Public Const Kpi As String = "kpi"
        Public Const Decor As String = "decor"
        Public Const Other As String = "other"

        Private Sub New()
        End Sub

        Public Shared Function Normalize(value As String) As String
            Dim normalized = If(value, "").Trim().ToLowerInvariant().Replace("_"c, "-"c)
            Select Case normalized
                Case Title, Subtitle, Body, Bullets, Table, Chart, Image, Caption, Footer, Kpi, Decor
                    Return normalized
                Case "text"
                    Return Body
                Case "list"
                    Return Bullets
                Case Else
                    Return Other
            End Select
        End Function
    End Class

    ''' <summary>
    ''' Описание одного слота макета: где на слайде лежит контент и сколько его влезает.
    ''' </summary>
    Public Class TemplateSlot
        Public Property Role As String = TemplateSlotRoles.Other

        ''' <summary>Индекс фигуры в шаблоне (1-based, как в PowerPoint).</summary>
        Public Property ShapeId As Integer

        ''' <summary>Имя фигуры для отладки и повторного поиска.</summary>
        Public Property ShapeName As String = ""

        ''' <summary>Предел текста по геометрии и шрифту (оценка).</summary>
        Public Property CapacityChars As Integer

        ''' <summary>Сколько элементов списка/строк таблицы/точек диаграммы влезает.</summary>
        Public Property ItemCount As Integer

        ''' <summary>text | bullets | table | chart | image | kpi | decor</summary>
        Public Property ContentKind As String = "text"

        <JsonProperty("optional")>
        Public Property IsOptional As Boolean

        ''' <summary>PpPlaceholderType, если слот пришёл из placeholder.</summary>
        Public Property PlaceholderType As String = ""

        ''' <summary>Верхняя граница фигуры в точках: помогает выбрать верхний title/body.</summary>
        Public Property Top As Double

        ''' <summary>Левая граница фигуры в точках.</summary>
        Public Property Left As Double
    End Class

    ''' <summary>
    ''' Откуда взят макет. Нужно для объяснений и повторного извлечения.
    ''' </summary>
    Public Class TemplateSourceInfo
        ''' <summary>slide | deck | import</summary>
        Public Property Kind As String = "slide"
        Public Property FileName As String = ""
        Public Property SlideIndex As Integer
        Public Property SlideCount As Integer
        Public Property AppVersion As String = ""
    End Class

    ''' <summary>
    ''' Манифест макета. Файл манифеста лежит в одной папке с template.pptx и preview.png.
    ''' </summary>
    Public Class SlideTemplateManifest
        Public Property SchemaVersion As Integer = TemplateSchema.CurrentVersion
        Public Property Id As String = ""
        Public Property Name As String = ""
        Public Property SetId As String = ""
        Public Property Archetype As String = TemplateArchetypes.Custom
        Public Property Slots As New List(Of TemplateSlot)()
        Public Property Fingerprint As String = ""
        Public Property ThemeTokens As New Dictionary(Of String, String)()
        Public Property Source As New TemplateSourceInfo()
        Public Property CreatedAt As DateTime = DateTime.UtcNow
        Public Property Tags As New List(Of String)()

        ''' <summary>Относительное имя файла шаблона внутри папки макета.</summary>
        Public Property TemplateFile As String = TemplateSchema.TemplateFileName

        ''' <summary>Относительное имя превью внутри папки макета (пусто — превью нет).</summary>
        Public Property PreviewPath As String = ""

        ''' <summary>Абсолютная папка на диске. Не пишется в JSON.</summary>
        <JsonIgnore>
        Public Property StorageFolder As String = ""

        <JsonIgnore>
        Public ReadOnly Property TemplateFullPath As String
            Get
                If String.IsNullOrWhiteSpace(StorageFolder) Then Return ""
                Return IO.Path.Combine(StorageFolder, If(String.IsNullOrWhiteSpace(TemplateFile), TemplateSchema.TemplateFileName, TemplateFile))
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property PreviewFullPath As String
            Get
                If String.IsNullOrWhiteSpace(StorageFolder) Then Return ""
                If String.IsNullOrWhiteSpace(PreviewPath) Then
                    Return IO.Path.Combine(StorageFolder, TemplateSchema.PreviewFileName)
                End If
                Return IO.Path.Combine(StorageFolder, PreviewPath)
            End Get
        End Property

        Public Function Clone() As SlideTemplateManifest
            Dim copy = DirectCast(Me.MemberwiseClone(), SlideTemplateManifest)
            copy.Slots = New List(Of TemplateSlot)(Slots)
            copy.Tags = New List(Of String)(Tags)
            copy.ThemeTokens = New Dictionary(Of String, String)(ThemeTokens)
            copy.Source = New TemplateSourceInfo With {
                .Kind = If(Source?.Kind, ""),
                .FileName = If(Source?.FileName, ""),
                .SlideIndex = If(Source?.SlideIndex, 0),
                .SlideCount = If(Source?.SlideCount, 0),
                .AppVersion = If(Source?.AppVersion, "")
            }
            Return copy
        End Function

        Public Sub Normalize()
            Archetype = TemplateArchetypes.Normalize(Archetype)
            If String.IsNullOrWhiteSpace(SetId) Then SetId = Id
            SetId = TemplateIds.Safe(SetId, Id)
            If String.IsNullOrWhiteSpace(Name) Then Name = Id
            If Slots Is Nothing Then Slots = New List(Of TemplateSlot)()
            For Each slot In Slots
                If slot Is Nothing Then Continue For
                slot.Role = TemplateSlotRoles.Normalize(slot.Role)
                slot.ContentKind = If(String.IsNullOrWhiteSpace(slot.ContentKind), slot.Role, slot.ContentKind.Trim().ToLowerInvariant())
            Next
            If Tags Is Nothing Then Tags = New List(Of String)()
            Tags = Tags.Where(Function(tag) Not String.IsNullOrWhiteSpace(tag)).
                Select(Function(tag) tag.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
            If ThemeTokens Is Nothing Then ThemeTokens = New Dictionary(Of String, String)()
            If Source Is Nothing Then Source = New TemplateSourceInfo()
        End Sub
    End Class

    ''' <summary>
    ''' Безопасные идентификаторы макетов/наборов: только буквы, цифры, дефис, точка, подчёркивание.
    ''' </summary>
    Public NotInheritable Class TemplateIds
        Private Shared ReadOnly SafePattern As New Text.RegularExpressions.Regex("^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")

        Private Sub New()
        End Sub

        Public Shared Function IsSafe(value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then Return False
            If value.Contains("..") Then Return False
            Return SafePattern.IsMatch(value)
        End Function

        Public Shared Function Safe(value As String, Optional fallback As String = "") As String
            If IsSafe(value) Then Return value.Trim()
            If IsSafe(fallback) Then Return fallback.Trim()
            Return "template-" & Guid.NewGuid().ToString("N").Substring(0, 12)
        End Function
    End Class

    ''' <summary>Компактная запись каталога для модели и UI. Без base64 и абсолютных путей.</summary>
    Public Class SlideTemplateSummary
        Public Property Id As String = ""
        Public Property Name As String = ""
        Public Property SetId As String = ""
        Public Property Archetype As String = TemplateArchetypes.Custom
        Public Property SlotCount As Integer
        Public Property SlotRoles As New List(Of String)()
        Public Property HasPreview As Boolean
        Public Property Tags As New List(Of String)()
        Public Property CreatedAt As DateTime = DateTime.UtcNow
    End Class

End Namespace
