' PowerPointAi/Design/DeckTemplate/DeckTemplateExtractor.vb
' Извлечение макета из реального слайда PowerPoint: изолированный template.pptx,
' манифест со слотами, fingerprint, theme tokens и PNG-превью.

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon.Agent.DeckTemplates

Namespace Design.DeckTemplate

    Public NotInheritable Class DeckTemplateExtractor

        Private Const MaxSlots As Integer = 60
        ' MsoThemeFontIndex отсутствует в этой interop-сборке; msoThemeFontLatin = 1.
        Private Const ThemeFontLatin As Integer = 1
        Private Const PreviewWidthPx As Integer = 640

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Сохраняет один слайд активной презентации как макет.
        ''' </summary>
        Public Shared Function ExtractFromSlide(app As PowerPoint.Application,
                                                slideIndex As Integer,
                                                name As String,
                                                setId As String,
                                                archetypeOverride As String,
                                                ByRef errorMessage As String,
                                                Optional warnings As List(Of String) = Nothing) As SlideTemplateManifest
            errorMessage = ""
            If warnings Is Nothing Then warnings = New List(Of String)()

            Dim presentation As PowerPoint.Presentation = Nothing
            Dim tempDir = CreateTempDirectory()
            Try
                presentation = app.ActivePresentation
                If presentation Is Nothing Then
                    errorMessage = "Нет активной презентации PowerPoint"
                    Return Nothing
                End If
                If slideIndex <= 0 OrElse slideIndex > presentation.Slides.Count Then
                    errorMessage = $"Слайд {slideIndex} вне диапазона 1..{presentation.Slides.Count}"
                    Return Nothing
                End If

                Dim snapshotPath = Path.Combine(tempDir, "snapshot.pptx")
                presentation.SaveCopyAs(snapshotPath)

                Dim manifest = ExtractOneSlide(app, snapshotPath, tempDir, slideIndex,
                                               name, setId, archetypeOverride, warnings,
                                               presentation.Name, presentation.Slides.Count)
                If manifest Is Nothing Then
                    errorMessage = "Не удалось извлечь макет из слайда " & slideIndex
                End If
                Return manifest
            Catch ex As Exception
                errorMessage = "Извлечение макета: " & ex.Message
                Return Nothing
            Finally
                ReleaseCom(presentation)
                TryDeleteDirectory(tempDir)
            End Try
        End Function

        ''' <summary>
        ''' Сохраняет каждый непустой слайд активной презентации как отдельный макет одного набора.
        ''' </summary>
        Public Shared Function ExtractFromDeck(app As PowerPoint.Application,
                                               setName As String,
                                               maxSlides As Integer,
                                               skipEmpty As Boolean,
                                               ByRef errorMessage As String,
                                               ByRef warnings As List(Of String)) As List(Of SlideTemplateManifest)
            errorMessage = ""
            warnings = New List(Of String)()
            Dim result As New List(Of SlideTemplateManifest)()
            If maxSlides <= 0 Then maxSlides = 20
            If maxSlides > 40 Then maxSlides = 40

            Dim presentation As PowerPoint.Presentation = Nothing
            Dim tempDir = CreateTempDirectory()
            Try
                presentation = app.ActivePresentation
                If presentation Is Nothing Then
                    errorMessage = "Нет активной презентации PowerPoint"
                    Return result
                End If

                Dim snapshotPath = Path.Combine(tempDir, "snapshot.pptx")
                presentation.SaveCopyAs(snapshotPath)

                Dim setId = "deck-" & Guid.NewGuid().ToString("N").Substring(0, 8)
                Dim total = presentation.Slides.Count
                Dim extracted = 0
                For index = 1 To total
                    If extracted >= maxSlides Then
                        warnings.Add($"Ограничение {maxSlides} макетов: обработаны не все слайды ({total}).")
                        Exit For
                    End If
                    If skipEmpty AndAlso IsEmptySlide(presentation.Slides(index)) Then
                        Continue For
                    End If

                    Dim slideName = If(String.IsNullOrWhiteSpace(setName), $"Слайд {index}", $"{setName} — {index}")
                    Dim manifest = ExtractOneSlide(app, snapshotPath, tempDir, index,
                                                   slideName, setId, "", warnings,
                                                   presentation.Name, presentation.Slides.Count,
                                                   TemplateSchema.CurrentVersion)
                    If manifest IsNot Nothing Then
                        extracted += 1
                        result.Add(manifest)
                    End If
                Next

                If result.Count = 0 Then
                    errorMessage = "Не удалось извлечь ни один макет: все слайды пустые или ошибки извлечения."
                End If
                Return result
            Catch ex As Exception
                errorMessage = "Извлечение набора макетов: " & ex.Message
                Return result
            Finally
                ReleaseCom(presentation)
                TryDeleteDirectory(tempDir)
            End Try
        End Function

        ''' <summary>
        ''' Копирует snapshot, изолирует один слайд, строит манифест и сохраняет макет в хранилище.
        ''' </summary>
        Private Shared Function ExtractOneSlide(app As PowerPoint.Application,
                                                snapshotPath As String,
                                                tempDir As String,
                                                slideIndex As Integer,
                                                name As String,
                                                requestedSetId As String,
                                                archetypeOverride As String,
                                                warnings As List(Of String),
                                                sourceFileName As String,
                                                sourceSlideCount As Integer,
                                                Optional schemaVersion As Integer = 0) As SlideTemplateManifest
            Dim token = Guid.NewGuid().ToString("N").Substring(0, 8)
            Dim tplPath = Path.Combine(tempDir, $"slide-{slideIndex}-{token}.pptx")
            Dim previewPath = Path.Combine(tempDir, $"slide-{slideIndex}-{token}.png")
            Dim tempPresentation As PowerPoint.Presentation = Nothing
            Dim liveSlide As PowerPoint.Slide = Nothing
            Try
                File.Copy(snapshotPath, tplPath, True)
                tempPresentation = app.Presentations.Open(tplPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse)

                Dim sourceCount = tempPresentation.Slides.Count
                If slideIndex <= 0 OrElse slideIndex > sourceCount Then Return Nothing

                For index = sourceCount To 1 Step -1
                    If index <> slideIndex Then tempPresentation.Slides(index).Delete()
                Next

                Dim slide As PowerPoint.Slide = tempPresentation.Slides(1)
                Try
                    Dim manifestId = If(String.IsNullOrWhiteSpace(requestedSetId),
                                        "tpl-" & token,
                                        TemplateIds.Safe(requestedSetId) & "-s" & slideIndex.ToString(CultureInfo.InvariantCulture))
                    Dim setName = If(String.IsNullOrWhiteSpace(requestedSetId), manifestId, TemplateIds.Safe(requestedSetId))
                    Dim manifest = BuildManifest(app, tempPresentation, slide, slideIndex, name,
                                                 manifestId, setName, archetypeOverride, warnings,
                                                 sourceFileName, sourceSlideCount)
                    If schemaVersion > 0 Then manifest.SchemaVersion = schemaVersion

                    Dim exported = ExportPreview(tempPresentation, slide, previewPath, warnings)
                    If Not exported Then
                        ' Резерв: превью из живой презентации, если скрытая не отдаёт экспорт.
                        Try
                            liveSlide = app.ActivePresentation.Slides(slideIndex)
                            ExportPreview(app.ActivePresentation, liveSlide, previewPath, warnings)
                        Catch ex As Exception
                            warnings.Add("Превью не создано: " & ex.Message)
                        End Try
                    End If

                    Try
                        tempPresentation.Save()
                    Catch ex As Exception
                        warnings.Add("Не удалось пересохранить изолированный макет: " & ex.Message)
                    End Try

                    Return TemplateStorage.Save(manifest, tplPath,
                                                If(File.Exists(previewPath), previewPath, ""))
                Finally
                    ReleaseCom(slide)
                End Try
            Catch ex As Exception
                warnings.Add($"Слайд {slideIndex}: {ex.Message}")
                Return Nothing
            Finally
                ' Close обязателен: ReleaseComObject сам по себе не закрывает скрытую
                ' презентацию и оставляет файл заблокированным.
                Try
                    If tempPresentation IsNot Nothing Then tempPresentation.Close()
                Catch
                End Try
                ReleaseCom(tempPresentation)
                ReleaseCom(liveSlide)
                TryDeleteFile(tplPath)
                TryDeleteFile(previewPath)
            End Try
        End Function

        Private Shared Function BuildManifest(app As PowerPoint.Application,
                                              presentation As PowerPoint.Presentation,
                                              slide As PowerPoint.Slide,
                                              sourceSlideIndex As Integer,
                                              name As String,
                                              id As String,
                                              setId As String,
                                              archetypeOverride As String,
                                              warnings As List(Of String),
                                              sourceFileName As String,
                                              sourceSlideCount As Integer) As SlideTemplateManifest
            Dim slots = DescribeSlots(slide, warnings)
            ' Двухаргументный If для String — null-coalescing, а не проверка пустой строки.
            Dim archetype As String
            If String.IsNullOrWhiteSpace(archetypeOverride) Then
                archetype = ClassifyArchetype(slide, slots)
            Else
                archetype = archetypeOverride
            End If
            archetype = TemplateArchetypes.Normalize(archetype)
            Dim manifest As New SlideTemplateManifest With {
                .Id = TemplateIds.Safe(id),
                .Name = If(String.IsNullOrWhiteSpace(name), "Слайд " & sourceSlideIndex, name.Trim()),
                .SetId = TemplateIds.Safe(setId, id),
                .Archetype = archetype,
                .Slots = slots,
                .Fingerprint = ComputeFingerprint(slide),
                .ThemeTokens = ReadThemeTokens(presentation),
                .Source = New TemplateSourceInfo With {
                    .Kind = "slide",
                    .FileName = If(sourceFileName, ""),
                    .SlideIndex = sourceSlideIndex,
                    .SlideCount = sourceSlideCount,
                    .AppVersion = SafeAppVersion(app)
                }
            }
            If manifest.Tags Is Nothing Then manifest.Tags = New List(Of String)()
            If Not String.IsNullOrWhiteSpace(archetype) Then manifest.Tags.Add(archetype)
            manifest.Normalize()
            Return manifest
        End Function

        ''' <summary>
        ''' Слоты слайда без сохранения на диск. Нужны renderer'у и implicit-каталогу текущей колоды.
        ''' </summary>
        Public Shared Function DescribeSlotsForSlide(slide As PowerPoint.Slide) As List(Of TemplateSlot)
            Return DescribeSlots(slide, New List(Of String)())
        End Function

        ''' <summary>
        ''' In-memory описание слайда как макета (без файлов и хранилища).
        ''' </summary>
        Public Shared Function DescribeSlideAsManifest(presentation As PowerPoint.Presentation,
                                                       slideIndex As Integer) As SlideTemplateManifest
            If presentation Is Nothing Then Return Nothing
            If slideIndex <= 0 OrElse slideIndex > presentation.Slides.Count Then Return Nothing

            Dim slide As PowerPoint.Slide = Nothing
            Try
                slide = presentation.Slides(slideIndex)
                Dim slots = DescribeSlotsForSlide(slide)
                Return New SlideTemplateManifest With {
                    .Id = $"slide-{slideIndex}",
                    .Name = $"Слайд {slideIndex}",
                    .SetId = "current-deck",
                    .Archetype = TemplateArchetypes.Normalize(ClassifyArchetype(slide, slots)),
                    .Slots = slots,
                    .Fingerprint = ComputeFingerprint(slide),
                    .Source = New TemplateSourceInfo With {
                        .Kind = "slide",
                        .FileName = If(presentation.Name, ""),
                        .SlideIndex = slideIndex,
                        .SlideCount = If(presentation.Slides Is Nothing, 0, presentation.Slides.Count)
                    }
                }
            Finally
                ReleaseCom(slide)
            End Try
        End Function

        Private Shared Function DescribeSlots(slide As PowerPoint.Slide, warnings As List(Of String)) As List(Of TemplateSlot)
            Dim slots As New List(Of TemplateSlot)()
            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                Dim slideWidth = Math.Max(1.0, CDbl(slide.Parent.PageSetup.SlideWidth))
                Dim slideHeight = Math.Max(1.0, CDbl(slide.Parent.PageSetup.SlideHeight))
                Dim count = Math.Min(shapes.Count, MaxSlots)
                For index = 1 To count
                    Dim shape As PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(index)
                        Dim slot = DescribeShape(slide, shape, slideWidth, slideHeight)
                        If slot IsNot Nothing Then slots.Add(slot)
                    Catch ex As Exception
                        warnings.Add($"Фигура {index}: {ex.Message}")
                    Finally
                        ReleaseCom(shape)
                    End Try
                Next
            Catch ex As Exception
                warnings.Add("Слоты не прочитаны: " & ex.Message)
            Finally
                ReleaseCom(shapes)
            End Try
            Return slots
        End Function

        Private Shared Function DescribeShape(slide As PowerPoint.Slide,
                                              shape As PowerPoint.Shape,
                                              slideWidth As Double,
                                              slideHeight As Double) As TemplateSlot
            Dim slot As New TemplateSlot With {
                .ShapeId = shape.Id,
                .ShapeName = If(shape.Name, ""),
                .IsOptional = False
            }
            Try
                slot.Top = CDbl(shape.Top)
                slot.Left = CDbl(shape.Left)
            Catch
            End Try

            Dim hasChart = shape.HasChart = MsoTriState.msoTrue
            Dim hasTable = shape.HasTable = MsoTriState.msoTrue
            Dim isPicture = shape.Type = MsoShapeType.msoPicture
            Dim hasText = shape.HasTextFrame = MsoTriState.msoTrue AndAlso
                          shape.TextFrame.HasText = MsoTriState.msoTrue

            If hasChart Then
                slot.Role = TemplateSlotRoles.Chart
                slot.ContentKind = TemplateSlotRoles.Chart
                slot.ItemCount = SafeChartPointCount(shape)
                slot.CapacityChars = 0
                Return slot
            End If
            If hasTable Then
                slot.Role = TemplateSlotRoles.Table
                slot.ContentKind = TemplateSlotRoles.Table
                slot.ItemCount = SafeTableRowCount(shape)
                slot.CapacityChars = 0
                Return slot
            End If
            If isPicture Then
                slot.Role = TemplateSlotRoles.Image
                slot.ContentKind = TemplateSlotRoles.Image
                slot.CapacityChars = 0
                slot.ItemCount = 1
                Return slot
            End If

            If shape.Type = MsoShapeType.msoPlaceholder Then
                Dim placeholderRole = RoleFromPlaceholder(shape)
                If Not String.IsNullOrWhiteSpace(placeholderRole) Then slot.Role = placeholderRole
            End If

            If hasText Then
                Dim text = Convert.ToString(shape.TextFrame.TextRange.Text)
                text = If(text, "").Trim()
                If slot.Role = TemplateSlotRoles.Other Then
                    slot.Role = If(LooksLikeTitle(shape, slideHeight, text), TemplateSlotRoles.Title, TemplateSlotRoles.Body)
                End If
                slot.ContentKind = If(slot.Role = TemplateSlotRoles.Bullets, TemplateSlotRoles.Bullets, TemplateSlotRoles.Body)
                If slot.Role = TemplateSlotRoles.Title OrElse slot.Role = TemplateSlotRoles.Subtitle Then
                    slot.ContentKind = TemplateSlotRoles.Body
                End If
                Dim fontSize = SafeFontSize(shape)
                slot.CapacityChars = EstimateTextCapacity(shape, fontSize)
                slot.ItemCount = If(slot.Role = TemplateSlotRoles.Bullets,
                                    EstimateLineCount(shape, fontSize),
                                    Math.Min(6, EstimateLineCount(shape, fontSize)))
                slot.IsOptional = slot.Role = TemplateSlotRoles.Footer OrElse
                                  slot.Role = TemplateSlotRoles.Caption OrElse
                                  slot.Role = TemplateSlotRoles.Decor
                Return slot
            End If

            ' Без текста, таблицы и диаграммы — декор (линии, фоны, рамки).
            slot.Role = TemplateSlotRoles.Decor
            slot.ContentKind = TemplateSlotRoles.Decor
            slot.CapacityChars = 0
            slot.ItemCount = 0
            slot.IsOptional = True
            Return slot
        End Function

        Private Shared Function RoleFromPlaceholder(shape As PowerPoint.Shape) As String
            Try
                Select Case CInt(shape.PlaceholderFormat.Type)
                    Case 1, 3, 5
                        Return TemplateSlotRoles.Title
                    Case 4
                        Return TemplateSlotRoles.Subtitle
                    Case 2, 6, 7, 17
                        Return TemplateSlotRoles.Body
                    Case 8
                        Return TemplateSlotRoles.Chart
                    Case 12
                        Return TemplateSlotRoles.Table
                    Case 18
                        Return TemplateSlotRoles.Image
                    Case 13, 14, 15, 16
                        Return TemplateSlotRoles.Footer
                    Case Else
                        Return ""
                End Select
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function LooksLikeTitle(shape As PowerPoint.Shape, slideHeight As Double, text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return False
            Try
                If CDbl(shape.Top) < slideHeight * 0.3 AndAlso text.Length <= 120 Then Return True
            Catch
            End Try
            Return False
        End Function

        Private Shared Function SafeFontSize(shape As PowerPoint.Shape) As Double
            Try
                Dim size = CDbl(shape.TextFrame2.TextRange.Font.Size)
                If size >= 6 AndAlso size <= 200 Then Return size
            Catch
            End Try
            Return 18.0
        End Function

        Private Shared Function EstimateTextCapacity(shape As PowerPoint.Shape, fontSize As Double) As Integer
            Try
                Dim widthPt = Math.Max(24.0, CDbl(shape.Width))
                Dim heightPt = Math.Max(16.0, CDbl(shape.Height))
                Dim charsPerLine = Math.Max(8, CInt(Math.Floor(widthPt / (fontSize * 0.55))))
                Dim lines = Math.Max(1, CInt(Math.Floor(heightPt / (fontSize * 1.25))))
                Return Math.Min(4000, charsPerLine * lines)
            Catch
                Return 200
            End Try
        End Function

        Private Shared Function EstimateLineCount(shape As PowerPoint.Shape, fontSize As Double) As Integer
            Try
                Dim heightPt = Math.Max(16.0, CDbl(shape.Height))
                Return Math.Min(20, Math.Max(1, CInt(Math.Floor(heightPt / (fontSize * 1.4)))))
            Catch
                Return 1
            End Try
        End Function

        Private Shared Function SafeChartPointCount(shape As PowerPoint.Shape) As Integer
            Try
                Dim series = shape.Chart.SeriesCollection
                If series IsNot Nothing AndAlso CInt(series.Count) > 0 Then
                    Dim categories = TryCast(series(1).XValues, Array)
                    If categories IsNot Nothing Then Return categories.Length
                End If
            Catch
            End Try
            Return 0
        End Function

        Private Shared Function SafeTableRowCount(shape As PowerPoint.Shape) As Integer
            Try
                Return Math.Max(0, CInt(shape.Table.Rows.Count) - 1)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function ClassifyArchetype(slide As PowerPoint.Slide, slots As List(Of TemplateSlot)) As String
            If slots Is Nothing Then Return TemplateArchetypes.Custom

            Dim hasChart = slots.Any(Function(slot) slot.Role = TemplateSlotRoles.Chart)
            Dim hasTable = slots.Any(Function(slot) slot.Role = TemplateSlotRoles.Table)
            Dim hasImage = slots.Any(Function(slot) slot.Role = TemplateSlotRoles.Image)
            Dim hasTitle = slots.Any(Function(slot) slot.Role = TemplateSlotRoles.Title)
            Dim subtitleCount = slots.Where(Function(slot) slot.Role = TemplateSlotRoles.Subtitle).Count()
            Dim bodyCount = slots.Where(Function(slot) slot.Role = TemplateSlotRoles.Body OrElse
                                                       slot.Role = TemplateSlotRoles.Bullets).Count()

            Dim texts = ReadSlideTexts(slide)
            If texts.Any(Function(text) IsClosingText(text)) Then Return TemplateArchetypes.Closing
            If hasChart Then Return TemplateArchetypes.Chart
            If hasTable Then Return TemplateArchetypes.Table
            If hasImage AndAlso (hasTitle OrElse bodyCount > 0) Then Return TemplateArchetypes.ImageText

            Dim kpiCount = texts.Where(Function(text) IsKpiText(text)).Count()
            If kpiCount >= 3 AndAlso Not hasChart AndAlso Not hasTable Then Return TemplateArchetypes.Kpi
            If bodyCount >= 2 AndAlso AreInColumns(slide, slots) Then Return TemplateArchetypes.TwoColumn
            If bodyCount = 1 AndAlso texts.Any(Function(text) text.Length >= 80) AndAlso slots.Count <= 6 Then
                Return TemplateArchetypes.Quote
            End If
            If bodyCount > 0 Then Return TemplateArchetypes.Bullets
            If hasTitle AndAlso subtitleCount > 0 Then Return TemplateArchetypes.Cover
            If hasTitle Then Return TemplateArchetypes.Section
            Return TemplateArchetypes.Custom
        End Function

        Private Shared Function ReadSlideTexts(slide As PowerPoint.Slide) As List(Of String)
            Dim texts As New List(Of String)()
            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For index = 1 To Math.Min(shapes.Count, MaxSlots)
                    Dim shape As PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(index)
                        If shape.HasTextFrame = MsoTriState.msoTrue AndAlso shape.TextFrame.HasText = MsoTriState.msoTrue Then
                            Dim text = Convert.ToString(shape.TextFrame.TextRange.Text)
                            If Not String.IsNullOrWhiteSpace(text) Then texts.Add(text.Trim())
                        End If
                    Catch
                    Finally
                        ReleaseCom(shape)
                    End Try
                Next
            Catch
            Finally
                ReleaseCom(shapes)
            End Try
            Return texts
        End Function

        Private Shared Function IsClosingText(text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return False
            Dim lower = text.ToLowerInvariant()
            Dim markers = {"спасибо", "вопросы", "вопрос", "контакты", "контакт", "thank you", "questions", "contact"}
            Return markers.Any(Function(marker) lower.Contains(marker))
        End Function

        Private Shared Function IsKpiText(text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) OrElse text.Length > 24 Then Return False
            Return text.Any(Function(ch) Char.IsDigit(ch)) OrElse text.Contains("%") OrElse text.Contains("₽") OrElse text.Contains("$")
        End Function

        Private Shared Function AreInColumns(slide As PowerPoint.Slide, slots As List(Of TemplateSlot)) As Boolean
            Dim bodySlots = slots.Where(Function(slot) slot.Role = TemplateSlotRoles.Body OrElse
                                                    slot.Role = TemplateSlotRoles.Bullets).ToList()
            If bodySlots.Count < 2 Then Return False
            Try
                Dim boxes As New List(Of Tuple(Of Double, Double))()
                Dim shapes = slide.Shapes
                For Each slot In bodySlots
                    Try
                        Dim shape = shapes(slot.ShapeId)
                        boxes.Add(Tuple.Create(CDbl(shape.Left), CDbl(shape.Left) + CDbl(shape.Width)))
                    Catch
                    End Try
                Next
                If boxes.Count < 2 Then Return False
                For i = 0 To boxes.Count - 2
                    For j = i + 1 To boxes.Count - 1
                        Dim overlap = Math.Min(boxes(i).Item2, boxes(j).Item2) - Math.Max(boxes(i).Item1, boxes(j).Item1)
                        Dim narrower = Math.Min(boxes(i).Item2 - boxes(i).Item1, boxes(j).Item2 - boxes(j).Item1)
                        If narrower > 0 AndAlso overlap > narrower * 0.7 Then Return False
                    Next
                Next
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ComputeFingerprint(slide As PowerPoint.Slide) As String
            Dim parts As New List(Of String)()
            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                Dim slideWidth = Math.Max(1.0, CDbl(slide.Parent.PageSetup.SlideWidth))
                Dim slideHeight = Math.Max(1.0, CDbl(slide.Parent.PageSetup.SlideHeight))
                Dim entries As New List(Of Tuple(Of Double, Double, String))()
                For index = 1 To Math.Min(shapes.Count, MaxSlots)
                    Dim shape As PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(index)
                        Dim kind = ShapeKind(shape)
                        Dim left = CDbl(shape.Left) / slideWidth
                        Dim top = CDbl(shape.Top) / slideHeight
                        Dim width = CDbl(shape.Width) / slideWidth
                        Dim height = CDbl(shape.Height) / slideHeight
                        Dim descriptor = String.Format(CultureInfo.InvariantCulture,
                            "{0}|{1:F3}|{2:F3}|{3:F3}|{4:F3}", kind, left, top, width, height)
                        entries.Add(Tuple.Create(top, left, descriptor))
                    Catch
                    Finally
                        ReleaseCom(shape)
                    End Try
                Next
                For Each entry In entries.OrderBy(Function(item) item.Item1).ThenBy(Function(item) item.Item2)
                    parts.Add(entry.Item3)
                Next
            Catch
            Finally
                ReleaseCom(shapes)
            End Try

            If parts.Count = 0 Then Return ""
            Using sha = SHA256.Create()
                Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(String.Join(";", parts)))
                Return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Function ShapeKind(shape As PowerPoint.Shape) As String
            Try
                If shape.HasChart = MsoTriState.msoTrue Then Return "chart"
                If shape.HasTable = MsoTriState.msoTrue Then Return "table"
                If shape.Type = MsoShapeType.msoPicture Then Return "picture"
                If shape.Type = MsoShapeType.msoGroup Then Return "group"
                If shape.Type = MsoShapeType.msoPlaceholder Then Return "placeholder"
                If shape.HasTextFrame = MsoTriState.msoTrue AndAlso shape.TextFrame.HasText = MsoTriState.msoTrue Then Return "text"
                Return "decor"
            Catch
                Return "other"
            End Try
        End Function

        Private Shared Function ReadThemeTokens(presentation As PowerPoint.Presentation) As Dictionary(Of String, String)
            Dim tokens As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim theme As Object = Nothing
            Try
                theme = presentation.SlideMaster.Theme
                Dim colors = theme.ThemeColorScheme
                AddThemeColor(tokens, "accent1", colors, MsoThemeColorIndex.msoThemeColorAccent1)
                AddThemeColor(tokens, "accent2", colors, MsoThemeColorIndex.msoThemeColorAccent2)
                AddThemeColor(tokens, "accent3", colors, MsoThemeColorIndex.msoThemeColorAccent3)
                AddThemeColor(tokens, "accent4", colors, MsoThemeColorIndex.msoThemeColorAccent4)
                AddThemeColor(tokens, "accent5", colors, MsoThemeColorIndex.msoThemeColorAccent5)
                AddThemeColor(tokens, "accent6", colors, MsoThemeColorIndex.msoThemeColorAccent6)
                AddThemeColor(tokens, "text1", colors, MsoThemeColorIndex.msoThemeColorText1)
                AddThemeColor(tokens, "text2", colors, MsoThemeColorIndex.msoThemeColorText2)
                AddThemeColor(tokens, "background1", colors, MsoThemeColorIndex.msoThemeColorBackground1)
                AddThemeColor(tokens, "background2", colors, MsoThemeColorIndex.msoThemeColorBackground2)

                Dim fonts = theme.ThemeFontScheme
                Try
                    tokens("majorFont") = Convert.ToString(fonts.MajorFont(ThemeFontLatin).Name)
                Catch
                End Try
                Try
                    tokens("minorFont") = Convert.ToString(fonts.MinorFont(ThemeFontLatin).Name)
                Catch
                End Try
            Catch
                ' Тема недоступна — макет всё равно пригоден, tokens остаются частичными.
            Finally
                ReleaseCom(theme)
            End Try
            Return tokens
        End Function

        Private Shared Sub AddThemeColor(tokens As Dictionary(Of String, String),
                                         key As String,
                                         colors As Object,
                                         index As MsoThemeColorIndex)
            Try
                Dim color = colors.Colors(index)
                tokens(key) = ToHexColor(CInt(color.RGB))
            Catch
            End Try
        End Sub

        Private Shared Function ToHexColor(bgr As Integer) As String
            Dim r = bgr And &HFF
            Dim g = (bgr >> 8) And &HFF
            Dim b = (bgr >> 16) And &HFF
            Return $"#{r:X2}{g:X2}{b:X2}"
        End Function

        Private Shared Function ExportPreview(presentation As PowerPoint.Presentation,
                                              slide As PowerPoint.Slide,
                                              previewPath As String,
                                              warnings As List(Of String)) As Boolean
            Try
                Dim slideWidth = Math.Max(1.0, CDbl(presentation.PageSetup.SlideWidth))
                Dim slideHeight = Math.Max(1.0, CDbl(presentation.PageSetup.SlideHeight))
                Dim widthPx = PreviewWidthPx
                Dim heightPx = Math.Max(1, CInt(Math.Round(PreviewWidthPx * slideHeight / slideWidth)))
                slide.Export(previewPath, "PNG", widthPx, heightPx)
                Return File.Exists(previewPath)
            Catch ex As Exception
                warnings.Add("Slide.Export: " & ex.Message)
                Return False
            End Try
        End Function

        Private Shared Function IsEmptySlide(slide As PowerPoint.Slide) As Boolean
            Try
                Dim textLength = 0
                Dim meaningful = False
                Dim shapes = slide.Shapes
                For index = 1 To Math.Min(shapes.Count, 30)
                    Dim shape = shapes(index)
                    Try
                        If shape.HasChart = MsoTriState.msoTrue OrElse shape.HasTable = MsoTriState.msoTrue Then
                            meaningful = True
                            Exit For
                        End If
                        If shape.HasTextFrame = MsoTriState.msoTrue AndAlso shape.TextFrame.HasText = MsoTriState.msoTrue Then
                            textLength += Convert.ToString(shape.TextFrame.TextRange.Text).Trim().Length
                        End If
                    Finally
                        ReleaseCom(shape)
                    End Try
                Next
                Return Not meaningful AndAlso textLength < 10
            Catch
                Return False
            End Try
        End Function

        Private Shared Function SafeAppVersion(app As PowerPoint.Application) As String
            Try
                Return Convert.ToString(app.Version)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function CreateTempDirectory() As String
            Dim folder = Path.Combine(Path.GetTempPath(), "office-ai-decktpl", Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Private Shared Sub TryDeleteFile(filePath As String)
            Try
                If Not String.IsNullOrWhiteSpace(filePath) AndAlso File.Exists(filePath) Then File.Delete(filePath)
            Catch ex As Exception
                Debug.WriteLine($"[DeckTemplateExtractor] Не удалось удалить {filePath}: {ex.Message}")
            End Try
        End Sub

        Private Shared Sub TryDeleteDirectory(folderPath As String)
            Try
                If Not String.IsNullOrWhiteSpace(folderPath) Then
                    If Directory.Exists(folderPath) Then Directory.Delete(folderPath, True)
                    Dim parent = Path.GetDirectoryName(folderPath)
                    If Not String.IsNullOrWhiteSpace(parent) AndAlso
                       Directory.Exists(parent) AndAlso
                       Not Directory.EnumerateFileSystemEntries(parent).Any() Then
                        Directory.Delete(parent)
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[DeckTemplateExtractor] Не удалось удалить {folderPath}: {ex.Message}")
            End Try
        End Sub

        Private Shared Sub ReleaseCom(comObject As Object)
            Try
                ShareRibbon.ComObjectHelper.ReleaseComObject(comObject)
            Catch
            End Try
        End Sub

    End Class

End Namespace
