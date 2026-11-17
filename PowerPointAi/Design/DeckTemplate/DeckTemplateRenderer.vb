' PowerPointAi/Design/DeckTemplate/DeckTemplateRenderer.vb
' Вставка макета в активную колоду и заполнение слотов контентом (clone-and-fill).

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon.Agent.DeckTemplates

Namespace Design.DeckTemplate

    Public Class DeckTemplateRenderResult
        Public Property Success As Boolean
        Public Property InsertedSlideIndex As Integer
        Public Property InsertedSlideId As Integer
        Public Property SlotsFilled As Integer
        Public Property Warnings As New List(Of String)()
        Public Property ErrorMessage As String = ""
    End Class

    Public Class DeckBuildResult
        Public Property Success As Boolean
        Public Property InsertedSlideIndexes As New List(Of Integer)()
        Public Property InsertedSlideIds As New List(Of Integer)()
        Public Property Warnings As New List(Of String)()
        Public Property ErrorMessage As String = ""
    End Class

    Public NotInheritable Class DeckTemplateRenderer
        Private Sub New()
        End Sub

        ''' <summary>Вставляет сохранённый макет в активную колоду и заполняет слоты.</summary>
        Public Shared Function InsertFromTemplate(app As PowerPoint.Application,
                                                  target As PowerPoint.Presentation,
                                                  template As SlideTemplateManifest,
                                                  spec As SlideSpec,
                                                  insertAt As Integer) As DeckTemplateRenderResult
            Dim result As New DeckTemplateRenderResult()
            If template Is Nothing Then
                result.ErrorMessage = "Макет не передан"
                Return result
            End If
            If Not File.Exists(template.TemplateFullPath) Then
                result.ErrorMessage = $"Файл макета недоступен: {template.Id}"
                Return result
            End If

            Dim source As PowerPoint.Presentation = Nothing
            Try
                source = app.Presentations.Open(template.TemplateFullPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse)
                If source.Slides.Count = 0 Then
                    result.ErrorMessage = "В файле макета нет слайдов"
                    Return result
                End If
                source.Slides(1).Copy()
                Return PasteAndFill(target, spec, insertAt, result)
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "Вставка макета: " & ex.Message
                Return result
            Finally
                Try
                    If source IsNot Nothing Then source.Close()
                Catch
                End Try
                ReleaseCom(source)
            End Try
        End Function

        ''' <summary>Клонирует слайд текущей колоды и заполняет его контентом.</summary>
        Public Shared Function InsertFromSlide(app As PowerPoint.Application,
                                               target As PowerPoint.Presentation,
                                               sourceSlideIndex As Integer,
                                               spec As SlideSpec,
                                               insertAt As Integer) As DeckTemplateRenderResult
            Dim result As New DeckTemplateRenderResult()
            If target Is Nothing Then
                result.ErrorMessage = "Нет активной презентации"
                Return result
            End If
            If sourceSlideIndex <= 0 OrElse sourceSlideIndex > target.Slides.Count Then
                result.ErrorMessage = $"Слайд-источник {sourceSlideIndex} вне диапазона 1..{target.Slides.Count}"
                Return result
            End If

            Try
                target.Slides(sourceSlideIndex).Copy()
                Return PasteAndFill(target, spec, insertAt, result)
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "Клонирование слайда: " & ex.Message
                Return result
            End Try
        End Function

        Private Shared Function PasteAndFill(target As PowerPoint.Presentation,
                                             spec As SlideSpec,
                                             insertAt As Integer,
                                             result As DeckTemplateRenderResult) As DeckTemplateRenderResult
            Dim beforeCount = CInt(target.Slides.Count)
            If insertAt <= 0 Then insertAt = beforeCount + 1
            If insertAt > beforeCount + 1 Then insertAt = beforeCount + 1

            Dim pasted As Object = Nothing
            Try
                pasted = target.Slides.Paste(insertAt)
            Catch
                pasted = target.Slides.Paste()
            End Try

            Dim newIndex = insertAt
            Try
                If pasted IsNot Nothing AndAlso CInt(pasted.Count) > 0 Then
                    newIndex = CInt(pasted(1).SlideIndex)
                End If
            Catch
            End Try
            If newIndex <= 0 OrElse newIndex > target.Slides.Count Then newIndex = target.Slides.Count

            Dim slide As PowerPoint.Slide = Nothing
            Dim slideId As Integer = 0
            Try
                slide = target.Slides(newIndex)
                slideId = CInt(slide.SlideID)
                result.InsertedSlideIndex = newIndex
                result.InsertedSlideId = slideId

                Dim filled As Integer = 0
                Dim ok = FillSlide(slide, spec, result.Warnings, filled, result.ErrorMessage)
                result.SlotsFilled = filled
                If Not ok Then
                    DeleteInsertedSlide(target, slideId)
                    result.Success = False
                    result.InsertedSlideIndex = 0
                    result.InsertedSlideId = 0
                    return result
                End If

                Try
                    target.Application.ActiveWindow.View.GotoSlide(newIndex)
                Catch
                End Try
                result.Success = True
                Return result
            Catch ex As Exception
                DeleteInsertedSlide(target, slideId)
                result.Success = False
                result.InsertedSlideIndex = 0
                result.InsertedSlideId = 0
                result.ErrorMessage = "Заполнение макета: " & ex.Message
                Return result
            Finally
                ReleaseCom(slide)
            End Try
        End Function

        Private Shared Function FillSlide(slide As PowerPoint.Slide,
                                          spec As SlideSpec,
                                          warnings As List(Of String),
                                          ByRef filled As Integer,
                                          ByRef errorMessage As String) As Boolean
            errorMessage = ""
            If spec Is Nothing Then spec = New SlideSpec()
            spec.Normalize()
            filled = 0

            Dim slots = DeckTemplateExtractor.DescribeSlotsForSlide(slide)
            If slots.Count = 0 Then
                errorMessage = "В макете не найдено ни одного слота"
                Return False
            End If

            If Not spec.HasAnyContent() Then
                warnings.Add("Контент не передан: макет вставлен без заполнения")
                Return True
            End If

            If Not String.IsNullOrWhiteSpace(spec.Title) Then
                Dim titleSlot = FindSlot(slots, TemplateSlotRoles.Title)
                If titleSlot Is Nothing Then
                    errorMessage = "В макете нет слота заголовка (title)"
                    Return False
                End If
                If Not SetSlotText(slide, titleSlot, spec.Title, warnings, False, filled) Then
                    errorMessage = "Слот заголовка не заполнен"
                    Return False
                End If
            End If

            If Not String.IsNullOrWhiteSpace(spec.Subtitle) Then
                Dim subtitleSlot = FindSlot(slots, TemplateSlotRoles.Subtitle)
                If subtitleSlot Is Nothing Then
                    warnings.Add("В макете нет слота subtitle: подзаголовок не размещён")
                Else
                    SetSlotText(slide, subtitleSlot, spec.Subtitle, warnings, False, filled)
                End If
            End If

            Dim bulletsText = If(spec.Bullets IsNot Nothing, String.Join(vbCr, spec.Bullets), "")
            Dim bodyContent = If(String.IsNullOrWhiteSpace(bulletsText), spec.Body, bulletsText)
            If Not String.IsNullOrWhiteSpace(bodyContent) Then
                Dim bulletsSlot = FindSlot(slots, TemplateSlotRoles.Bullets)
                Dim bodySlot = FindSlot(slots, TemplateSlotRoles.Body)
                Dim targetSlot = If(spec.Bullets.Count > 0 AndAlso bulletsSlot IsNot Nothing, bulletsSlot, If(bodySlot, bulletsSlot))
                If targetSlot Is Nothing Then
                    errorMessage = "В макете нет слота для текста (body/bullets)"
                    Return False
                End If

                Dim text = bodyContent
                If spec.Bullets.Count > 0 AndAlso targetSlot.ItemCount > 0 AndAlso spec.Bullets.Count > targetSlot.ItemCount Then
                    text = String.Join(vbCr, spec.Bullets.Take(targetSlot.ItemCount))
                    warnings.Add($"Список сокращён до {targetSlot.ItemCount} пунктов по вместимости макета")
                End If
                If Not SetSlotText(slide, targetSlot, text, warnings, spec.Bullets.Count > 0, filled) Then
                    errorMessage = "Слот текста не заполнен"
                    Return False
                End If
            End If

            If spec.Table IsNot Nothing Then
                Dim tableSlot = FindSlot(slots, TemplateSlotRoles.Table)
                If tableSlot Is Nothing Then
                    errorMessage = "В макете нет слота таблицы (table)"
                    Return False
                End If
                If Not FillTable(slide, tableSlot, spec.Table, warnings) Then
                    errorMessage = "Не удалось заполнить таблицу макета"
                    Return False
                End If
                filled += 1
            End If

            If spec.Chart IsNot Nothing Then
                Dim chartSlot = FindSlot(slots, TemplateSlotRoles.Chart)
                If chartSlot Is Nothing Then
                    errorMessage = "В макете нет слота диаграммы (chart)"
                    Return False
                End If
                Dim chartError As String = ""
                If Not FillChart(slide, chartSlot, spec.Chart, warnings, chartError) Then
                    errorMessage = chartError
                    Return False
                End If
                filled += 1
            End If

            If spec.Kpis IsNot Nothing AndAlso spec.Kpis.Count > 0 Then
                Dim kpiSlots = slots.Where(Function(slot) slot.Role = TemplateSlotRoles.Kpi OrElse
                                                      slot.Role = TemplateSlotRoles.Body).ToList()
                If kpiSlots.Count = 0 Then
                    errorMessage = "В макете нет слотов для KPI"
                    Return False
                End If
                Dim applied = 0
                For Each kpi In spec.Kpis
                    If applied >= kpiSlots.Count Then
                        warnings.Add("KPI: не хватило слотов, часть значений пропущена")
                        Exit For
                    End If
                    Dim text = $"{kpi.Label}: {kpi.Value}".Trim()
                    If text.Length = 0 OrElse text = ":" Then Continue For
                    SetSlotText(slide, kpiSlots(applied), text, warnings, False, filled)
                    applied += 1
                Next
            End If

            If Not String.IsNullOrWhiteSpace(spec.Caption) Then
                Dim captionSlot = FindSlot(slots, TemplateSlotRoles.Caption)
                If captionSlot Is Nothing Then
                    warnings.Add("В макете нет слота caption")
                Else
                    SetSlotText(slide, captionSlot, spec.Caption, warnings, False, filled)
                End If
            End If

            If Not String.IsNullOrWhiteSpace(spec.Footer) Then
                Dim footerSlot = FindSlot(slots, TemplateSlotRoles.Footer)
                If footerSlot Is Nothing Then
                    warnings.Add("В макете нет слота footer")
                Else
                    SetSlotText(slide, footerSlot, spec.Footer, warnings, False, filled)
                End If
            End If

            If spec.Image IsNot Nothing Then
                warnings.Add("Вставка изображения в слот пока не поддерживается: слот оставлен как в макете")
            End If

            Return True
        End Function

        Private Shared Function SetSlotText(slide As PowerPoint.Slide,
                                            slot As TemplateSlot,
                                            value As String,
                                            warnings As List(Of String),
                                            isBullets As Boolean,
                                            ByRef filled As Integer) As Boolean
            Dim shape = FindShapeById(slide, slot.ShapeId)
            If shape Is Nothing Then
                warnings.Add($"Слот {slot.Role}: фигура {slot.ShapeId} не найдена")
                Return False
            End If

            Dim text = If(value, "").Trim()
            If slot.CapacityChars > 0 AndAlso text.Length > slot.CapacityChars Then
                text = TruncateAtWord(text, slot.CapacityChars)
                warnings.Add($"Текст слота {slot.Role} сокращён до {slot.CapacityChars} символов")
            End If

            Try
                shape.TextFrame2.TextRange.Text = text
                If isBullets Then
                    Try
                        shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible = MsoTriState.msoTrue
                    Catch
                    End Try
                End If
                filled += 1
                Return True
            Catch ex As Exception
                warnings.Add($"Слот {slot.Role}: не удалось записать текст ({ex.Message})")
                Return False
            Finally
                ReleaseCom(shape)
            End Try
        End Function

        Private Shared Function FillTable(slide As PowerPoint.Slide,
                                          slot As TemplateSlot,
                                          tableSpec As SlideTableSpec,
                                          warnings As List(Of String)) As Boolean
            Dim shape = FindShapeById(slide, slot.ShapeId)
            Dim table As Object = Nothing
            Try
                If shape Is Nothing OrElse shape.HasTable <> MsoTriState.msoTrue Then Return False
                table = shape.Table
                Dim maxRows = Math.Max(0, CInt(table.Rows.Count) - 1)
                Dim maxCols = CInt(table.Columns.Count)
                If tableSpec.RowCount > maxRows Then
                    warnings.Add($"Таблица: {tableSpec.RowCount} строк не помещаются в {maxRows}, лишние пропущены")
                End If
                If tableSpec.ColumnCount > maxCols Then
                    warnings.Add($"Таблица: {tableSpec.ColumnCount} столбцов не помещаются в {maxCols}, лишние пропущены")
                End If

                Dim rowLimit = Math.Min(tableSpec.RowCount, maxRows)
                Dim colLimit = Math.Min(tableSpec.ColumnCount, maxCols)
                For column = 1 To Math.Min(If(tableSpec.Headers Is Nothing, 0, tableSpec.Headers.Count), colLimit)
                    Try
                        table.Cell(1, column).Shape.TextFrame.TextRange.Text = If(tableSpec.Headers(column - 1), "")
                    Catch
                    End Try
                Next
                For row = 1 To rowLimit
                    Dim values = tableSpec.Rows(row - 1)
                    For column = 1 To Math.Min(If(values Is Nothing, 0, values.Count), colLimit)
                        Try
                            table.Cell(row + 1, column).Shape.TextFrame.TextRange.Text = If(values(column - 1), "")
                        Catch
                        End Try
                    Next
                Next
                Return True
            Finally
                ReleaseCom(table)
                ReleaseCom(shape)
            End Try
        End Function

        Private Shared Function FillChart(slide As PowerPoint.Slide,
                                          slot As TemplateSlot,
                                          chartSpec As SlideChartSpec,
                                          warnings As List(Of String),
                                          ByRef errorMessage As String) As Boolean
            errorMessage = ""
            Dim shape = FindShapeById(slide, slot.ShapeId)
            Dim chart As Object = Nothing
            Dim seriesCollection As Object = Nothing
            Dim series As Object = Nothing
            Try
                If shape Is Nothing OrElse shape.HasChart <> MsoTriState.msoTrue Then
                    errorMessage = "Слот диаграммы не содержит Chart"
                    Return False
                End If

                Dim values = chartSpec.Values
                If values Is Nothing OrElse values.Count = 0 Then
                    errorMessage = "Chart spec не содержит values"
                    Return False
                End If

                chart = shape.Chart
                seriesCollection = chart.SeriesCollection
                If CInt(seriesCollection.Count) = 0 Then
                    errorMessage = "У диаграммы нет рядов"
                    Return False
                End If
                series = seriesCollection(1)

                If chartSpec.Categories IsNot Nothing AndAlso chartSpec.Categories.Count > 0 Then
                    If chartSpec.Categories.Count <> values.Count Then
                        errorMessage = $"Chart: {chartSpec.Categories.Count} категорий и {values.Count} значений"
                        Return False
                    End If
                    series.XValues = chartSpec.Categories.ToArray()
                Else
                    Dim existingCount As Integer = 0
                    Try
                        Dim existing = TryCast(series.XValues, Array)
                        If existing IsNot Nothing Then existingCount = existing.Length
                    Catch
                    End Try
                    If existingCount > 0 AndAlso existingCount <> values.Count Then
                        errorMessage = $"Chart: у макета {existingCount} категорий, передано {values.Count} значений"
                        Return False
                    End If
                End If

                series.Values = values.Select(Function(value) CObj(value)).ToArray()
                If Not String.IsNullOrWhiteSpace(chartSpec.SeriesName) Then series.Name = chartSpec.SeriesName
                Return True
            Finally
                ReleaseCom(series)
                ReleaseCom(seriesCollection)
                ReleaseCom(chart)
                ReleaseCom(shape)
            End Try
        End Function

        ''' <summary>
        ''' Строит колоду по плану: последовательно клонирует макеты/слайды и заполняет их.
        ''' При любой ошибке удаляет все уже вставленные слайды (rollback пачки).
        ''' </summary>
        Public Shared Function BuildDeck(app As PowerPoint.Application,
                                         target As PowerPoint.Presentation,
                                         steps As List(Of DeckPlanStep)) As DeckBuildResult
            Dim result As New DeckBuildResult()
            If target Is Nothing Then
                result.ErrorMessage = "Нет целевой презентации"
                Return result
            End If
            If steps Is Nothing OrElse steps.Count = 0 Then
                result.ErrorMessage = "План колоды пуст"
                Return result
            End If

            For Each planStep In steps
                Dim template = If(planStep.Match Is Nothing, Nothing, planStep.Match.Template)
                If template Is Nothing Then
                    RollbackInserted(target, result.InsertedSlideIds)
                    result.Success = False
                    result.ErrorMessage = $"Слайд {planStep.Index}: не найден подходящий макет"
                    Return result
                End If

                Dim stepResult As DeckTemplateRenderResult
                If Not String.IsNullOrWhiteSpace(template.StorageFolder) Then
                    stepResult = InsertFromTemplate(app, target, template, planStep.Spec, CInt(target.Slides.Count) + 1)
                Else
                    Dim sourceIndex = If(template.Source Is Nothing, 0, template.Source.SlideIndex)
                    If sourceIndex <= 0 Then
                        RollbackInserted(target, result.InsertedSlideIds)
                        result.Success = False
                        result.ErrorMessage = $"Слайд {planStep.Index}: неизвестный слайд-источник"
                        Return result
                    End If
                    stepResult = InsertFromSlide(app, target, sourceIndex, planStep.Spec, CInt(target.Slides.Count) + 1)
                End If

                If stepResult Is Nothing OrElse Not stepResult.Success Then
                    RollbackInserted(target, result.InsertedSlideIds)
                    result.Success = False
                    result.ErrorMessage = $"Слайд {planStep.Index}: " &
                                          If(stepResult Is Nothing, "не удалось вставить", stepResult.ErrorMessage)
                    Return result
                End If

                result.InsertedSlideIds.Add(stepResult.InsertedSlideId)
                result.InsertedSlideIndexes.Add(stepResult.InsertedSlideIndex)
                If stepResult.Warnings IsNot Nothing Then result.Warnings.AddRange(stepResult.Warnings)
            Next

            result.Success = True
            Return result
        End Function

        Private Shared Sub RollbackInserted(target As PowerPoint.Presentation, slideIds As List(Of Integer))
            If target Is Nothing OrElse slideIds Is Nothing Then Return
            For Each slideId In slideIds
                DeleteInsertedSlide(target, slideId)
            Next
            slideIds.Clear()
        End Sub

        Private Shared Function FindSlot(slots As List(Of TemplateSlot), role As String) As TemplateSlot
            Dim candidates = slots.Where(Function(slot) String.Equals(slot.Role, role, StringComparison.OrdinalIgnoreCase)).ToList()
            If candidates.Count = 0 Then Return Nothing
            ' Несколько title/body-слотов в ручных слайдах: берём самый верхний,
            ' иначе заголовок может попасть в подпись категории.
            If role = TemplateSlotRoles.Title OrElse
               role = TemplateSlotRoles.Subtitle OrElse
               role = TemplateSlotRoles.Body OrElse
               role = TemplateSlotRoles.Bullets Then
                Return candidates.OrderBy(Function(slot) slot.Top).ThenBy(Function(slot) slot.Left).First()
            End If
            Return candidates.First()
        End Function

        Private Shared Function FindShapeById(slide As PowerPoint.Slide, shapeId As Integer) As PowerPoint.Shape
            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For index = 1 To CInt(shapes.Count)
                    Dim shape As PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(index)
                        If CInt(shape.Id) = shapeId Then Return shape
                    Catch
                    End Try
                Next
                Return Nothing
            Finally
                ReleaseCom(shapes)
            End Try
        End Function

        Private Shared Sub DeleteInsertedSlide(target As PowerPoint.Presentation, slideId As Integer)
            If slideId <= 0 Then Return
            For index = CInt(target.Slides.Count) To 1 Step -1
                Dim slide As PowerPoint.Slide = Nothing
                Try
                    slide = target.Slides(index)
                    If CInt(slide.SlideID) = slideId Then
                        slide.Delete()
                        Return
                    End If
                Catch
                Finally
                    ReleaseCom(slide)
                End Try
            Next
        End Sub

        Private Shared Function TruncateAtWord(value As String, limit As Integer) As String
            If String.IsNullOrEmpty(value) OrElse value.Length <= limit Then Return value
            Dim cut = value.Substring(0, Math.Max(1, limit))
            Dim space = cut.LastIndexOf(" "c)
            If space > limit \ 2 Then cut = cut.Substring(0, space)
            Return cut.TrimEnd() & "…"
        End Function

        Private Shared Sub ReleaseCom(comObject As Object)
            Try
                ShareRibbon.ComObjectHelper.ReleaseComObject(comObject)
            Catch
            End Try
        End Sub

    End Class

End Namespace
