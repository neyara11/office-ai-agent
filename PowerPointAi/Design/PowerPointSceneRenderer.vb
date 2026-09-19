Imports System.Drawing
Imports System.IO
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon

Namespace Design

    Public Class SceneRenderResult
        Public Property Slide As PowerPoint.Slide
        Public Property CreatedShapeCount As Integer
        Public Property Warnings As New List(Of String)()
    End Class

    Public NotInheritable Class PowerPointSceneRenderer
        Private Const ShapeTagPrefix As String = "office-ai-design:"

        Private Sub New()
        End Sub

        Public Shared Function Render(presentation As PowerPoint.Presentation,
                                      spec As SlideDesignSpec,
                                      plan As SlideRenderPlan,
                                      insertIndex As Integer,
                                      tokens As DesignTokens) As SceneRenderResult
            If presentation Is Nothing Then Throw New ArgumentNullException(NameOf(presentation))
            Dim result As New SceneRenderResult()
            Dim slides As PowerPoint.Slides = Nothing
            Dim slide As PowerPoint.Slide = Nothing
            Try
                slides = presentation.Slides
                slide = slides.Add(insertIndex, PowerPoint.PpSlideLayout.ppLayoutBlank)
                result.Slide = slide
                ApplyBackground(slide, plan.Background)

                For Each node In plan.Nodes
                    Dim created = RenderNode(slide, node, tokens, result.Warnings)
                    If created Then result.CreatedShapeCount += 1
                Next
                If Not WriteNotes(slide, plan.Notes) Then
                    result.Warnings.Add("Speaker notes skipped because the notes placeholder was unavailable")
                End If
                Return result
            Catch
                If slide IsNot Nothing Then
                    Try
                        slide.Delete()
                    Catch
                    Finally
                        ComObjectHelper.ReleaseComObject(slide)
                        result.Slide = Nothing
                    End Try
                End If
                Throw
            Finally
                ComObjectHelper.ReleaseComObject(slides)
            End Try
        End Function

        Private Shared Function RenderNode(slide As PowerPoint.Slide,
                                           node As SceneNode,
                                           tokens As DesignTokens,
                                           warnings As List(Of String)) As Boolean
            If node Is Nothing OrElse node.Bounds Is Nothing Then Return False
            Dim shapes As PowerPoint.Shapes = Nothing
            Dim shape As PowerPoint.Shape = Nothing
            Try
                shapes = slide.Shapes
                Select Case node.Kind
                    Case "round-rect"
                        shape = shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle,
                                                node.Bounds.X, node.Bounds.Y, node.Bounds.Width, node.Bounds.Height)
                        ApplyCornerRadius(shape, node)
                        ApplyShapeStyle(shape, node, addShadow:=node.Shadow)
                    Case "rect"
                        shape = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                node.Bounds.X, node.Bounds.Y, node.Bounds.Width, node.Bounds.Height)
                        ApplyShapeStyle(shape, node, addShadow:=False)
                    Case "circle"
                        shape = shapes.AddShape(MsoAutoShapeType.msoShapeOval,
                                                node.Bounds.X, node.Bounds.Y, node.Bounds.Width, node.Bounds.Height)
                        ApplyShapeStyle(shape, node, addShadow:=False)
                    Case "line"
                        shape = shapes.AddLine(node.Bounds.X,
                                               node.Bounds.Y,
                                               node.Bounds.X + node.Bounds.Width,
                                               node.Bounds.Y + node.Bounds.Height)
                        ApplyLineStyle(shape, node, tokens)
                    Case "image"
                        If String.IsNullOrWhiteSpace(node.ImagePath) OrElse Not File.Exists(node.ImagePath) Then
                            warnings.Add($"Image skipped because file is unavailable: {node.ImagePath}")
                            Return False
                        End If
                        shape = shapes.AddPicture(node.ImagePath,
                                                  MsoTriState.msoFalse,
                                                  MsoTriState.msoTrue,
                                                  node.Bounds.X, node.Bounds.Y,
                                                  -1, -1)
                        CropImageToFill(shape, node.Bounds, warnings)
                    Case Else
                        shape = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                                  node.Bounds.X, node.Bounds.Y, node.Bounds.Width, node.Bounds.Height)
                        ApplyTextStyle(shape, node, tokens)
                End Select

                If shape Is Nothing Then Return False
                Try
                    shape.Name = BuildShapeName(node.Id)
                Catch
                End Try
                Try
                    shape.AlternativeText = ShapeTagPrefix & If(node.Id, "node")
                Catch
                End Try
                Return True
            Finally
                ComObjectHelper.ReleaseComObject(shape)
                ComObjectHelper.ReleaseComObject(shapes)
            End Try
        End Function

        Private Shared Sub ApplyBackground(slide As PowerPoint.Slide, colorValue As String)
            Dim background As PowerPoint.ShapeRange = Nothing
            Dim fill As Object = Nothing
            Dim color As Object = Nothing
            Try
                slide.FollowMasterBackground = MsoTriState.msoFalse
                background = slide.Background
                ' WPS may expose Fill only through IDispatch and reject Office.FillFormat QI.
                Dim backgroundObject As Object = background
                fill = backgroundObject.Fill
                fill.Solid()
                color = fill.ForeColor
                color.RGB = ToOle(colorValue, "#FFFFFF")
            Finally
                ComObjectHelper.ReleaseComObject(color)
                ComObjectHelper.ReleaseComObject(fill)
                ComObjectHelper.ReleaseComObject(background)
            End Try
        End Sub

        Private Shared Sub CropImageToFill(shape As PowerPoint.Shape,
                                           bounds As SceneRect,
                                           warnings As List(Of String))
            If shape Is Nothing OrElse bounds Is Nothing OrElse
               shape.Width <= 0 OrElse shape.Height <= 0 OrElse
               bounds.Width <= 0 OrElse bounds.Height <= 0 Then Return

            Dim sourceWidth = shape.Width
            Dim sourceHeight = shape.Height
            Dim sourceAspect = sourceWidth / sourceHeight
            Dim targetAspect = bounds.Width / bounds.Height
            Dim pictureFormat As Object = Nothing
            Try
                Dim shapeObject As Object = shape
                pictureFormat = shapeObject.PictureFormat
                If sourceAspect > targetAspect Then
                    Dim visibleWidth = sourceHeight * targetAspect
                    Dim crop = Math.Max(0, (sourceWidth - visibleWidth) / 2.0F)
                    pictureFormat.CropLeft = crop
                    pictureFormat.CropRight = crop
                ElseIf sourceAspect < targetAspect Then
                    Dim visibleHeight = sourceWidth / targetAspect
                    Dim crop = Math.Max(0, (sourceHeight - visibleHeight) / 2.0F)
                    pictureFormat.CropTop = crop
                    pictureFormat.CropBottom = crop
                End If
                shape.LockAspectRatio = MsoTriState.msoFalse
                shape.Left = bounds.X
                shape.Top = bounds.Y
                shape.Width = bounds.Width
                shape.Height = bounds.Height
            Catch ex As Exception
                warnings.Add($"Image crop was unavailable; aspect-fit fallback was used: {ex.Message}")
                FitImageWithinBounds(shape, bounds)
            Finally
                ComObjectHelper.ReleaseComObject(pictureFormat)
            End Try
        End Sub

        Private Shared Sub FitImageWithinBounds(shape As PowerPoint.Shape, bounds As SceneRect)
            Dim sourceWidth = Math.Max(1, shape.Width)
            Dim sourceHeight = Math.Max(1, shape.Height)
            Dim scale = Math.Min(bounds.Width / sourceWidth, bounds.Height / sourceHeight)
            shape.LockAspectRatio = MsoTriState.msoTrue
            shape.Width = sourceWidth * scale
            shape.Left = bounds.X + (bounds.Width - shape.Width) / 2.0F
            shape.Top = bounds.Y + (bounds.Height - shape.Height) / 2.0F
        End Sub

        Private Shared Sub ApplyShapeStyle(shape As PowerPoint.Shape, node As SceneNode, addShadow As Boolean)
            Dim fill As Object = Nothing
            Dim fillColor As Object = Nothing
            Dim line As Object = Nothing
            Dim lineColor As Object = Nothing
            Dim shadow As Object = Nothing
            Dim shadowColor As Object = Nothing
            Try
                ' Keep Fill late-bound for WPS implementations that do not expose
                ' Microsoft.Office.Core.FillFormat through QueryInterface.
                Dim shapeObject As Object = shape
                fill = shapeObject.Fill
                fill.Visible = MsoTriState.msoTrue
                fill.Solid()
                fillColor = fill.ForeColor
                fillColor.RGB = ToOle(node.FillColor, "#FFFFFF")
                fill.Transparency = Math.Max(0, Math.Min(1, node.FillTransparency))

                line = shapeObject.Line
                If String.IsNullOrWhiteSpace(node.LineColor) Then
                    line.Visible = MsoTriState.msoFalse
                Else
                    line.Visible = MsoTriState.msoTrue
                    lineColor = line.ForeColor
                    lineColor.RGB = ToOle(node.LineColor, node.FillColor)
                    line.Weight = Math.Max(0.5F, node.LineWeight)
                    line.Transparency = 0.25F
                End If

                If addShadow Then
                    Try
                        shadow = shapeObject.Shadow
                        shadow.Visible = MsoTriState.msoTrue
                        shadowColor = shadow.ForeColor
                        shadowColor.RGB = ToOle("#000000", "#000000")
                        shadow.Transparency = 0.82F
                        shadow.Blur = 5
                        shadow.OffsetX = 0
                        shadow.OffsetY = 2
                    Catch
                    End Try
                End If
            Finally
                ComObjectHelper.ReleaseComObject(shadowColor)
                ComObjectHelper.ReleaseComObject(shadow)
                ComObjectHelper.ReleaseComObject(lineColor)
                ComObjectHelper.ReleaseComObject(line)
                ComObjectHelper.ReleaseComObject(fillColor)
                ComObjectHelper.ReleaseComObject(fill)
            End Try
        End Sub

        Private Shared Sub ApplyLineStyle(shape As PowerPoint.Shape,
                                          node As SceneNode,
                                          tokens As DesignTokens)
            Dim line As Object = Nothing
            Dim color As Object = Nothing
            Try
                Dim shapeObject As Object = shape
                line = shapeObject.Line
                color = line.ForeColor
                color.RGB = ToOle(node.LineColor, tokens.Divider)
                line.Weight = Math.Max(0.5F, node.LineWeight)
            Finally
                ComObjectHelper.ReleaseComObject(color)
                ComObjectHelper.ReleaseComObject(line)
            End Try
        End Sub

        Private Shared Sub ApplyCornerRadius(shape As PowerPoint.Shape, node As SceneNode)
            If shape Is Nothing OrElse node Is Nothing OrElse node.Bounds Is Nothing Then Return
            Dim adjustments As Object = Nothing
            Try
                Dim shapeObject As Object = shape
                adjustments = shapeObject.Adjustments
                Dim shortestSide = Math.Max(1.0F, Math.Min(node.Bounds.Width, node.Bounds.Height))
                adjustments.Item(1) = Math.Max(0.01F, Math.Min(0.2F, node.CornerRadius / shortestSide))
            Catch
            Finally
                ComObjectHelper.ReleaseComObject(adjustments)
            End Try
        End Sub

        Private Shared Sub ApplyTextStyle(shape As PowerPoint.Shape, node As SceneNode, tokens As DesignTokens)
            Dim shapeFill As Object = Nothing
            Dim shapeLine As Object = Nothing
            Dim frame As Object = Nothing
            Dim range As Object = Nothing
            Dim font As Object = Nothing
            Dim legacyFrame As Object = Nothing
            Dim legacyRange As Object = Nothing
            Dim legacyFont As Object = Nothing
            Dim fontColor As Object = Nothing
            Dim paragraph As Object = Nothing
            Dim frameAvailable As Boolean = False
            Try
                Dim shapeObject As Object = shape
                shapeFill = shapeObject.Fill
                shapeFill.Visible = MsoTriState.msoFalse
                shapeLine = shapeObject.Line
                shapeLine.Visible = MsoTriState.msoFalse

                ' Часть хостов (WPS, отдельные сборки PowerPoint) не отдаёт Office.TextFrame2.
                ' Тогда работаем через старый TextFrame и пропускаем TextFrame2-настройки.
                Try
                    frame = shapeObject.TextFrame2
                    frameAvailable = frame IsNot Nothing
                Catch
                    frame = Nothing
                    frameAvailable = False
                End Try

                If frameAvailable Then
                    frame.MarginLeft = 1
                    frame.MarginRight = 1
                    frame.MarginTop = 1
                    frame.MarginBottom = 1
                    frame.WordWrap = MsoTriState.msoTrue
                    frame.AutoSize = MsoAutoSize.msoAutoSizeNone
                    frame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop
                    range = frame.TextRange
                Else
                    legacyFrame = shapeObject.TextFrame
                    legacyRange = legacyFrame.TextRange
                    range = legacyRange
                End If

                range.Text = If(node.Text, "")
                font = range.Font
                font.Name = tokens.FontFamily
                Try
                    font.NameFarEast = tokens.FontFamily
                Catch
                End Try
                font.Size = If(node.FontSize > 0, node.FontSize, tokens.BodySize)
                font.Bold = If(node.Bold, MsoTriState.msoTrue, MsoTriState.msoFalse)
                ' WPS/部分 PowerPoint 兼容层返回的 Font2.Fill 不支持 Office.FillFormat IID，
                ' 使用旧文本对象模型设置颜色，避免 E_NOINTERFACE。
                If legacyFrame Is Nothing Then
                    legacyFrame = shapeObject.TextFrame
                    legacyRange = legacyFrame.TextRange
                End If
                legacyFont = legacyRange.Font
                fontColor = legacyFont.Color
                fontColor.RGB = ToOle(node.TextColor, tokens.TextPrimary)

                If frameAvailable Then
                    paragraph = range.ParagraphFormat
                    Select Case If(node.Alignment, "left").ToLowerInvariant()
                        Case "center"
                            paragraph.Alignment = MsoParagraphAlignment.msoAlignCenter
                        Case "right"
                            paragraph.Alignment = MsoParagraphAlignment.msoAlignRight
                        Case Else
                            paragraph.Alignment = MsoParagraphAlignment.msoAlignLeft
                    End Select
                    paragraph.SpaceWithin = 1.05F
                End If
            Finally
                ComObjectHelper.ReleaseComObject(paragraph)
                ComObjectHelper.ReleaseComObject(fontColor)
                ComObjectHelper.ReleaseComObject(legacyFont)
                ComObjectHelper.ReleaseComObject(legacyRange)
                ComObjectHelper.ReleaseComObject(legacyFrame)
                ComObjectHelper.ReleaseComObject(font)
                ' В fallback-ветке range и legacyRange — один и тот же объект, повторно не освобождаем.
                If frameAvailable Then ComObjectHelper.ReleaseComObject(range)
                ComObjectHelper.ReleaseComObject(frame)
                ComObjectHelper.ReleaseComObject(shapeLine)
                ComObjectHelper.ReleaseComObject(shapeFill)
            End Try
        End Sub

        Private Shared Function WriteNotes(slide As PowerPoint.Slide, notes As String) As Boolean
            If String.IsNullOrWhiteSpace(notes) Then Return True
            Dim notesPage As Object = Nothing
            Dim noteShapes As Object = Nothing
            Dim placeholders As Object = Nothing
            Dim placeholder As Object = Nothing
            Dim frame As Object = Nothing
            Dim range As Object = Nothing
            Try
                notesPage = slide.NotesPage
                noteShapes = notesPage.Shapes
                placeholders = noteShapes.Placeholders
                placeholder = FindNotesBodyPlaceholder(placeholders)
                If placeholder Is Nothing Then Return False
                frame = placeholder.TextFrame
                range = frame.TextRange
                range.Text = notes
                Dim renderedNotes = Convert.ToString(range.Text)
                Return String.Equals(NormalizeNoteText(renderedNotes), NormalizeNoteText(notes), StringComparison.Ordinal)
            Catch
                Return False
            Finally
                ComObjectHelper.ReleaseComObject(range)
                ComObjectHelper.ReleaseComObject(frame)
                ComObjectHelper.ReleaseComObject(placeholder)
                ComObjectHelper.ReleaseComObject(placeholders)
                ComObjectHelper.ReleaseComObject(noteShapes)
                ComObjectHelper.ReleaseComObject(notesPage)
            End Try
        End Function

        Private Shared Function FindNotesBodyPlaceholder(placeholders As Object) As Object
            If placeholders Is Nothing Then Return Nothing
            Dim count As Integer
            Try
                count = CInt(placeholders.Count)
            Catch
                Return Nothing
            End Try

            For index = 1 To count
                Dim candidate As Object = Nothing
                Dim placeholderFormat As Object = Nothing
                Try
                    candidate = placeholders(index)
                    placeholderFormat = candidate.PlaceholderFormat
                    Dim placeholderType = CInt(placeholderFormat.Type)
                    If placeholderType = CInt(PowerPoint.PpPlaceholderType.ppPlaceholderBody) Then
                        Dim selected = candidate
                        candidate = Nothing
                        Return selected
                    End If
                Catch
                Finally
                    ComObjectHelper.ReleaseComObject(placeholderFormat)
                    ComObjectHelper.ReleaseComObject(candidate)
                End Try
            Next
            Return Nothing
        End Function

        Private Shared Function NormalizeNoteText(value As String) As String
            Return If(value, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Trim()
        End Function

        Private Shared Function ToOle(value As String, fallback As String) As Integer
            Try
                Return ColorTranslator.ToOle(ColorTranslator.FromHtml(If(String.IsNullOrWhiteSpace(value), fallback, value)))
            Catch
                Return ColorTranslator.ToOle(ColorTranslator.FromHtml(fallback))
            End Try
        End Function

        Private Shared Function BuildShapeName(nodeId As String) As String
            Dim safe = New String(If(nodeId, "node").Where(Function(ch) Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "-"c).ToArray())
            If safe.Length > 48 Then safe = safe.Substring(0, 48)
            Return "AI_" & safe
        End Function
    End Class

End Namespace
