Imports System.Drawing
Imports System.IO
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon

Namespace Design

    Public NotInheritable Partial Class PowerPointVisualVerifier
        Private Const ShapeTagPrefix As String = "office-ai-design:"

        Private Sub New()
        End Sub

        Public Shared Function PreflightAndRepair(plan As SlideRenderPlan,
                                                  slideWidth As Single,
                                                  slideHeight As Single) As VisualVerificationReport
            Dim report As New VisualVerificationReport()
            If plan Is Nothing Then
                report.Issues.Add(New VisualIssue With {.Code = "PLAN_MISSING", .Severity = "error", .Message = "Slide render plan is missing"})
                Return report
            End If

            Dim invalidIds = plan.Nodes.Where(Function(node) node Is Nothing OrElse String.IsNullOrWhiteSpace(node.Id)).Count()
            If invalidIds > 0 Then
                report.Issues.Add(New VisualIssue With {
                    .Code = "SCENE_NODE_ID_MISSING",
                    .Severity = "error",
                    .Message = $"{invalidIds} scene node(s) are missing stable IDs"
                })
            End If
            For Each duplicate In plan.Nodes.Where(Function(node) node IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(node.Id)).
                    GroupBy(Function(node) node.Id, StringComparer.OrdinalIgnoreCase).
                    Where(Function(group) group.Count() > 1)
                report.Issues.Add(New VisualIssue With {
                    .Code = "SCENE_NODE_ID_DUPLICATE",
                    .NodeId = duplicate.Key,
                    .Severity = "error",
                    .Message = $"Scene node ID is duplicated: {duplicate.Key}"
                })
            Next

            For Each node In plan.Nodes
                If node.Bounds Is Nothing Then Continue For
                If HasInvalidBounds(node.Bounds, node.Kind) Then
                    report.Issues.Add(New VisualIssue With {
                        .Code = "INVALID_BOUNDS",
                        .NodeId = node.Id,
                        .Severity = "error",
                        .Message = "Element bounds are invalid for the scene node type"
                    })
                    Continue For
                End If
                If node.Bounds.X < 0 OrElse node.Bounds.Y < 0 OrElse
                   node.Bounds.X + node.Bounds.Width > slideWidth OrElse
                   node.Bounds.Y + node.Bounds.Height > slideHeight Then
                    node.Bounds.X = Math.Max(0, Math.Min(slideWidth - node.Bounds.Width, node.Bounds.X))
                    node.Bounds.Y = Math.Max(0, Math.Min(slideHeight - node.Bounds.Height, node.Bounds.Y))
                    report.Issues.Add(New VisualIssue With {.Code = "OUT_OF_BOUNDS", .NodeId = node.Id, .Severity = "error", .Message = "Element exceeded slide bounds", .Repaired = True})
                    report.RepairCount += 1
                End If

                If node.Kind = "text" AndAlso Not String.IsNullOrWhiteSpace(node.Text) Then
                    Dim minFont = MinimumFontForNode(node.Id, node.StyleRole)
                    If node.FontSize < minFont Then
                        node.FontSize = minFont
                        report.Issues.Add(New VisualIssue With {.Code = "FONT_TOO_SMALL", .NodeId = node.Id, .Severity = "error", .Message = "Font was below minimum readable size", .Repaired = True})
                        report.RepairCount += 1
                    End If
                    While Not FitsEstimatedText(node) AndAlso node.FontSize > minFont
                        node.FontSize = Math.Max(minFont, node.FontSize - 0.75F)
                        report.RepairCount += 1
                    End While
                    If Not FitsEstimatedText(node) Then
                        ' Оценка приблизительная; реальное переполнение проверяется после отрисовки,
                        ' поэтому не валим всю колоду — сообщаем предупреждением и снижаем оценку.
                        report.Issues.Add(New VisualIssue With {
                            .Code = "TEXT_OVERFLOW_PREDICTED",
                            .NodeId = node.Id,
                            .Severity = "warning",
                            .Message = "Оценочно текст не помещается при минимальном допустимом кегле; требуется сокращение текста или изменение компоновки"
                        })
                        report.AestheticScore = Math.Min(report.AestheticScore, 80)
                    End If
                    RepairTextContrast(plan, node, report)
                End If
            Next

            Dim collisionNodes = plan.Nodes.Where(Function(item) item.Collision AndAlso item.Bounds IsNot Nothing).ToList()
            For leftIndex = 0 To collisionNodes.Count - 2
                For rightIndex = leftIndex + 1 To collisionNodes.Count - 1
                    If OverlapRatio(collisionNodes(leftIndex).Bounds, collisionNodes(rightIndex).Bounds) > 0.04F Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "UNEXPECTED_OVERLAP",
                            .NodeId = collisionNodes(rightIndex).Id,
                            .Severity = "error",
                            .Message = $"Layout elements overlap: {collisionNodes(leftIndex).Id} / {collisionNodes(rightIndex).Id}"
                        })
                    End If
                Next
            Next
            VerifySceneTextCollisions(plan, report)
            ScoreComposition(plan, slideWidth, slideHeight, report)
            Return report
        End Function

        Private Shared Sub RepairTextContrast(plan As SlideRenderPlan,
                                              textNode As SceneNode,
                                              report As VisualVerificationReport)
            If plan Is Nothing OrElse textNode Is Nothing OrElse textNode.Bounds Is Nothing Then Return
            Dim background = FindUnderlyingColor(plan, textNode)
            Dim current = If(String.IsNullOrWhiteSpace(textNode.TextColor), "#111827", textNode.TextColor)
            Dim minimum = If(textNode.FontSize >= 18 OrElse (textNode.Bold AndAlso textNode.FontSize >= 14), 3.0R, 4.5R)
            If ContrastRatio(current, background) >= minimum Then Return

            Dim lightContrast = ContrastRatio("#FFFFFF", background)
            Dim darkContrast = ContrastRatio("#111827", background)
            textNode.TextColor = If(lightContrast >= darkContrast, "#FFFFFF", "#111827")
            report.Issues.Add(New VisualIssue With {
                .Code = "TEXT_CONTRAST_REPAIRED",
                .NodeId = textNode.Id,
                .Severity = "error",
                .Message = $"Text contrast was below {minimum}:1 and was repaired",
                .Repaired = True
            })
            report.RepairCount += 1
        End Sub

        Private Shared Function FindUnderlyingColor(plan As SlideRenderPlan, textNode As SceneNode) As String
            Dim centerX = textNode.Bounds.X + textNode.Bounds.Width / 2.0F
            Dim centerY = textNode.Bounds.Y + textNode.Bounds.Height / 2.0F
            Dim container = plan.Nodes.Where(Function(node)
                Return node IsNot Nothing AndAlso node IsNot textNode AndAlso node.Bounds IsNot Nothing AndAlso
                       (node.Kind = "rect" OrElse node.Kind = "round-rect" OrElse node.Kind = "circle") AndAlso
                       centerX >= node.Bounds.X AndAlso centerX <= node.Bounds.X + node.Bounds.Width AndAlso
                       centerY >= node.Bounds.Y AndAlso centerY <= node.Bounds.Y + node.Bounds.Height AndAlso
                       node.FillTransparency < 0.45F
            End Function).OrderBy(Function(node) node.Bounds.Width * node.Bounds.Height).FirstOrDefault()
            If container Is Nothing OrElse String.IsNullOrWhiteSpace(container.FillColor) Then Return plan.Background
            Return container.FillColor
        End Function

        Private Shared Function ContrastRatio(foreground As String, background As String) As Double
            Try
                Dim foregroundLuminance = RelativeLuminance(ColorTranslator.FromHtml(foreground))
                Dim backgroundLuminance = RelativeLuminance(ColorTranslator.FromHtml(background))
                Dim lighter = Math.Max(foregroundLuminance, backgroundLuminance)
                Dim darker = Math.Min(foregroundLuminance, backgroundLuminance)
                Return (lighter + 0.05R) / (darker + 0.05R)
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function RelativeLuminance(color As Color) As Double
            Dim red = LinearColor(color.R / 255.0R)
            Dim green = LinearColor(color.G / 255.0R)
            Dim blue = LinearColor(color.B / 255.0R)
            Return 0.2126R * red + 0.7152R * green + 0.0722R * blue
        End Function

        Private Shared Function LinearColor(value As Double) As Double
            If value <= 0.03928R Then Return value / 12.92R
            Return Math.Pow((value + 0.055R) / 1.055R, 2.4R)
        End Function

        Private Shared Sub ScoreComposition(plan As SlideRenderPlan,
                                            slideWidth As Single,
                                            slideHeight As Single,
                                            report As VisualVerificationReport)
            If plan Is Nothing OrElse report Is Nothing Then Return
            Dim score As Integer = 100
            Dim slideArea = Math.Max(1, slideWidth * slideHeight)
            Dim visualNodes = plan.Nodes.Where(Function(node)
                Return node.Bounds IsNot Nothing AndAlso
                       node.Kind <> "text" AndAlso node.Kind <> "line"
            End Function).ToList()
            Dim visualArea = visualNodes.Sum(Function(node) CDbl(node.Bounds.Width * node.Bounds.Height))
            Dim density = visualArea / slideArea
            Dim textNodes = plan.Nodes.Where(Function(node) node.Kind = "text" AndAlso Not String.IsNullOrWhiteSpace(node.Text)).ToList()
            Dim characterCount = textNodes.Sum(Function(node) node.Text.Length)
            Dim fontLevels = textNodes.Select(Function(node) Math.Round(node.FontSize, 1)).Distinct().Count()

            If density > 0.82 Then score -= 22
            If characterCount > 900 Then
                score -= 18
            ElseIf characterCount > 650 Then
                score -= 9
            End If
            If textNodes.Count > 18 Then
                score -= 10
            ElseIf textNodes.Count > 13 Then
                score -= 5
            End If
            If fontLevels < 2 AndAlso textNodes.Count >= 4 Then score -= 10

            If RequiresVisualStructure(plan.SlideType, textNodes.Count) AndAlso visualNodes.Count = 0 Then
                score -= 12
                report.Issues.Add(New VisualIssue With {
                    .Code = "VISUAL_STRUCTURE_MISSING",
                    .Severity = "warning",
                    .Message = "The composition contains no rendered visual structure beyond text"
                })
            End If
            Dim titleBodyRatio = TitleBodyFontRatio(textNodes)
            If RequiresTitleHierarchy(plan.SlideType) AndAlso
               titleBodyRatio > 0 AndAlso titleBodyRatio < 1.2R Then
                score -= 8
                report.Issues.Add(New VisualIssue With {
                    .Code = "TYPOGRAPHIC_HIERARCHY_WEAK",
                    .Severity = "warning",
                    .Message = $"Title-to-body font ratio {Math.Round(titleBodyRatio, 2)} is below the professional hierarchy threshold"
                })
            End If

            Dim weightedCenterX As Double = slideWidth / 2.0R
            If visualArea > 0 Then
                weightedCenterX = visualNodes.Sum(Function(node)
                    Dim area = CDbl(node.Bounds.Width * node.Bounds.Height)
                    Return (node.Bounds.X + node.Bounds.Width / 2.0R) * area
                End Function) / visualArea
            End If
            Dim horizontalImbalance = Math.Abs(weightedCenterX - slideWidth / 2.0R) / Math.Max(1, slideWidth / 2.0R)
            If horizontalImbalance > 0.55 Then
                score -= 12
            ElseIf horizontalImbalance > 0.38 Then
                score -= 6
            End If

            Dim unrepairedErrors = report.Issues.Where(Function(issue) issue.Severity = "error" AndAlso Not issue.Repaired).Count()
            Dim warnings = report.Issues.Where(Function(issue) issue.Severity = "warning").Count()
            score -= unrepairedErrors * 15
            score -= Math.Min(10, warnings * 2)
            score = Math.Max(0, Math.Min(100, score))

            report.AestheticScore = score
            report.Metrics("visualDensity") = Math.Round(density, 3)
            report.Metrics("textCharacters") = characterCount
            report.Metrics("textBlocks") = textNodes.Count
            report.Metrics("fontLevels") = fontLevels
            report.Metrics("horizontalImbalance") = Math.Round(horizontalImbalance, 3)
            report.Metrics("visualNodeCount") = visualNodes.Count
            report.Metrics("titleBodyFontRatio") = Math.Round(titleBodyRatio, 3)
            If score < 78 Then
                ' Оценка эстетики эвристическая; не валим из-за неё всю колоду.
                ' Реальные дефекты (переполнение, наложения, контраст) проверяются отдельно.
                report.Issues.Add(New VisualIssue With {
                    .Code = "AESTHETIC_SCORE_LOW",
                    .Severity = "warning",
                    .Message = $"Эстетическая оценка слайда {score} ниже рекомендуемого порога 78"
                })
            ElseIf score < 85 Then
                report.Issues.Add(New VisualIssue With {
                    .Code = "AESTHETIC_SCORE_WARNING",
                    .Severity = "warning",
                    .Message = "Эстетическую оценку слайда стоит улучшить"
                })
            End If
        End Sub

        Public Shared Function VerifyAndRepairRenderedSlide(slide As PowerPoint.Slide,
                                                            slideWidth As Single,
                                                            slideHeight As Single,
                                                            Optional plan As SlideRenderPlan = Nothing,
                                                            Optional tokens As DesignTokens = Nothing) As VisualVerificationReport
            Dim report As New VisualVerificationReport()
            If slide Is Nothing Then
                report.Issues.Add(New VisualIssue With {.Code = "SLIDE_MISSING", .Severity = "error", .Message = "Rendered slide is missing"})
                Return report
            End If

            Dim expectedNodes As New Dictionary(Of String, SceneNode)(StringComparer.OrdinalIgnoreCase)
            If plan IsNot Nothing Then
                For Each node In plan.Nodes
                    If node IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(node.Id) Then expectedNodes(node.Id) = node
                Next
            End If
            Dim seenNodeIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim actualBounds As New Dictionary(Of String, SceneRect)(StringComparer.OrdinalIgnoreCase)
            Dim actualTextCharacters As Integer = 0
            Dim actualTextBlocks As Integer = 0
            Dim actualFontLevels As New HashSet(Of Double)()

            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For index = 1 To shapes.Count
                    Dim shape As PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(index)
                        Dim tag = If(shape.AlternativeText, "")
                        If Not tag.StartsWith(ShapeTagPrefix, StringComparison.OrdinalIgnoreCase) Then Continue For
                        Dim nodeId = tag.Substring(ShapeTagPrefix.Length)
                        seenNodeIds.Add(nodeId)

                        Dim repairedBounds As Boolean = False
                        If shape.Left < 0 Then shape.Left = 0 : repairedBounds = True
                        If shape.Top < 0 Then shape.Top = 0 : repairedBounds = True
                        If shape.Left + shape.Width > slideWidth Then shape.Left = Math.Max(0, slideWidth - shape.Width) : repairedBounds = True
                        If shape.Top + shape.Height > slideHeight Then shape.Top = Math.Max(0, slideHeight - shape.Height) : repairedBounds = True
                        If repairedBounds Then
                            report.RepairCount += 1
                            report.Issues.Add(New VisualIssue With {.Code = "RENDER_BOUNDS_REPAIRED", .NodeId = nodeId, .Severity = "error", .Message = "Rendered element exceeded slide bounds", .Repaired = True})
                        End If

                        Dim expectedNode As SceneNode = Nothing
                        If expectedNodes.TryGetValue(nodeId, expectedNode) AndAlso expectedNode.Bounds IsNot Nothing Then
                            If BoundsDrifted(shape, expectedNode.Bounds) Then
                                shape.Left = expectedNode.Bounds.X
                                shape.Top = expectedNode.Bounds.Y
                                shape.Width = expectedNode.Bounds.Width
                                shape.Height = expectedNode.Bounds.Height
                                report.RepairCount += 1
                                report.Issues.Add(New VisualIssue With {
                                    .Code = "RENDER_BOUNDS_DRIFT_REPAIRED",
                                    .NodeId = nodeId,
                                    .Severity = "error",
                                    .Message = "PowerPoint changed the planned element bounds; the intended geometry was restored",
                                    .Repaired = True
                                })
                            End If
                        End If
                        actualBounds(nodeId) = New SceneRect(shape.Left, shape.Top, shape.Width, shape.Height)

                        If shape.HasTextFrame = MsoTriState.msoTrue Then
                            Dim actualFrame As TextFrame2 = Nothing
                            Dim actualRange As TextRange2 = Nothing
                            Try
                                Try
                                    actualFrame = shape.TextFrame2
                                Catch textFrameEx As Exception
                                    ' Часть хостов (WPS, отдельные сборки PowerPoint) не отдаёт Office.TextFrame2.
                                    ' Это не повод откатывать всю колоду: пропускаем проверку текста с предупреждением.
                                    report.Issues.Add(New VisualIssue With {
                                        .Code = "RENDER_TEXT_VERIFY_UNAVAILABLE",
                                        .NodeId = nodeId,
                                        .Severity = "warning",
                                        .Message = "Не удалось прочитать TextFrame2 у фигуры; проверка текста пропущена: " & textFrameEx.Message
                                    })
                                    actualFrame = Nothing
                                End Try
                                If actualFrame IsNot Nothing AndAlso actualFrame.HasText = MsoTriState.msoTrue Then
                                    RepairTextOverflow(shape, nodeId, report)
                                    actualRange = actualFrame.TextRange
                                    actualTextCharacters += If(actualRange.Text, "").Length
                                    actualTextBlocks += 1
                                    Dim renderedFontSize = GetRenderedFontSize(actualRange)
                                    If renderedFontSize > 0 Then actualFontLevels.Add(Math.Round(renderedFontSize, 1))
                                    VerifyRenderedTypography(actualRange, nodeId, tokens, report)
                                End If
                            Finally
                                ComObjectHelper.ReleaseComObject(actualRange)
                                ComObjectHelper.ReleaseComObject(actualFrame)
                            End Try
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            If plan IsNot Nothing Then
                For Each node In plan.Nodes
                    If node Is Nothing OrElse String.IsNullOrWhiteSpace(node.Id) OrElse node.Bounds Is Nothing Then Continue For
                    If Not seenNodeIds.Contains(node.Id) Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "RENDER_NODE_MISSING",
                            .NodeId = node.Id,
                            .Severity = "error",
                            .Message = "A planned scene element was not created in PowerPoint"
                        })
                    End If
                Next
                report.Metrics("plannedNodes") = plan.Nodes.Count
                report.Metrics("renderedNodes") = seenNodeIds.Count
                VerifyRenderedCollisions(plan, actualBounds, report)
                VerifyRenderedTextCollisions(plan, actualBounds, report)
                ScoreRenderedComposition(plan, actualBounds, actualTextCharacters, actualTextBlocks,
                                         actualFontLevels.Count, slideWidth, slideHeight, report)
            End If
            VerifyRenderedPixels(slide, report)
            Return report
        End Function

    End Class

End Namespace
