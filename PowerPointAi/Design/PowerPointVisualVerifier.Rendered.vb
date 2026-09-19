Imports System.Drawing
Imports System.IO
Imports Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon

Namespace Design

    Public NotInheritable Partial Class PowerPointVisualVerifier

        Private Shared Sub VerifyRenderedPixels(slide As PowerPoint.Slide,
                                                report As VisualVerificationReport)
            If slide Is Nothing OrElse report Is Nothing Then Return
            Dim exportPath = Path.Combine(Path.GetTempPath(),
                                          "office-ai-slide-" & Guid.NewGuid().ToString("N") & ".png")
            Try
                slide.Export(exportPath, "PNG", 960, 540)
                If Not File.Exists(exportPath) OrElse New FileInfo(exportPath).Length = 0 Then
                    report.Issues.Add(New VisualIssue With {
                        .Code = "RENDER_EXPORT_EMPTY",
                        .Severity = "error",
                        .Message = "PowerPoint exported an empty slide preview"
                    })
                    report.AestheticScore = 0
                    Return
                End If

                Using bitmap As New Bitmap(exportPath)
                    Dim quantizedColors As New HashSet(Of Integer)()
                    Dim sampleCount As Integer = 0
                    Dim luminanceSum As Double = 0
                    Dim luminanceSquaredSum As Double = 0
                    Dim stepX = Math.Max(1, bitmap.Width \ 64)
                    Dim stepY = Math.Max(1, bitmap.Height \ 36)
                    For y = Math.Min(bitmap.Height - 1, stepY \ 2) To bitmap.Height - 1 Step stepY
                        For x = Math.Min(bitmap.Width - 1, stepX \ 2) To bitmap.Width - 1 Step stepX
                            Dim pixel = bitmap.GetPixel(x, y)
                            Dim quantized = ((pixel.R \ 16) << 8) Or ((pixel.G \ 16) << 4) Or (pixel.B \ 16)
                            quantizedColors.Add(quantized)
                            Dim luminance = 0.2126R * pixel.R + 0.7152R * pixel.G + 0.0722R * pixel.B
                            luminanceSum += luminance
                            luminanceSquaredSum += luminance * luminance
                            sampleCount += 1
                        Next
                    Next

                    Dim mean = If(sampleCount = 0, 0, luminanceSum / sampleCount)
                    Dim variance = If(sampleCount = 0, 0,
                        Math.Max(0, luminanceSquaredSum / sampleCount - mean * mean))
                    Dim deviation = Math.Sqrt(variance)
                    report.Metrics("pixelColorBins") = quantizedColors.Count
                    report.Metrics("pixelLuminanceDeviation") = Math.Round(deviation, 3)
                    report.Metrics("pixelVerificationAvailable") = 1

                    If quantizedColors.Count <= 2 AndAlso deviation < 2.0R Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "RENDER_PIXELS_FLAT",
                            .Severity = "error",
                            .Message = "The exported slide is blank or visually indistinguishable from a solid fill"
                        })
                        report.AestheticScore = Math.Min(report.AestheticScore, 20)
                    ElseIf quantizedColors.Count <= 4 AndAlso deviation < 6.0R Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "RENDER_PIXELS_LOW_VARIATION",
                            .Severity = "warning",
                            .Message = "The exported slide has unusually low visual variation; inspect the rendered composition"
                        })
                        report.AestheticScore = Math.Min(report.AestheticScore, 76)
                    End If
                End Using
            Catch ex As Exception
                report.Metrics("pixelVerificationAvailable") = 0
                report.Issues.Add(New VisualIssue With {
                    .Code = "RENDER_PIXEL_VERIFY_UNAVAILABLE",
                    .Severity = "warning",
                    .Message = $"PowerPoint slide export verification was unavailable: {ex.Message}"
                })
            Finally
                Try
                    If File.Exists(exportPath) Then File.Delete(exportPath)
                Catch
                End Try
            End Try
        End Sub

        Private Shared Function BoundsDrifted(shape As PowerPoint.Shape, expected As SceneRect) As Boolean
            Const tolerance As Single = 1.5F
            Return Math.Abs(shape.Left - expected.X) > tolerance OrElse
                   Math.Abs(shape.Top - expected.Y) > tolerance OrElse
                   Math.Abs(shape.Width - expected.Width) > tolerance OrElse
                   Math.Abs(shape.Height - expected.Height) > tolerance
        End Function

        Private Shared Sub VerifyRenderedCollisions(plan As SlideRenderPlan,
                                                    actualBounds As Dictionary(Of String, SceneRect),
                                                    report As VisualVerificationReport)
            Dim collisionNodes = plan.Nodes.Where(Function(node)
                Return node IsNot Nothing AndAlso node.Collision AndAlso
                       Not String.IsNullOrWhiteSpace(node.Id) AndAlso actualBounds.ContainsKey(node.Id)
            End Function).ToList()
            For leftIndex = 0 To collisionNodes.Count - 2
                For rightIndex = leftIndex + 1 To collisionNodes.Count - 1
                    Dim leftNode = collisionNodes(leftIndex)
                    Dim rightNode = collisionNodes(rightIndex)
                    If OverlapRatio(actualBounds(leftNode.Id), actualBounds(rightNode.Id)) > 0.04F Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "RENDER_UNEXPECTED_OVERLAP",
                            .NodeId = rightNode.Id,
                            .Severity = "error",
                            .Message = $"Rendered elements overlap: {leftNode.Id} / {rightNode.Id}"
                        })
                    End If
                Next
            Next
        End Sub

        Private Shared Sub VerifySceneTextCollisions(plan As SlideRenderPlan,
                                                     report As VisualVerificationReport)
            Dim textNodes = plan.Nodes.Where(Function(node)
                Return node IsNot Nothing AndAlso node.Kind = "text" AndAlso node.Bounds IsNot Nothing AndAlso
                       Not IsDecorativeTextNode(node.Id)
            End Function).ToList()
            For leftIndex = 0 To textNodes.Count - 2
                For rightIndex = leftIndex + 1 To textNodes.Count - 1
                    If OverlapRatio(textNodes(leftIndex).Bounds, textNodes(rightIndex).Bounds) > 0.08F Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "TEXT_BOX_OVERLAP",
                            .NodeId = textNodes(rightIndex).Id,
                            .Severity = "error",
                            .Message = $"Text boxes overlap: {textNodes(leftIndex).Id} / {textNodes(rightIndex).Id}"
                        })
                    End If
                Next
            Next
        End Sub

        Private Shared Sub VerifyRenderedTextCollisions(plan As SlideRenderPlan,
                                                        actualBounds As Dictionary(Of String, SceneRect),
                                                        report As VisualVerificationReport)
            Dim textNodes = plan.Nodes.Where(Function(node)
                Return node IsNot Nothing AndAlso node.Kind = "text" AndAlso
                       Not String.IsNullOrWhiteSpace(node.Id) AndAlso Not IsDecorativeTextNode(node.Id) AndAlso
                       actualBounds.ContainsKey(node.Id)
            End Function).ToList()
            For leftIndex = 0 To textNodes.Count - 2
                For rightIndex = leftIndex + 1 To textNodes.Count - 1
                    Dim leftNode = textNodes(leftIndex)
                    Dim rightNode = textNodes(rightIndex)
                    If OverlapRatio(actualBounds(leftNode.Id), actualBounds(rightNode.Id)) > 0.08F Then
                        report.Issues.Add(New VisualIssue With {
                            .Code = "RENDER_TEXT_BOX_OVERLAP",
                            .NodeId = rightNode.Id,
                            .Severity = "error",
                            .Message = $"Rendered text boxes overlap: {leftNode.Id} / {rightNode.Id}"
                        })
                    End If
                Next
            Next
        End Sub

        Private Shared Function IsDecorativeTextNode(nodeId As String) As Boolean
            Return String.Equals(If(nodeId, ""), "quote-mark", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Sub ScoreRenderedComposition(plan As SlideRenderPlan,
                                                    actualBounds As Dictionary(Of String, SceneRect),
                                                    textCharacters As Integer,
                                                    textBlocks As Integer,
                                                    fontLevels As Integer,
                                                    slideWidth As Single,
                                                    slideHeight As Single,
                                                    report As VisualVerificationReport)
            Dim visualNodes = plan.Nodes.Where(Function(node)
                Return node IsNot Nothing AndAlso node.Kind <> "text" AndAlso node.Kind <> "line" AndAlso
                       Not String.IsNullOrWhiteSpace(node.Id) AndAlso actualBounds.ContainsKey(node.Id)
            End Function).ToList()
            Dim slideArea = Math.Max(1.0R, CDbl(slideWidth) * slideHeight)
            Dim visualArea = visualNodes.Sum(Function(node)
                Dim bounds = actualBounds(node.Id)
                Return CDbl(bounds.Width) * bounds.Height
            End Function)
            Dim density = visualArea / slideArea
            Dim score As Integer = 100
            If density > 0.82R Then score -= 22
            If textCharacters > 900 Then
                score -= 18
            ElseIf textCharacters > 650 Then
                score -= 9
            End If
            If textBlocks > 18 Then
                score -= 10
            ElseIf textBlocks > 13 Then
                score -= 5
            End If
            If fontLevels < 2 AndAlso textBlocks >= 4 Then score -= 10
            If RequiresVisualStructure(plan.SlideType, textBlocks) AndAlso visualNodes.Count = 0 Then
                score -= 12
                report.Issues.Add(New VisualIssue With {
                    .Code = "RENDERED_VISUAL_STRUCTURE_MISSING",
                    .Severity = "error",
                    .Message = $"The rendered {plan.SlideType} slide contains no visible visual structure beyond text"
                })
            End If

            Dim weightedCenterX As Double = slideWidth / 2.0R
            If visualArea > 0 Then
                weightedCenterX = visualNodes.Sum(Function(node)
                    Dim bounds = actualBounds(node.Id)
                    Dim area = CDbl(bounds.Width) * bounds.Height
                    Return (bounds.X + bounds.Width / 2.0R) * area
                End Function) / visualArea
            End If
            Dim horizontalImbalance = Math.Abs(weightedCenterX - slideWidth / 2.0R) / Math.Max(1, slideWidth / 2.0R)
            If horizontalImbalance > 0.55R Then
                score -= 12
            ElseIf horizontalImbalance > 0.38R Then
                score -= 6
            End If

            Dim unrepairedErrors = report.Issues.FindAll(Function(issue) issue.Severity = "error" AndAlso Not issue.Repaired).Count
            Dim warnings = report.Issues.FindAll(Function(issue) issue.Severity = "warning").Count
            score -= unrepairedErrors * 15
            score -= Math.Min(10, warnings * 2)
            score = Math.Max(0, Math.Min(100, score))
            report.AestheticScore = score
            report.Metrics("renderedVisualDensity") = Math.Round(density, 3)
            report.Metrics("renderedTextCharacters") = textCharacters
            report.Metrics("renderedTextBlocks") = textBlocks
            report.Metrics("renderedFontLevels") = fontLevels
            report.Metrics("renderedHorizontalImbalance") = Math.Round(horizontalImbalance, 3)
            report.Metrics("renderedVisualNodeCount") = visualNodes.Count

            If score < 78 Then
                report.Issues.Add(New VisualIssue With {
                    .Code = "RENDERED_AESTHETIC_SCORE_LOW",
                    .Severity = "warning",
                    .Message = $"Rendered slide aesthetic score {score} is below the professional delivery threshold 78"
                })
            ElseIf score < 85 Then
                report.Issues.Add(New VisualIssue With {
                    .Code = "RENDERED_AESTHETIC_SCORE_WARNING",
                    .Severity = "warning",
                    .Message = $"Rendered slide aesthetic score {score} should be improved"
                })
            End If
        End Sub

        Private Shared Function RequiresVisualStructure(slideType As String,
                                                        textBlockCount As Integer) As Boolean
            Select Case If(slideType, "").Trim().ToLowerInvariant()
                Case "comparison", "kpi", "process", "architecture", "matrix"
                    Return True
                Case "content", "two-column"
                    Return textBlockCount >= 4
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function RequiresTitleHierarchy(slideType As String) As Boolean
            Select Case If(slideType, "").Trim().ToLowerInvariant()
                Case "content", "two-column", "comparison", "kpi", "process", "architecture", "matrix"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function TitleBodyFontRatio(textNodes As List(Of SceneNode)) As Double
            If textNodes Is Nothing OrElse textNodes.Count < 2 Then Return 0
            Dim title = textNodes.FirstOrDefault(Function(node)
                Dim id = If(node.Id, "").Trim().ToLowerInvariant()
                Return id = "title" OrElse id = "section-title" OrElse id = "closing-title"
            End Function)
            If title Is Nothing OrElse title.FontSize <= 0 Then Return 0
            Dim bodySizes = textNodes.Where(Function(node)
                Dim id = If(node.Id, "").Trim().ToLowerInvariant()
                Return node IsNot title AndAlso node.FontSize > 0 AndAlso
                       Not id.Contains("metric-value") AndAlso Not id.EndsWith("-value") AndAlso
                       id <> "quote-mark"
            End Function).Select(Function(node) CDbl(node.FontSize)).ToList()
            If bodySizes.Count = 0 Then Return 0
            Return title.FontSize / Math.Max(1.0R, bodySizes.Max())
        End Function

        Private Shared Sub RepairTextOverflow(shape As PowerPoint.Shape, nodeId As String, report As VisualVerificationReport)
            Dim frame As TextFrame2 = Nothing
            Dim range As TextRange2 = Nothing
            Dim font As Font2 = Nothing
            Try
                frame = shape.TextFrame2
                range = frame.TextRange
                font = range.Font
                Dim availableHeight = Math.Max(4, shape.Height - frame.MarginTop - frame.MarginBottom)
                Dim availableWidth = Math.Max(4, shape.Width - frame.MarginLeft - frame.MarginRight)
                Dim repaired As Boolean = False
                Dim minimumFont = MinimumFontForNode(nodeId, "")
                While TextOverflows(range, availableWidth, availableHeight) AndAlso font.Size > minimumFont
                    font.Size = Math.Max(minimumFont, font.Size - 0.75F)
                    repaired = True
                End While
                If repaired Then
                    report.RepairCount += 1
                    report.Issues.Add(New VisualIssue With {.Code = "TEXT_OVERFLOW_REPAIRED", .NodeId = nodeId, .Severity = "error", .Message = "Text overflow was repaired by fitting typography", .Repaired = True})
                End If
                If TextOverflows(range, availableWidth, availableHeight) Then
                    report.Issues.Add(New VisualIssue With {.Code = "TEXT_OVERFLOW", .NodeId = nodeId, .Severity = "error", .Message = "Text still overflows its rendered bounds after repair"})
                End If
            Catch ex As Exception
                report.Issues.Add(New VisualIssue With {.Code = "TEXT_VERIFY_FAILED", .NodeId = nodeId, .Severity = "warning", .Message = ex.Message})
            Finally
                ComObjectHelper.ReleaseComObject(font)
                ComObjectHelper.ReleaseComObject(range)
                ComObjectHelper.ReleaseComObject(frame)
            End Try
        End Sub

        Private Shared Function TextOverflows(range As TextRange2,
                                              availableWidth As Single,
                                              availableHeight As Single) As Boolean
            If range Is Nothing Then Return False
            Const tolerance As Single = 1.5F
            Return range.BoundHeight > availableHeight + tolerance OrElse
                   range.BoundWidth > availableWidth + tolerance
        End Function

        Private Shared Sub VerifyRenderedTypography(range As TextRange2,
                                                    nodeId As String,
                                                    tokens As DesignTokens,
                                                    report As VisualVerificationReport)
            If range Is Nothing Then Return
            Dim font As Font2 = Nothing
            Try
                font = range.Font
                Dim minimumFont = MinimumFontForNode(nodeId, "")
                If font.Size > 0 AndAlso font.Size < minimumFont - 0.1F Then
                    report.Issues.Add(New VisualIssue With {
                        .Code = "RENDERED_FONT_TOO_SMALL",
                        .NodeId = nodeId,
                        .Severity = "error",
                        .Message = $"Rendered font size {Math.Round(font.Size, 1)} is below the semantic minimum {minimumFont}"
                    })
                End If
                If tokens Is Nothing OrElse String.IsNullOrWhiteSpace(tokens.FontFamily) Then Return

                Dim latinName = If(font.Name, "")
                Dim farEastName As String = ""
                Try
                    farEastName = If(font.NameFarEast, "")
                Catch
                End Try
                If Not FontNameMatches(latinName, tokens.FontFamily) AndAlso
                   Not FontNameMatches(farEastName, tokens.FontFamily) Then
                    report.Issues.Add(New VisualIssue With {
                        .Code = "FONT_SUBSTITUTED",
                        .NodeId = nodeId,
                        .Severity = "warning",
                        .Message = $"PowerPoint substituted '{tokens.FontFamily}' with '{FirstAvailableFontName(latinName, farEastName)}'"
                    })
                End If
            Finally
                ComObjectHelper.ReleaseComObject(font)
            End Try
        End Sub

        Private Shared Function GetRenderedFontSize(range As TextRange2) As Single
            If range Is Nothing Then Return 0
            Dim font As Font2 = Nothing
            Try
                font = range.Font
                Return font.Size
            Finally
                ComObjectHelper.ReleaseComObject(font)
            End Try
        End Function

        Private Shared Function FontNameMatches(actual As String, expected As String) As Boolean
            If String.IsNullOrWhiteSpace(actual) OrElse String.IsNullOrWhiteSpace(expected) Then Return False
            Return String.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function FirstAvailableFontName(ParamArray values As String()) As String
            For Each value In values
                If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
            Next
            Return "unknown font"
        End Function

        Private Shared Function MinimumFontForNode(nodeId As String, styleRole As String) As Single
            Dim id = If(nodeId, "").Trim().ToLowerInvariant()
            If id = "title" OrElse id = "statement" OrElse id = "section-title" OrElse
               id = "closing-title" OrElse id = "quote" Then Return 22.0F
            If id.Contains("metric-value") Then Return 24.0F
            If id.EndsWith("-title") OrElse id.Contains("feature-title") Then Return 13.0F
            If id.StartsWith("footer-") OrElse id.Contains("-source") OrElse
               id.StartsWith("chart-scale-") OrElse id.StartsWith("chart-category-") OrElse
               String.Equals(styleRole, "caption", StringComparison.OrdinalIgnoreCase) Then Return 9.5F
            Return 10.5F
        End Function

        Private Shared Function FitsEstimatedText(node As SceneNode) As Boolean
            If node Is Nothing OrElse node.Bounds Is Nothing OrElse String.IsNullOrWhiteSpace(node.Text) Then Return True
            Dim size = Math.Max(9.5F, node.FontSize)
            Dim unitsPerLine = Math.Max(2.0R, Math.Floor(node.Bounds.Width / (size * 0.95F)))
            Dim availableLines = Math.Max(1, CInt(Math.Floor(node.Bounds.Height / (size * 1.35F))))
            Dim normalized = node.Text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim requiredLines As Integer = 0
            For Each line In normalized.Split({vbLf}, StringSplitOptions.None)
                requiredLines += Math.Max(1, CInt(Math.Ceiling(EstimateTextUnits(line) / unitsPerLine)))
            Next
            Return requiredLines <= availableLines
        End Function

        Private Shared Function EstimateTextUnits(value As String) As Double
            Dim units As Double = 0
            For Each character In If(value, "")
                If Char.IsWhiteSpace(character) Then
                    units += 0.32R
                ElseIf AscW(character) >= &H2E80 Then
                    units += 1.0R
                ElseIf Char.IsUpper(character) Then
                    units += 0.68R
                Else
                    units += 0.55R
                End If
            Next
            Return units
        End Function

        Private Shared Function OverlapRatio(left As SceneRect, right As SceneRect) As Single
            Dim overlapWidth = Math.Max(0, Math.Min(left.X + left.Width, right.X + right.Width) - Math.Max(left.X, right.X))
            Dim overlapHeight = Math.Max(0, Math.Min(left.Y + left.Height, right.Y + right.Height) - Math.Max(left.Y, right.Y))
            Dim overlap = overlapWidth * overlapHeight
            Dim minArea = Math.Min(left.Width * left.Height, right.Width * right.Height)
            If minArea <= 0 Then Return 0
            Return overlap / minArea
        End Function

        Private Shared Function HasInvalidBounds(bounds As SceneRect, nodeKind As String) As Boolean
            If bounds Is Nothing Then Return True
            Dim invalidNumber = Single.IsNaN(bounds.X) OrElse Single.IsInfinity(bounds.X) OrElse
                                Single.IsNaN(bounds.Y) OrElse Single.IsInfinity(bounds.Y) OrElse
                                Single.IsNaN(bounds.Width) OrElse Single.IsInfinity(bounds.Width) OrElse
                                Single.IsNaN(bounds.Height) OrElse Single.IsInfinity(bounds.Height)
            If invalidNumber Then Return True
            If String.Equals(nodeKind, "line", StringComparison.OrdinalIgnoreCase) Then
                Return Math.Abs(bounds.Width) < 0.01F AndAlso Math.Abs(bounds.Height) < 0.01F
            End If
            Return bounds.Width <= 0 OrElse bounds.Height <= 0
        End Function
    End Class

End Namespace
