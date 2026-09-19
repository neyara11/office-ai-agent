Imports System.IO
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace Design

    Public NotInheritable Partial Class ProfessionalDeckExecutor

        Private Shared Sub VerifyDeckComposition(plans As List(Of SlideRenderPlan),
                                                 reports As List(Of VisualVerificationReport))
            If plans Is Nothing OrElse reports Is Nothing OrElse plans.Count <> reports.Count Then Return
            Dim signatures = plans.Select(Function(plan) BuildCompositionSignature(plan)).ToList()
            For index = 2 To signatures.Count - 1
                If String.IsNullOrWhiteSpace(signatures(index)) Then Continue For
                If String.Equals(signatures(index), signatures(index - 1), StringComparison.Ordinal) AndAlso
                   String.Equals(signatures(index), signatures(index - 2), StringComparison.Ordinal) Then
                    reports(index).Issues.Add(New VisualIssue With {
                        .Code = "DECK_COMPOSITION_REPETITION",
                        .Severity = "warning",
                        .Message = "Три подряд идущих слайда используют одну композицию; стоит разнообразить иерархию, фокус или вариант Scene"
                    })
                    reports(index).AestheticScore = Math.Min(reports(index).AestheticScore, 72)
                End If
            Next
        End Sub

        Private Shared Function BuildCompositionSignature(plan As SlideRenderPlan) As String
            If plan Is Nothing Then Return ""
            Dim slideType = If(plan.SlideType, "").Trim().ToLowerInvariant()
            If slideType = "cover" OrElse slideType = "section" OrElse slideType = "closing" Then Return ""
            Dim nodes = plan.Nodes.Where(Function(node)
                Return node IsNot Nothing AndAlso node.Bounds IsNot Nothing AndAlso
                       node.Kind <> "line" AndAlso Not String.IsNullOrWhiteSpace(node.Id)
            End Function).Select(Function(node)
                Dim xBand = CInt(Math.Round(node.Bounds.X / 48.0F))
                Dim yBand = CInt(Math.Round(node.Bounds.Y / 36.0F))
                Dim widthBand = CInt(Math.Round(node.Bounds.Width / 48.0F))
                Dim heightBand = CInt(Math.Round(node.Bounds.Height / 36.0F))
                Return $"{node.Kind}:{xBand},{yBand},{widthBand},{heightBand}"
            End Function).OrderBy(Function(value) value)
            Return slideType & "|" & String.Join(";", nodes)
        End Function

        Private Shared Function ExecutePreview(spec As DeckDesignSpec,
                                               tokens As DesignTokens,
                                               slideWidth As Single,
                                               slideHeight As Single,
                                               initialCount As Integer) As ToolResult
            Const toolId As String = "CreateSlides"
            Dim slideResults As New JArray()
            Dim warnings As New List(Of String) From {
                "Preview validates the editable Scene plan only; no PowerPoint shapes were created"
            }
            Dim plans As New List(Of SlideRenderPlan)()
            Dim reports As New List(Of VisualVerificationReport)()

            For index = 0 To spec.Slides.Count - 1
                Dim slideSpec = spec.Slides(index)
                Try
                    Dim plan = SlideLayoutEngine.Compile(slideSpec, tokens, slideWidth, slideHeight,
                                                         index, spec.Slides.Count)
                    Dim report = PowerPointVisualVerifier.PreflightAndRepair(plan, slideWidth, slideHeight)
                    plans.Add(plan)
                    reports.Add(report)
                    slideResults.Add(BuildSlideResult(index, slideSpec,
                                                       If(report.Passed, "preview_ready", "failed"),
                                                       report, Nothing,
                                                       If(report.Passed, "", "LAYOUT_VERIFY_FAILED")))
                    If Not report.Passed Then
                        Return BuildPreviewFailure(spec, tokens, initialCount, slideResults, warnings,
                                                   "Предпроверка профессионального макета не пройдена" & DescribeReportIssues(report),
                                                   ExceptionClassifier.CodeVerifyFailed)
                    End If
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", index + 1}, {"title", slideSpec.Title}, {"slideType", slideSpec.SlideType},
                        {"status", "failed"}, {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "preview_compile_preflight"}
                    })
                    Return BuildPreviewFailure(spec, tokens, initialCount, slideResults, warnings,
                                               ex.Message, classified.ErrorCode)
                End Try
            Next

            VerifyDeckComposition(plans, reports)
            If reports.Any(Function(report) Not report.Passed) Then
                slideResults = New JArray()
                For index = 0 To spec.Slides.Count - 1
                    Dim report = reports(index)
                    slideResults.Add(BuildSlideResult(index, spec.Slides(index),
                                                       If(report.Passed, "preview_ready", "failed"),
                                                       report, Nothing,
                                                       If(report.Passed, "", "DECK_COMPOSITION_VERIFY_FAILED")))
                Next
                Return BuildPreviewFailure(spec, tokens, initialCount, slideResults, warnings,
                                           "Ритм композиционных изменений презентации не прошёл проверку",
                                           ExceptionClassifier.CodeVerifyFailed)
            End If

            Dim observation = BuildPreviewObservation(spec, initialCount, slideResults, warnings, True)
            Dim data As New JObject From {
                {"preview", True}, {"rendered", False}, {"designSystem", tokens.Name},
                {"previewedSlides", spec.Slides.Count}, {"createdSlides", 0},
                {"slideResults", slideResults.DeepClone()}
            }
            Return ToolResult.Succeed(toolId,
                                      $"Предпроверка Scene профессиональных слайдов выполнена: {spec.Slides.Count}; презентация не изменялась",
                                      data:=data, observation:=observation)
        End Function

        Private Shared Function BuildPreviewFailure(spec As DeckDesignSpec,
                                                    tokens As DesignTokens,
                                                    initialCount As Integer,
                                                    slideResults As JArray,
                                                    warnings As List(Of String),
                                                    message As String,
                                                    errorCode As String) As ToolResult
            Return ToolResult.Failed("CreateSlides", message,
                                     data:=New JObject From {
                                         {"preview", True}, {"rendered", False},
                                         {"designSystem", tokens.Name}, {"createdSlides", 0},
                                         {"slideResults", slideResults.DeepClone()}
                                     },
                                     errorCode:=errorCode,
                                     userMessage:="Предпроверка профессиональных слайдов не пройдена; Agent исправит Scene по результатам визуальной проверки",
                                     recoverable:=True,
                                     observation:=BuildPreviewObservation(spec, initialCount, slideResults,
                                                                          warnings, False))
        End Function

        Private Shared Function BuildPreviewObservation(spec As DeckDesignSpec,
                                                        initialCount As Integer,
                                                        slideResults As JArray,
                                                        warnings As List(Of String),
                                                        success As Boolean) As JObject
            Dim issueCount = slideResults.OfType(Of JObject)().Sum(Function(item) If(item("issueCount")?.Value(Of Integer)(), 0))
            Dim repairCount = slideResults.OfType(Of JObject)().Sum(Function(item) If(item("repairCount")?.Value(Of Integer)(), 0))
            Dim scores = slideResults.OfType(Of JObject)().
                Where(Function(item) item("aestheticScore") IsNot Nothing).
                Select(Function(item) item("aestheticScore").Value(Of Integer)()).ToList()
            Return New JObject From {
                {"kind", "preview"},
                {"summary", If(success,
                    $"Проверено Scene профессиональных слайдов: {spec.Slides.Count}; презентация не изменялась",
                    $"Сбой предпроверки Scene профессиональных слайдов; проверено {slideResults.Count}/{spec.Slides.Count}")},
                {"changed", False}, {"preview", True}, {"rendered", False},
                {"slideCountBefore", initialCount}, {"slideCountAfter", initialCount},
                {"targetRefs", New JArray()}, {"slideResults", slideResults.DeepClone()},
                {"visualVerification", New JObject From {
                    {"issueCount", issueCount}, {"repairCount", repairCount},
                    {"averageAestheticScore", If(scores.Count = 0, 0, CInt(Math.Round(scores.Average())))},
                    {"minimumAestheticScore", If(scores.Count = 0, 0, scores.Min())},
                    {"passed", success}, {"scope", "scene_preflight"}
                }},
                {"warnings", JArray.FromObject(warnings)}
            }
        End Function

        ''' <summary>
        ''' Кратко перечисляет неисправленные ошибки отчёта, чтобы причина сбоя была видна пользователю.
        ''' </summary>
        Private Shared Function DescribeReportIssues(report As VisualVerificationReport) As String
            If report Is Nothing OrElse report.Issues Is Nothing OrElse report.Issues.Count = 0 Then Return ""

            Dim errors = report.Issues.
                Where(Function(item) item IsNot Nothing AndAlso
                                     String.Equals(item.Severity, "error", StringComparison.OrdinalIgnoreCase) AndAlso
                                     Not item.Repaired).
                Take(3).
                Select(Function(item) $"{item.Code}: {item.Message}").
                ToList()

            If errors.Count = 0 Then Return ""
            Return " [" & String.Join("; ", errors) & "]"
        End Function

        Private Shared Function BuildFailure(presentation As PowerPoint.Presentation,
                                             spec As DeckDesignSpec,
                                             initialCount As Integer,
                                             createdCount As Integer,
                                             targetRefs As List(Of String),
                                             slideResults As JArray,
                                             warnings As List(Of String),
                                             message As String,
                                             errorCode As String) As ToolResult
            Dim detectedCreatedCount = Math.Max(createdCount,
                                                CountGeneratedSlidesAfter(presentation, initialCount))
            Dim remainingCreatedCount = detectedCreatedCount
            Dim rollbackError As String = ""
            Dim rolledBack = TryRollbackCreatedSlides(presentation, initialCount, remainingCreatedCount, rollbackError)

            If rolledBack Then
                If detectedCreatedCount > 0 Then
                    warnings.Add($"Deck generation failed and {detectedCreatedCount} newly created slide(s) were rolled back")
                    MarkSlideResultsRolledBack(slideResults)
                End If
                targetRefs.Clear()
                remainingCreatedCount = 0
            Else
                warnings.Add($"Deck generation rollback was incomplete: {rollbackError}")
                RebuildTargetRefs(presentation, targetRefs, initialCount)
                errorCode = ExceptionClassifier.CodePartialApply
            End If
            Dim failureObservation = BuildObservation(spec, initialCount, remainingCreatedCount,
                                                      targetRefs, slideResults, warnings, False)
            failureObservation("rollback") = New JObject From {
                {"succeeded", rolledBack},
                {"detectedSlides", detectedCreatedCount},
                {"rolledBackSlides", Math.Max(0, detectedCreatedCount - remainingCreatedCount)},
                {"remainingSlides", remainingCreatedCount},
                {"error", rollbackError}
            }
            If rolledBack AndAlso detectedCreatedCount > 0 Then
                failureObservation("summary") = $"Сбой генерации профессиональных слайдов; все {detectedCreatedCount} слайдов, добавленные в этом раунде, откачены"
            End If
            Dim interfaceUnavailable = String.Equals(errorCode, ExceptionClassifier.CodeCom, StringComparison.OrdinalIgnoreCase) AndAlso
                ExceptionClassifier.IsComInterfaceUnavailableMessage(message)

            Return ToolResult.Failed("CreateSlides", message,
                                     data:=New JObject From {
                                         {"createdSlides", remainingCreatedCount},
                                         {"rolledBackSlides", Math.Max(0, detectedCreatedCount - remainingCreatedCount)},
                                         {"rollbackSucceeded", rolledBack},
                                         {"rollbackError", rollbackError},
                                         {"targetRefs", JArray.FromObject(targetRefs)},
                                         {"slideResults", slideResults.DeepClone()}
                                     },
                                     errorCode:=errorCode,
                                     userMessage:=If(interfaceUnavailable,
                                                     "В текущем PowerPoint/WPS отсутствует требуемый COM-интерфейс: " & message & "; повторное выполнение остановлено",
                                                     "Генерация профессиональных слайдов выполнена не полностью: " & message),
                                     recoverable:=Not interfaceUnavailable,
                                     observation:=failureObservation)
        End Function

        Private Shared Function TryRollbackCreatedSlides(presentation As PowerPoint.Presentation,
                                                         initialCount As Integer,
                                                         ByRef remainingCreatedCount As Integer,
                                                         ByRef rollbackError As String) As Boolean
            rollbackError = ""
            If presentation Is Nothing Then
                rollbackError = "Active presentation is unavailable"
                Return False
            End If

            Try
                Dim slides As PowerPoint.Slides = Nothing
                Dim deletedCount As Integer = 0
                Try
                    slides = presentation.Slides
                    For index = slides.Count To initialCount + 1 Step -1
                        Dim slide As PowerPoint.Slide = Nothing
                        Try
                            slide = slides(index)
                            If IsGeneratedSlide(slide) Then
                                slide.Delete()
                                deletedCount += 1
                            End If
                        Finally
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next
                Finally
                    ComObjectHelper.ReleaseComObject(slides)
                End Try
                remainingCreatedCount = Math.Max(0, remainingCreatedCount - deletedCount)
                If remainingCreatedCount = 0 Then Return True
                rollbackError = $"Only {deletedCount} of the expected generated slide(s) could be identified safely"
                Return False
            Catch ex As Exception
                remainingCreatedCount = Math.Max(remainingCreatedCount,
                                                 CountGeneratedSlidesAfter(presentation, initialCount))
                rollbackError = ex.Message
                Return False
            End Try
        End Function

        Private Shared Function IsGeneratedSlide(slide As PowerPoint.Slide) As Boolean
            If slide Is Nothing Then Return False
            Dim shapes As PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For index = 1 To shapes.Count
                    Dim shape As PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(index)
                        Dim tag = If(shape.AlternativeText, "")
                        If tag.StartsWith(GeneratedShapeTagPrefix, StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
                Return False
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try
        End Function

        Private Shared Function CountGeneratedSlidesAfter(presentation As PowerPoint.Presentation,
                                                          initialCount As Integer) As Integer
            If presentation Is Nothing Then Return 0
            Dim slides As PowerPoint.Slides = Nothing
            Dim count As Integer = 0
            Try
                slides = presentation.Slides
                For index = Math.Max(1, initialCount + 1) To slides.Count
                    Dim slide As PowerPoint.Slide = Nothing
                    Try
                        slide = slides(index)
                        If IsGeneratedSlide(slide) Then count += 1
                    Finally
                        ComObjectHelper.ReleaseComObject(slide)
                    End Try
                Next
                Return count
            Finally
                ComObjectHelper.ReleaseComObject(slides)
            End Try
        End Function

        Private Shared Function GetRequiredSlideCount(presentation As PowerPoint.Presentation) As Integer
            If presentation Is Nothing Then Throw New ArgumentNullException(NameOf(presentation))
            Dim slides As PowerPoint.Slides = Nothing
            Try
                slides = presentation.Slides
                Return slides.Count
            Finally
                ComObjectHelper.ReleaseComObject(slides)
            End Try
        End Function

        Private Shared Sub MarkSlideResultsRolledBack(slideResults As JArray)
            For Each result In slideResults.OfType(Of JObject)()
                If String.Equals(result.Value(Of String)("status"), "succeeded", StringComparison.OrdinalIgnoreCase) Then
                    result("status") = "rolled_back"
                End If
                If result("targetRef") IsNot Nothing Then result("targetRef") = ""
            Next
        End Sub

        Private Shared Sub RebuildTargetRefs(presentation As PowerPoint.Presentation,
                                             targetRefs As List(Of String),
                                             initialCount As Integer)
            targetRefs.Clear()
            If presentation Is Nothing Then Return
            Dim slides As PowerPoint.Slides = Nothing
            Try
                slides = presentation.Slides
                For index = Math.Max(1, initialCount + 1) To slides.Count
                    Dim slide As PowerPoint.Slide = Nothing
                    Try
                        slide = slides(index)
                        If IsGeneratedSlide(slide) Then
                            targetRefs.Add($"PowerPoint:presentations/active/slides/{slide.SlideIndex}")
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(slide)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(slides)
            End Try
        End Sub

        Private Shared Function BuildFailureWithVisualEvidence(evidence As AgentVisualEvidence,
                                                               presentation As PowerPoint.Presentation,
                                                               spec As DeckDesignSpec,
                                                               initialCount As Integer,
                                                               createdCount As Integer,
                                                               targetRefs As List(Of String),
                                                               slideResults As JArray,
                                                               warnings As List(Of String),
                                                               message As String,
                                                               errorCode As String) As ToolResult
            Dim result = BuildFailure(presentation, spec, initialCount, createdCount, targetRefs,
                                      slideResults, warnings, message, errorCode)
            If result IsNot Nothing AndAlso evidence IsNot Nothing Then result.VisualEvidence.Add(evidence)
            Return result
        End Function

        Private Shared Function CaptureVisualEvidence(slide As PowerPoint.Slide,
                                                       slideIndex As Integer,
                                                       slideWidth As Single,
                                                       slideHeight As Single) As AgentVisualEvidence
            If slide Is Nothing Then Return Nothing
            Const maxEvidenceBytes As Integer = 2 * 1024 * 1024
            Dim exportWidth As Integer = 960
            Dim exportHeight As Integer = 540
            If slideWidth > 0 AndAlso slideHeight > 0 Then
                exportHeight = Math.Max(1, CInt(Math.Round(exportWidth * CDbl(slideHeight) / CDbl(slideWidth))))
                If exportHeight > 720 Then
                    exportHeight = 720
                    exportWidth = Math.Max(1, CInt(Math.Round(exportHeight * CDbl(slideWidth) / CDbl(slideHeight))))
                End If
            End If
            Dim exportPath = Path.Combine(Path.GetTempPath(),
                                          $"office-ai-ppt-evidence-{Guid.NewGuid():N}.png")
            Try
                slide.Export(exportPath, "PNG", exportWidth, exportHeight)
                If Not File.Exists(exportPath) Then Return Nothing
                Dim bytes = File.ReadAllBytes(exportPath)
                If bytes.Length = 0 OrElse bytes.Length > maxEvidenceBytes Then Return Nothing
                Return New AgentVisualEvidence With {
                    .MimeType = "image/png",
                    .DataUrl = "data:image/png;base64," & Convert.ToBase64String(bytes),
                    .Source = "powerpoint-rendered-slide",
                    .ItemIndex = slideIndex,
                    .Width = exportWidth,
                    .Height = exportHeight,
                    .ByteLength = bytes.Length
                }
            Catch ex As Exception
                AppLogger.Warn("ProfessionalDeckExecutor",
                               $"Visual evidence capture unavailable: {AppLogger.Redact(ex.Message)}")
                Return Nothing
            Finally
                Try
                    If File.Exists(exportPath) Then File.Delete(exportPath)
                Catch
                End Try
            End Try
        End Function

        Private Shared Function BuildObservation(spec As DeckDesignSpec,
                                                 initialCount As Integer,
                                                 createdCount As Integer,
                                                 targetRefs As List(Of String),
                                                 slideResults As JArray,
                                                 warnings As List(Of String),
                                                 success As Boolean) As JObject
            Dim issueCount = slideResults.OfType(Of JObject)().Sum(Function(item) If(item("issueCount")?.Value(Of Integer)(), 0))
            Dim repairCount = slideResults.OfType(Of JObject)().Sum(Function(item) If(item("repairCount")?.Value(Of Integer)(), 0))
            Dim scores = slideResults.OfType(Of JObject)().
                Where(Function(item) item("aestheticScore") IsNot Nothing).
                Select(Function(item) item("aestheticScore").Value(Of Integer)()).ToList()
            Dim averageScore = If(scores.Count = 0, 0, CInt(Math.Round(scores.Average())))
            Dim minimumScore = If(scores.Count = 0, 0, scores.Min())
            Return New JObject From {
                {"kind", "write"},
                {"summary", If(success, $"Профессиональный движок дизайна создал слайдов: {createdCount}, визуальная проверка выполнена", $"Профессиональный движок дизайна создал слайдов: {createdCount}/{spec.Slides.Count}")},
                {"changed", createdCount > 0}, {"designSystem", spec.DesignSystem},
                {"slideCountBefore", initialCount}, {"slideCountAfter", initialCount + createdCount},
                {"targetRefs", JArray.FromObject(targetRefs)}, {"slideResults", slideResults.DeepClone()},
                {"visualVerification", New JObject From {
                    {"issueCount", issueCount}, {"repairCount", repairCount},
                    {"averageAestheticScore", averageScore}, {"minimumAestheticScore", minimumScore},
                    {"passed", success}
                }},
                {"warnings", JArray.FromObject(warnings)}
            }
        End Function

        Private Shared Function BuildSlideResult(index As Integer,
                                                spec As SlideDesignSpec,
                                                status As String,
                                                report As VisualVerificationReport,
                                                targetRef As String,
                                                errorCode As String) As JObject
            Return New JObject From {
                {"index", index + 1}, {"title", spec.Title}, {"slideType", spec.SlideType}, {"status", status},
                {"targetRef", If(targetRef, "")}, {"issueCount", If(report?.Issues?.Count, 0)},
                {"repairCount", If(report?.RepairCount, 0)},
                {"aestheticScore", If(report?.AestheticScore, 0)},
                {"metrics", JObject.FromObject(If(report?.Metrics, New Dictionary(Of String, Double)()))},
                {"issues", JArray.FromObject(If(report?.Issues, New List(Of VisualIssue)()))}, {"errorCode", If(errorCode, "")}
            }
        End Function

        Private Shared Sub MergeReport(target As VisualVerificationReport, source As VisualVerificationReport)
            If target Is Nothing OrElse source Is Nothing Then Return
            target.Issues.AddRange(source.Issues)
            target.RepairCount += source.RepairCount
            target.AestheticScore = Math.Min(target.AestheticScore, source.AestheticScore)
            For Each metric In source.Metrics
                target.Metrics(metric.Key) = metric.Value
            Next
        End Sub
    End Class

End Namespace
