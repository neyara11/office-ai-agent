Imports System.IO
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace Design

    Public NotInheritable Partial Class ProfessionalDeckExecutor
        Private Const GeneratedShapeTagPrefix As String = "office-ai-design:"

        Private Sub New()
        End Sub

        Public Shared Function ExecuteAsToolResult(params As JObject,
                                                   Optional preview As Boolean = False) As ToolResult
            Const toolId As String = "CreateSlides"
            Dim presentation As PowerPoint.Presentation = Nothing
            Try
                Dim previewRequested = preview
                If params IsNot Nothing AndAlso params("preview") IsNot Nothing Then
                    If params("preview").Type <> JTokenType.Boolean Then
                        Return ToolResult.Failed(toolId,
                                                 "CreateSlides preview must be a boolean",
                                                 errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                                 userMessage:="preview должен быть true или false",
                                                 recoverable:=True)
                    End If
                    previewRequested = previewRequested OrElse params.Value(Of Boolean)("preview")
                End If
                presentation = Globals.ThisAddIn.Application.ActivePresentation
                If presentation Is Nothing Then
                    Return ToolResult.Failed(toolId,
                                             "No active PowerPoint presentation",
                                             errorCode:=ExceptionClassifier.CodeDocMissing,
                                             userMessage:="Сначала откройте или создайте презентацию PowerPoint",
                                             recoverable:=False)
                End If
                Return ExecuteInternal(params, presentation, previewRequested)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            Finally
                ComObjectHelper.ReleaseComObject(presentation)
            End Try
        End Function

        ''' <summary>
        ''' Разрешает http(s)-ссылки в imagePath в локальные файлы (скачивает в локальный кэш).
        ''' Возвращает Nothing при успехе либо готовый failed ToolResult, чтобы Loop мог заменить картинку.
        ''' </summary>
        Private Shared Function ResolveSlideImages(spec As DeckDesignSpec) As ToolResult
            Const toolId As String = "CreateSlides"

            For index = 0 To spec.Slides.Count - 1
                Dim slideSpec = spec.Slides(index)
                If String.IsNullOrWhiteSpace(slideSpec.ImagePath) Then Continue For
                If Not ImageAcquisitionService.IsRemoteUrl(slideSpec.ImagePath) Then Continue For

                Dim download = ImageAcquisitionService.DownloadAsync(slideSpec.ImagePath).GetAwaiter().GetResult()
                If Not download.Success Then
                    Return ToolResult.Failed(toolId,
                                             $"Slide {index + 1}: cannot download image '{slideSpec.ImagePath}': {download.ErrorMessage}",
                                             errorCode:="IMAGE_FETCH_FAILED",
                                             userMessage:=$"Не удалось получить изображение для слайда {index + 1} по ссылке {slideSpec.ImagePath}: {download.ErrorMessage}",
                                             recoverable:=True)
                End If

                slideSpec.ImagePath = download.LocalPath
            Next

            Return Nothing
        End Function

        Private Shared Function ExecuteInternal(params As JObject,
                                                presentation As PowerPoint.Presentation,
                                                preview As Boolean) As ToolResult
            Const toolId As String = "CreateSlides"
            Dim spec = DeckDesignSpec.Parse(params)
            If spec.Slides.Count = 0 Then
                Return ToolResult.Failed(toolId, "CreateSlides requires at least one slide spec",
                                         errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                         userMessage:="Нет спецификаций дизайна слайдов для создания", recoverable:=True)
            End If
            If spec.Slides.Count > 50 Then
                Return ToolResult.Failed(toolId, "CreateSlides supports at most 50 slides per run",
                                         errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                         userMessage:="За один раз можно создать не более 50 слайдов", recoverable:=True)
            End If
            For index = 0 To spec.Slides.Count - 1
                If String.IsNullOrWhiteSpace(spec.Slides(index).Title) Then
                    Return ToolResult.Failed(toolId, $"Slide {index + 1} is missing a title",
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:=$"У слайда {index + 1} отсутствует заголовок", recoverable:=True)
                End If
                Dim sceneError = ValidateSceneSpec(spec.Slides(index))
                If Not String.IsNullOrWhiteSpace(sceneError) Then
                    Return ToolResult.Failed(toolId,
                                             $"Slide {index + 1}: {sceneError}",
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:=$"Недостаточно информации Scene для слайда {index + 1}: {sceneError}",
                                             recoverable:=True)
                End If
            Next

            ' imagePath может быть http(s)-ссылкой (в том числе на внутренний сайт).
            ' Скачиваем до компиляции плана: рендер умеет только локальные файлы.
            Dim imageResolveError = ResolveSlideImages(spec)
            If imageResolveError IsNot Nothing Then Return imageResolveError

            If Not DesignSystemCatalog.IsSupported(spec.DesignSystem) AndAlso spec.DesignTokens Is Nothing Then
                Return ToolResult.Failed(toolId,
                                         $"Unknown designSystem '{spec.DesignSystem}' without designTokens",
                                         errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                         userMessage:="Для неизвестной системы дизайна нужно предоставить полный designTokens или использовать зарегистрированную систему дизайна",
                                         recoverable:=True)
            End If

            Dim tokens = DesignSystemCatalog.Resolve(spec.DesignSystem, spec.DesignTokens)
            Dim pageSetup As PowerPoint.PageSetup = Nothing
            Dim slideWidth As Single
            Dim slideHeight As Single
            Try
                pageSetup = presentation.PageSetup
                slideWidth = pageSetup.SlideWidth
                slideHeight = pageSetup.SlideHeight
            Finally
                ComObjectHelper.ReleaseComObject(pageSetup)
            End Try
            Dim initialCount = GetRequiredSlideCount(presentation)
            Dim targetRefs As New List(Of String)()
            Dim slideResults As New JArray()
            Dim warnings As New List(Of String)()
            Dim createdCount As Integer = 0

            If preview Then
                Return ExecutePreview(spec, tokens, slideWidth, slideHeight, initialCount)
            End If

            Dim deckPlans As New List(Of SlideRenderPlan)()
            Dim deckPreflightReports As New List(Of VisualVerificationReport)()
            For preflightIndex = 0 To spec.Slides.Count - 1
                Dim preflightSpec = spec.Slides(preflightIndex)
                Try
                    Dim precompiledPlan = SlideLayoutEngine.Compile(preflightSpec, tokens, slideWidth, slideHeight,
                                                                    preflightIndex, spec.Slides.Count)
                    Dim precompiledReport = PowerPointVisualVerifier.PreflightAndRepair(precompiledPlan,
                                                                                        slideWidth, slideHeight)
                    deckPlans.Add(precompiledPlan)
                    deckPreflightReports.Add(precompiledReport)
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", preflightIndex + 1}, {"title", preflightSpec.Title},
                        {"slideType", preflightSpec.SlideType}, {"status", "failed"},
                        {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "deck_compile_preflight"}
                    })
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        ex.Message, classified.ErrorCode)
                End Try
                If Not deckPreflightReports(preflightIndex).Passed Then
                    slideResults.Add(BuildSlideResult(preflightIndex, preflightSpec, "failed",
                                                       deckPreflightReports(preflightIndex), Nothing,
                                                       "LAYOUT_VERIFY_FAILED"))
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        "Предпроверка профессионального макета не пройдена" & DescribeReportIssues(deckPreflightReports(preflightIndex)),
                                        ExceptionClassifier.CodeVerifyFailed)
                End If
            Next

            VerifyDeckComposition(deckPlans, deckPreflightReports)
            If deckPreflightReports.Any(Function(report) Not report.Passed) Then
                For preflightIndex = 0 To spec.Slides.Count - 1
                    Dim report = deckPreflightReports(preflightIndex)
                    slideResults.Add(BuildSlideResult(preflightIndex, spec.Slides(preflightIndex),
                                                       If(report.Passed, "not_rendered", "failed"),
                                                       report, Nothing,
                                                       If(report.Passed, "", "DECK_COMPOSITION_VERIFY_FAILED")))
                Next
                Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                    "Ритм композиционных изменений презентации не прошёл проверку",
                                    ExceptionClassifier.CodeVerifyFailed)
            End If

            For index = 0 To spec.Slides.Count - 1
                Dim slideSpec = spec.Slides(index)
                Dim plan As SlideRenderPlan = Nothing
                Dim preflight As VisualVerificationReport = Nothing
                Try
                    plan = SlideLayoutEngine.Compile(slideSpec, tokens, slideWidth, slideHeight, index, spec.Slides.Count)
                    preflight = PowerPointVisualVerifier.PreflightAndRepair(plan, slideWidth, slideHeight)
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", index + 1}, {"title", slideSpec.Title}, {"slideType", slideSpec.SlideType},
                        {"status", "failed"}, {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "compile_preflight"}
                    })
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        ex.Message, classified.ErrorCode)
                End Try
                If Not preflight.Passed Then
                    slideResults.Add(BuildSlideResult(index, slideSpec, "failed", preflight, Nothing, "LAYOUT_VERIFY_FAILED"))
                    Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        "Предпроверка профессионального макета не пройдена" & DescribeReportIssues(preflight),
                                        ExceptionClassifier.CodeVerifyFailed)
                End If

                Dim renderResult As SceneRenderResult = Nothing
                Try
                    renderResult = PowerPointSceneRenderer.Render(presentation, slideSpec, plan,
                                                                  initialCount + createdCount + 1, tokens)
                    createdCount += 1
                    Dim slideIndex = renderResult.Slide.SlideIndex
                    Dim targetRef = $"PowerPoint:presentations/active/slides/{slideIndex}"
                    targetRefs.Add(targetRef)
                    warnings.AddRange(renderResult.Warnings)
                    Dim renderedReport = PowerPointVisualVerifier.VerifyAndRepairRenderedSlide(
                        renderResult.Slide, slideWidth, slideHeight, plan, tokens)
                    If Not String.IsNullOrWhiteSpace(slideSpec.ImagePath) AndAlso
                       renderResult.Warnings.Any(Function(warning)
                           Return warning.IndexOf("Image skipped", StringComparison.OrdinalIgnoreCase) >= 0
                       End Function) Then
                        renderedReport.Issues.Add(New VisualIssue With {
                            .Code = "IMAGE_ARTIFACT_MISSING",
                            .Severity = "error",
                            .Message = $"Requested image could not be rendered: {slideSpec.ImagePath}"
                        })
                    End If
                    If Not String.IsNullOrWhiteSpace(slideSpec.Notes) AndAlso
                       renderResult.Warnings.Any(Function(warning)
                           Return warning.IndexOf("Speaker notes skipped", StringComparison.OrdinalIgnoreCase) >= 0
                       End Function) Then
                        renderedReport.Issues.Add(New VisualIssue With {
                            .Code = "NOTES_ARTIFACT_MISSING",
                            .Severity = "error",
                            .Message = "Requested speaker notes could not be rendered"
                        })
                    End If
                    MergeReport(preflight, renderedReport)
                    slideResults.Add(BuildSlideResult(index, slideSpec,
                                                      If(preflight.Passed, "succeeded", "failed"),
                                                      preflight, targetRef,
                                                      If(preflight.Passed, "", "VISUAL_VERIFY_FAILED")))
                    If Not preflight.Passed Then
                        Return BuildFailureWithVisualEvidence(CaptureVisualEvidence(renderResult.Slide, index + 1,
                                                                                     slideWidth, slideHeight),
                                                              presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                            "Проверка визуального качества после отрисовки не пройдена" & DescribeReportIssues(preflight),
                                            ExceptionClassifier.CodeVerifyFailed)
                    End If
                Catch ex As Exception
                    Dim classified = ExceptionClassifier.Classify(ex)
                    slideResults.Add(New JObject From {
                        {"index", index + 1}, {"title", slideSpec.Title}, {"slideType", slideSpec.SlideType},
                        {"status", "failed"}, {"errorCode", classified.ErrorCode}, {"message", ex.Message},
                        {"phase", "render_verify"}
                    })
                    Dim visualEvidence As AgentVisualEvidence = Nothing
                    If renderResult IsNot Nothing AndAlso renderResult.Slide IsNot Nothing Then
                        visualEvidence = CaptureVisualEvidence(renderResult.Slide, index + 1, slideWidth, slideHeight)
                    End If
                    Return BuildFailureWithVisualEvidence(visualEvidence,
                                                          presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                        ex.Message, classified.ErrorCode)
                Finally
                    If renderResult IsNot Nothing Then ComObjectHelper.ReleaseComObject(renderResult.Slide)
                End Try
            Next

            Try
                Dim observation = BuildObservation(spec, initialCount, createdCount, targetRefs, slideResults, warnings, True)
                Dim data As New JObject From {
                    {"designSystem", tokens.Name}, {"createdSlides", createdCount},
                    {"targetRefs", JArray.FromObject(targetRefs)}, {"slideResults", slideResults.DeepClone()}
                }
                Return ToolResult.Succeed("CreateSlides",
                                          $"С использованием системы дизайна {tokens.Name} создано профессиональных слайдов: {createdCount}",
                                          data:=data, observation:=observation,
                                          artifacts:=New JObject From {{"slides", JArray.FromObject(targetRefs)}})
            Catch ex As Exception
                Dim classified = ExceptionClassifier.Classify(ex)
                slideResults.Add(New JObject From {
                    {"index", spec.Slides.Count}, {"status", "failed"},
                    {"errorCode", classified.ErrorCode}, {"message", ex.Message}, {"phase", "finalize_result"}
                })
                Return BuildFailure(presentation, spec, initialCount, createdCount, targetRefs, slideResults, warnings,
                                    ex.Message, classified.ErrorCode)
            End Try
        End Function



        Private Shared Function ValidateSceneSpec(spec As SlideDesignSpec) As String
            If spec Is Nothing Then Return "Scene пуст"
            If Not spec.SlideTypeRecognized Then
                Return $"Unsupported slideType '{spec.RequestedSlideType}'; select a registered Scene archetype"
            End If
            Dim variantError = ValidateVariant(spec)
            If Not String.IsNullOrWhiteSpace(variantError) Then Return variantError
            If Not String.IsNullOrWhiteSpace(spec.ImagePath) AndAlso
               spec.SlideType <> "cover" AndAlso spec.SlideType <> "content" Then
                Return $"Страница {spec.SlideType} не использует imagePath; используйте макет cover/content с изображением либо удалите неиспользуемый imagePath"
            End If
            If spec.Chart IsNot Nothing Then
                If spec.SlideType <> "content" Then Return "chart сейчас допустим только на странице content"
                If Not String.IsNullOrWhiteSpace(spec.ImagePath) OrElse spec.Table IsNot Nothing Then
                    Return "Одна страница content может объявлять только один основной визуальный элемент из imagePath, chart, table"
                End If
                Dim chartError = ValidateChartSpec(spec.Chart)
                If Not String.IsNullOrWhiteSpace(chartError) Then Return chartError
            End If
            If spec.Table IsNot Nothing Then
                If spec.SlideType <> "content" Then Return "table сейчас допустим только на странице content"
                If Not String.IsNullOrWhiteSpace(spec.ImagePath) OrElse spec.Chart IsNot Nothing Then
                    Return "Одна страница content может объявлять только один основной визуальный элемент из imagePath, chart, table"
                End If
                Dim tableError = ValidateTableSpec(spec.Table)
                If Not String.IsNullOrWhiteSpace(tableError) Then Return tableError
            End If
            Select Case spec.SlideType
                Case "statement"
                    If spec.Items.Count > 3 Then Return "Страница statement вмещает не более 3 доказательств; разбейте страницу или используйте content"
                Case "content"
        If spec.Items.Count < 1 Then Return "Странице content нужен хотя бы 1 item"
        If spec.Items.Count > 6 Then Return "Страница content вмещает не более 6 items; разбейте страницу"
                    If spec.Chart IsNot Nothing AndAlso spec.Items.Count > 4 Then
                        Return "Страница content с chart вмещает не более 4 items-выводов"
                    End If
                    If spec.Table IsNot Nothing AndAlso spec.Items.Count > 4 Then
                        Return "Страница content с table вмещает не более 4 items-выводов"
                    End If
                    If Not String.IsNullOrWhiteSpace(spec.ImagePath) AndAlso spec.Items.Count > 4 Then
                        Return "Страница content с imagePath вмещает не более 4 items"
                    End If
                    If (String.Equals(spec.LayoutVariant, "feature-left", StringComparison.OrdinalIgnoreCase) OrElse
                        spec.Items.Any(Function(item) item.Emphasis)) AndAlso spec.Items.Count > 5 Then
                        Return "Страница feature-left content вмещает не более 5 items"
                    End If
                Case "two-column"
                    If spec.Items.Count <> 2 Then Return "Страница two-column должна содержать ровно 2 items"
                Case "comparison"
                    Dim isTable = spec.Items.Count >= 3 AndAlso
                                  spec.Items.All(Function(item) item.Features IsNot Nothing AndAlso item.Features.Count >= 2)
                    If isTable AndAlso spec.Items.Count > 5 Then Return "Таблица comparison вмещает не более 5 строк; разбейте страницу"
                    If isTable AndAlso (spec.ColumnHeaders Is Nothing OrElse spec.ColumnHeaders.Count <> 3 OrElse
                                        spec.ColumnHeaders.Any(Function(header) String.IsNullOrWhiteSpace(header))) Then
                        Return "Таблица comparison должна содержать 3 непустых columnHeaders: [критерий сравнения, левый вариант, правый вариант]"
                    End If
                    If Not isTable AndAlso spec.Items.Count <> 2 Then Return "Странице comparison нужно ровно 2 объекта сравнения или 3-5 строк двухколоночных features"
                Case "kpi"
        If spec.Metrics.Count < 2 AndAlso spec.Items.Count < 2 Then Return "Странице kpi нужно не менее 2 metrics"
        If spec.Metrics.Count > 4 OrElse spec.Items.Count > 4 Then Return "Страница kpi вмещает не более 4 показателей; разбейте страницу"
                    If spec.Metrics.Count >= 2 AndAlso
                       spec.Metrics.Any(Function(metric) String.IsNullOrWhiteSpace(metric.Value) OrElse String.IsNullOrWhiteSpace(metric.Label)) Then
                        Return "kpi metrics должны содержать и value, и label; не подменяйте показатели заглушками"
                    End If
                    If spec.Metrics.Count = 0 AndAlso
                       spec.Items.Any(Function(item) String.IsNullOrWhiteSpace(item.Value) OrElse String.IsNullOrWhiteSpace(item.Title)) Then
                        Return "kpi items должны содержать и value, и title; при отсутствии надёжных чисел используйте content/statement"
                    End If
                Case "process"
        If spec.Items.Count < 3 Then Return "Странице process нужно не менее 3 шагов"
        If spec.Items.Count > 6 Then Return "Страница process вмещает не более 6 шагов; разбейте процесс"
                Case "architecture"
        If spec.Items.Count < 2 Then Return "Странице architecture нужно не менее 2 уровней"
        If spec.Items.Count > 5 Then Return "Страница architecture вмещает не более 5 уровней; разбейте архитектуру"
                Case "matrix"
                    If spec.Items.Count <> 4 Then Return "Страница matrix должна содержать ровно 4 items-квадранта"
                    If String.IsNullOrWhiteSpace(spec.XAxisLabel) OrElse String.IsNullOrWhiteSpace(spec.YAxisLabel) Then
                        Return "Страница matrix должна предоставлять xAxisLabel и yAxisLabel; нельзя предполагать фиксированные бизнес-измерения"
                    End If
            End Select
            Return ""
        End Function

        Private Shared Function ValidateVariant(spec As SlideDesignSpec) As String
            If spec Is Nothing OrElse String.IsNullOrWhiteSpace(spec.LayoutVariant) Then Return ""
            Dim layoutVariant = spec.LayoutVariant.Trim().ToLowerInvariant()
            If layoutVariant = "default" OrElse layoutVariant = "standard" Then Return ""
            Select Case spec.SlideType
                Case "content"
                    If layoutVariant = "feature-left" Then Return ""
                Case "kpi"
                    If layoutVariant = "hero-left" Then Return ""
                Case "process"
                    If layoutVariant = "vertical" Then Return ""
                Case "architecture"
                    If layoutVariant = "hub-spoke" AndAlso spec.Items.Count >= 3 Then Return ""
            End Select
            If spec.SlideType = "architecture" AndAlso layoutVariant = "hub-spoke" Then
                Return "architecture hub-spoke requires one core item and at least two spoke items"
            End If
            Return $"variant '{spec.LayoutVariant}' is not supported by the {spec.SlideType} Scene"
        End Function

        Private Shared Function ValidateChartSpec(chart As DesignChart) As String
            If chart Is Nothing Then Return ""
            Dim chartType = If(chart.ChartType, "column").Trim().ToLowerInvariant()
            If chartType <> "column" AndAlso chartType <> "line" Then
                Return "chart.chartType поддерживает только column или line"
            End If
            If chart.Categories.Count < 2 OrElse chart.Categories.Count > 8 Then
                Return "chart.categories должны содержать 2-8 категорий"
            End If
            If chart.Categories.Any(Function(category) String.IsNullOrWhiteSpace(category)) Then
                Return "chart.categories не должны содержать пустые метки"
            End If
            If chart.Series.Count < 1 OrElse chart.Series.Count > 3 Then
                Return "chart.series должны содержать 1-3 серии"
            End If
            If chart.Series.Count > 1 AndAlso chart.Series.Any(Function(series) String.IsNullOrWhiteSpace(series.Name)) Then
                Return "Каждая series многосерийного chart должна содержать name"
            End If
            Dim hasNonZeroValue As Boolean = False
            For Each series In chart.Series
                If series.Values.Count <> chart.Categories.Count Then
                    Return "Количество values в каждой chart series должно совпадать с categories"
                End If
                For Each value In series.Values
                    If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                        Return "chart values должны быть конечными числами"
                    End If
                    If Math.Abs(value) > 0.000001R Then hasNonZeroValue = True
                Next
            Next
            If Not hasNonZeroValue Then Return "chart требует хотя бы одно ненулевое значение"
            Return ""
        End Function

        Private Shared Function ValidateTableSpec(table As DesignTable) As String
            If table Is Nothing Then Return ""
            If table.Headers.Count < 2 OrElse table.Headers.Count > 5 Then
                Return "table.headers должны содержать 2-5 столбцов"
            End If
            If table.Headers.Any(Function(header) String.IsNullOrWhiteSpace(header)) Then
                Return "table.headers не должны содержать пустые заголовки"
            End If
            If table.Rows.Count < 1 OrElse table.Rows.Count > 6 Then
                Return "table.rows должны содержать 1-6 строк"
            End If
            If table.Rows.Any(Function(row) row Is Nothing OrElse row.Count <> table.Headers.Count) Then
                Return "Количество ячеек в каждой строке table должно совпадать с headers"
            End If
            If table.HighlightColumn < -1 OrElse table.HighlightColumn >= table.Headers.Count Then
                Return "table.highlightColumn должен быть допустимым нулевым индексом столбца"
            End If
            Return ""
        End Function
    End Class

End Namespace
