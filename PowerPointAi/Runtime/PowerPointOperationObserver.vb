Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Office.Core
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend NotInheritable Class PowerPointOperationObserver
        Private Sub New()
        End Sub

        Public Shared Function CaptureState(batch As OfficeOperationBatch,
                                            Optional additionalTargetRefs As IEnumerable(Of String) = Nothing) As JObject
            Dim state As New JObject()
            Dim refs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If batch?.Operations IsNot Nothing Then
                For Each operation In batch.Operations
                    If operation IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(operation.TargetRef) Then refs.Add(operation.TargetRef)
                Next
            End If
            If additionalTargetRefs IsNot Nothing Then
                For Each targetRef In additionalTargetRefs
                    If Not String.IsNullOrWhiteSpace(targetRef) Then refs.Add(targetRef)
                Next
            End If

            Dim activePresentation As PowerPoint.Presentation = Nothing
            Try
                Try
                    activePresentation = Globals.ThisAddIn.Application.ActivePresentation
                Catch
                End Try
                If activePresentation IsNot Nothing Then
                    state("presentation") = If(activePresentation.Name, "")
                    Dim slides As PowerPoint.Slides = Nothing
                    Try
                        slides = activePresentation.Slides
                        state("slideCount") = slides.Count
                    Finally
                        ReleaseCom(slides)
                    End Try
                End If
            Catch ex As Exception
                state("captureError") = ex.Message
            Finally
                ReleaseCom(activePresentation)
            End Try

            Dim targets As New JObject()
            For Each targetRef In refs.OrderBy(Function(item) item, StringComparer.OrdinalIgnoreCase)
                Try
                    Using resolved = PowerPointObjectResolver.Resolve(targetRef)
                        targets(targetRef) = CaptureResolvedValue(resolved.Value, resolved.ObjectKind)
                    End Using
                Catch ex As Exception
                    targets(targetRef) = New JObject From {
                        {"exists", False},
                        {"error", ex.Message}
                    }
                End Try
            Next
            state("targets") = targets
            Return state
        End Function

        Public Shared Function BuildObservation(batch As OfficeOperationBatch,
                                                operationResults As JArray,
                                                beforeState As JObject,
                                                afterState As JObject,
                                                targetRefs As IEnumerable(Of String),
                                                warnings As IEnumerable(Of String)) As JObject
            Dim changed = Not JToken.DeepEquals(If(beforeState, New JObject()), If(afterState, New JObject()))
            Dim targetArray As New JArray()
            If targetRefs IsNot Nothing Then
                For Each targetRef In targetRefs.Where(Function(item) Not String.IsNullOrWhiteSpace(item)).Distinct(StringComparer.OrdinalIgnoreCase)
                    targetArray.Add(targetRef)
                Next
            End If

            Dim warningArray As New JArray()
            If warnings IsNot Nothing Then
                For Each warning In warnings.Where(Function(item) Not String.IsNullOrWhiteSpace(item))
                    warningArray.Add(warning)
                Next
            End If
            If Not changed Then warningArray.Add("Operation batch completed without an observable PowerPoint state change")

            Dim succeededCount = If(operationResults, New JArray()).OfType(Of JObject)().Count(
                Function(item) String.Equals(item("status")?.ToString(), "succeeded", StringComparison.OrdinalIgnoreCase))
            Dim totalCount = If(operationResults?.Count, 0)
            Dim writeExpected = batch?.Operations IsNot Nothing AndAlso
                                batch.Operations.Any(Function(operation) operation IsNot Nothing AndAlso
                                    Not String.Equals(operation.Action, "get", StringComparison.OrdinalIgnoreCase))
            Return New JObject From {
                {"kind", "office_operation_batch"},
                {"summary", $"Декларативные операции PowerPoint выполнены: {succeededCount}/{totalCount}"},
                {"changed", changed},
                {"writeExpected", writeExpected},
                {"appType", "PowerPoint"},
                {"atomic", If(batch?.Atomic, True)},
                {"targetRefs", targetArray},
                {"operations", If(operationResults, New JArray())},
                {"before", If(beforeState, New JObject())},
                {"after", If(afterState, New JObject())},
                {"diff", BuildDiff(beforeState, afterState)},
                {"warnings", warningArray}
            }
        End Function

        Public Shared Function VerifyExpectedEffects(batch As OfficeOperationBatch,
                                                     operationResults As JArray,
                                                     afterState As JObject) As JArray
            Dim verification As New JArray()
            If batch Is Nothing OrElse afterState Is Nothing Then Return verification
            Dim targets = TryCast(afterState("targets"), JObject)
            If targets Is Nothing Then Return verification

            If batch.Operations IsNot Nothing Then
                For Each operation In batch.Operations
                    If operation Is Nothing OrElse operation.ExpectedEffects Is Nothing OrElse Not operation.ExpectedEffects.HasValues Then Continue For
                    Dim operationResult = operationResults?.OfType(Of JObject)().FirstOrDefault(
                        Function(item) String.Equals(item("id")?.ToString(), operation.Id, StringComparison.OrdinalIgnoreCase))
                    Dim resultRef = operationResult?("resultRef")?.ToString()
                    Dim verifyRef = If(String.IsNullOrWhiteSpace(resultRef), operation.TargetRef, resultRef)
                    Dim snapshot = TryCast(targets(verifyRef), JObject)
                    For Each expectedProperty In operation.ExpectedEffects.Properties()
                        Dim actual = If(snapshot Is Nothing, Nothing, snapshot(expectedProperty.Name))
                        Dim passed = TokensEquivalent(actual, expectedProperty.Value)
                        verification.Add(New JObject From {
                            {"id", $"{operation.Id}:{expectedProperty.Name}"},
                            {"required", True},
                            {"targetRef", verifyRef},
                            {"property", expectedProperty.Name},
                            {"status", If(passed, "passed", "failed")},
                            {"expected", expectedProperty.Value.DeepClone()},
                            {"actual", If(actual?.DeepClone(), JValue.CreateNull())}
                        })
                    Next
                Next
            End If

            If batch.SuccessCriteria IsNot Nothing Then
                For index = 0 To batch.SuccessCriteria.Count - 1
                    Dim criterion = batch.SuccessCriteria(index)
                    If criterion Is Nothing OrElse String.IsNullOrWhiteSpace(criterion.TargetRef) Then Continue For
                    Dim snapshot = TryCast(targets(criterion.TargetRef), JObject)
                    Dim actual = If(snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(criterion.PropertyName),
                                    Nothing,
                                    snapshot(criterion.PropertyName))
                    Dim passed = EvaluateCriterion(actual, criterion)
                    verification.Add(New JObject From {
                        {"id", If(criterion.Id, $"criterion-{index + 1}")},
                        {"required", criterion.Required},
                        {"targetRef", criterion.TargetRef},
                        {"property", If(criterion.PropertyName, "")},
                        {"operator", If(criterion.Operator, "equals")},
                        {"status", If(passed, "passed", "failed")},
                        {"expected", If(criterion.ExpectedValue?.DeepClone(), JValue.CreateNull())},
                        {"actual", If(actual?.DeepClone(), JValue.CreateNull())}
                    })
                Next
            End If
            Return verification
        End Function

        Public Shared Function HasRequiredVerificationFailure(verification As JArray) As Boolean
            If verification Is Nothing Then Return False
            Return verification.OfType(Of JObject)().Any(
                Function(item) If(item("required")?.Value(Of Boolean)(), True) AndAlso
                    String.Equals(item("status")?.ToString(), "failed", StringComparison.OrdinalIgnoreCase))
        End Function

        Private Shared Function EvaluateCriterion(actual As JToken, criterion As OperationCriterion) As Boolean
            Dim op = If(criterion?.Operator, "equals").Trim().ToLowerInvariant()
            Select Case op
                Case "exists"
                    Return actual IsNot Nothing AndAlso actual.Type <> JTokenType.Null
                Case "not_equals"
                    Return Not TokensEquivalent(actual, criterion.ExpectedValue)
                Case "contains"
                    Return actual IsNot Nothing AndAlso criterion.ExpectedValue IsNot Nothing AndAlso
                           actual.ToString().IndexOf(criterion.ExpectedValue.ToString(), StringComparison.OrdinalIgnoreCase) >= 0
                Case "gte", "lte"
                    Dim actualNumber As Decimal
                    Dim expectedNumber As Decimal
                    If Not Decimal.TryParse(If(actual?.ToString(), ""), actualNumber) OrElse
                       Not Decimal.TryParse(If(criterion.ExpectedValue?.ToString(), ""), expectedNumber) Then Return False
                    Return If(op = "gte", actualNumber >= expectedNumber, actualNumber <= expectedNumber)
                Case Else
                    Return TokensEquivalent(actual, criterion.ExpectedValue)
            End Select
        End Function

        Private Shared Function TokensEquivalent(actual As JToken, expected As JToken) As Boolean
            If actual Is Nothing OrElse expected Is Nothing Then Return actual Is Nothing AndAlso expected Is Nothing
            If actual.Type = JTokenType.Array AndAlso expected.Type = JTokenType.Array Then
                Dim actualArray = DirectCast(actual, JArray)
                Dim expectedArray = DirectCast(expected, JArray)
                If actualArray.Count <> expectedArray.Count Then Return False
                For index = 0 To actualArray.Count - 1
                    If Not TokensEquivalent(actualArray(index), expectedArray(index)) Then Return False
                Next
                Return True
            End If
            If actual.Type = JTokenType.String OrElse expected.Type = JTokenType.String Then
                Return String.Equals(NormalizeOfficeText(actual.ToString()),
                                     NormalizeOfficeText(expected.ToString()),
                                     StringComparison.Ordinal)
            End If
            Return JToken.DeepEquals(actual, expected)
        End Function

        Private Shared Function NormalizeOfficeText(value As String) As String
            Return If(value, "").Replace(vbCr, "").Trim()
        End Function

        Private Shared Function CaptureResolvedValue(value As Object, objectKind As String) As JObject
            Dim snapshot As New JObject From {
                {"exists", value IsNot Nothing},
                {"objectKind", If(objectKind, "")}
            }
            If value Is Nothing Then Return snapshot

            If TypeOf value Is PowerPoint.Presentation Then
                Dim presentation = DirectCast(value, PowerPoint.Presentation)
                snapshot("name") = If(presentation.Name, "")
                Dim slides As PowerPoint.Slides = Nothing
                Try
                    slides = presentation.Slides
                    snapshot("slideCount") = slides.Count
                Finally
                    ReleaseCom(slides)
                End Try
            ElseIf TypeOf value Is PowerPoint.Slides Then
                snapshot("count") = DirectCast(value, PowerPoint.Slides).Count
            ElseIf TypeOf value Is PowerPoint.Slide Then
                Dim slide = DirectCast(value, PowerPoint.Slide)
                snapshot("slideIndex") = slide.SlideIndex
                Dim shapes As PowerPoint.Shapes = Nothing
                Try
                    shapes = slide.Shapes
                    snapshot("shapeCount") = shapes.Count
                Finally
                    ReleaseCom(shapes)
                End Try
            ElseIf TypeOf value Is PowerPoint.Shapes Then
                snapshot("shapeCount") = DirectCast(value, PowerPoint.Shapes).Count
            ElseIf TypeOf value Is PowerPoint.Shape Then
                CaptureShape(DirectCast(value, PowerPoint.Shape), snapshot)
            ElseIf TypeOf value Is SmartArt Then
                CaptureSmartArt(DirectCast(value, SmartArt), snapshot)
            ElseIf TypeOf value Is SmartArtNodes Then
                CaptureSmartArtNodes(DirectCast(value, SmartArtNodes), snapshot)
            ElseIf TypeOf value Is SmartArtNode Then
                snapshot("text") = GetNodeText(DirectCast(value, SmartArtNode))
                snapshot("textHash") = ComputeHash(snapshot("text")?.ToString())
            ElseIf TypeOf value Is TextRange2 Then
                Dim text = If(DirectCast(value, TextRange2).Text, "")
                snapshot("text") = text
                snapshot("textHash") = ComputeHash(text)
            ElseIf TypeOf value Is PowerPoint.TextRange Then
                Dim text = If(DirectCast(value, PowerPoint.TextRange).Text, "")
                snapshot("text") = text
                snapshot("textHash") = ComputeHash(text)
            End If
            Return snapshot
        End Function

        Private Shared Sub CaptureShape(shape As PowerPoint.Shape, snapshot As JObject)
            snapshot("shapeId") = shape.Id
            snapshot("name") = If(shape.Name, "")
            snapshot("shapeType") = CInt(shape.Type)
            Dim hasSmartArt As Boolean = False
            Try
                hasSmartArt = shape.HasSmartArt = MsoTriState.msoTrue
            Catch
            End Try
            snapshot("hasSmartArt") = hasSmartArt
            If Not hasSmartArt Then Return

            Dim smartArt As SmartArt = Nothing
            Try
                smartArt = shape.SmartArt
                If smartArt IsNot Nothing Then CaptureSmartArt(smartArt, snapshot)
            Finally
                ReleaseCom(smartArt)
            End Try
        End Sub

        Private Shared Sub CaptureSmartArt(smartArt As SmartArt, snapshot As JObject)
            Dim layout As SmartArtLayout = Nothing
            Dim nodes As SmartArtNodes = Nothing
            Try
                layout = smartArt.Layout
                If layout IsNot Nothing Then snapshot("smartArtLayoutId") = If(layout.Id, "")
            Catch
            Finally
                ReleaseCom(layout)
            End Try

            Try
                nodes = smartArt.AllNodes
                If nodes IsNot Nothing Then CaptureSmartArtNodes(nodes, snapshot)
            Finally
                ReleaseCom(nodes)
            End Try
        End Sub

        Private Shared Sub CaptureSmartArtNodes(nodes As SmartArtNodes, snapshot As JObject)
            Dim texts As New List(Of String)()
            snapshot("nodeCount") = nodes.Count
            For index = 1 To nodes.Count
                Dim node As SmartArtNode = Nothing
                Try
                    node = nodes.Item(index)
                    texts.Add(GetNodeText(node))
                Catch
                    texts.Add("")
                Finally
                    ReleaseCom(node)
                End Try
            Next
            Dim combined = String.Join(vbLf, texts)
            snapshot("nodeTexts") = JArray.FromObject(texts)
            snapshot("nodeTextHash") = ComputeHash(combined)
        End Sub

        Private Shared Function GetNodeText(node As SmartArtNode) As String
            If node Is Nothing Then Return ""
            Dim frame As TextFrame2 = Nothing
            Dim range As TextRange2 = Nothing
            Try
                frame = node.TextFrame2
                If frame Is Nothing Then Return ""
                range = frame.TextRange
                Return If(range?.Text, "")
            Catch
                Return ""
            Finally
                ReleaseCom(range)
                ReleaseCom(frame)
            End Try
        End Function

        Private Shared Function BuildDiff(beforeState As JObject, afterState As JObject) As JObject
            Return New JObject From {
                {"slideCountDelta", GetInteger(afterState, "slideCount") - GetInteger(beforeState, "slideCount")},
                {"targetStateChanged", Not JToken.DeepEquals(beforeState?("targets"), afterState?("targets"))}
            }
        End Function

        Private Shared Function GetInteger(source As JObject, name As String) As Integer
            If source Is Nothing Then Return 0
            Return If(source(name)?.Value(Of Integer)(), 0)
        End Function

        Private Shared Function ComputeHash(value As String) As String
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(If(value, "")))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Sub ReleaseCom(value As Object)
            If value Is Nothing Then Return
            Try
                If Marshal.IsComObject(value) Then Marshal.ReleaseComObject(value)
            Catch
            End Try
        End Sub
    End Class

End Namespace
