Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon.Agent.Context
Imports System.Diagnostics

Namespace Context
    Public Class PowerPointContextProvider
        Implements IContextProvider

        Private ReadOnly _app As PowerPoint.Application

        Public Sub New(app As PowerPoint.Application)
            _app = app
        End Sub

        Public Function GetContext() As OfficeContext Implements IContextProvider.GetContext
            Dim ctx As New OfficeContext With {.AppType = "PowerPoint"}

            Try
                If _app.ActivePresentation IsNot Nothing Then
                    Dim pres As PowerPoint.Presentation = _app.ActivePresentation
                    Dim slideCount As Integer = pres.Slides.Count
                    Dim currentSlideIndex As Integer = 0

                    ' 获取当前幻灯片
                    Try
                        If _app.ActiveWindow IsNot Nothing AndAlso _app.ActiveWindow.View IsNot Nothing Then
                            ' View.Slide 的 COM 返回值在部分 PowerPoint 视图中不是可直接强转的
                            ' Slide RCW，会触发 first-chance InvalidCastException。显式晚绑定只读取
                            ' SlideIndex，避免无意义的接口强制转换。
                            Dim currentSlide As Object = _app.ActiveWindow.View.Slide
                            If currentSlide IsNot Nothing Then
                                currentSlideIndex = CInt(CallByName(currentSlide, "SlideIndex", CallType.Get))
                            End If
                        End If
                    Catch ex As Exception
                        Debug.WriteLine("获取当前幻灯片索引失败: " & ex.GetType().Name & ": " & ex.Message)
                        currentSlideIndex = 1
                    End Try

                    ' 选区信息
                    Dim selectionInfo As String = $"Слайд {currentSlideIndex}/{slideCount}"
                    Dim selectedShapeCount As Integer = 0
                    Dim formatDesc As String = ""

                    ' 获取选中的形状
                    Try
                        If _app.ActiveWindow IsNot Nothing AndAlso _app.ActiveWindow.Selection IsNot Nothing Then
                            Dim sel = _app.ActiveWindow.Selection
                            If sel.Type = PowerPoint.PpSelectionType.ppSelectionShapes Then
                                selectedShapeCount = sel.ShapeRange.Count

                                If selectedShapeCount > 0 Then
                                    Dim shape = sel.ShapeRange(1)
                                    selectionInfo = $"Слайд {currentSlideIndex}/{slideCount}, выбрано объектов: {selectedShapeCount}"

                                    ' 如果是文本框，获取格式信息
                                    If shape.HasTextFrame Then
                                        Try
                                            Dim textRange = shape.TextFrame.TextRange
                                            Dim fontSize As Single = textRange.Font.Size
                                            Dim fontName As String = textRange.Font.Name
                                            Dim isBold As Boolean = (textRange.Font.Bold = -1)
                                            Dim isItalic As Boolean = (textRange.Font.Italic = -1)

                                            formatDesc = $"Размер шрифта: {fontSize}pt, шрифт: {fontName}"
                                            If isBold Then formatDesc &= ", полужирный"
                                            If isItalic Then formatDesc &= ", курсив"

                                            ' 文本内容预览
                                            Dim textContent As String = textRange.Text
                                            If Not String.IsNullOrEmpty(textContent) Then
                                                If textContent.Length > 50 Then
                                                    textContent = textContent.Substring(0, 50) & "..."
                                                End If
                                                formatDesc &= $"{vbCrLf}Содержимое: {textContent}"
                                            End If

                                        Catch formatEx As Exception
                                            Debug.WriteLine("获取文本格式失败: " & formatEx.Message)
                                        End Try
                                    End If
                                End If
                            End If
                        End If
                    Catch selEx As Exception
                        Debug.WriteLine("获取选区失败: " & selEx.Message)
                    End Try

                    ctx.Selection = New SelectionInfo With {
                        .Address = selectionInfo,
                        .ItemCount = selectedShapeCount,
                        .DataType = If(selectedShapeCount > 0, "Фигуры/текстовое поле", "Слайд"),
                        .Preview = formatDesc
                    }

                    ' 文档结构信息
                    ctx.DocStructure = New DocumentStructure With {
                        .Summary = $"Презентация PowerPoint, слайдов: {slideCount}, текущий: {currentSlideIndex}"
                    }
                End If

            Catch ex As Exception
                Debug.WriteLine("获取PowerPoint上下文失败: " & ex.Message)
            End Try

            Return ctx
        End Function
    End Class
End Namespace
