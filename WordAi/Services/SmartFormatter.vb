' WordAi/Services/SmartFormatter.vb
' 智能格式化器 - 根据查询直接设置格式，无需先选中

Imports Word = Microsoft.Office.Interop.Word
Imports System.Diagnostics

Namespace Services

    ''' <summary>
    ''' 智能格式化器 - 根据自然语言描述直接格式化内容
    ''' </summary>
    Public Class SmartFormatter

        Private ReadOnly _app As Word.Application
        Private ReadOnly _locator As ContentLocator
        Public Sub New(app As Word.Application)
            _app = app
            _locator = New ContentLocator(app)
        End Sub

        ''' <summary>
        ''' 根据查询设置格式
        ''' </summary>
        ''' <param name="query">目标查询，如"标题"、"第3段"</param>
        ''' <param name="action">操作描述，如"调大"、"加粗"、"改为红色"</param>
        Public Function FormatByQuery(query As String, action As String) As Boolean
            Try
                Debug.WriteLine($"[SmartFormatter] 查询: {query}, 操作: {action}")

                ' 1. 先定位到目标
                If Not _locator.FindAndSelect(query) Then
                    Debug.WriteLine($"[SmartFormatter] 无法定位到: {query}")
                    Return False
                End If

                ' 2. 获取当前选区
                Dim sel As Word.Selection = _app.Selection
                If sel Is Nothing Then
                    Return False
                End If

                ' 3. 解析并执行操作
                Return ExecuteFormatAction(sel, action)

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] FormatByQuery 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 执行格式化操作
        ''' </summary>
        Private Function ExecuteFormatAction(sel As Word.Selection, action As String) As Boolean
            Try
                Dim actionLower As String = action.ToLower().Trim()

                ' 字号调整
                If actionLower.Contains("调大") OrElse actionLower.Contains("增大") OrElse actionLower.Contains("变大") OrElse
                   actionLower.Contains("увелич") OrElse actionLower.Contains("крупнее") OrElse
                   actionLower.Contains("больше") Then
                    Return IncreaseFontSize(sel)
                End If

                If actionLower.Contains("调小") OrElse actionLower.Contains("缩小") OrElse actionLower.Contains("变小") OrElse
                   actionLower.Contains("уменьш") OrElse actionLower.Contains("мельче") OrElse
                   actionLower.Contains("меньше") Then
                    Return DecreaseFontSize(sel)
                End If

                ' 加粗/斜体/下划线
                If actionLower.Contains("加粗") OrElse actionLower.Contains("粗体") OrElse
                   actionLower.Contains("жирн") Then
                    sel.Font.Bold = -1
                    Debug.WriteLine("[SmartFormatter] 设置加粗")
                    Return True
                End If

                If actionLower.Contains("取消加粗") OrElse actionLower.Contains("不加粗") OrElse
                   actionLower.Contains("убрать жирн") OrElse actionLower.Contains("не жирн") Then
                    sel.Font.Bold = 0
                    Debug.WriteLine("[SmartFormatter] 取消加粗")
                    Return True
                End If

                If actionLower.Contains("斜体") OrElse actionLower.Contains("курсив") Then
                    sel.Font.Italic = -1
                    Debug.WriteLine("[SmartFormatter] 设置斜体")
                    Return True
                End If

                If actionLower.Contains("下划线") OrElse actionLower.Contains("подчерк") Then
                    sel.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                    Debug.WriteLine("[SmartFormatter] 设置下划线")
                    Return True
                End If

                ' 颜色
                If actionLower.Contains("红色") OrElse actionLower.Contains("красн") Then
                    sel.Font.Color = Word.WdColor.wdColorRed
                    Debug.WriteLine("[SmartFormatter] 设置红色")
                    Return True
                End If

                If actionLower.Contains("蓝色") OrElse actionLower.Contains("син") Then
                    sel.Font.Color = Word.WdColor.wdColorBlue
                    Debug.WriteLine("[SmartFormatter] 设置蓝色")
                    Return True
                End If

                If actionLower.Contains("黑色") OrElse actionLower.Contains("черн") Then
                    sel.Font.Color = Word.WdColor.wdColorBlack
                    Debug.WriteLine("[SmartFormatter] 设置黑色")
                    Return True
                End If

                ' 对齐
                If actionLower.Contains("居中") OrElse actionLower.Contains("по центру") OrElse
                   actionLower.Contains("центр") Then
                    sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Debug.WriteLine("[SmartFormatter] 设置居中")
                    Return True
                End If

                If actionLower.Contains("左对齐") OrElse actionLower.Contains("по левому") OrElse
                   actionLower.Contains("влево") Then
                    sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    Debug.WriteLine("[SmartFormatter] 设置左对齐")
                    Return True
                End If

                If actionLower.Contains("右对齐") OrElse actionLower.Contains("по правому") OrElse
                   actionLower.Contains("вправо") Then
                    sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    Debug.WriteLine("[SmartFormatter] 设置右对齐")
                    Return True
                End If

                Debug.WriteLine($"[SmartFormatter] 无法识别操作: {action}")
                Return False

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] ExecuteFormatAction 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 增大字号（智能递增）
        ''' </summary>
        Private Function IncreaseFontSize(sel As Word.Selection) As Boolean
            Try
                Dim currentSize As Single = sel.Font.Size
                Dim newSize As Single

                ' 智能递增：小字号+2，大字号+4
                If currentSize < 12 Then
                    newSize = currentSize + 2
                ElseIf currentSize < 18 Then
                    newSize = currentSize + 2
                ElseIf currentSize < 24 Then
                    newSize = currentSize + 4
                Else
                    newSize = currentSize + 4
                End If

                sel.Font.Size = newSize
                Debug.WriteLine($"[SmartFormatter] 字号 {currentSize}pt → {newSize}pt")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] IncreaseFontSize 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 减小字号（智能递减）
        ''' </summary>
        Private Function DecreaseFontSize(sel As Word.Selection) As Boolean
            Try
                Dim currentSize As Single = sel.Font.Size
                Dim newSize As Single

                ' 智能递减：大字号-4，小字号-2
                If currentSize <= 12 Then
                    newSize = Math.Max(8, currentSize - 2)
                ElseIf currentSize <= 18 Then
                    newSize = currentSize - 2
                ElseIf currentSize <= 24 Then
                    newSize = currentSize - 4
                Else
                    newSize = currentSize - 4
                End If

                sel.Font.Size = newSize
                Debug.WriteLine($"[SmartFormatter] 字号 {currentSize}pt → {newSize}pt")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] DecreaseFontSize 失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 判断是否是可直接执行的 Word 格式命令。
        ''' </summary>
        Public Shared Function LooksLikeDirectFormattingCommand(message As String) As Boolean
            Return FormattingIntentCompiler.LooksLikeDirectFormattingCommand(message)
        End Function

        ''' <summary>
        ''' 直接执行自然语言格式命令，并返回结构化执行结果。
        ''' </summary>
        Public Function ApplyNaturalLanguageFormatDetailed(command As String) As FormattingExecutionResult
            Dim result As New FormattingExecutionResult()
            If String.IsNullOrWhiteSpace(command) Then
                result.ErrorMessage = "Команда форматирования пуста"
                Return result
            End If

            Dim compiler As New FormattingIntentCompiler()
            Dim plan = compiler.Compile(command, HasUsableSelection())
            result.Plan = plan
            If plan Is Nothing OrElse Not plan.HasOperations Then
                result.ErrorMessage = "Не распознано ни одной выполняемой операции форматирования"
                Return result
            End If

            Dim undoStarted As Boolean = False
            Try
                Try
                    If _app.UndoRecord IsNot Nothing Then
                        _app.UndoRecord.StartCustomRecord("Настройка формата ИИ")
                        undoStarted = True
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[SmartFormatter] UndoRecord start failed: {ex.Message}")
                End Try

                result = ApplyPlanDetailed(plan)
                If result.Success Then
                    Debug.WriteLine($"[SmartFormatter] 已执行格式计划: {plan.ToHumanReadableSummary()}")
                Else
                    Debug.WriteLine($"[SmartFormatter] 格式计划无可应用目标: {command}")
                End If

                Return result
            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] ApplyNaturalLanguageFormat 失败: {ex.Message}")
                result.Success = False
                result.ErrorMessage = ex.Message
                Return result
            Finally
                If undoStarted Then
                    Try
                        _app.UndoRecord.EndCustomRecord()
                    Catch ex As Exception
                        Debug.WriteLine($"[SmartFormatter] UndoRecord end failed: {ex.Message}")
                    End Try
                End If
            End Try
        End Function

        Private Function HasUsableSelection() As Boolean
            Try
                Dim sel = _app.Selection
                If sel Is Nothing OrElse sel.Range Is Nothing Then Return False
                Dim selectedText = If(sel.Range.Text, "").Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
                Return selectedText.Length > 0
            Catch
                Return False
            End Try
        End Function

        Private Function ApplyPlanDetailed(plan As FormattingIntentPlan) As FormattingExecutionResult
            Dim result As New FormattingExecutionResult With {
                .Plan = plan
            }

            Dim ranges = ResolveTargetRanges(plan.Scope)
            If ranges Is Nothing OrElse ranges.Count = 0 Then
                result.ErrorMessage = "Не удалось определить целевой диапазон"
                Return result
            End If

            Dim changed As Boolean = False
            For Each rng In ranges
                If rng Is Nothing Then Continue For
                Dim rangeChanged As Boolean = False
                For Each op In plan.Operations
                    If ApplyOperationToRange(rng, op) Then
                        rangeChanged = True
                        result.AppliedOperationCount += 1
                    End If
                Next
                If rangeChanged Then result.AppliedRangeCount += 1
                changed = rangeChanged OrElse changed
            Next
            result.Success = changed
            If Not changed Then result.ErrorMessage = "Целевой диапазон существует, но ни одна операция не была успешно применена"
            Return result
        End Function

        Private Function ResolveTargetRanges(scope As FormattingTargetScope) As List(Of Word.Range)
            Dim result As New List(Of Word.Range)()
            Dim doc = _app.ActiveDocument
            If doc Is Nothing Then Return result

            Select Case scope
                Case FormattingTargetScope.Selection
                    If HasUsableSelection() Then result.Add(_app.Selection.Range)
                Case FormattingTargetScope.CurrentParagraph
                    If _app.Selection IsNot Nothing AndAlso _app.Selection.Paragraphs.Count > 0 Then
                        result.Add(_app.Selection.Paragraphs(1).Range)
                    End If
                Case FormattingTargetScope.Headings
                    For Each para As Word.Paragraph In doc.Paragraphs
                        If IsHeadingParagraph(para) Then result.Add(para.Range)
                    Next
                Case FormattingTargetScope.Body
                    For Each para As Word.Paragraph In doc.Paragraphs
                        If Not IsHeadingParagraph(para) Then result.Add(para.Range)
                    Next
                Case Else
                    result.Add(doc.Content)
            End Select

            If result.Count = 0 Then result.Add(doc.Content)
            Return result
        End Function

        Private Function IsHeadingParagraph(para As Word.Paragraph) As Boolean
            Try
                If para.OutlineLevel <> Word.WdOutlineLevel.wdOutlineLevelBodyText Then Return True
                Dim styleName = para.Style?.NameLocal?.ToString()
                If Not String.IsNullOrWhiteSpace(styleName) AndAlso
                   (styleName.Contains("标题") OrElse styleName.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    styleName.IndexOf("Заголовок", StringComparison.OrdinalIgnoreCase) >= 0) Then
                    Return True
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] IsHeadingParagraph failed: {ex.Message}")
            End Try
            Return False
        End Function

        Private Function ApplyOperationToRange(targetRange As Word.Range, op As FormattingOperation) As Boolean
            Select Case op.Kind
                Case FormattingOperationKind.FontSizeDelta
                    ApplyFontSizeDelta(targetRange, op.NumericValue)
                Case FormattingOperationKind.FontSizeGradeDelta
                    ApplyFontSizeGradeDelta(targetRange, op.NumericValue)
                Case FormattingOperationKind.FontSizeAbsolute
                    targetRange.Font.Size = CSng(op.NumericValue)
                Case FormattingOperationKind.FontFamily
                    targetRange.Font.Name = op.TextValue
                    targetRange.Font.NameFarEast = op.TextValue
                Case FormattingOperationKind.Bold
                    targetRange.Font.Bold = If(op.BooleanValue, -1, 0)
                Case FormattingOperationKind.Italic
                    targetRange.Font.Italic = If(op.BooleanValue, -1, 0)
                Case FormattingOperationKind.Underline
                    targetRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                Case FormattingOperationKind.FontColor
                    ApplyFontColor(targetRange, op.TextValue)
                Case FormattingOperationKind.Alignment
                    ApplyAlignment(targetRange, op.TextValue)
                Case FormattingOperationKind.LineSpacing
                    targetRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
                    targetRange.ParagraphFormat.LineSpacing = CSng(12 * op.NumericValue)
                Case FormattingOperationKind.FirstLineIndent
                    targetRange.ParagraphFormat.FirstLineIndent = CSng(op.NumericValue * targetRange.Font.Size)
                Case Else
                    Return False
            End Select
            Return True
        End Function

        Private Sub ApplyFontColor(targetRange As Word.Range, colorName As String)
            Select Case If(colorName, "").ToLowerInvariant()
                Case "red"
                    targetRange.Font.Color = Word.WdColor.wdColorRed
                Case "blue"
                    targetRange.Font.Color = Word.WdColor.wdColorBlue
                Case "green"
                    targetRange.Font.Color = Word.WdColor.wdColorGreen
                Case Else
                    targetRange.Font.Color = Word.WdColor.wdColorBlack
            End Select
        End Sub

        Private Sub ApplyAlignment(targetRange As Word.Range, alignment As String)
            Select Case If(alignment, "").ToLowerInvariant()
                Case "center"
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Case "right"
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                Case "justify"
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                Case Else
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            End Select
        End Sub

        Private Sub ApplyFontSizeDelta(targetRange As Word.Range, delta As Double)
            Try
                If targetRange.Paragraphs IsNot Nothing AndAlso targetRange.Paragraphs.Count > 0 Then
                    For Each para As Word.Paragraph In targetRange.Paragraphs
                        Dim rng = para.Range
                        Dim current = NormalizeFontSize(rng.Font.Size)
                        rng.Font.Size = CSng(Math.Max(8, Math.Min(72, current + delta)))
                    Next
                Else
                    Dim current = NormalizeFontSize(targetRange.Font.Size)
                    targetRange.Font.Size = CSng(Math.Max(8, Math.Min(72, current + delta)))
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] ApplyFontSizeDelta 失败: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Sub ApplyFontSizeGradeDelta(targetRange As Word.Range, gradeDelta As Double)
            Try
                If targetRange.Paragraphs IsNot Nothing AndAlso targetRange.Paragraphs.Count > 0 Then
                    For Each para As Word.Paragraph In targetRange.Paragraphs
                        Dim rng = para.Range
                        rng.Font.Size = CSng(ShiftChineseFontSizeGrade(NormalizeFontSize(rng.Font.Size), gradeDelta))
                    Next
                Else
                    targetRange.Font.Size = CSng(ShiftChineseFontSizeGrade(NormalizeFontSize(targetRange.Font.Size), gradeDelta))
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SmartFormatter] ApplyFontSizeGradeDelta 失败: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Function ShiftChineseFontSizeGrade(currentSize As Double, gradeDelta As Double) As Double
            Dim grades As Double() = {5, 5.5, 6.5, 7.5, 9, 10.5, 12, 14, 15, 16, 18, 22, 24, 26, 36, 42}
            Dim nearestIndex As Integer = 0
            Dim nearestDistance As Double = Double.MaxValue

            For i = 0 To grades.Length - 1
                Dim distance = Math.Abs(grades(i) - currentSize)
                If distance < nearestDistance Then
                    nearestDistance = distance
                    nearestIndex = i
                End If
            Next

            Dim targetIndex = nearestIndex + CInt(Math.Round(gradeDelta))
            targetIndex = Math.Max(0, Math.Min(grades.Length - 1, targetIndex))
            Return grades(targetIndex)
        End Function

        Private Function NormalizeFontSize(value As Single) As Double
            If value > 0 AndAlso value < 200 Then Return value
            Return 12
        End Function

    End Class

End Namespace
