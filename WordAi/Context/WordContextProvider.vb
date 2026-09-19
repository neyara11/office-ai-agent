Imports Word = Microsoft.Office.Interop.Word
Imports ShareRibbon.Agent.Context
Imports System.Diagnostics

Namespace Context
    Public Class WordContextProvider
        Implements IContextProvider

        Private ReadOnly _app As Word.Application

        Public Sub New(app As Word.Application)
            _app = app
        End Sub

        Public Function GetContext() As OfficeContext Implements IContextProvider.GetContext
            Dim ctx As New OfficeContext With {.AppType = "Word"}

            Try
                Dim doc As Word.Document = _app.ActiveDocument
                Dim sel As Word.Selection = _app.Selection

                ' === 1. 获取选中内容（如果有） ===
                If sel IsNot Nothing Then
                    Dim selectedText As String = sel.Text
                    Dim textLength As Integer = If(selectedText IsNot Nothing, selectedText.Length, 0)

                    ctx.Selection = New SelectionInfo With {
                        .Address = $"Выбрано символов: {textLength}",
                        .ItemCount = textLength,
                        .DataType = If(textLength > 0, "Текст", "Позиция курсора")
                    }

                    ' 格式信息
                    If textLength > 0 Then
                        Try
                            Dim fontSize As Single = sel.Font.Size
                            Dim fontName As String = sel.Font.Name
                            Dim isBold As Boolean = (sel.Font.Bold = -1)
                            Dim isItalic As Boolean = (sel.Font.Italic = -1)
                            Dim isUnderline As Boolean = (sel.Font.Underline <> Word.WdUnderline.wdUnderlineNone)
                            Dim styleName As String = ""
                            Try
                                styleName = sel.Style.NameLocal
                            Catch
                                styleName = "Основной текст"
                            End Try

                            Dim formatDesc As String = $"Размер шрифта: {fontSize}pt, шрифт: {fontName}"
                            If isBold Then formatDesc &= ", полужирный"
                            If isItalic Then formatDesc &= ", курсив"
                            If isUnderline Then formatDesc &= ", подчёркивание"
                            If Not String.IsNullOrEmpty(styleName) Then formatDesc &= $", стиль: {styleName}"

                            Dim preview As String = selectedText
                            If preview.Length > 100 Then
                                preview = preview.Substring(0, 100) & "..."
                            End If

                            ctx.Selection.Preview = $"Содержимое: {preview}{vbCrLf}Формат: {formatDesc}"

                        Catch formatEx As Exception
                            Debug.WriteLine("获取格式信息失败: " & formatEx.Message)
                        End Try
                    End If
                End If

                ' === 2. 获取文档结构（即使没有选中，也要提供上下文） ===
                Dim structureInfo As New Text.StringBuilder()
                structureInfo.AppendLine($"Документ Word, абзацев: {doc.Paragraphs.Count}")

                ' 获取标题结构
                Dim headings As New List(Of String)()
                Try
                    For Each para As Word.Paragraph In doc.Paragraphs
                        If para.Range.Words.Count > 0 Then
                            Dim styleName As String = para.Style.NameLocal
                            If styleName.Contains("标题") Then
                                Dim headingText As String = para.Range.Text.Trim()
                                If headingText.Length > 50 Then
                                    headingText = headingText.Substring(0, 50) & "..."
                                End If
                                headings.Add($"- {styleName}: {headingText}")
                                If headings.Count >= 10 Then Exit For ' 最多显示 10 个标题
                            End If
                        End If
                    Next

                    If headings.Count > 0 Then
                        structureInfo.AppendLine()
                        structureInfo.AppendLine("Заголовки документа:")
                        For Each heading In headings
                            structureInfo.AppendLine(heading)
                        Next
                    End If
                Catch headingEx As Exception
                    Debug.WriteLine("获取标题失败: " & headingEx.Message)
                End Try

                ' 获取当前段落上下文（前后各2段）
                Try
                    If sel.Paragraphs.Count > 0 Then
                        Dim currentPara As Word.Paragraph = sel.Paragraphs(1)

                        ' 查找当前段落在文档中的索引
                        Dim paraIndex As Integer = 1
                        For Each para As Word.Paragraph In doc.Paragraphs
                            If para.Range.Start = currentPara.Range.Start Then
                                Exit For
                            End If
                            paraIndex += 1
                        Next

                        structureInfo.AppendLine()
                        structureInfo.AppendLine($"Текущая позиция: абзац {paraIndex}")
                        structureInfo.AppendLine()
                        structureInfo.AppendLine("Контекст:")

                        ' 前2段
                        For i As Integer = Math.Max(1, paraIndex - 2) To paraIndex - 1
                            If i <= doc.Paragraphs.Count Then
                                Dim paraText As String = doc.Paragraphs(i).Range.Text.Trim()
                                If paraText.Length > 30 Then paraText = paraText.Substring(0, 30) & "..."
                                If Not String.IsNullOrEmpty(paraText) Then
                                    structureInfo.AppendLine($"  [{i}] {paraText}")
                                End If
                            End If
                        Next

                        ' 当前段
                        Dim currentText As String = currentPara.Range.Text.Trim()
                        If currentText.Length > 50 Then currentText = currentText.Substring(0, 50) & "..."
                        structureInfo.AppendLine($"→ [{paraIndex}] {currentText} (текущий)")

                        ' 后2段
                        For i As Integer = paraIndex + 1 To Math.Min(doc.Paragraphs.Count, paraIndex + 2)
                            Dim paraText As String = doc.Paragraphs(i).Range.Text.Trim()
                            If paraText.Length > 30 Then paraText = paraText.Substring(0, 30) & "..."
                            If Not String.IsNullOrEmpty(paraText) Then
                                structureInfo.AppendLine($"  [{i}] {paraText}")
                            End If
                        Next
                    End If

                Catch contextEx As Exception
                    Debug.WriteLine("获取上下文失败: " & contextEx.Message)
                End Try

                ctx.DocStructure = New DocumentStructure With {
                    .Summary = structureInfo.ToString()
                }

            Catch ex As Exception
                Debug.WriteLine("获取Word上下文失败: " & ex.Message)
            End Try

            Return ctx
        End Function
    End Class
End Namespace
