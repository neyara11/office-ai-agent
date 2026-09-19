Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports ShareRibbon
Imports Word = Microsoft.Office.Interop.Word

''' <summary>
''' Word文档翻译服务 - 支持一键翻译和沉浸式翻译
''' </summary>
Public Class WordDocumentTranslateService
    Inherits DocumentTranslateService

    Private _wordApp As Word.Application
    Private _document As Document
    ' OpenXML 级翻译器（新版）
    Private _openXmlTranslator As OpenXmlWordTranslator
    Private _translationBlocks As List(Of OpenXmlWordTranslator.TranslationBlock)

    Public Sub New(wordApp As Word.Application)
        MyBase.New()
        _wordApp = wordApp
        _document = wordApp.ActiveDocument
    End Sub

    ''' <summary>
    ''' 获取文档所有段落（使用 OpenXML 级扫描）
    ''' </summary>
    Public Overrides Function GetAllParagraphs() As List(Of String)
        _openXmlTranslator = New OpenXmlWordTranslator()
        _translationBlocks = _openXmlTranslator.ScanDocument(_document)

        Return _translationBlocks.Select(Function(b) b.OriginalText).ToList()
    End Function

    ''' <summary>
    ''' 获取选中的段落（使用 OpenXML 级扫描）
    ''' </summary>
    Public Overrides Function GetSelectedParagraphs() As List(Of String)
        _openXmlTranslator = New OpenXmlWordTranslator()
        _translationBlocks = _openXmlTranslator.ScanSelection(_document, _wordApp.Selection)

        Return _translationBlocks.Select(Function(b) b.OriginalText).ToList()
    End Function

    ''' <summary>
    ''' 应用翻译结果到整个文档（使用 OpenXML 级写入）
    ''' </summary>
    Public Overrides Sub ApplyTranslation(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        If outputMode = TranslateOutputMode.NewDocument Then
            ApplyToNewDocument(results)
            Return
        End If

        If outputMode = TranslateOutputMode.SidePanel Then
            ' 侧栏模式由调用方处理
            Return
        End If

        If _openXmlTranslator IsNot Nothing AndAlso _translationBlocks IsNot Nothing Then
            Dim settings = TranslateSettings.Load()
            _openXmlTranslator.ApplyTranslation(_document, _translationBlocks, results, outputMode, settings)
        End If
    End Sub

    ''' <summary>
    ''' 应用翻译结果到选中区域（使用 OpenXML 级写入）
    ''' </summary>
    Public Overrides Sub ApplyTranslationToSelection(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        If outputMode = TranslateOutputMode.NewDocument Then
            ApplyToNewDocument(results)
            Return
        End If

        If outputMode = TranslateOutputMode.SidePanel Then
            ' 侧栏模式由调用方处理
            Return
        End If

        If _openXmlTranslator IsNot Nothing AndAlso _translationBlocks IsNot Nothing Then
            Dim settings = TranslateSettings.Load()
            _openXmlTranslator.ApplyTranslation(_document, _translationBlocks, results, outputMode, settings)
        End If
    End Sub

    ''' <summary>
    ''' 创建新文档并写入翻译结果（保留原文档格式）
    ''' </summary>
    Private Sub ApplyToNewDocument(results As List(Of TranslateParagraphResult))
        Try
            ' 复制原文档到新文档（保留所有格式）
            _document.Content.Copy()
            Dim newDoc = _wordApp.Documents.Add()
            newDoc.Content.Paste()

            ' 替换每个段落的文本内容（保留格式）
            Dim paras = newDoc.Paragraphs
            For i = 0 To Math.Min(results.Count, paras.Count) - 1
                Dim result = results(i)
                If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                    Dim para = paras(i + 1)
                    ' 仅替换文本，不改变格式
                    Dim paraRange = para.Range
                    Dim originalEnd = paraRange.End
                    
                    ' 保存原始格式
                    Dim fontName = paraRange.Font.Name
                    Dim fontSize = paraRange.Font.Size
                    Dim fontBold = paraRange.Font.Bold
                    Dim fontItalic = paraRange.Font.Italic
                    Dim fontColor = paraRange.Font.Color
                    Dim paraAlignment = para.Alignment
                    Dim firstLineIndent = para.FirstLineIndent
                    Dim leftIndent = para.LeftIndent
                    Dim rightIndent = para.RightIndent
                    Dim lineSpacing = para.LineSpacing
                    
                    ' 替换文本
                    paraRange.Text = result.TranslatedText & vbCr
                    
                    ' 恢复格式
                    paraRange.Font.Name = fontName
                    paraRange.Font.Size = fontSize
                    paraRange.Font.Bold = fontBold
                    paraRange.Font.Italic = fontItalic
                    paraRange.Font.Color = fontColor
                    para.Alignment = paraAlignment
                    para.FirstLineIndent = firstLineIndent
                    para.LeftIndent = leftIndent
                    para.RightIndent = rightIndent
                    para.LineSpacing = lineSpacing
                End If
            Next

            newDoc.Activate()
        Catch ex As Exception
            MessageBox.Show("Ошибка при создании нового документа: " & ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 生成翻译结果的格式化文本（用于侧栏显示）
    ''' </summary>
    Public Function FormatResultsForDisplay(results As List(Of TranslateParagraphResult), showOriginal As Boolean) As String
        Dim sb As New StringBuilder()

        For Each result In results
            If showOriginal Then
                sb.AppendLine("【Оригинал】")
                sb.AppendLine(result.OriginalText)
                sb.AppendLine()
                sb.AppendLine("【Перевод】")
            End If

            If result.Success Then
                sb.AppendLine(result.TranslatedText)
            Else
                sb.AppendLine($"[Ошибка перевода: {result.ErrorMessage}]")
                sb.AppendLine(result.OriginalText)
            End If

            If showOriginal Then
                sb.AppendLine(New String("-"c, 40))
            End If
            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function
End Class
