' ShareRibbon\Services\Reformat\TempDocumentService.vb
' 临时文档服务 - 管理排版/校对过程中的临时文件

Imports System.IO

''' <summary>
''' 临时文档服务 - 原文档 ↔ 临时文档的复制与清理
''' 所有排版/校对操作在临时文档上进行，确认后才合并回原文档
''' </summary>
Public Class TempDocumentService

    Private Shared ReadOnly _tempDir As String = Path.Combine(
        Path.GetTempPath(), "OfficeAI_Reformat")

    Private Shared _lock As New Object()

    ''' <summary>
    ''' 创建原文档的临时副本
    ''' </summary>
    ''' <param name="sourceDocPath">原文档完整路径</param>
    ''' <returns>临时文档路径</returns>
    Public Shared Function CreateTempDocument(sourceDocPath As String) As String
        EnsureTempDirectory()

        If String.IsNullOrWhiteSpace(sourceDocPath) OrElse Not File.Exists(sourceDocPath) Then
            Throw New FileNotFoundException($"Недопустимый путь к исходному документу или файл не существует: {sourceDocPath}", sourceDocPath)
        End If

        Dim docName = Path.GetFileNameWithoutExtension(sourceDocPath)
        Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim guidStr As String = Guid.NewGuid().ToString("N").Substring(0, 8)
        Dim tempFileName = $"{docName}_{timestamp}_{guidStr}.docx"
        Dim tempPath = Path.Combine(_tempDir, tempFileName)

        File.Copy(sourceDocPath, tempPath, overwrite:=True)

        Return tempPath
    End Function

    ''' <summary>
    ''' 从Word Document对象创建临时副本（获取路径后复制）
    ''' </summary>
    Public Shared Function CreateTempDocument(sourceDoc As Object) As String
        EnsureTempDirectory()

        Dim docPath As String = ""
        Try
            docPath = sourceDoc.FullName
        Catch
            docPath = ""
        End Try

        If Not String.IsNullOrWhiteSpace(docPath) AndAlso File.Exists(docPath) Then
            Return CreateTempDocument(docPath)
        End If

        ' 未保存文档的 FullName 通常是“文档1”，不是文件路径。用 Word 临时文档承接 FormattedText，
        ' 避免 SaveAs2 改变用户当前文档的保存位置。
        Dim tempName = $"UnsavedDoc_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("N").Substring(0, 8)}.docx"
        Dim tempPath = Path.Combine(_tempDir, tempName)
        Dim tempDoc As Object = Nothing

        Try
            Dim wordApp = sourceDoc.Application
            tempDoc = wordApp.Documents.Add()
            CopyDocumentContent(sourceDoc, tempDoc)
            tempDoc.SaveAs2(FileName:=tempPath, FileFormat:=12, AddToRecentFiles:=False)
        Catch ex As Exception
            Throw New InvalidOperationException($"Не удалось создать временную копию несохранённого документа: {ex.Message}", ex)
        Finally
            If tempDoc IsNot Nothing Then
                Try
                    tempDoc.Close(SaveChanges:=False)
                Catch
                End Try
            End If
        End Try

        Return tempPath
    End Function

    Private Shared Sub CopyDocumentContent(sourceDoc As Object, tempDoc As Object)
        Dim sourceRange As Object = Nothing
        Dim targetRange As Object = Nothing

        Try
            sourceRange = sourceDoc.Content
            targetRange = tempDoc.Content

            ' Word 的 Range.FormattedText 是可赋值属性，但链式 late binding
            ' tempDoc.Content.FormattedText = sourceDoc.Content.FormattedText
            ' 在部分宿主中会被 VB 运行时判定为“不是一种引用属性”。
            targetRange.FormattedText = sourceRange.FormattedText
        Catch ex As Exception
            Debug.WriteLine($"TempDocumentService.CopyDocumentContent FormattedText 复制失败，退回纯文本: {ex.Message}")
            Try
                If targetRange Is Nothing Then targetRange = tempDoc.Content
                If sourceRange Is Nothing Then sourceRange = sourceDoc.Content
                targetRange.Text = sourceRange.Text
            Catch fallbackEx As Exception
                Throw New InvalidOperationException($"Не удалось скопировать содержимое несохранённого документа: {fallbackEx.Message}", fallbackEx)
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 安全删除临时文件
    ''' </summary>
    Public Shared Sub Cleanup(tempPath As String)
        If String.IsNullOrEmpty(tempPath) Then Return

        Try
            If File.Exists(tempPath) Then
                File.Delete(tempPath)
            End If
        Catch ex As Exception
            Debug.WriteLine($"TempDocumentService.Cleanup 删除失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 确保临时目录存在
    ''' </summary>
    Private Shared Sub EnsureTempDirectory()
        SyncLock _lock
            If Not Directory.Exists(_tempDir) Then
                Directory.CreateDirectory(_tempDir)
            End If
        End SyncLock
    End Sub

End Class
