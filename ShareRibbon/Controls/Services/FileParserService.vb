' ShareRibbon\Controls\Services\FileParserService.vb
' 文件解析服务：文本、CSV、编码检测等

Imports System.IO
Imports System.Text

''' <summary>
''' 文件解析服务，负责解析文本文件、CSV文件和检测文件编码
''' </summary>
Public Class FileParserService

#Region "文本文件解析"

        ''' <summary>
        ''' 解析文本文件
        ''' </summary>
        Public Function ParseTextFile(filePath As String) As FileContentResult
            Try
                Dim extension As String = Path.GetExtension(filePath).ToLower()

                ' 对 CSV 文件使用专门的处理逻辑
                If extension = ".csv" Then
                    Return ParseCsvFile(filePath)
                End If

                ' 对普通文本文件进行编码检测
                Dim encoding As Encoding = DetectFileEncoding(filePath)
                Dim content As String = File.ReadAllText(filePath, encoding)

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "Text",
                    .ParsedContent = content,
                    .RawData = content
                }
            Catch ex As Exception
                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "Text",
                    .ParsedContent = $"[Ошибка разбора текстового файла: {ex.Message}]"
                }
            End Try
        End Function

#End Region

#Region "CSV 文件解析"

        ''' <summary>
        ''' Разбирает CSV по RFC 4180: кавычки, удвоенные кавычки и переносы строк внутри полей.
        ''' </summary>
        Public Function ParseCsvFile(filePath As String) As FileContentResult
            Try
                Dim encoding As Encoding = DetectFileEncoding(filePath)
                Dim csvContent As String = File.ReadAllText(filePath, encoding)

                Dim formattedContent As New StringBuilder()
                formattedContent.AppendLine($"CSV-файл: {Path.GetFileName(filePath)} (кодировка: {encoding.EncodingName})")
                formattedContent.AppendLine()

                Dim delimiter As Char = DetectCsvDelimiterByRecords(csvContent)
                Dim records = ParseCsvRecords(csvContent, delimiter)

                If records.Count = 0 Then
                    formattedContent.AppendLine("[CSV-файл пуст]")
                Else
                    Dim headers = records(0)
                    formattedContent.AppendLine($"Заголовки ({headers.Count} колонок, разделитель '{delimiter}'):")
                    formattedContent.AppendLine(FormatCsvRecord(headers))
                    formattedContent.AppendLine()
                    formattedContent.AppendLine($"Данные (всего записей: {records.Count - 1}):")

                    Dim budget As Integer = 12000
                    Dim shown As Integer = 0
                    For i = 1 To records.Count - 1
                        Dim line = FormatCsvRecord(records(i))
                        If formattedContent.Length + line.Length > budget Then Exit For
                        formattedContent.AppendLine(line)
                        shown += 1
                    Next

                    If shown < records.Count - 1 Then
                        formattedContent.AppendLine("...")
                        formattedContent.AppendLine($"[Показано {shown} из {records.Count - 1} строк из-за лимита контекста]")
                    End If
                End If

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "CSV",
                    .ParsedContent = formattedContent.ToString(),
                    .RawData = csvContent
                }
            Catch ex As Exception
                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "CSV",
                    .ParsedContent = $"[Ошибка разбора CSV-файла: {ex.Message}]"
                }
            End Try
        End Function

        Private Shared Function FormatCsvRecord(record As List(Of String)) As String
            Return String.Join(" | ", record)
        End Function

        ''' <summary>
        ''' Потоковый разбор CSV с поддержкой кавычек и многострочных полей.
        ''' </summary>
        Private Shared Function ParseCsvRecords(text As String, delimiter As Char) As List(Of List(Of String))
            Dim records As New List(Of List(Of String))()
            If String.IsNullOrEmpty(text) Then Return records

            Dim record As New List(Of String)()
            Dim field As New StringBuilder()
            Dim inQuotes As Boolean = False
            Dim i As Integer = 0

            While i < text.Length
                Dim c = text(i)
                If inQuotes Then
                    If c = """"c Then
                        If i + 1 < text.Length AndAlso text(i + 1) = """"c Then
                            field.Append(""""c)
                            i += 2
                        Else
                            inQuotes = False
                            i += 1
                        End If
                    Else
                        field.Append(c)
                        i += 1
                    End If
                    Continue While
                End If

                If c = """"c AndAlso field.Length = 0 Then
                    inQuotes = True
                    i += 1
                ElseIf c = delimiter Then
                    record.Add(field.ToString())
                    field.Clear()
                    i += 1
                ElseIf c = vbCr OrElse c = vbLf Then
                    record.Add(field.ToString())
                    field.Clear()
                    AddCsvRecord(records, record)
                    record = New List(Of String)()
                    If c = vbCr AndAlso i + 1 < text.Length AndAlso text(i + 1) = vbLf Then i += 1
                    i += 1
                Else
                    field.Append(c)
                    i += 1
                End If
            End While

            If field.Length > 0 OrElse record.Count > 0 Then
                record.Add(field.ToString())
                AddCsvRecord(records, record)
            End If
            Return records
        End Function

        Private Shared Sub AddCsvRecord(records As List(Of List(Of String)), record As List(Of String))
            If record Is Nothing OrElse record.Count = 0 Then Return
            If record.Count = 1 AndAlso String.IsNullOrWhiteSpace(record(0)) Then Return
            records.Add(record)
        End Sub

        ''' <summary>
        ''' Выбирает разделитель по устойчивости числа колонок в первых записях.
        ''' </summary>
        Private Shared Function DetectCsvDelimiterByRecords(text As String) As Char
            Dim candidates As Char() = {","c, ";"c, vbTab, "|"c}
            Dim best As Char = ","c
            Dim bestScore As Double = -1

            Dim lines = text.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
            Dim sampleLines As New List(Of String)()
            For i = 0 To Math.Min(lines.Length, 10) - 1
                If Not String.IsNullOrWhiteSpace(lines(i)) Then sampleLines.Add(lines(i))
            Next
            Dim sample = String.Join(vbLf, sampleLines)

            For Each delimiter In candidates
                Dim records = ParseCsvRecords(sample, delimiter)
                If records.Count = 0 Then Continue For

                Dim widthCounts As New Dictionary(Of Integer, Integer)()
                For Each record In records
                    Dim width = record.Count
                    If width < 2 Then Continue For
                    If widthCounts.ContainsKey(width) Then
                        widthCounts(width) += 1
                    Else
                        widthCounts(width) = 1
                    End If
                Next

                If widthCounts.Count = 0 Then Continue For

                Dim bestWidth = 0
                Dim bestFrequency = 0
                For Each pair In widthCounts
                    If pair.Value > bestFrequency Then
                        bestFrequency = pair.Value
                        bestWidth = pair.Key
                    End If
                Next

                Dim score = bestWidth * (bestFrequency / Math.Max(1, records.Count))
                If score > bestScore Then
                    bestScore = score
                    best = delimiter
                End If
            Next

            Return best
        End Function

        ''' <summary>
        ''' Совместимый helper: определяет разделитель по одной строке.
        ''' </summary>
        Public Function DetectCsvDelimiter(sampleLine As String) As Char
            Dim possibleDelimiters As Char() = {","c, ";"c, vbTab, "|"c}
            Dim bestDelimiter As Char = ","c
            Dim maxCount As Integer = 0

            For Each delimiter In possibleDelimiters
                Dim count As Integer = sampleLine.Count(Function(c) c = delimiter)
                If count > maxCount Then
                    maxCount = count
                    bestDelimiter = delimiter
                End If
            Next

            Return bestDelimiter
        End Function

#End Region

#Region "编码检测"

        ''' <summary>
        ''' 检测文件编码
        ''' </summary>
        Public Function DetectFileEncoding(filePath As String) As Encoding
            Try
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                    ' 读取前几个字节来检测 BOM
                    Dim bom(3) As Byte
                    Dim bytesRead As Integer = fs.Read(bom, 0, bom.Length)

                    ' UTF-8 with BOM
                    If bytesRead >= 3 AndAlso bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
                        Return New UTF8Encoding(True)
                    End If

                    ' UTF-16 Big Endian
                    If bytesRead >= 2 AndAlso bom(0) = &HFE AndAlso bom(1) = &HFF Then
                        Return Encoding.BigEndianUnicode
                    End If

                    ' UTF-16 Little Endian
                    If bytesRead >= 2 AndAlso bom(0) = &HFF AndAlso bom(1) = &HFE Then
                        If bytesRead >= 4 AndAlso bom(2) = 0 AndAlso bom(3) = 0 Then
                            Return Encoding.UTF32
                        Else
                            Return Encoding.Unicode
                        End If
                    End If

                    ' UTF-32 Big Endian
                    If bytesRead >= 4 AndAlso bom(0) = 0 AndAlso bom(1) = 0 AndAlso bom(2) = &HFE AndAlso bom(3) = &HFF Then
                        Return New UTF32Encoding(True, True)
                    End If
                End Using

                ' Кодировка определяется по содержимому: UTF-8 проверяется первым,
                ' GB18030 рассматривается только если байты не являются валидным UTF-8.
                Return DetectEncodingByContent(filePath)

            Catch ex As Exception
                Return Encoding.Default
            End Try
        End Function

        ''' <summary>
        ''' Выбирает кодировку по содержимому: сначала строгая проверка UTF-8, затем GB18030.
        ''' Русский UTF-8 CSV без BOM больше не определяется как GBK.
        ''' </summary>
        Private Function DetectEncodingByContent(filePath As String) As Encoding
            Dim sampleBytes = ReadSampleBytes(filePath, 65536)
            If sampleBytes.Length = 0 Then Return Encoding.Default

            If IsValidUtf8(sampleBytes) Then Return New UTF8Encoding(False)

            Try
                Dim strictGbk = Encoding.GetEncoding("GB18030",
                                                    EncoderFallback.ExceptionFallback,
                                                    DecoderFallback.ExceptionFallback)
                Dim text = strictGbk.GetString(TrimIncompleteTail(sampleBytes))
                If Not String.IsNullOrEmpty(text) Then Return Encoding.GetEncoding("GB18030")
            Catch
            End Try

            Return Encoding.Default
        End Function

        Private Shared Function ReadSampleBytes(filePath As String, maxBytes As Integer) As Byte()
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                Dim length = CInt(Math.Min(fs.Length, maxBytes))
                If length <= 0 Then Return New Byte() {}
                Dim buffer(length - 1) As Byte
                Dim read = 0
                While read < length
                    Dim chunk = fs.Read(buffer, read, length - read)
                    If chunk <= 0 Then Exit While
                    read += chunk
                End While
                If read = length Then Return buffer
                Dim result(read - 1) As Byte
                Array.Copy(buffer, result, read)
                Return result
            End Using
        End Function

        Private Shared Function IsValidUtf8(bytes As Byte()) As Boolean
            Try
                Dim strictUtf8 As New UTF8Encoding(False, True)
                strictUtf8.GetString(TrimIncompleteTail(bytes))
                Return True
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Отбрасывает до 3 последних байт, чтобы обрезанный на границе сэмпла
        ''' многобайтовый символ не делал валидный UTF-8 невалидным.
        ''' </summary>
        Private Shared Function TrimIncompleteTail(bytes As Byte()) As Byte()
            If bytes Is Nothing OrElse bytes.Length <= 4 Then Return bytes
            Dim length = bytes.Length
            Dim i = length - 1
            Dim walked = 0
            While i >= 0 AndAlso walked < 4 AndAlso (bytes(i) And &HC0) = &H80
                walked += 1
                i -= 1
            End While
            If i < 0 Then Return bytes
            Dim lead = bytes(i)
            Dim expected As Integer = 1
            If (lead And &HE0) = &HC0 Then
                expected = 2
            ElseIf (lead And &HF0) = &HE0 Then
                expected = 3
            ElseIf (lead And &HF8) = &HF0 Then
                expected = 4
            End If
            If expected > 1 AndAlso (length - i) < expected Then
                Dim result(i - 1) As Byte
                Array.Copy(bytes, result, i)
                Return result
            End If
            Return bytes
        End Function

#End Region

    End Class
