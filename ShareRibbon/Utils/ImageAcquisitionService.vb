Imports System.IO
Imports System.Net.Http
Imports System.Threading
Imports System.Threading.Tasks

''' <summary>
''' Скачивание изображений по http(s)-ссылке в локальный файл.
''' Используется захватом страниц и рендером презентаций: Scene.imagePath может быть
''' ссылкой, в том числе на внутренний сайт без выхода в интернет.
''' </summary>
Public NotInheritable Class ImageAcquisitionService

    ''' <summary>Максимальный размер скачиваемого изображения.</summary>
    Public Const MaxImageBytes As Long = 15L * 1024 * 1024

    Public Shared ReadOnly DefaultTimeout As TimeSpan = TimeSpan.FromSeconds(20)

    Private Sub New()
    End Sub

    Public Class ImageDownloadResult
        Public Property Success As Boolean
        Public Property LocalPath As String
        Public Property ContentType As String
        Public Property ErrorMessage As String
        ''' <summary>Ошибка цепочки сертификата (самоподписанный или внутренний CA).</summary>
        Public Property TlsFailure As Boolean
    End Class

    ''' <summary>Является ли значение абсолютной http(s)-ссылкой.</summary>
    Public Shared Function IsRemoteUrl(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return False

        Dim trimmed = value.Trim()
        Return trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
               trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Каталог локального кэша картинок (переживает повторный рендер/repair).</summary>
    Public Shared Function GetImageCacheDirectory() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OfficeAI",
            "images")
    End Function

    ''' <summary>
    ''' Возвращает локальный путь к изображению: http(s)-ссылка скачивается,
    ''' относительная ссылка достраивается от baseUrl, локальный путь проверяется на существование.
    ''' </summary>
    Public Shared Async Function ResolveLocalPathAsync(value As String,
                                                       Optional baseUrl As String = Nothing) As Task(Of ImageDownloadResult)
        Dim result As New ImageDownloadResult()
        If String.IsNullOrWhiteSpace(value) Then
            result.ErrorMessage = "пустой путь к изображению"
            Return result
        End If

        Dim candidate = value.Trim()

        If IsRemoteUrl(candidate) Then
            Return Await DownloadAsync(candidate).ConfigureAwait(False)
        End If

        If Not String.IsNullOrWhiteSpace(baseUrl) Then
            Dim resolved As Uri = Nothing
            If Uri.TryCreate(baseUrl, UriKind.Absolute, resolved) AndAlso
               Uri.TryCreate(resolved, candidate, resolved) AndAlso
               IsRemoteUrl(resolved.ToString()) Then
                Return Await DownloadAsync(resolved.ToString()).ConfigureAwait(False)
            End If
        End If

        If File.Exists(candidate) Then
            result.Success = True
            result.LocalPath = candidate
            Return result
        End If

        result.ErrorMessage = $"файл не найден: {candidate}"
        Return result
    End Function

    ''' <summary>Скачивает изображение и возвращает байты (для хостов, которым нужен поток, а не путь).</summary>
    Public Shared Async Function DownloadBytesAsync(url As String,
                                                    Optional timeout As TimeSpan? = Nothing) As Task(Of Byte())
        Dim download = Await DownloadAsync(url, timeout).ConfigureAwait(False)
        If Not download.Success OrElse String.IsNullOrEmpty(download.LocalPath) Then Return Nothing
        If Not File.Exists(download.LocalPath) Then Return Nothing

        Return File.ReadAllBytes(download.LocalPath)
    End Function

    ''' <summary>
    ''' Скачивает изображение по абсолютному http(s) URL в локальный файл.
    ''' Сначала выполняется обычная проверка сертификата; если она не прошла,
    ''' делается одна повторная попытка без валидации цепочки — во внутреннем контуре
    ''' у сайта может быть самоподписанный сертификат или внутренний CA на любом домене.
    ''' </summary>
    Public Shared Async Function DownloadAsync(url As String,
                                              Optional timeout As TimeSpan? = Nothing) As Task(Of ImageDownloadResult)
        Dim strict = Await DownloadCoreAsync(url, allowInvalidCertificate:=False, timeout:=timeout).ConfigureAwait(False)
        If strict.Success OrElse Not strict.TlsFailure Then Return strict

        Dim permissive = Await DownloadCoreAsync(url, allowInvalidCertificate:=True, timeout:=timeout).ConfigureAwait(False)
        If permissive.Success Then
            Debug.WriteLine($"[ImageAcquisition] Сертификат сайта не прошёл проверку, изображение получено без валидации цепочки: {url}")
            Return permissive
        End If

        ' Если без проверки цепочки запрос прошёл TLS, отдаём его ошибку — она точнее (404, не картинка, лимит).
        If Not permissive.TlsFailure Then Return permissive

        Return strict
    End Function

    ''' <summary>Скачивает изображение и возвращает локальный файл в кэш.</summary>
    Private Shared Async Function DownloadCoreAsync(url As String,
                                                    allowInvalidCertificate As Boolean,
                                                    timeout As TimeSpan?) As Task(Of ImageDownloadResult)
        Dim result As New ImageDownloadResult()
        Dim client As HttpClient = Nothing
        Try
            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(url, UriKind.Absolute, uri) Then
                result.ErrorMessage = $"некорректный URL: {url}"
                Return result
            End If

            Dim effectiveTimeout = If(timeout.HasValue, timeout.Value, DefaultTimeout)

            client = CreateImageClient(uri, allowInvalidCertificate)
            Using cts As New CancellationTokenSource(effectiveTimeout)
                Using request As New HttpRequestMessage(HttpMethod.Get, uri)
                    request.Headers.TryAddWithoutValidation("Accept", "image/*,*/*;q=0.8")
                    ' Часть сайтов (в т.ч. Wikimedia) отдаёт 403 без User-Agent.
                    request.Headers.TryAddWithoutValidation("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OfficeAI/1.0")
                    Using response = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cts.Token).ConfigureAwait(False)
                        If Not response.IsSuccessStatusCode Then
                            result.ErrorMessage = $"HTTP {CInt(response.StatusCode)} {response.ReasonPhrase}"
                            Return result
                        End If

                        Dim declaredLength = response.Content.Headers.ContentLength
                        If declaredLength.HasValue AndAlso declaredLength.Value > MaxImageBytes Then
                            result.ErrorMessage = $"файл слишком большой ({declaredLength.Value} байт)"
                            Return result
                        End If

                        ' На .NET Framework чтение тела ответа может игнорировать токен отмены
                        ' (сервер держит соединение открытым), поэтому ограничиваем чтение по времени
                        ' и при таймауте принудительно закрываем ответ — это рвёт сокет.
                        Dim readTask = ReadWithLimitAsync(response, cts.Token)
                        Dim finished = Await Task.WhenAny(readTask, Task.Delay(effectiveTimeout)).ConfigureAwait(False)
                        If finished IsNot readTask Then
                            ' Не оставляем необработанное исключение от брошенного чтения.
                            readTask.ContinueWith(Sub(t)
                                                      Dim ignored = t.Exception
                                                  End Sub, TaskContinuationOptions.OnlyOnFaulted)
                            Try
                                response.Dispose()
                            Catch
                            End Try

                            result.ErrorMessage = $"превышено время ожидания ({effectiveTimeout.TotalSeconds:0} с)"
                            Return result
                        End If

                        Dim bytes = Await readTask.ConfigureAwait(False)
                        If bytes Is Nothing Then
                            result.ErrorMessage = $"превышен лимит размера {MaxImageBytes} байт"
                            Return result
                        End If

                        Dim extension = DetectImageExtension(bytes)
                        If extension Is Nothing Then
                            result.ErrorMessage = "ответ не является изображением (поддерживаются jpeg, png, gif, bmp, tiff)"
                            Return result
                        End If

                        Dim targetDir = GetImageCacheDirectory()
                        Directory.CreateDirectory(targetDir)
                        Dim targetPath = Path.Combine(targetDir, Guid.NewGuid().ToString("N") & extension)
                        File.WriteAllBytes(targetPath, bytes)

                        result.Success = True
                        result.LocalPath = targetPath
                        result.ContentType = If(response.Content.Headers.ContentType Is Nothing, Nothing, response.Content.Headers.ContentType.MediaType)
                        Return result
                    End Using
                End Using
            End Using
        Catch ex As OperationCanceledException
            result.ErrorMessage = $"превышено время ожидания ({DefaultTimeout.TotalSeconds:0} с)"
        Catch ex As Exception
            result.ErrorMessage = HttpClientFactory.DescribeError(ex)
            result.TlsFailure = IsTlsFailureException(ex)
        Finally
            If client IsNot Nothing Then client.Dispose()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Клиент для картинок. allowInvalidCertificate используется только на повторной попытке:
    ''' во внутреннем контуре у сайта может быть самоподписанный сертификат или внутренний CA.
    ''' </summary>
    Private Shared Function CreateImageClient(uri As Uri, allowInvalidCertificate As Boolean) As HttpClient
        Dim handler = HttpClientFactory.CreateHandler(uri.ToString())
        If allowInvalidCertificate AndAlso handler.ServerCertificateCustomValidationCallback Is Nothing Then
            handler.ServerCertificateCustomValidationCallback = Function(message, cert, chain, errors) True
        End If

        Dim client As New HttpClient(handler)
        client.Timeout = Timeout.InfiniteTimeSpan
        Return client
    End Function

    ''' <summary>Ошибка именно цепочки сертификата, а не сети/HTTP/формата: только её стоит повторять без валидации.</summary>
    Private Shared Function IsTlsFailureException(ex As Exception) As Boolean
        Dim current = ex
        While current IsNot Nothing
            If TypeOf current Is System.Security.Authentication.AuthenticationException Then Return True

            Dim webEx = TryCast(current, System.Net.WebException)
            If webEx IsNot Nothing AndAlso
               (webEx.Status = System.Net.WebExceptionStatus.TrustFailure OrElse
                webEx.Status = System.Net.WebExceptionStatus.SecureChannelFailure) Then
                Return True
            End If

            Dim message = If(current.Message, "")
            If message.IndexOf("SSL", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("TLS", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("certificate", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("сертификат", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("доверие", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If

            current = current.InnerException
        End While

        Return False
    End Function

    Private Shared Async Function ReadWithLimitAsync(response As HttpResponseMessage,
                                                     token As CancellationToken) As Task(Of Byte())
        Using stream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
            Using buffer As New MemoryStream()
                Dim chunk(8191) As Byte
                Dim total As Long = 0
                Do
                    Dim read = Await stream.ReadAsync(chunk, 0, chunk.Length, token).ConfigureAwait(False)
                    If read <= 0 Then Exit Do

                    total += read
                    If total > MaxImageBytes Then Return Nothing

                    buffer.Write(chunk, 0, read)
                Loop

                Return buffer.ToArray()
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Определяет формат по сигнатуре файла, а не по Content-Type.
    ''' Public: хостам нужен корректный суффикс, чтобы записать байты во временный файл
    ''' для вставки в документ.
    ''' </summary>
    Public Shared Function DetectImageExtension(bytes As Byte()) As String
        If bytes Is Nothing OrElse bytes.Length < 12 Then Return Nothing

        If bytes(0) = &HFF AndAlso bytes(1) = &HD8 Then Return ".jpg"
        If bytes(0) = &H89 AndAlso bytes(1) = &H50 AndAlso bytes(2) = &H4E AndAlso bytes(3) = &H47 Then Return ".png"
        If bytes(0) = &H47 AndAlso bytes(1) = &H49 AndAlso bytes(2) = &H46 Then Return ".gif"
        If bytes(0) = &H42 AndAlso bytes(1) = &H4D Then Return ".bmp"
        If (bytes(0) = &H49 AndAlso bytes(1) = &H49 AndAlso bytes(2) = &H2A) OrElse
           (bytes(0) = &H4D AndAlso bytes(1) = &H4D AndAlso bytes(2) = &H2A) Then Return ".tif"

        Return Nothing
    End Function

End Class
