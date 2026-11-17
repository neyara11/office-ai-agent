' ShareRibbon/Agent/DeckTemplates/TemplateStorage.vb
' Файловое хранилище макетов: %LOCALAPPDATA%\OfficeAiAppData\templates.
' Раскладка: <root>\<setId>\<templateId>\{template.pptx, manifest.json, preview.png}.
' Для одиночного макета, где setId = templateId, допускается <root>\<templateId>\.

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json

Namespace Agent.DeckTemplates

    Public NotInheritable Class TemplateStorage
        Private Shared ReadOnly JsonSettings As New JsonSerializerSettings With {
            .Formatting = Formatting.Indented,
            .NullValueHandling = NullValueHandling.Ignore,
            .DateParseHandling = DateParseHandling.DateTime
        }

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property RootPath As String
            Get
                Return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    ConfigSettings.OfficeAiAppDataFolder,
                    "templates")
            End Get
        End Property

        ''' <summary>Сохраняет/обновляет макет. Источники могут быть пустыми при обновлении манифеста.</summary>
        Public Shared Function Save(manifest As SlideTemplateManifest,
                                    Optional templateSourcePath As String = "",
                                    Optional previewSourcePath As String = "") As SlideTemplateManifest
            If manifest Is Nothing Then Throw New ArgumentNullException(NameOf(manifest))

            manifest.Normalize()
            manifest.Id = TemplateIds.Safe(manifest.Id)
            manifest.SetId = TemplateIds.Safe(manifest.SetId, manifest.Id)
            manifest.SchemaVersion = TemplateSchema.CurrentVersion
            manifest.TemplateFile = TemplateSchema.TemplateFileName

            Dim folder = GetFolderForManifest(manifest)
            Dim staging = Path.Combine(RootPath, ".staging", Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(staging)

            Try
                Dim existingFolder = FindExistingFolder(manifest.Id)
                If Not String.IsNullOrWhiteSpace(templateSourcePath) AndAlso File.Exists(templateSourcePath) Then
                    File.Copy(templateSourcePath, Path.Combine(staging, TemplateSchema.TemplateFileName), True)
                ElseIf Not String.IsNullOrWhiteSpace(existingFolder) Then
                    Dim existingTemplate = Path.Combine(existingFolder, TemplateSchema.TemplateFileName)
                    If File.Exists(existingTemplate) Then
                        File.Copy(existingTemplate, Path.Combine(staging, TemplateSchema.TemplateFileName), True)
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(previewSourcePath) AndAlso File.Exists(previewSourcePath) Then
                    File.Copy(previewSourcePath, Path.Combine(staging, TemplateSchema.PreviewFileName), True)
                    manifest.PreviewPath = TemplateSchema.PreviewFileName
                ElseIf Not String.IsNullOrWhiteSpace(existingFolder) Then
                    Dim existingPreview = Path.Combine(existingFolder, TemplateSchema.PreviewFileName)
                    If File.Exists(existingPreview) Then
                        File.Copy(existingPreview, Path.Combine(staging, TemplateSchema.PreviewFileName), True)
                        manifest.PreviewPath = TemplateSchema.PreviewFileName
                    End If
                End If

                WriteJsonAtomic(Path.Combine(staging, TemplateSchema.ManifestFileName), Serialize(manifest))
                ReplaceFolder(folder, staging)
                TryDeleteDirectory(staging)
            Catch
                TryDeleteDirectory(staging)
                Throw
            End Try

            manifest.StorageFolder = folder
            Return manifest
        End Function

        Public Shared Function SaveManifest(manifest As SlideTemplateManifest) As SlideTemplateManifest
            Return Save(manifest, Nothing, Nothing)
        End Function

        Public Shared Function LoadAll() As List(Of SlideTemplateManifest)
            Dim result As New List(Of SlideTemplateManifest)()
            Dim root = RootPath
            If Not Directory.Exists(root) Then Return result

            Dim byId As New Dictionary(Of String, SlideTemplateManifest)(StringComparer.OrdinalIgnoreCase)
            For Each manifestFile In Directory.GetFiles(root, TemplateSchema.ManifestFileName, SearchOption.AllDirectories)
                If IsInternalPath(manifestFile) Then Continue For
                Try
                    Dim manifest = Deserialize(File.ReadAllText(manifestFile, Encoding.UTF8))
                    If manifest Is Nothing OrElse Not TemplateIds.IsSafe(manifest.Id) Then Continue For
                    If manifest.SchemaVersion > TemplateSchema.CurrentVersion Then Continue For
                    manifest.Normalize()
                    manifest.StorageFolder = Path.GetDirectoryName(manifestFile)

                    Dim existing As SlideTemplateManifest = Nothing
                    If byId.TryGetValue(manifest.Id, existing) Then
                        If manifest.CreatedAt <= existing.CreatedAt Then Continue For
                    End If
                    byId(manifest.Id) = manifest
                Catch ex As Exception
                    Debug.WriteLine($"[TemplateStorage] Пропущен манифест {manifestFile}: {ex.Message}")
                End Try
            Next

            result.AddRange(byId.Values)
            Return result.OrderByDescending(Function(item) item.CreatedAt).ToList()
        End Function

        Public Shared Function Load(id As String) As SlideTemplateManifest
            If Not TemplateIds.IsSafe(id) Then Return Nothing
            Dim direct = Path.Combine(RootPath, id, TemplateSchema.ManifestFileName)
            If File.Exists(direct) Then
                Try
                    Dim manifest = Deserialize(File.ReadAllText(direct, Encoding.UTF8))
                    If manifest IsNot Nothing Then
                        manifest.Normalize()
                        manifest.StorageFolder = Path.GetDirectoryName(direct)
                        Return manifest
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[TemplateStorage] Не удалось прочитать {direct}: {ex.Message}")
                End Try
            End If
            Return LoadAll().FirstOrDefault(Function(item) String.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))
        End Function

        Public Shared Function Delete(id As String) As Boolean
            Dim manifest = Load(id)
            If manifest Is Nothing OrElse String.IsNullOrWhiteSpace(manifest.StorageFolder) Then Return False
            If Not IsInsideRoot(manifest.StorageFolder) Then Return False
            If Not Directory.Exists(manifest.StorageFolder) Then Return False
            Dim parent = Path.GetDirectoryName(manifest.StorageFolder)
            Directory.Delete(manifest.StorageFolder, True)
            Try
                If Not String.IsNullOrWhiteSpace(parent) AndAlso
                   Not String.Equals(Path.GetFullPath(parent), Path.GetFullPath(RootPath), StringComparison.OrdinalIgnoreCase) AndAlso
                   Directory.Exists(parent) AndAlso
                   Not Directory.EnumerateFileSystemEntries(parent).Any() Then
                    Directory.Delete(parent)
                End If
            Catch
            End Try
            Return True
        End Function

        ''' <summary>Экспортирует папку макета в zip для переноса между машинами.</summary>
        Public Shared Function Export(id As String, destinationZipPath As String) As String
            If String.IsNullOrWhiteSpace(destinationZipPath) Then Throw New ArgumentException("Не указан путь экспорта", NameOf(destinationZipPath))
            Dim manifest = Load(id)
            If manifest Is Nothing Then Throw New InvalidOperationException($"Макет не найден: {id}")
            If String.IsNullOrWhiteSpace(manifest.StorageFolder) OrElse Not Directory.Exists(manifest.StorageFolder) Then
                Throw New InvalidOperationException($"Папка макета недоступна: {id}")
            End If

            Dim parent = Path.GetDirectoryName(destinationZipPath)
            If Not String.IsNullOrWhiteSpace(parent) Then Directory.CreateDirectory(parent)
            If File.Exists(destinationZipPath) Then File.Delete(destinationZipPath)
            ZipFile.CreateFromDirectory(manifest.StorageFolder, destinationZipPath, CompressionLevel.Optimal, False)
            Return destinationZipPath
        End Function

        ''' <summary>Импортирует zip, созданный Export. Возвращает сохранённый манифест.</summary>
        Public Shared Function Import(zipPath As String) As SlideTemplateManifest
            If String.IsNullOrWhiteSpace(zipPath) OrElse Not File.Exists(zipPath) Then
                Throw New FileNotFoundException("Архив макета не найден", zipPath)
            End If

            Dim staging = Path.Combine(RootPath, ".import", Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(staging)
            Try
                Using archive = ZipFile.OpenRead(zipPath)
                    For Each entry In archive.Entries
                        If String.IsNullOrWhiteSpace(entry.Name) Then Continue For
                        Dim name = entry.Name
                        If name <> TemplateSchema.ManifestFileName AndAlso
                           name <> TemplateSchema.TemplateFileName AndAlso
                           name <> TemplateSchema.PreviewFileName Then
                            Continue For
                        End If
                        entry.ExtractToFile(Path.Combine(staging, name), True)
                    Next
                End Using

                Dim manifestPath = Path.Combine(staging, TemplateSchema.ManifestFileName)
                If Not File.Exists(manifestPath) Then
                    Throw New InvalidDataException("В архиве нет manifest.json")
                End If

                Dim manifest = Deserialize(File.ReadAllText(manifestPath, Encoding.UTF8))
                If manifest Is Nothing Then Throw New InvalidDataException("manifest.json пуст или повреждён")
                manifest.Normalize()
                manifest.Id = TemplateIds.Safe(manifest.Id)
                manifest.SetId = TemplateIds.Safe(manifest.SetId, manifest.Id)
                manifest.SchemaVersion = TemplateSchema.CurrentVersion
                manifest.TemplateFile = TemplateSchema.TemplateFileName
                If File.Exists(Path.Combine(staging, TemplateSchema.PreviewFileName)) Then
                    manifest.PreviewPath = TemplateSchema.PreviewFileName
                End If

                WriteJsonAtomic(manifestPath, Serialize(manifest))
                Dim folder = GetFolderForManifest(manifest)
                ReplaceFolder(folder, staging)
                TryDeleteDirectory(staging)
                manifest.StorageFolder = folder
                Return manifest
            Catch
                TryDeleteDirectory(staging)
                Throw
            End Try
        End Function

        Public Shared Function GetFolderForManifest(manifest As SlideTemplateManifest) As String
            If manifest Is Nothing Then Throw New ArgumentNullException(NameOf(manifest))
            If String.Equals(manifest.SetId, manifest.Id, StringComparison.OrdinalIgnoreCase) Then
                Return Path.Combine(RootPath, manifest.Id)
            End If
            Return Path.Combine(RootPath, manifest.SetId, manifest.Id)
        End Function

        Public Shared Function Serialize(manifest As SlideTemplateManifest) As String
            Return JsonConvert.SerializeObject(manifest, JsonSettings)
        End Function

        Public Shared Function Deserialize(json As String) As SlideTemplateManifest
            If String.IsNullOrWhiteSpace(json) Then Return Nothing
            Return JsonConvert.DeserializeObject(Of SlideTemplateManifest)(json, JsonSettings)
        End Function

        Private Shared Function FindExistingFolder(id As String) As String
            If Not TemplateIds.IsSafe(id) Then Return ""
            Dim direct = Path.Combine(RootPath, id)
            If File.Exists(Path.Combine(direct, TemplateSchema.ManifestFileName)) Then Return direct
            Return LoadAll().FirstOrDefault(Function(item) String.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))?.StorageFolder
        End Function

        Private Shared Sub WriteJsonAtomic(targetPath As String, json As String)
            Dim parent = Path.GetDirectoryName(targetPath)
            If Not String.IsNullOrWhiteSpace(parent) Then Directory.CreateDirectory(parent)
            Dim temp = targetPath & ".tmp"
            File.WriteAllText(temp, json, New UTF8Encoding(False))
            If File.Exists(targetPath) Then
                File.Replace(temp, targetPath, Nothing)
            Else
                File.Move(temp, targetPath)
            End If
        End Sub

        Private Shared Sub ReplaceFolder(target As String, staging As String)
            Dim parent = Path.GetDirectoryName(target)
            If Not String.IsNullOrWhiteSpace(parent) Then Directory.CreateDirectory(parent)
            If Not Directory.Exists(target) Then
                Directory.Move(staging, target)
                Return
            End If

            Dim trash = Path.Combine(RootPath, ".trash", Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(Path.GetDirectoryName(trash))
            Directory.Move(target, trash)
            Try
                Directory.Move(staging, target)
            Catch
                Try
                    If Not Directory.Exists(target) Then Directory.Move(trash, target)
                Catch
                End Try
                Throw
            End Try
            TryDeleteDirectory(trash)
        End Sub

        Private Shared Function IsInsideRoot(candidatePath As String) As Boolean
            If String.IsNullOrWhiteSpace(candidatePath) Then Return False
            Dim full = Path.GetFullPath(candidatePath)
            Dim root = Path.GetFullPath(RootPath)
            If Not root.EndsWith(Path.DirectorySeparatorChar) Then root &= Path.DirectorySeparatorChar
            Return full.StartsWith(root, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsInternalPath(candidatePath As String) As Boolean
            If String.IsNullOrWhiteSpace(candidatePath) Then Return False
            Dim root = RootPath
            If Not candidatePath.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then Return False
            Dim relative = candidatePath.Substring(root.Length)
            Return relative.Split(Path.DirectorySeparatorChar).
                Any(Function(segment) segment.StartsWith(".", StringComparison.Ordinal))
        End Function

        Private Shared Sub TryDeleteDirectory(folderPath As String)
            Try
                If Not String.IsNullOrWhiteSpace(folderPath) AndAlso Directory.Exists(folderPath) Then
                    Directory.Delete(folderPath, True)
                End If
                ' Убираем пустые служебные папки (.staging/.import/.trash), но не сам каталог макетов.
                Dim parent = Path.GetDirectoryName(folderPath)
                If Not String.IsNullOrWhiteSpace(parent) AndAlso
                   Path.GetFileName(parent).StartsWith(".", StringComparison.Ordinal) AndAlso
                   Directory.Exists(parent) AndAlso
                   Not Directory.EnumerateFileSystemEntries(parent).Any() Then
                    Directory.Delete(parent)
                End If
            Catch ex As Exception
                Debug.WriteLine($"[TemplateStorage] Не удалось удалить {folderPath}: {ex.Message}")
            End Try
        End Sub

    End Class

End Namespace
