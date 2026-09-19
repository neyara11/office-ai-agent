Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Agent
Imports ShareRibbon.Agent.OfficeOperations

Namespace OfficeRuntime

    ''' <summary>
    ''' Read-only PowerPoint API catalog. It reflects a bounded initial type set,
    ''' caches metadata in memory, and returns only task-relevant members.
    ''' It never resolves live COM objects and never invokes Office members.
    ''' </summary>
    Public NotInheritable Class PowerPointApiCatalogProvider
        Private Shared ReadOnly CatalogCache As New Lazy(Of CatalogSnapshot)(
            AddressOf BuildCatalog,
            LazyThreadSafetyMode.ExecutionAndPublication)

        Private Shared ReadOnly BlockedMemberNames As New HashSet(Of String)(
            {"quit", "close", "saveas", "savecopyas", "vbproject", "run", "execute", "executemso", "shell"},
            StringComparer.OrdinalIgnoreCase)

        Private Shared ReadOnly DestructiveTokens As String() = {"delete", "remove", "clear"}
        Private Sub New()
        End Sub

        Public Shared Function SearchAsToolResult(params As JObject) As ToolResult
            Const toolId As String = "DiscoverOfficeCapability"
            Try
                Dim request = If(params, New JObject()).ToObject(Of OfficeCapabilitySearchRequest)()
                Dim validation = OfficeOperationValidation.ValidateCapabilitySearch(request)
                If Not validation.IsValid Then
                    Return ToolResult.Failed(toolId,
                                             validation.ToErrorMessage(),
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:="Недопустимые параметры запроса возможностей",
                                             recoverable:=True)
                End If

                Dim result = Search(request)
                Dim observation = New JObject From {
                    {"kind", "office_capability_search"},
                    {"summary", $"Найдено связанных членов PowerPoint API: {result.Members.Count}"},
                    {"changed", False},
                    {"readOnly", True},
                    {"resultCount", result.Members.Count},
                    {"truncated", result.Truncated},
                    {"catalogFingerprint", result.CatalogFingerprint},
                    {"warnings", JArray.FromObject(result.Warnings)}
                }

                If result.Members.Count = 0 Then
                    Return ToolResult.Failed(toolId,
                                             "No matching PowerPoint capability was found",
                                             data:=result,
                                             errorCode:=ExceptionClassifier.CodeCapabilityNotFound,
                                             userMessage:="Связанные с целью возможности объектов PowerPoint не найдены",
                                             recoverable:=True,
                                             observation:=observation)
                End If

                Return ToolResult.Succeed(toolId,
                                          observation("summary").ToString(),
                                          data:=result,
                                          observation:=observation)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            End Try
        End Function

        Public Shared Function Search(request As OfficeCapabilitySearchRequest) As OfficeCapabilitySearchResult
            Dim validation = OfficeOperationValidation.ValidateCapabilitySearch(request)
            If Not validation.IsValid Then Throw New ArgumentException(validation.ToErrorMessage())

            Dim snapshot = CatalogCache.Value
            Dim normalizedQuery = NormalizeSearchText(request.Query)
            Dim normalizedTargetType = NormalizeSearchText(request.TargetType)
            Dim scored As New List(Of ScoredCatalogEntry)()

            For Each entry In snapshot.Entries
                If Not request.IncludeReadOnly AndAlso entry.IsReadOnly Then Continue For
                If Not String.IsNullOrWhiteSpace(normalizedTargetType) AndAlso
                   entry.NormalizedTypeName.IndexOf(normalizedTargetType, StringComparison.Ordinal) < 0 AndAlso
                   entry.NormalizedSearchText.IndexOf(normalizedTargetType, StringComparison.Ordinal) < 0 Then
                    Continue For
                End If

                Dim score = ScoreEntry(entry, normalizedQuery, request.Query, normalizedTargetType)
                If score > 0 Then scored.Add(New ScoredCatalogEntry With {.Entry = entry, .Score = score})
            Next

            Dim ordered = scored.
                OrderByDescending(Function(item) item.Score).
                ThenBy(Function(item) item.Entry.Member.MemberId, StringComparer.Ordinal).
                ToList()
            Dim selected = ordered.Take(request.MaxResults).Select(Function(item) item.Entry.Member).ToList()

            Return New OfficeCapabilitySearchResult With {
                .AppType = "PowerPoint",
                .Query = request.Query,
                .Members = selected,
                .CatalogFingerprint = snapshot.Fingerprint,
                .Truncated = ordered.Count > selected.Count,
                .Warnings = New List(Of String)()
            }
        End Function

        Friend Shared Function TryGetMemberBinding(memberId As String,
                                                   ByRef member As OfficeCapabilityMember,
                                                   ByRef reflectedMember As MemberInfo) As Boolean
            member = Nothing
            reflectedMember = Nothing
            If String.IsNullOrWhiteSpace(memberId) Then Return False

            Dim entry = CatalogCache.Value.Entries.FirstOrDefault(
                Function(item) String.Equals(item.Member.MemberId, memberId, StringComparison.Ordinal))
            If entry Is Nothing Then Return False
            member = entry.Member
            reflectedMember = entry.ReflectedMember
            Return reflectedMember IsNot Nothing
        End Function

        Private Shared Function BuildCatalog() As CatalogSnapshot
            Dim entries As New Dictionary(Of String, CatalogEntry)(StringComparer.Ordinal)
            For Each apiType In GetInitialCatalogTypes()
                If apiType Is Nothing Then Continue For
                AddProperties(apiType, entries)
                AddMethods(apiType, entries)
            Next

            Dim ordered = entries.Values.OrderBy(Function(item) item.Member.MemberId, StringComparer.Ordinal).ToList()
            Dim fingerprintSource = String.Join(vbLf, ordered.Select(Function(item) item.Member.MemberId))
            Return New CatalogSnapshot With {
                .Entries = ordered,
                .Fingerprint = ComputeHash(fingerprintSource)
            }
        End Function

        Private Shared Function GetInitialCatalogTypes() As IEnumerable(Of Type)
            Dim powerPointAssembly = ResolveInteropAssembly("Microsoft.Office.Interop.PowerPoint")
            Dim officeAssembly = ResolveInteropAssembly("office")
            Dim result As New List(Of Type) From {
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint._Presentation", GetType(PowerPoint.Presentation)),
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint.Slides", GetType(PowerPoint.Slides)),
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint._Slide", GetType(PowerPoint.Slide)),
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint.Shapes", GetType(PowerPoint.Shapes)),
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint.Shape", GetType(PowerPoint.Shape)),
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint.TextFrame", GetType(PowerPoint.TextFrame)),
                ResolveCatalogType(powerPointAssembly, "Microsoft.Office.Interop.PowerPoint.TextRange", GetType(PowerPoint.TextRange)),
                ResolveCatalogType(officeAssembly, "Microsoft.Office.Core.TextFrame2", GetType(Microsoft.Office.Core.TextFrame2)),
                ResolveCatalogType(officeAssembly, "Microsoft.Office.Core.TextRange2", GetType(Microsoft.Office.Core.TextRange2)),
                ResolveCatalogType(officeAssembly, "Microsoft.Office.Core.SmartArt", GetType(Microsoft.Office.Core.SmartArt)),
                ResolveCatalogType(officeAssembly, "Microsoft.Office.Core.SmartArtNodes", GetType(Microsoft.Office.Core.SmartArtNodes)),
                ResolveCatalogType(officeAssembly, "Microsoft.Office.Core.SmartArtNode", GetType(Microsoft.Office.Core.SmartArtNode))
            }
            Return result.Where(Function(item) item IsNot Nothing).
                          GroupBy(Function(item) item.FullName, StringComparer.Ordinal).
                          Select(Function(group) group.First()).ToList()
        End Function

        Private Shared Function ResolveCatalogType(interopAssembly As Assembly,
                                                   fullTypeName As String,
                                                   embeddedFallback As Type) As Type
            If interopAssembly IsNot Nothing Then
                Try
                    Dim resolved = interopAssembly.GetType(fullTypeName, throwOnError:=False, ignoreCase:=False)
                    If resolved IsNot Nothing Then Return resolved
                Catch
                End Try
            End If
            Return embeddedFallback
        End Function

        Private Shared Function ResolveInteropAssembly(simpleName As String) As Assembly
            For Each loaded In AppDomain.CurrentDomain.GetAssemblies()
                Try
                    If String.Equals(loaded.GetName().Name, simpleName, StringComparison.OrdinalIgnoreCase) Then Return loaded
                Catch
                End Try
            Next

            Try
                Dim fullName = If(String.Equals(simpleName, "Microsoft.Office.Interop.PowerPoint", StringComparison.OrdinalIgnoreCase),
                                  "Microsoft.Office.Interop.PowerPoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
                                  "office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c")
                Return Assembly.Load(fullName)
            Catch
            End Try

            For Each candidate In GetInteropAssemblyCandidates(simpleName)
                Try
                    If File.Exists(candidate) Then Return Assembly.LoadFrom(candidate)
                Catch ex As Exception
                    Debug.WriteLine($"[PowerPointApiCatalog] Unable to load interop assembly {candidate}: {ex.Message}")
                End Try
            Next
            Return Nothing
        End Function

        Private Shared Function GetInteropAssemblyCandidates(simpleName As String) As IEnumerable(Of String)
            Dim fileName = If(String.Equals(simpleName, "office", StringComparison.OrdinalIgnoreCase),
                              "Office.dll",
                              simpleName & ".dll")
            Dim candidates As New List(Of String)()
            Dim assemblyDir = Path.GetDirectoryName(GetType(PowerPointApiCatalogProvider).Assembly.Location)
            If Not String.IsNullOrWhiteSpace(assemblyDir) Then candidates.Add(Path.Combine(assemblyDir, fileName))

            Dim windowsDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
            If Not String.IsNullOrWhiteSpace(windowsDir) Then
                Dim gacRoots = New String() {
                    Path.Combine(windowsDir, "assembly", "GAC_MSIL", simpleName),
                    Path.Combine(windowsDir, "Microsoft.NET", "assembly", "GAC_MSIL", simpleName)
                }
                For Each gacRoot In gacRoots
                    If Not Directory.Exists(gacRoot) Then Continue For
                    Try
                        candidates.AddRange(Directory.GetFiles(gacRoot, fileName, SearchOption.TopDirectoryOnly))
                        For Each versionDir In Directory.GetDirectories(gacRoot)
                            candidates.AddRange(Directory.GetFiles(versionDir, fileName, SearchOption.TopDirectoryOnly))
                        Next
                    Catch ex As Exception
                        Debug.WriteLine($"[PowerPointApiCatalog] Unable to inspect GAC path {gacRoot}: {ex.Message}")
                    End Try
                Next
            End If
            Return candidates.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        End Function

        Private Shared Sub AddProperties(apiType As Type, entries As Dictionary(Of String, CatalogEntry))
            For Each prop In apiType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                Try
                    Dim parameters = prop.GetIndexParameters().Select(AddressOf BuildParameter).ToList()
                    Dim member = BuildMember(apiType,
                                             prop.Name,
                                             "property",
                                             parameters,
                                             FriendlyTypeName(prop.PropertyType),
                                             prop.CanRead AndAlso Not prop.CanWrite)
                    AddEntry(entries, member, prop, prop.CanRead AndAlso Not prop.CanWrite)
                Catch ex As Exception
                    Debug.WriteLine($"[PowerPointApiCatalog] Skip property {apiType.Name}.{prop.Name}: {ex.Message}")
                End Try
            Next
        End Sub

        Private Shared Sub AddMethods(apiType As Type, entries As Dictionary(Of String, CatalogEntry))
            For Each method In apiType.GetMethods(BindingFlags.Public Or BindingFlags.Instance)
                If method.IsSpecialName Then Continue For
                Try
                    Dim parameters = method.GetParameters().Select(AddressOf BuildParameter).ToList()
                    Dim member = BuildMember(apiType,
                                             method.Name,
                                             "method",
                                             parameters,
                                             FriendlyTypeName(method.ReturnType),
                                             IsReadOnlyMethod(method.Name))
                    AddEntry(entries, member, method, IsReadOnlyMethod(method.Name))
                Catch ex As Exception
                    Debug.WriteLine($"[PowerPointApiCatalog] Skip method {apiType.Name}.{method.Name}: {ex.Message}")
                End Try
            Next
        End Sub

        Private Shared Function BuildMember(apiType As Type,
                                            memberName As String,
                                            memberKind As String,
                                            parameters As List(Of OfficeCapabilityParameter),
                                            returnType As String,
                                            isReadOnly As Boolean) As OfficeCapabilityMember
            Dim logicalTypeName = apiType.Name.TrimStart("_"c)
            Dim declaringTypeName = If(apiType.FullName, logicalTypeName).Replace("._", ".")
            Dim parameterSignature = String.Join(",", parameters.Select(Function(item) item.ParameterType))
            Dim memberId = $"PowerPoint.{logicalTypeName}.{memberKind}.{memberName}({parameterSignature})->{returnType}"
            Dim executable = Not BlockedMemberNames.Contains(memberName)
            Dim unsupportedReason = If(executable, "", "Lifecycle, macro, or external execution member is disabled")

            Return New OfficeCapabilityMember With {
                .MemberId = memberId,
                .DeclaringType = declaringTypeName,
                .MemberName = memberName,
                .MemberKind = memberKind,
                .Parameters = parameters,
                .ReturnType = returnType,
                .RiskLevel = ClassifyRisk(memberName, isReadOnly),
                .Executable = executable,
                .UnsupportedReason = unsupportedReason,
                .Aliases = BuildAliases(logicalTypeName, memberName)
            }
        End Function

        Private Shared Sub AddEntry(entries As Dictionary(Of String, CatalogEntry),
                                    member As OfficeCapabilityMember,
                                    reflectedMember As MemberInfo,
                                    isReadOnly As Boolean)
            If member Is Nothing OrElse entries.ContainsKey(member.MemberId) Then Return
            Dim terms As New List(Of String) From {
                member.DeclaringType,
                member.MemberName,
                member.MemberKind,
                member.ReturnType
            }
            terms.AddRange(member.Aliases)
            terms.AddRange(member.Parameters.Select(Function(item) item.Name & " " & item.ParameterType))
            entries(member.MemberId) = New CatalogEntry With {
                .Member = member,
                .ReflectedMember = reflectedMember,
                .IsReadOnly = isReadOnly,
                .NormalizedTypeName = NormalizeSearchText(member.DeclaringType),
                .NormalizedSearchText = NormalizeSearchText(String.Join(" ", terms)),
                .NormalizedTerms = terms.Select(AddressOf NormalizeSearchText).
                                         Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
                                         Distinct().ToList()
            }
        End Sub

        Private Shared Function BuildParameter(parameter As ParameterInfo) As OfficeCapabilityParameter
            Dim description = If(parameter.IsOptional, "optional", "required")
            If String.Equals(parameter.ParameterType.FullName, "Microsoft.Office.Core.SmartArtLayout", StringComparison.Ordinal) Then
                description &= "; pass a SmartArt layout index, layout ID/name string, or {index|layoutId|id|name: value}"
            End If
            Return New OfficeCapabilityParameter With {
                .Name = parameter.Name,
                .ParameterType = FriendlyTypeName(parameter.ParameterType),
                .Required = Not parameter.IsOptional,
                .DefaultValue = GetSafeDefaultValue(parameter),
                .Description = description
            }
        End Function

        Private Shared Function GetSafeDefaultValue(parameter As ParameterInfo) As Object
            If parameter Is Nothing OrElse Not parameter.IsOptional Then Return Nothing
            Try
                Dim value = parameter.DefaultValue
                If value Is Nothing OrElse value Is DBNull.Value OrElse value Is Missing.Value Then Return Nothing
                If value.GetType().IsEnum Then Return value.ToString()
                If TypeOf value Is String OrElse TypeOf value Is Boolean OrElse TypeOf value Is Decimal OrElse
                   TypeOf value Is Double OrElse TypeOf value Is Single OrElse TypeOf value Is Integer OrElse
                   TypeOf value Is Long OrElse TypeOf value Is Short OrElse TypeOf value Is Byte Then
                    Return value
                End If
            Catch
            End Try
            Return Nothing
        End Function

        Private Shared Function FriendlyTypeName(value As Type) As String
            If value Is Nothing Then Return "Object"
            If value.IsByRef Then value = value.GetElementType()
            If value Is GetType(Void) Then Return "Void"
            If value.IsArray Then Return FriendlyTypeName(value.GetElementType()) & "[]"
            Return If(String.IsNullOrWhiteSpace(value.FullName), value.Name, value.FullName)
        End Function

        Private Shared Function BuildAliases(typeName As String, memberName As String) As List(Of String)
            Dim aliases As New List(Of String)()
            Select Case typeName
                Case "Presentation"
                    aliases.AddRange({"演示文稿", "文稿"})
                Case "Slides"
                    aliases.AddRange({"幻灯片集合", "页面集合"})
                Case "Slide"
                    aliases.AddRange({"幻灯片", "页面", "当前页"})
                Case "Shapes"
                    aliases.AddRange({"形状集合", "图形集合", "当前页对象"})
                Case "Shape"
                    aliases.AddRange({"形状", "图形", "对象"})
                Case "TextFrame", "TextFrame2"
                    aliases.AddRange({"文本框", "文字框"})
                Case "TextRange", "TextRange2"
                    aliases.AddRange({"文本", "文字"})
                Case "SmartArt"
                    aliases.AddRange({"智能图形", "智能图示", "智能艺术图", "三阶段"})
                Case "SmartArtNodes"
                    aliases.AddRange({"SmartArt节点集合", "智能图形节点集合", "节点集合"})
                Case "SmartArtNode"
                    aliases.AddRange({"SmartArt节点", "智能图形节点", "节点"})
            End Select

            Select Case memberName.ToLowerInvariant()
                Case "addsmartart"
                    aliases.AddRange({"SmartArt", "create SmartArt", "insert SmartArt", "创建SmartArt", "插入SmartArt", "智能图形", "添加智能图形", "创建智能图形", "三阶段", "三阶段SmartArt"})
                Case "addnode"
                    aliases.AddRange({"添加节点", "创建节点"})
                Case "allnodes", "nodes"
                    aliases.AddRange({"全部节点", "节点列表"})
                Case "text", "textrange", "textframe", "textframe2"
                    aliases.AddRange({"设置文字", "写入文本", "节点文字"})
                Case "layout"
                    aliases.AddRange({"SmartArt布局", "智能图形布局"})
                Case "delete"
                    aliases.Add("删除")
            End Select
            Return aliases
        End Function

        Private Shared Function ClassifyRisk(memberName As String, isReadOnly As Boolean) As String
            If BlockedMemberNames.Contains(memberName) Then Return "risky"
            Dim normalized = memberName.ToLowerInvariant()
            If DestructiveTokens.Any(Function(token) normalized.Contains(token)) Then Return "risky"
            If isReadOnly Then Return "safe"
            Return "medium"
        End Function

        Private Shared Function IsReadOnlyMethod(memberName As String) As Boolean
            Dim normalized = If(memberName, "").ToLowerInvariant()
            Return normalized.StartsWith("get") OrElse
                   normalized.StartsWith("find") OrElse
                   normalized.StartsWith("item") OrElse
                   normalized.StartsWith("can") OrElse
                   normalized.StartsWith("is")
        End Function

        Private Shared Function ScoreEntry(entry As CatalogEntry,
                                           normalizedQuery As String,
                                           rawQuery As String,
                                           normalizedTargetType As String) As Integer
            If entry Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedQuery) Then Return 0
            Dim score As Integer = 0
            Dim normalizedMemberName = NormalizeSearchText(entry.Member.MemberName)
            If normalizedMemberName = normalizedQuery Then score += 120
            If entry.NormalizedSearchText.IndexOf(normalizedQuery, StringComparison.Ordinal) >= 0 Then score += 60

            For Each term In entry.NormalizedTerms
                If term.Length < 2 Then Continue For
                If normalizedQuery.IndexOf(term, StringComparison.Ordinal) >= 0 Then
                    score += Math.Min(45, 12 + term.Length)
                ElseIf term.IndexOf(normalizedQuery, StringComparison.Ordinal) >= 0 Then
                    score += 15
                End If
            Next

            If Not String.IsNullOrWhiteSpace(normalizedTargetType) AndAlso
               entry.NormalizedTypeName.IndexOf(normalizedTargetType, StringComparison.Ordinal) >= 0 Then score += 40
            If score <= 0 Then Return 0

            Dim lowerQuery = If(rawQuery, "").ToLowerInvariant()
            Dim lowerMember = If(entry.Member.MemberName, "").ToLowerInvariant()
            If ContainsAny(lowerQuery, {"create", "add", "insert", "创建", "添加", "插入"}) AndAlso
               ContainsAny(lowerMember, {"create", "add", "insert"}) Then score += 30
            If ContainsAny(lowerQuery, {"text", "文字", "文本"}) AndAlso
               ContainsAny(lowerMember, {"text", "range"}) Then score += 20
            If ContainsAny(lowerQuery, {"node", "节点"}) AndAlso
               ContainsAny(entry.NormalizedSearchText, {"node", "节点"}) Then score += 25
            If entry.Member.Executable Then score += 2
            Return score
        End Function

        Private Shared Function ContainsAny(value As String, tokens As IEnumerable(Of String)) As Boolean
            Dim source = If(value, "")
            Return tokens.Any(Function(token) source.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
        End Function

        Private Shared Function NormalizeSearchText(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim builder As New StringBuilder()
            For Each ch In value.Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(ch) Then builder.Append(ch)
            Next
            Return builder.ToString()
        End Function

        Private Shared Function ComputeHash(value As String) As String
            Using sha = SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(If(value, "")))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Class CatalogSnapshot
            Public Property Entries As List(Of CatalogEntry)
            Public Property Fingerprint As String
        End Class

        Private Class CatalogEntry
            Public Property Member As OfficeCapabilityMember
            Public Property ReflectedMember As MemberInfo
            Public Property IsReadOnly As Boolean
            Public Property NormalizedTypeName As String
            Public Property NormalizedSearchText As String
            Public Property NormalizedTerms As List(Of String)
        End Class

        Private Class ScoredCatalogEntry
            Public Property Entry As CatalogEntry
            Public Property Score As Integer
        End Class
    End Class

End Namespace
