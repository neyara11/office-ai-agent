Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 自动补全服务：处理输入补全请求、FIM/Chat 两种补全模式、补全历史记录
''' </summary>
Public Class AutocompleteService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _getContextSnapshot As Func(Of JObject)
    Private ReadOnly _getAppType As Func(Of String)

    Public Sub New(
        executeScript As Func(Of String, Task),
        getContextSnapshot As Func(Of JObject),
        getAppType As Func(Of String))

        _executeScript = executeScript
        _getContextSnapshot = getContextSnapshot
        _getAppType = getAppType
    End Sub

    ''' <summary>
    ''' 处理自动补全请求
    ''' </summary>
    Public Async Function HandleRequestCompletion(jsonDoc As JObject) As Task
        Try
            If Not ChatSettings.EnableAutocomplete Then Return

            Dim inputText As String = If(jsonDoc("input")?.ToString(), "")
            Dim timestamp As Long = If(jsonDoc("timestamp")?.Value(Of Long)(), 0)

            If String.IsNullOrWhiteSpace(inputText) OrElse inputText.Length < 2 Then Return

            Dim contextSnapshot = _getContextSnapshot()
            Dim completions = Await RequestCompletionsFromLLM(inputText, contextSnapshot)

            Dim resultJson As New JObject()
            resultJson("completions") = JArray.FromObject(completions)
            resultJson("timestamp") = timestamp

            Await _executeScript($"showCompletions({resultJson.ToString(Newtonsoft.Json.Formatting.None)});")

        Catch ex As Exception
            Debug.WriteLine($"HandleRequestCompletion 出错: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理补全采纳记录
    ''' </summary>
    Public Sub HandleAcceptCompletion(jsonDoc As JObject)
        Try
            Dim inputText As String = If(jsonDoc("input")?.ToString(), "")
            Dim completion As String = If(jsonDoc("completion")?.ToString(), "")
            Dim context As String = If(jsonDoc("context")?.ToString(), "")
            RecordCompletionHistory(inputText, completion, context)
        Catch ex As Exception
            Debug.WriteLine($"HandleAcceptCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 为意图识别/规划阶段丰富上下文：内容区引用摘要 + RAG 相关记忆
    ''' </summary>
    Public Sub EnrichContextForIntent(snapshot As JObject,
                                      question As String,
                                      filePaths As List(Of String),
                                      selectedContents As List(Of SendMessageReferenceContentItem))
        If snapshot Is Nothing Then Return
        Dim refParts As New List(Of String)()
        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            refParts.Add($"用户引用了 {filePaths.Count} 个文件")
        End If
        If selectedContents IsNot Nothing AndAlso selectedContents.Count > 0 Then
            refParts.Add($"{selectedContents.Count} 段选中内容")
            For i = 0 To Math.Min(selectedContents.Count - 1, 4)
                Dim item = selectedContents(i)
                Dim desc = If(String.IsNullOrEmpty(item.sheetName), item.address, $"{item.sheetName}: {item.address}")
                If desc.Length > 60 Then desc = desc.Substring(0, 57) & "..."
                refParts.Add($"  - {desc}")
            Next
        End If
        If refParts.Count > 0 Then
            snapshot("referenceSummary") = String.Join("；" & vbCrLf, refParts)
        End If
        If Not String.IsNullOrWhiteSpace(question) Then
            Try
                Dim memories = MemoryService.GetRelevantMemories(question, 2, Nothing, Nothing, _getAppType())
                If memories IsNot Nothing AndAlso memories.Count > 0 Then
                    Dim lines As New List(Of String)()
                    For Each m In memories
                        Dim c = If(m.Content, "").Trim()
                        If c.Length > 200 Then c = c.Substring(0, 197) & "..."
                        If Not String.IsNullOrEmpty(c) Then lines.Add(c)
                    Next
                    If lines.Count > 0 Then snapshot("ragSnippets") = String.Join(vbCrLf & "---" & vbCrLf, lines)
                End If
            Catch ex As Exception
                Debug.WriteLine($"EnrichContextForIntent RAG: {ex.Message}")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' 调用大模型获取补全建议
    ''' </summary>
    Private Async Function RequestCompletionsFromLLM(inputText As String, contextSnapshot As JObject) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        Try
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then Return completions

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)

            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key

            Dim useFimMode = selectedModel.fimSupported AndAlso Not String.IsNullOrEmpty(selectedModel.fimUrl)

            If useFimMode Then
                completions = Await RequestCompletionsWithFIM(inputText, contextSnapshot, selectedModel, apiKey)
            Else
                completions = Await RequestCompletionsWithChat(inputText, contextSnapshot, cfg, selectedModel, apiKey)
            End If
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsFromLLM 出错: {ex.Message}")
        End Try
        Return completions
    End Function

    ''' <summary>
    ''' 使用 FIM (Fill-In-the-Middle) API 获取补全
    ''' </summary>
    Private Async Function RequestCompletionsWithFIM(inputText As String, contextSnapshot As JObject,
                                                      model As ConfigManager.ConfigItemModel, apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        Try
            Dim fimUrl = model.fimUrl
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("prompt") = inputText
            requestObj("suffix") = ""
            requestObj("max_tokens") = 50
            requestObj("temperature") = 0.3
            requestObj("stream") = False
            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            Dim client = HttpClientPool.GetClient(fimUrl)
            Using request As New HttpRequestMessage(HttpMethod.Post, fimUrl)
                request.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using timeoutCts As New System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(10))
                    Using response = Await client.SendAsync(request, timeoutCts.Token)
                        response.EnsureSuccessStatusCode()
                        Dim responseBody = Await response.Content.ReadAsStringAsync()
                Dim jObj = JObject.Parse(responseBody)
                Dim text = jObj("choices")?(0)?("text")?.ToString()
                If Not String.IsNullOrWhiteSpace(text) Then
                    text = text.Trim().Split({vbCr, vbLf, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)(0)
                    If text.Length <= 50 Then completions.Add(text)
                End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsWithFIM 出错: {ex.Message}")
        End Try
        Return completions
    End Function

    ''' <summary>
    ''' 使用 Chat Completion API 获取补全
    ''' </summary>
    Private Async Function RequestCompletionsWithChat(inputText As String, contextSnapshot As JObject,
                                                       cfg As ConfigManager.ConfigItem, model As ConfigManager.ConfigItemModel,
                                                       apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        Try
            Dim apiUrl = cfg.url
            Dim modelName = model.modelName
            Dim appType = If(contextSnapshot("appType")?.ToString(), "Office")
            Dim selectionText = If(contextSnapshot("selection")?.ToString(), "")
            Dim systemPrompt = GetCompletionSystemPrompt(appType)

            Dim userContent As New StringBuilder()
            userContent.AppendLine($"Текущее приложение: {appType}")
            userContent.AppendLine($"Пользователь ввёл: ""{inputText}""")
            If Not String.IsNullOrWhiteSpace(selectionText) Then
                userContent.AppendLine($"Выделенный текст: ""{selectionText.Substring(0, Math.Min(200, selectionText.Length))}""")
            End If
            If contextSnapshot("sheetName") IsNot Nothing Then
                userContent.AppendLine($"Текущий лист: {contextSnapshot("sheetName")}")
            End If
            If contextSnapshot("slideIndex") IsNot Nothing Then
                userContent.AppendLine($"Текущий слайд: {contextSnapshot("slideIndex")}")
            End If
            userContent.AppendLine()
            userContent.AppendLine("Дай варианты автодополнения (в формате JSON).")

            Dim messages As New JArray()
            messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject() From {{"role", "user"}, {"content", userContent.ToString()}})

            Dim gatewayResponse = Await AiGateway.SendChatAsync(New AiRequestOptions With {
                .ApiUrl = apiUrl,
                .ApiKey = apiKey,
                .ModelName = modelName,
                .Platform = cfg.platform,
                .ReasoningMode = ReasoningRequestHelper.ReasoningDisabled,
                .Messages = messages,
                .Temperature = 0.3R,
                .TimeoutSeconds = 10
            })

            If gatewayResponse Is Nothing OrElse Not gatewayResponse.Success Then
                Debug.WriteLine($"RequestCompletionsWithChat API失败: {If(gatewayResponse Is Nothing, "empty response", gatewayResponse.ErrorMessage)}")
                Return completions
            End If

            AddCompletionsFromMessage(gatewayResponse.Content, completions)
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsWithChat 出错: {ex.Message}")
        End Try
        Return completions
    End Function

    Private Sub AddCompletionsFromMessage(msg As String, completions As List(Of String))
        If String.IsNullOrEmpty(msg) OrElse completions Is Nothing Then Return

        Try
            Dim cleanedMsg = msg.Trim()
            If cleanedMsg.StartsWith("```") Then
                Dim firstNewLine = cleanedMsg.IndexOf(vbLf)
                If firstNewLine > 0 Then cleanedMsg = cleanedMsg.Substring(firstNewLine + 1)
            End If
            If cleanedMsg.EndsWith("```") Then
                cleanedMsg = cleanedMsg.Substring(0, cleanedMsg.Length - 3)
            End If
            cleanedMsg = cleanedMsg.Trim()
            Dim jsonStart = cleanedMsg.IndexOf("{")
            Dim jsonEnd = cleanedMsg.LastIndexOf("}")
            If jsonStart >= 0 AndAlso jsonEnd > jsonStart Then
                cleanedMsg = cleanedMsg.Substring(jsonStart, jsonEnd - jsonStart + 1)
            End If
            Dim resultObj = JObject.Parse(cleanedMsg)
            Dim completionsArray = resultObj("completions")
            If completionsArray IsNot Nothing Then
                For Each item In completionsArray
                    Dim c = item.ToString().Trim()
                    If Not String.IsNullOrWhiteSpace(c) Then completions.Add(c)
                Next
            End If
        Catch parseEx As Exception
            Debug.WriteLine($"解析补全JSON失败: {parseEx.Message}")
            If Not String.IsNullOrWhiteSpace(msg) AndAlso msg.Length < 50 Then
                completions.Add(msg.Trim())
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 记录补全历史
    ''' </summary>
    Private Sub RecordCompletionHistory(inputText As String, completion As String, context As String)
        Try
            Dim historyPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                ConfigSettings.OfficeAiAppDataFolder,
                "autocomplete_history.json")

            Dim history As JObject
            If File.Exists(historyPath) Then
                Dim json = File.ReadAllText(historyPath)
                history = JObject.Parse(json)
            Else
                history = New JObject()
                history("version") = 1
                history("history") = New JArray()
            End If

            Dim historyArray = CType(history("history"), JArray)
            Dim existingItem = historyArray.FirstOrDefault(Function(item)
                                                               Return item("input")?.ToString() = inputText AndAlso
                                                                      item("completion")?.ToString() = completion
                                                           End Function)
            If existingItem IsNot Nothing Then
                existingItem("count") = existingItem("count").Value(Of Integer)() + 1
                existingItem("lastUsed") = DateTime.UtcNow.ToString("o")
            Else
                Dim newItem As New JObject()
                newItem("input") = inputText
                newItem("completion") = completion
                newItem("context") = context
                newItem("count") = 1
                newItem("lastUsed") = DateTime.UtcNow.ToString("o")
                historyArray.Add(newItem)
                While historyArray.Count > 100
                    historyArray.RemoveAt(0)
                End While
            End If

            Dim dir = Path.GetDirectoryName(historyPath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            File.WriteAllText(historyPath, history.ToString(Newtonsoft.Json.Formatting.Indented))
        Catch ex As Exception
            Debug.WriteLine($"RecordCompletionHistory 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 根据 Office 应用类型获取场景化的补全系统提示词
    ''' </summary>
    Private Function GetCompletionSystemPrompt(appType As String) As String
        Dim baseRules = "
Правила:
1. Возвращай только оставшуюся часть дополнения, не повторяй уже введённый пользователем текст
2. Формат JSON: {""completions"": [""дополнение 1"", ""дополнение 2"", ""дополнение 3""]}
3. Не более 3 вариантов
4. Дополнения должны быть краткими, обычно не более 6–8 слов
5. Отвечай только на русском языке"

        Select Case appType.ToLower()
            Case "excel"
                Return $"Ты движок автодополнения для AI-помощника Excel. По текущему вводу пользователя и контексту Excel предскажи, какое действие он хочет выполнить.

Примеры типичных дополнений Excel:
- ""помоги"" → ""посчитать сумму столбца"", ""отфильтровать дубликаты"", ""создать сводную таблицу""
- ""сделай"" → ""преобразовать выделенное в таблицу"", ""удалить дубликаты"", ""объединить столбцы A и B""
- ""посчитай"" → ""количество по категориям"", ""среднюю выручку"", ""темп роста по месяцам""
- ""формула"" → ""разница двух столбцов"", ""найти соответствие"", ""сумма по условию""
- ""формат"" → ""денежный формат"", ""условное форматирование"", ""ширина столбцов""
- ""диаграмма"" → ""построить столбчатую"", ""линия тренда"", ""подписи данных""
{baseRules}"

            Case "word"
                Return $"Ты движок автодополнения для AI-помощника Word. По текущему вводу пользователя и контексту Word предскажи, какое действие он хочет выполнить.

Примеры типичных дополнений Word:
- ""помоги"" → ""отредактировать абзац"", ""перевести выделенное"", ""составить план статьи""
- ""сделай"" → ""официальный тон"", ""сделать заголовок первого уровня"", ""настроить отступы""
- ""сократи"" → ""ключевые тезисы статьи"", ""протокол встречи"", ""главная мысль""
- ""допиши"" → ""этот абзац"", ""подробнее раскрыть мысль"", ""добавить пример""
- ""формат"" → ""единые интервалы"", ""добавить колонтитулы"", ""настроить оглавление""
- ""проверь"" → ""грамматические ошибки"", ""опечатки"", ""пунктуацию""
{baseRules}"

            Case "powerpoint"
                Return $"Ты движок автодополнения для AI-помощника PowerPoint. По текущему вводу пользователя и контексту PPT предскажи, какое действие он хочет выполнить.

Примеры типичных дополнений PPT:
- ""помоги"" → ""оформить этот слайд"", ""подготовить текст выступления"", ""добавить переходы""
- ""сделай"" → ""преобразовать текст в SmartArt"", ""обрезать картинку кругом"", ""градиентный фон""
- ""создай"" → ""презентацию по проекту"", ""слайд с продуктом"", ""слайд о команде""
- ""добавь"" → ""диаграмму с данными"", ""таймлайн"", ""блок-схему""
- ""оформи"" → ""единый шрифт"", ""цветовую схему"", ""макет из образца""
- ""сократи"" → ""ключевые тезисы"", ""главные цифры"", ""слайд с выводами""
{baseRules}"

            Case Else
                Return $"Ты движок автодополнения для AI-помощника Office. По текущему вводу пользователя и контексту Office предскажи, что он хочет ввести.
{baseRules}
5. Учитывай контекст Office (выделенный текст, тип документа)"
        End Select
    End Function

End Class
