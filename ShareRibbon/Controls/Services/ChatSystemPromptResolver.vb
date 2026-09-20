Public Class ChatSystemPromptResolver
    ''' <summary>
    ''' Обязательное правило дисциплины инструментов. Добавляется к любому системному промпту,
    ''' включая пользовательский: без него модель завершает ход обещанием («проверю», «сейчас добавлю»)
    ''' вместо вызова инструмента, и задача остаётся невыполненной.
    ''' </summary>
    Private Shared ReadOnly ToolDisciplineRule As String =
        "Дисциплина инструментов: если для выполнения задачи нужен инструмент, вызывай его в этом же ответе. " &
        "Не заканчивай ход обещанием («проверю», «сейчас добавлю», «затем», «после этого») — либо выполни следующий шаг инструментом, либо дай окончательный ответ. " &
        "Не повторяй один и тот же поиск: как только данных достаточно, переходи к действию."

    Public Function ResolveSystemPrompt(
        requestedSystemPrompt As String,
        appInfo As ApplicationInfo,
        intentResult As IntentResult,
        responseMode As String) As String

        Dim basePrompt = ResolveBasePrompt(requestedSystemPrompt, appInfo, intentResult, responseMode)
        Return basePrompt.TrimEnd() & Environment.NewLine & Environment.NewLine & ToolDisciplineRule
    End Function

    Private Function ResolveBasePrompt(
        requestedSystemPrompt As String,
        appInfo As ApplicationInfo,
        intentResult As IntentResult,
        responseMode As String) As String

        If Not String.IsNullOrWhiteSpace(requestedSystemPrompt) Then
            Return requestedSystemPrompt
        End If

        Dim appType = If(appInfo IsNot Nothing, appInfo.Type.ToString(), "Excel")
        Dim context As New PromptContext With {
            .ApplicationType = appType,
            .IntentResult = intentResult,
            .FunctionMode = responseMode
        }

        Dim resolved = PromptManager.Instance.GetCombinedPrompt(context)
        If Not String.IsNullOrWhiteSpace(resolved) Then Return resolved

        If Not String.IsNullOrWhiteSpace(ConfigSettings.propmtContent) Then
            Return ConfigSettings.propmtContent
        End If

        Return "Ты ассистент Office AI. Отвечай по запросам пользователя кратко, точно и только на русском языке."
    End Function
End Class
