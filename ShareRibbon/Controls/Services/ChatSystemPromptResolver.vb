Public Class ChatSystemPromptResolver
    Public Function ResolveSystemPrompt(
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
