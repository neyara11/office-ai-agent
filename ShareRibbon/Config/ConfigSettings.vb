' 存储配置的api大模型和api key
Public Class ConfigSettings
    Private Sub New()
    End Sub

    Public Shared Property platform As String
    Public Shared Property ApiUrl As String
    Public Shared Property ApiKey As String
    Public Shared Property ModelName As String
    Public Shared Property mcpable As Boolean
    Public Shared Property ReasoningMode As String = "default"

    ' 允许使用自签名/不受信任的 TLS 证书（仅针对当前选中的服务商）
    Public Shared Property AllowInsecureTls As Boolean = False

    ' Embedding 模型配置
    Public Shared Property EmbeddingModel As String = ""

    ' FIM (Fill-In-the-Middle) 补全能力
    Public Shared Property fimSupported As Boolean = False
    Public Shared Property fimUrl As String = ""

    ' 提示词相关配置
    Public Shared Property propmtName As String
    Public Shared Property propmtContent As String

    Public Const OfficeAiAppDataFolder As String = "OfficeAiAppData"
End Class
