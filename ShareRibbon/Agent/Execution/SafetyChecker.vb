Namespace Agent.Execution

    ''' <summary>
    ''' 代码安全检查器 - 拦截危险操作
    ''' </summary>
    Public Class SafetyChecker

        ''' <summary>
        ''' 危险操作黑名单（绝对禁止执行）
        ''' </summary>
        Private Shared ReadOnly DANGEROUS() As String = {
            "Kill ",
            "Shell(",
            "Shell ",
            "CreateObject(""WScript",
            "CreateObject(""Scripting.FileSystemObject",
            ".DeleteFile",
            ".DeleteFolder",
            "Environ(",
            "RegWrite",
            "RegDelete",
            "Application.Quit",
            "ActiveWorkbook.Close False",
            "Documents.Close",
            "Presentation.Close"
        }

        ''' <summary>
        ''' 需要用户确认的操作（允许但需确认）
        ''' </summary>
        Private Shared ReadOnly NEEDS_CONFIRM() As String = {
            ".Delete",
            ".Clear",
            ".ClearContents",
            "Workbooks.Close",
            ".SaveAs",
            "ActiveDocument.Close",
            "Selection.Delete",
            "Range.Delete",
            "Worksheet.Delete",
            "Slide.Delete"
        }

        ''' <summary>
        ''' 检查代码安全性
        ''' </summary>
        ''' <param name="code">待检查的代码</param>
        ''' <returns>安全检查结果</returns>
        Public Shared Function Check(code As String) As SafetyResult
            If String.IsNullOrWhiteSpace(code) Then
                Return SafetyResult.Safe()
            End If

            ' 1. 检测危险操作（绝对禁止）
            For Each danger In DANGEROUS
                If code.IndexOf(danger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return SafetyResult.Blocked($"Обнаружена опасная операция: {danger}")
                End If
            Next

            ' 2. 检测需要确认的操作
            For Each item In NEEDS_CONFIRM
                If code.IndexOf(item, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return SafetyResult.NeedConfirm($"Будет выполнено: {item}")
                End If
            Next

            ' 3. 通过安全检查
            Return SafetyResult.Safe()
        End Function

    End Class

    ''' <summary>
    ''' 安全检查结果
    ''' </summary>
    Public Class SafetyResult

        ''' <summary>
        ''' 是否安全（可以执行）
        ''' </summary>
        Public Property IsSafe As Boolean

        ''' <summary>
        ''' 是否需要用户确认
        ''' </summary>
        Public Property NeedsConfirm As Boolean

        ''' <summary>
        ''' 原因描述
        ''' </summary>
        Public Property Reason As String

        ''' <summary>
        ''' 检测到的危险操作列表
        ''' </summary>
        Public Property DetectedIssues As List(Of String)

        Public Sub New()
            DetectedIssues = New List(Of String)()
        End Sub

        ''' <summary>
        ''' 创建"安全通过"结果
        ''' </summary>
        Public Shared Function Safe() As SafetyResult
            Return New SafetyResult With {
                .IsSafe = True,
                .NeedsConfirm = False,
                .Reason = "Проверка безопасности пройдена"
            }
        End Function

        ''' <summary>
        ''' 创建"被拦截"结果
        ''' </summary>
        Public Shared Function Blocked(reason As String) As SafetyResult
            Return New SafetyResult With {
                .IsSafe = False,
                .NeedsConfirm = False,
                .Reason = reason
            }
        End Function

        ''' <summary>
        ''' 创建"需要确认"结果
        ''' </summary>
        Public Shared Function NeedConfirm(reason As String) As SafetyResult
            Return New SafetyResult With {
                .IsSafe = True,
                .NeedsConfirm = True,
                .Reason = reason
            }
        End Function

        ''' <summary>
        ''' 转为用户友好的提示文本
        ''' </summary>
        Public Function ToUserMessage() As String
            If Not IsSafe Then
                Return $"⛔ Блокировка безопасности: {Reason}"
            ElseIf NeedsConfirm Then
                Return $"⚠️ Требуется подтверждение: {Reason}"
            Else
                Return $"✅ {Reason}"
            End If
        End Function

    End Class

End Namespace
