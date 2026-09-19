Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports System.Windows.Forms
Imports PowerPointAi.Handlers

Namespace Extensions

    ''' <summary>
    ''' PptGenerationHandler 集成扩展
    ''' 为 ChatControl 提供幻灯片大纲检测和生成功能
    ''' </summary>
    Public Module PptGenerationHandlerExtension

        ''' <summary>
        ''' 检测 AI 响应是否包含幻灯片大纲
        ''' </summary>
        Public Function DetectSlideOutline(aiResponse As String) As Boolean
            If String.IsNullOrWhiteSpace(aiResponse) Then
                Return False
            End If

            ' 检测 Markdown 标题（# 开头）
            Dim hasHeading As Boolean = aiResponse.Contains("#")

            ' 检测要点标记（- 或 * 或数字编号）
            Dim hasBullet As Boolean = False
            Dim lines As String() = aiResponse.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                Dim trimmed = line.Trim()
                If trimmed.StartsWith("-") OrElse
                   trimmed.StartsWith("*") OrElse
                   System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^\d+\.") Then
                    hasBullet = True
                    Exit For
                End If
            Next

            ' 检测"幻灯片"关键字
            Dim hasSlideKeyword As Boolean = aiResponse.Contains("幻灯片") OrElse
                                              aiResponse.ToLower().Contains("slide")

            ' 至少满足以下条件之一：
            ' 1. 有标题 + 有要点
            ' 2. 有"幻灯片"关键字 + 有标题
            Return (hasHeading AndAlso hasBullet) OrElse (hasSlideKeyword AndAlso hasHeading)
        End Function

        ''' <summary>
        ''' 尝试从 AI 响应中生成幻灯片
        ''' </summary>
        Public Function TryGenerateSlides(aiResponse As String, pptApp As PowerPoint.Application) As Boolean
            Try
                ' 检测是否包含幻灯片大纲
                If Not DetectSlideOutline(aiResponse) Then
                    Return False
                End If

                ' 询问用户确认
                Dim result As DialogResult = MessageBox.Show(
                    "Обнаружен план слайдов. Сгенерировать слайды автоматически?" & vbCrLf & vbCrLf &
                    "Новые слайды будут вставлены после текущего.",
                    "Генерация слайдов",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If result = DialogResult.Yes Then
                    ' 创建 PptGenerationHandler
                    Dim handler As New PptGenerationHandler(pptApp)

                    ' 生成幻灯片
                    Dim count As Integer = handler.GenerateSlidesFromAI(aiResponse, insertAfterCurrent:=True)

                    If count > 0 Then
                        MessageBox.Show(
                            String.Format("Успешно создано слайдов: {0}.", count),
                            "Готово",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

                        System.Diagnostics.Debug.WriteLine("[PptGenerationHandlerExtension] 成功生成 " & count & " 张幻灯片")
                        Return True
                    Else
                        MessageBox.Show(
                            "Не удалось создать слайды. Проверьте формат плана.",
                            "Не удалось создать слайды",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
                        Return False
                    End If
                End If

                Return False

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[PptGenerationHandlerExtension] 错误: " & ex.Message)
                MessageBox.Show(
                    "Ошибка при создании слайдов: " & ex.Message,
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Return False
            End Try
        End Function

    End Module

End Namespace
