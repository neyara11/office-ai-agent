Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports ExcelAi.Handlers

Namespace Extensions

    ''' <summary>
    ''' FormulaHandler 集成扩展
    ''' 为 ChatControl 提供公式检测和应用功能
    ''' </summary>
    Public Module FormulaHandlerExtension

        ''' <summary>
        ''' 检测 AI 响应是否包含 Excel 公式
        ''' </summary>
        Public Function DetectFormula(aiResponse As String) As Boolean
            If String.IsNullOrWhiteSpace(aiResponse) Then
                Return False
            End If

            ' 检测公式特征
            Dim hasEqualSign As Boolean = aiResponse.Contains("=")

            ' 检测常见 Excel 函数
            Dim commonFunctions As String() = {
                "SUM", "AVERAGE", "COUNT", "MAX", "MIN",
                "IF", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
                "SUMIF", "COUNTIF", "AVERAGEIF",
                "LEFT", "RIGHT", "MID", "LEN", "CONCAT",
                "公式", "函数"
            }

            Dim hasFunction As Boolean = False
            For Each func In commonFunctions
                If aiResponse.ToUpper().Contains(func) Then
                    hasFunction = True
                    Exit For
                End If
            Next

            Return hasEqualSign AndAlso hasFunction
        End Function

        ''' <summary>
        ''' 尝试从 AI 响应中应用公式
        ''' </summary>
        ''' <returns>是否成功应用公式</returns>
        Public Function TryApplyFormula(aiResponse As String, excelApp As Excel.Application) As Boolean
            Try
                ' 检测是否包含公式
                If Not DetectFormula(aiResponse) Then
                    Return False
                End If

                ' 获取目标单元格
                Dim targetRange As Excel.Range = Nothing
                Try
                    targetRange = CType(excelApp.ActiveCell, Excel.Range)
                    If targetRange Is Nothing Then
                        Return False
                    End If
                Catch
                    Return False
                End Try

                ' 创建 FormulaHandler
                Dim handler As New FormulaHandler()

                ' 尝试应用公式
                Dim success As Boolean = handler.ApplyFormulaFromAI(aiResponse, targetRange)

                If success Then
                    ' 成功应用
                    MessageBox.Show(
                        "Формула успешно применена к ячейке " & targetRange.Address(False, False),
                        "Формула применена",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    System.Diagnostics.Debug.WriteLine("[FormulaHandler集成] 公式已应用: " & targetRange.Address)
                    Return True
                Else
                    ' 应用失败，尝试提取公式让用户手动复制
                    Dim formula As String = ExtractFormulaFromResponse(aiResponse)
                    If Not String.IsNullOrEmpty(formula) Then
                        Dim result As DialogResult = MessageBox.Show(
                            "Обнаружена формула, но автоматическое применение не удалось." & vbCrLf & vbCrLf &
                            "Извлечённая формула: " & formula & vbCrLf & vbCrLf &
                            "Скопировать её в буфер обмена?",
                            "Извлечение формулы",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question)

                        If result = DialogResult.Yes Then
                            Clipboard.SetText(formula)
                            MessageBox.Show("Формула скопирована в буфер обмена, вставьте её в ячейку", "Подсказка", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If
                    Return False
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[FormulaHandler集成] 错误: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 从 AI 响应中提取公式
        ''' </summary>
        Private Function ExtractFormulaFromResponse(aiResponse As String) As String
            Try
                ' 匹配 =开头的公式
                Dim formulaPattern As String = "=[\w\s\(\)\+\-\*/,:$]+(?:\([^\)]*\))*"
                Dim match As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(aiResponse, formulaPattern)

                If match.Success Then
                    Return match.Value.Trim()
                End If

                ' 如果没有匹配到，尝试查找代码块中的公式
                Dim codeBlockPattern As String = "```[\s\S]*?```"
                Dim codeMatch As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(aiResponse, codeBlockPattern)

                If codeMatch.Success Then
                    Dim codeBlock As String = codeMatch.Value
                    Dim innerMatch As System.Text.RegularExpressions.Match =
                        System.Text.RegularExpressions.Regex.Match(codeBlock, formulaPattern)

                    If innerMatch.Success Then
                        Return innerMatch.Value.Trim()
                    End If
                End If

                Return String.Empty

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[FormulaHandler集成] 提取公式失败: " & ex.Message)
                Return String.Empty
            End Try
        End Function

    End Module

End Namespace
