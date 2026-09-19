Imports System.Collections.Generic
Imports System.Text
Imports Newtonsoft.Json

Namespace Agent.Context

    ''' <summary>
    ''' Planner、Observe 与 UI trace 共用的最小可序列化上下文合同。
    ''' v1 只包装现有三端 Reader；后续可在不改变 Harness API 的前提下扩充宿主字段。
    ''' </summary>
    Public Class ContextPack
        Public Property SchemaVersion As String = "1.0"
        Public Property CapturedAt As DateTime = DateTime.Now
        Public Property AppType As String = ""
        Public Property Scope As String = "active"
        Public Property Document As New ContextDocument()
        Public Property Selection As New ContextSelection()
        Public Property [Structure] As New ContextStructure()
        Public Property Budget As New ContextBudget()
        Public Property ReaderErrors As New List(Of ContextReaderError)()

        Public Shared Function FromOfficeContext(context As OfficeContext,
                                                 Optional hostContextText As String = "",
                                                 Optional maxChars As Integer = 12000) As ContextPack
            Dim pack As New ContextPack()
            pack.Budget.MaxChars = Math.Max(2000, maxChars)

            If context Is Nothing Then
                pack.ReaderErrors.Add(New ContextReaderError With {
                    .Reader = "OfficeContext",
                    .ErrorCode = "CONTEXT_MISSING",
                    .Message = "Контекст хоста недоступен"
                })
                pack.Document.Preview = Truncate(hostContextText, pack.Budget.MaxChars, pack.Budget.Truncated)
                pack.Budget.UsedChars = pack.Document.Preview.Length
                Return pack
            End If

            pack.AppType = If(context.AppType, "")
            If context.Selection IsNot Nothing Then
                pack.Selection.Ref = BuildSelectionRef(pack.AppType, context.Selection.Address)
                pack.Selection.Address = If(context.Selection.Address, "")
                pack.Selection.ItemCount = context.Selection.ItemCount
                pack.Selection.DataType = If(context.Selection.DataType, "")
            End If

            If context.DocStructure IsNot Nothing Then
                pack.[Structure].HasHeadings = context.DocStructure.HasHeadings
                pack.[Structure].HeadingCount = context.DocStructure.HeadingCount
            End If

            ' selection_first：先保留选区，再给结构和宿主正文分配剩余预算。
            Dim remaining = pack.Budget.MaxChars
            pack.Selection.Preview = Truncate(If(context.Selection?.Preview, ""), Math.Min(remaining, 6000), pack.Budget.Truncated)
            remaining -= pack.Selection.Preview.Length
            pack.[Structure].Summary = Truncate(If(context.DocStructure?.Summary, ""), Math.Min(Math.Max(0, remaining), 3000), pack.Budget.Truncated)
            remaining -= pack.[Structure].Summary.Length
            pack.Document.Preview = Truncate(If(hostContextText, ""), Math.Max(0, remaining), pack.Budget.Truncated)
            pack.Budget.UsedChars = pack.Selection.Preview.Length + pack.[Structure].Summary.Length + pack.Document.Preview.Length
            Return pack
        End Function

        Public Function ToPromptText() As String
            Dim sb As New StringBuilder()
            sb.AppendLine("## ContextPack")
            sb.AppendLine($"schemaVersion: {SchemaVersion}")
            sb.AppendLine($"appType: {AppType}")
            sb.AppendLine($"scope: {Scope}")
            If Not String.IsNullOrWhiteSpace(Selection.Ref) Then sb.AppendLine($"selectionRef: {Selection.Ref}")
            If Not String.IsNullOrWhiteSpace(Selection.DataType) Then sb.AppendLine($"selectionType: {Selection.DataType}")
            If Selection.ItemCount > 0 Then sb.AppendLine($"selectionItems: {Selection.ItemCount}")
            If Not String.IsNullOrWhiteSpace(Selection.Preview) Then
                sb.AppendLine("### Selection")
                sb.AppendLine(Selection.Preview)
            End If
            If Not String.IsNullOrWhiteSpace([Structure].Summary) Then
                sb.AppendLine("### Structure")
                sb.AppendLine([Structure].Summary)
            End If
            If Not String.IsNullOrWhiteSpace(Document.Preview) Then
                sb.AppendLine("### Active document context")
                sb.AppendLine(Document.Preview)
            End If
            If ReaderErrors.Count > 0 Then
                sb.AppendLine("### Reader errors")
                For Each item In ReaderErrors
                    sb.AppendLine($"- [{item.ErrorCode}] {item.Reader}: {item.Message}")
                Next
            End If
            sb.AppendLine($"budget: {Budget.UsedChars}/{Budget.MaxChars}; truncated={Budget.Truncated}")
            Return sb.ToString()
        End Function

        Public Function ToJson() As String
            Return JsonConvert.SerializeObject(Me, Formatting.None)
        End Function

        Private Shared Function BuildSelectionRef(appType As String, address As String) As String
            Dim app = If(String.IsNullOrWhiteSpace(appType), "Office", appType.Trim())
            If String.IsNullOrWhiteSpace(address) Then Return $"{app}:Selection"
            Return $"{app}:{address.Trim()}"
        End Function

        Private Shared Function Truncate(value As String, maxLength As Integer, ByRef truncated As Boolean) As String
            Dim text = If(value, "")
            If maxLength <= 0 Then
                If text.Length > 0 Then truncated = True
                Return ""
            End If
            If text.Length <= maxLength Then Return text
            truncated = True
            Return text.Substring(0, maxLength)
        End Function
    End Class

    Public Class ContextDocument
        Public Property Ref As String = "ActiveDocument"
        Public Property Name As String = ""
        Public Property Preview As String = ""
    End Class

    Public Class ContextSelection
        Public Property Ref As String = ""
        Public Property Address As String = ""
        Public Property ItemCount As Integer
        Public Property DataType As String = ""
        Public Property Preview As String = ""
    End Class

    Public Class ContextStructure
        Public Property Summary As String = ""
        Public Property HasHeadings As Boolean
        Public Property HeadingCount As Integer
    End Class

    Public Class ContextBudget
        Public Property Strategy As String = "selection_first"
        Public Property MaxChars As Integer = 12000
        Public Property UsedChars As Integer
        Public Property Truncated As Boolean
    End Class

    Public Class ContextReaderError
        Public Property Reader As String = ""
        Public Property ErrorCode As String = ""
        Public Property Message As String = ""
    End Class

End Namespace
