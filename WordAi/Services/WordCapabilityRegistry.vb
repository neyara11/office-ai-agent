' WordAi\Services\WordCapabilityRegistry.vb
' Word capability metadata used by WordActionHarness planning.

Namespace Services

    Public Enum WordCapabilityRiskLevel
        Low
        Medium
        High
    End Enum

    Public Class WordCapabilityDescriptor
        Public Property Id As String = ""
        Public Property DisplayName As String = ""
        Public Property Kind As WordActionKind = WordActionKind.None
        Public Property Description As String = ""
        Public Property InputSchema As String = ""
        Public Property RiskLevel As WordCapabilityRiskLevel = WordCapabilityRiskLevel.Medium
        Public Property SupportsPreview As Boolean
        Public Property SupportsUndo As Boolean = True
        Public Property ObserveContract As String = ""
        Public Property RepairContract As String = ""
        Public Property ExplainContract As String = ""
        Public Property ExampleRequests As New List(Of String)()

        Public Function ToHumanReadableSummary() As String
            Dim parts As New List(Of String) From {
                $"{DisplayName} ({Id})",
                $"Риск: {RiskLevel}",
                $"Предпросмотр: {If(SupportsPreview, "да", "нет")}",
                $"Отмена: {If(SupportsUndo, "попадает в стек отмены Word", "описывается исполнителем")}"
            }

            If Not String.IsNullOrWhiteSpace(ObserveContract) Then parts.Add("Наблюдение: " & ObserveContract)
            If Not String.IsNullOrWhiteSpace(RepairContract) Then parts.Add("Исправление: " & RepairContract)
            Return String.Join("; ", parts)
        End Function
    End Class

    Public NotInheritable Class WordCapabilityRegistry

        Private Shared ReadOnly _capabilities As Lazy(Of List(Of WordCapabilityDescriptor)) =
            New Lazy(Of List(Of WordCapabilityDescriptor))(AddressOf BuildCapabilities)

        Private Sub New()
        End Sub

        Public Shared Function All() As IReadOnlyList(Of WordCapabilityDescriptor)
            Return _capabilities.Value
        End Function

        Public Shared Function Find(kind As WordActionKind) As WordCapabilityDescriptor
            Return _capabilities.Value.FirstOrDefault(Function(item) item.Kind = kind)
        End Function

        Public Shared Function Require(kind As WordActionKind) As WordCapabilityDescriptor
            Dim descriptor = Find(kind)
            If descriptor Is Nothing Then
                Return New WordCapabilityDescriptor With {
                    .Id = "word.unknown",
                    .DisplayName = "Неизвестная возможность Word",
                    .Kind = kind,
                    .Description = "Незарегистрированная возможность Word",
                    .RiskLevel = WordCapabilityRiskLevel.High,
                    .SupportsPreview = False,
                    .SupportsUndo = False
                }
            End If
            Return descriptor
        End Function

        Private Shared Function BuildCapabilities() As List(Of WordCapabilityDescriptor)
            Return New List(Of WordCapabilityDescriptor) From {
                New WordCapabilityDescriptor With {
                    .Id = "word.proofread",
                    .DisplayName = "Вычитка Word",
                    .Kind = WordActionKind.Proofread,
                    .Description = "Прочитать текущее выделение или весь документ и построить план проверки опечаток, пунктуации, грамматики, терминологии и единообразия стиля.",
                    .InputSchema = "ProofreadIntentPlan(scope, issueTypes, applyMode, showSidePanel)",
                    .RiskLevel = WordCapabilityRiskLevel.Low,
                    .SupportsPreview = True,
                    .SupportsUndo = True,
                    .ObserveContract = "Результаты проверки должны попадать на боковую панель или в поток правок с высокой уверенностью, сохраняя возможность подтверждения пользователем.",
                    .RepairContract = "Если не удаётся прочитать документ или выделение, перейти к плану проверки всего текста или сообщить об отсутствии обрабатываемого документа.",
                    .ExplainContract = "Объяснить область проверки, типы проблем и применение автоправок с высокой уверенностью.",
                    .ExampleRequests = New List(Of String) From {"Проверить весь текст", "Проверить опечатки в выделенном абзаце", "Проверить пунктуацию и терминологию в этом документе"}
                },
                New WordCapabilityDescriptor With {
                    .Id = "word.direct-formatting",
                    .DisplayName = "Прямое форматирование Word",
                    .Kind = WordActionKind.DirectFormatting,
                    .Description = "Преобразовать явные запросы на размер шрифта, шрифт, полужирный, цвет, выравнивание и другое в FormattingIntentPlan и применить их к документу Word.",
                    .InputSchema = "FormattingIntentPlan(scope, operations, confidence)",
                    .RiskLevel = WordCapabilityRiskLevel.Medium,
                    .SupportsPreview = False,
                    .SupportsUndo = True,
                    .ObserveContract = "После выполнения выборочно прочитать целевой Range и проверить, что размер шрифта, шрифт, выравнивание и другие операции применились.",
                    .RepairContract = "При неудачном наблюдении вернуть сводку ошибки, а верхний уровень переходит к обычному чату или последующему repair loop.",
                    .ExplainContract = "Объяснить область действия, применённые операции форматирования, результат наблюдения и способ отмены.",
                    .ExampleRequests = New List(Of String) From {"Увеличить весь шрифт на 2 пункта", "Сделать выделенный текст шрифтом SimSun и полужирным", "Установить межстрочный интервал основного текста 1.5"}
                },
                New WordCapabilityDescriptor With {
                    .Id = "word.numbering",
                    .DisplayName = "Перенумерация Word",
                    .Kind = WordActionKind.Numbering,
                    .Description = "Распознать абзацы с автонумерацией Word и перенумеровать их в непрерывную последовательность 1,2,3....",
                    .InputSchema = "NumberingRequest(scope, sequenceGoal)",
                    .RiskLevel = WordCapabilityRiskLevel.Medium,
                    .SupportsPreview = False,
                    .SupportsUndo = True,
                    .ObserveContract = "После выполнения прочитать несколько первых нумерованных абзацев и показать ListString и предпросмотр текста.",
                    .RepairContract = "Если абзацев с автонумерацией нет, не преобразовывать обычную текстовую нумерацию и сообщить пользователю, что текущая область не может быть обработана.",
                    .ExplainContract = "Объяснить область обработки, число обнаруженных и применённых элементов, предпросмотр наблюдения и способ отмены.",
                    .ExampleRequests = New List(Of String) From {"Заменить предыдущие номера на 12345", "Перенумеровать список в непрерывную последовательность", "Упорядочить автонумерацию во всём тексте"}
                },
                New WordCapabilityDescriptor With {
                    .Id = "word.semantic-reformat",
                    .DisplayName = "Семантическая вёрстка Word",
                    .Kind = WordActionKind.SemanticReformat,
                    .Description = "План интеллектуальной вёрстки заголовков, уровней, оглавления и структурных абзацев с приоритетом предпросмотра и подтверждения.",
                    .InputSchema = "WordFormattingTaskPlan(scopeSummary, standardName, targetSummary, operations)",
                    .RiskLevel = WordCapabilityRiskLevel.High,
                    .SupportsPreview = True,
                    .SupportsUndo = True,
                    .ObserveContract = "После применения подсчитать ожидаемые, изменённые и исправленные абзацы и объяснить причины промахов.",
                    .RepairContract = "Если структурные абзацы не найдены, скорректировать нацеливание по результатам observe или оставить их как элементы для подтверждения.",
                    .ExplainContract = "Объяснить стандарт вёрстки, число предпросмотров, применённых и исправленных элементов и способ отмены.",
                    .ExampleRequests = New List(Of String) From {"Упорядочить уровни заголовков", "Оптимизировать вёрстку всего текста по формату официального документа", "Привести в порядок оглавление и структуру нумерации"}
                }
            }
        End Function

    End Class

End Namespace
