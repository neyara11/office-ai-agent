Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 翻译领域模板 - 支持不同专业领域的翻译配置
''' </summary>
Public Class TranslateDomainTemplate
    ''' <summary>领域名称</summary>
    Public Property Name As String = ""

    ''' <summary>领域描述</summary>
    Public Property Description As String = ""

    ''' <summary>系统提示词模板</summary>
    Public Property SystemPrompt As String = ""

    ''' <summary>是否为内置模板</summary>
    Public Property IsBuiltIn As Boolean = False

    ''' <summary>专业术语列表（可选）</summary>
    Public Property Glossary As Dictionary(Of String, String) = New Dictionary(Of String, String)()

    Public Sub New()
    End Sub

    Public Sub New(name As String, description As String, systemPrompt As String, Optional isBuiltIn As Boolean = False)
        Me.Name = name
        Me.Description = description
        Me.SystemPrompt = systemPrompt
        Me.IsBuiltIn = isBuiltIn
    End Sub
End Class

''' <summary>
''' 翻译领域模板管理器
''' </summary>
Public Class TranslateDomainManager
    Private Shared ReadOnly fileName As String = "translate_domains.json"
    Private Shared ReadOnly filePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                                              ConfigSettings.OfficeAiAppDataFolder, fileName)

    Private Shared _templates As List(Of TranslateDomainTemplate)

    ''' <summary>
    ''' 获取所有领域模板
    ''' </summary>
    Public Shared ReadOnly Property Templates As List(Of TranslateDomainTemplate)
        Get
            If _templates Is Nothing Then
                Load()
            End If
            Return _templates
        End Get
    End Property

    ''' <summary>
    ''' 获取内置领域模板
    ''' </summary>
    Private Shared Function GetBuiltInTemplates() As List(Of TranslateDomainTemplate)
        Return New List(Of TranslateDomainTemplate) From {
            New TranslateDomainTemplate(
                "Универсальный",
                "Универсальный перевод для повседневных документов",
                "Ты профессиональный переводчик. Точно переведи приведённый ниже текст, сохраняя формат, тон и стиль оригинала. При переводе учитывай:
1. Сохраняй структуру абзацев
2. Имена собственные можно оставлять в оригинале или указывать в скобках
3. Формат чисел и дат сохраняй единообразным
4. Фразы должны быть естественными и соответствовать нормам целевого языка
5. Переводи только на целевой язык, указанный пользователем; не добавляй пояснений от себя",
                True
            ),
            New TranslateDomainTemplate(
                "Финансы и экономика",
                "Профессиональный перевод в сфере финансов, бухгалтерии и инвестиций",
                "Ты профессиональный переводчик в сфере финансов и экономики. Переведи приведённый ниже текст, учитывая:
1. Финансовые термины передавай стандартными эквивалентами (например: equity — капитал/доля, derivative — производный инструмент, hedge — хеджирование)
2. Статьи финансовой отчётности передавай нормативной бухгалтерской терминологией
3. Числа, валюты и проценты сохраняй точно
4. Названия компаний и организаций при первом упоминании можно приводить в оригинале
5. Сохраняй профессиональный и строгий тон
Частые термины: ROI — возврат на инвестиции, P/E ratio — коэффициент цена/прибыль, liquidity — ликвидность, leverage — леверидж
6. Переводи только на целевой язык и не добавляй пояснений",
                True
            ),
            New TranslateDomainTemplate(
                "Инженерия и техника",
                "Профессиональный перевод в сфере инженерии, техники и производства",
                "Ты профессиональный переводчик в сфере инженерии и техники. Переведи приведённый ниже текст, учитывая:
1. Технические термины передавай отраслевыми стандартными эквивалентами
2. Единицы измерения сохраняй или приводи к привычным для целевого языка (например, имперские/метрические)
3. Формулы и технические параметры сохраняй точно
4. Обозначения моделей и стандартов (например, ISO, GB) оставляй в оригинале
5. Сохраняй строгость и точность технической документации
6. Шаги и последовательности операций описывай чётко
7. Переводи только на целевой язык и не добавляй пояснений",
                True
            ),
            New TranslateDomainTemplate(
                "Право и договоры",
                "Профессиональный перевод в сфере права, договоров и нормативных актов",
                "Ты профессиональный переводчик в сфере права. Переведи приведённый ниже текст, учитывая:
1. Юридические термины передавай нормативными эквивалентами (например: jurisdiction — юрисдикция, liability — ответственность/обязательство, indemnify — возмещать убытки)
2. Структуру договорных пунктов сохраняй полностью
3. Даты, суммы и наименования сторон передавай без ошибок
4. Формат ссылок на правовые нормы соблюдай
5. Сохраняй строгость и авторитетность юридического текста
6. Не опускай и не добавляй содержание по своему усмотрению
7. Переводи только на целевой язык и не добавляй пояснений",
                True
            ),
            New TranslateDomainTemplate(
                "Медицина",
                "Профессиональный перевод в сфере медицины, здравоохранения и фармацевтики",
                "Ты профессиональный переводчик в сфере медицины. Переведи приведённый ниже текст, учитывая:
1. Медицинские термины передавай стандартными эквивалентами, при необходимости с латинским/английским оригиналом
2. Названия лекарств передавай международным непатентованным наименованием, можно с торговым названием
3. Дозировки, способ применения и кратность указывай без ошибок
4. Анатомические и патологические термины передавай нормативно
5. Клинические проявления, диагнозы и схемы лечения переводи точно
6. Сохраняй профессиональность и строгость медицинской литературы
7. Переводи только на целевой язык и не добавляй пояснений",
                True
            ),
            New TranslateDomainTemplate(
                "Научные работы",
                "Перевод научных статей и исследовательских отчётов",
                "Ты профессиональный переводчик в научной сфере. Переведи приведённый ниже текст, учитывая:
1. Научные термины передавай принятыми в дисциплине эквивалентами
2. Структуру аннотации, ключевых слов и основного текста сохраняй полностью
3. Формат цитат и списка литературы сохраняй единообразным
4. Подписи к рисункам и таблицам, примечания переводи точно
5. Методы исследования и анализ данных описывай чётко
6. Сохраняй строгость и объективность научного стиля
7. Переводи только на целевой язык и не добавляй пояснений",
                True
            ),
            New TranslateDomainTemplate(
                "Деловая переписка",
                "Перевод деловых писем, отчётов и коммерческих предложений",
                "Ты профессиональный переводчик в деловой сфере. Переведи приведённый ниже текст, учитывая:
1. Сохраняй профессионализм и вежливость делового письма
2. Названия компаний, должностей и подразделений передавай нормативно
3. Деловую терминологию используй уместно
4. Тон официальный, но не безразличный
5. Формат дат, чисел и контактных данных соблюдай
6. Структуру письма, обращения и подписи сохраняй полностью
7. Переводи только на целевой язык и не добавляй пояснений",
                True
            ),
            New TranslateDomainTemplate(
                "Литература и креатив",
                "Перевод литературных и креативных текстов",
                "Ты профессиональный литературный переводчик. Переведи приведённый ниже текст, учитывая:
1. Сохраняй художественность и образность оригинала
2. Стилистические приёмы, метафоры и аллюзии уместно адаптируй
3. Речь персонажей сохраняй с их характерными чертами
4. Культурные реалии передавай разумно
5. Ритм и звучание по возможности сохраняй
6. Перевод должен быть плавным, выразительным и естественным
7. Переводи только на целевой язык и не добавляй пояснений",
                True
            )
        }
    End Function

    ''' <summary>
    ''' 加载领域模板
    ''' </summary>
    Public Shared Sub Load()
        Try
            If File.Exists(filePath) Then
                Dim json As String = File.ReadAllText(filePath)
                _templates = JsonConvert.DeserializeObject(Of List(Of TranslateDomainTemplate))(json)
            End If
        Catch
            _templates = Nothing
        End Try

        ' 确保内置模板存在
        If _templates Is Nothing Then
            _templates = New List(Of TranslateDomainTemplate)()
        End If

        Dim builtInTemplates = GetBuiltInTemplates()
        For Each builtIn In builtInTemplates
            Dim existing = _templates.FirstOrDefault(Function(t) t.Name = builtIn.Name AndAlso t.IsBuiltIn)
            If existing Is Nothing Then
                _templates.Insert(0, builtIn)
            Else
                ' 更新内置模板的提示词
                existing.Description = builtIn.Description
                existing.SystemPrompt = builtIn.SystemPrompt
            End If
        Next

        Save()
    End Sub

    ''' <summary>
    ''' 保存领域模板
    ''' </summary>
    Public Shared Sub Save()
        Try
            Dim dir = Path.GetDirectoryName(filePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            Dim json As String = JsonConvert.SerializeObject(_templates, Formatting.Indented)
            File.WriteAllText(filePath, json)
        Catch
            ' 忽略保存错误
        End Try
    End Sub

    ''' <summary>
    ''' 添加自定义领域模板
    ''' </summary>
    Public Shared Sub AddTemplate(template As TranslateDomainTemplate)
        template.IsBuiltIn = False
        _templates.Add(template)
        Save()
    End Sub

    ''' <summary>
    ''' 删除领域模板（仅可删除非内置模板）
    ''' </summary>
    Public Shared Function RemoveTemplate(name As String) As Boolean
        Dim template = _templates.FirstOrDefault(Function(t) t.Name = name AndAlso Not t.IsBuiltIn)
        If template IsNot Nothing Then
            _templates.Remove(template)
            Save()
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' 根据名称获取模板
    ''' </summary>
    Public Shared Function GetTemplate(name As String) As TranslateDomainTemplate
        Return _templates.FirstOrDefault(Function(t) t.Name = name)
    End Function
End Class
