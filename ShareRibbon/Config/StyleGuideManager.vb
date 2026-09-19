' ShareRibbon\Config\StyleGuideManager.vb
' 排版规范管理器（单例模式）

Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 排版规范管理器（单例模式）
''' </summary>
Public Class StyleGuideManager
    Private Shared _instance As StyleGuideManager
    Private _styleGuides As List(Of StyleGuideResource)
    Private ReadOnly _configPath As String

    ''' <summary>获取单例实例</summary>
    Public Shared ReadOnly Property Instance As StyleGuideManager
        Get
            If _instance Is Nothing Then
                _instance = New StyleGuideManager()
            End If
            Return _instance
        End Get
    End Property

    ''' <summary>获取所有规范</summary>
    Public ReadOnly Property StyleGuides As List(Of StyleGuideResource)
        Get
            Return _styleGuides
        End Get
    End Property

    Private Sub New()
        _configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "styleguides.json")
        LoadStyleGuides()
    End Sub

    ''' <summary>加载规范配置</summary>
    Private Sub LoadStyleGuides()
        _styleGuides = New List(Of StyleGuideResource)()

        If File.Exists(_configPath) Then
            Try
                Dim json = File.ReadAllText(_configPath, Text.Encoding.UTF8)
                Dim loadedGuides = JsonConvert.DeserializeObject(Of List(Of StyleGuideResource))(json)

                ' 合并预置规范和用户规范
                MergePresetsAndUserGuides(loadedGuides)
            Catch ex As Exception
                Debug.WriteLine($"加载规范配置失败: {ex.Message}")
                LoadPresetStyleGuides()
            End Try
        Else
            ' 首次使用，加载预置规范
            LoadPresetStyleGuides()
            SaveStyleGuides()
        End If
    End Sub

    ''' <summary>合并预置规范和用户规范</summary>
    Private Sub MergePresetsAndUserGuides(userGuides As List(Of StyleGuideResource))
        If userGuides Is Nothing OrElse userGuides.Count = 0 Then
            LoadPresetStyleGuides()
            Return
        End If

        ' 获取预置规范ID列表
        Dim presets = GetPresetStyleGuides()
        Dim presetIds = presets.Select(Function(p) p.Id).ToHashSet()

        ' 先处理预置规范：如果用户数据中有同ID的（可能被修改过），使用用户版本；否则使用默认预置
        For Each preset In presets
            Dim userVersion = userGuides.FirstOrDefault(Function(g) g.Id = preset.Id)
            If userVersion IsNot Nothing Then
                ' 使用用户版本（可能被修改过），确保标记为预置
                userVersion.IsPreset = True
                _styleGuides.Add(userVersion)
            Else
                ' 用户数据中没有此预置规范（可能是新增的预置），添加默认版本
                _styleGuides.Add(preset)
            End If
        Next

        ' 再添加用户自定义规范（非预置的）
        For Each userGuide In userGuides
            ' 跳过预置规范ID（已在上面处理过）
            If presetIds.Contains(userGuide.Id) Then
                Continue For
            End If
            ' 添加用户自定义规范
            If Not _styleGuides.Any(Function(g) g.Id = userGuide.Id) Then
                userGuide.IsPreset = False ' 确保用户规范不被标记为预置
                _styleGuides.Add(userGuide)
            End If
        Next
    End Sub

    ''' <summary>加载预置规范</summary>
    Private Sub LoadPresetStyleGuides()
        _styleGuides.AddRange(GetPresetStyleGuides())
    End Sub

    ''' <summary>保存规范配置</summary>
    Public Sub SaveStyleGuides()
        Try
            Dim dir = Path.GetDirectoryName(_configPath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json = JsonConvert.SerializeObject(_styleGuides, Formatting.Indented)
            File.WriteAllText(_configPath, json, Text.Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"保存规范配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>添加规范</summary>
    Public Sub AddStyleGuide(guide As StyleGuideResource)
        If String.IsNullOrEmpty(guide.Id) Then
            guide.Id = Guid.NewGuid().ToString()
        End If
        guide.CreatedAt = DateTime.Now
        guide.LastModified = DateTime.Now
        guide.IsPreset = False
        _styleGuides.Add(guide)
        SaveStyleGuides()
    End Sub

    ''' <summary>更新规范</summary>
    Public Sub UpdateStyleGuide(guide As StyleGuideResource)
        Dim existing = _styleGuides.FirstOrDefault(Function(g) g.Id = guide.Id)
        If existing IsNot Nothing Then
            guide.LastModified = DateTime.Now
            Dim index = _styleGuides.IndexOf(existing)
            _styleGuides(index) = guide
            SaveStyleGuides()
        End If
    End Sub

    ''' <summary>删除规范</summary>
    Public Function DeleteStyleGuide(guideId As String) As Boolean
        Dim guide = _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
        If guide IsNot Nothing Then
            If guide.IsPreset Then
                Return False ' 预置规范不可删除
            End If
            _styleGuides.Remove(guide)
            SaveStyleGuides()
            Return True
        End If
        Return False
    End Function

    ''' <summary>复制规范</summary>
    Public Function DuplicateStyleGuide(guideId As String, newName As String) As StyleGuideResource
        Dim original = _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
        If original IsNot Nothing Then
            ' 深拷贝
            Dim json = JsonConvert.SerializeObject(original)
            Dim duplicate = JsonConvert.DeserializeObject(Of StyleGuideResource)(json)
            duplicate.Id = Guid.NewGuid().ToString()
            duplicate.Name = If(String.IsNullOrEmpty(newName), original.Name & " (副本)", newName)
            duplicate.IsPreset = False
            duplicate.CreatedAt = DateTime.Now
            duplicate.LastModified = DateTime.Now
            _styleGuides.Add(duplicate)
            SaveStyleGuides()
            Return duplicate
        End If
        Return Nothing
    End Function

    ''' <summary>导出规范到文件</summary>
    Public Function ExportStyleGuide(guideId As String, filePath As String) As Boolean
        Try
            Dim guide = _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
            If guide IsNot Nothing Then
                ' 导出为原始格式（txt/md）
                Dim extension = If(String.IsNullOrEmpty(guide.SourceFileExtension), ".md", guide.SourceFileExtension)
                Dim actualPath = If(filePath.EndsWith(extension), filePath, filePath & extension)
                File.WriteAllText(actualPath, guide.GuideContent, Text.Encoding.UTF8)
                Return True
            End If
        Catch ex As Exception
            Debug.WriteLine($"导出规范失败: {ex.Message}")
        End Try
        Return False
    End Function

    ''' <summary>根据ID获取规范</summary>
    Public Function GetStyleGuideById(guideId As String) As StyleGuideResource
        Return _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
    End Function

    ''' <summary>获取所有分类</summary>
    Public Function GetAllCategories() As List(Of String)
        Dim categories = _styleGuides.Select(Function(g) g.Category).Distinct().ToList()
        categories.Insert(0, "全部")
        Return categories
    End Function

    ''' <summary>获取所有规范（兼容方法名）</summary>
    Public Function GetAllStyleGuides() As List(Of StyleGuideResource)
        Return _styleGuides.ToList()
    End Function

    ''' <summary>刷新规范列表（重新从文件加载）</summary>
    Public Sub Refresh()
        LoadStyleGuides()
    End Sub

#Region "预置规范"

    ''' <summary>获取预置规范列表</summary>
    Private Function GetPresetStyleGuides() As List(Of StyleGuideResource)
        Dim presets As New List(Of StyleGuideResource)()

        ' 预置规范1：GB/T 9704-2012 党政机关公文格式规范
        presets.Add(CreateOfficialDocumentGuide())

        ' 预置规范2：学术论文排版规范
        presets.Add(CreateAcademicPaperGuide())

        ' 预置规范3：商务报告排版规范
        presets.Add(CreateBusinessReportGuide())

        Return presets
    End Function

    ''' <summary>创建党政机关公文格式规范</summary>
    Private Function CreateOfficialDocumentGuide() As StyleGuideResource
        Dim guide As New StyleGuideResource With {
            .Id = "preset-guide-official",
            .Name = "Стандарт оформления официальных документов КНР",
            .Description = "Формат официальных документов по стандарту GB/T 9704-2012 (КНР)",
            .Category = "Официальные документы",
            .TargetApp = "Word",
            .IsPreset = True,
            .SourceFileName = "GB_T_9704-2012.md",
            .SourceFileExtension = ".md",
            .FileEncoding = "UTF-8"
        }

        guide.GuideContent = "# 党政机关公文格式规范（GB/T 9704-2012）

## 一、页面设置

### 1. 纸张规格
- **尺寸**：A4（210mm × 297mm）
- **方向**：纵向

### 2. 页边距
- **上边距**：37mm（约3.7cm）
- **下边距**：35mm（约3.5cm）
- **左边距**：28mm（约2.8cm）
- **右边距**：26mm（约2.6cm）

### 3. 版心尺寸
- **宽度**：156mm
- **高度**：225mm

## 二、字体要求

### 1. 发文机关标志
- **字体**：方正小标宋简体
- **字号**：由发文机关自定，但应醒目美观
- **颜色**：红色（一般为RGB: 192, 0, 0）

### 2. 发文字号
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **对齐**：居中

### 3. 标题
- **字体**：方正小标宋简体
- **字号**：二号（22pt）
- **对齐**：居中
- **加粗**：是

### 4. 主送机关
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **对齐**：左对齐，顶格

### 5. 正文
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **对齐**：两端对齐
- **首行缩进**：2字符

### 6. 一级标题
- **字体**：黑体
- **字号**：三号（16pt）
- **编号格式**：一、二、三...

### 7. 二级标题
- **字体**：楷体_GB2312
- **字号**：三号（16pt）
- **编号格式**：（一）（二）（三）...

### 8. 三级标题
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **编号格式**：1. 2. 3. ...
- **加粗**：是

## 三、段落格式

### 1. 行距
- **正文行距**：以每面22行、每行28字并撑满版心为目标；Word中通常使用固定值约28磅
- **段前间距**：0
- **段后间距**：0

### 2. 缩进
- **首行缩进**：2字符（约0.85cm）
- **左缩进**：0
- **右缩进**：0

## 四、特殊元素

### 1. 红色分隔线
- **位置**：发文机关标志下方
- **宽度**：与版心等宽（156mm）
- **粗细**：约2pt
- **颜色**：红色

### 2. 页码
- **位置**：页脚居中
- **格式**：- X -（如：- 1 -）
- **字体**：宋体
- **字号**：四号（14pt）

### 3. 成文日期
- **格式**：XXXX年X月X日
- **位置**：正文下方右侧
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）

### 4. 附注与版记
- **附注**：如联系人、联系电话等，一般居左空二字加圆括号
- **抄送机关**：置于版记，使用「抄送：×××。」格式
- **印发机关和印发日期**：置于版记末行，左右分列

## 五、注意事项

1. 公文用纸一般使用60g/m²—80g/m²胶版印刷纸或复印纸
2. 公文如需标注紧急程度，应在公文首页左上角标注
3. 附件说明位于正文下方、成文日期上方
4. 印章应端正、居中，上不压正文、下要骑年盖月
"

        guide.Tags = New List(Of String) From {"официальный", "документ", "GB/T 9704", "госстандарт", "公文", "行政", "政府文件"}

        Return guide
    End Function

    ''' <summary>创建学术论文排版规范</summary>
    Private Function CreateAcademicPaperGuide() As StyleGuideResource
        Dim guide As New StyleGuideResource With {
            .Id = "preset-guide-academic",
            .Name = "Стандарт оформления научной статьи",
            .Description = "Универсальный стандарт оформления научных статей (ориентир — ГОСТ и GB/T 7713)",
            .Category = "Научный",
            .TargetApp = "Word",
            .IsPreset = True,
            .SourceFileName = "academic_paper_guide.md",
            .SourceFileExtension = ".md",
            .FileEncoding = "UTF-8"
        }

        guide.GuideContent = "# Стандарт оформления научной статьи

## 1. Параметры страницы

### 1.1. Бумага
- **Размер**: A4 (210mm × 297mm)
- **Ориентация**: книжная

### 1.2. Поля
- **Верхнее**: 2,54cm (1 дюйм)
- **Нижнее**: 2,54cm (1 дюйм)
- **Левое**: 3,18cm (1,25 дюйма)
- **Правое**: 3,18cm (1,25 дюйма)

## 2. Иерархия заголовков

### 2.1. Название работы
- **Шрифт**: Arial
- **Размер**: 18pt
- **Выравнивание**: по центру
- **Полужирный**: да
- **Интервал после**: 1 строка

### 2.2. Сведения об авторах
- **Шрифт**: Times New Roman
- **Размер**: 12pt
- **Выравнивание**: по центру
- **Формат**: ФИО (организация, город, индекс)

### 2.3. Аннотация
- **Шрифт заголовка**: Arial
- **Размер заголовка**: 10,5pt
- **Шрифт текста**: Times New Roman
- **Размер текста**: 10,5pt
- **Междустрочный интервал**: 1,5
- **Ключевые слова**: 3–5, через точку с запятой

### 2.4. Заголовок 1-го уровня
- **Шрифт**: Arial
- **Размер**: 14pt
- **Выравнивание**: по левому краю
- **Полужирный**: да
- **Нумерация**: 1, 2, 3 … или I, II, III …
- **Интервал до**: 0,5 строки
- **Интервал после**: 0,5 строки

### 2.5. Заголовок 2-го уровня
- **Шрифт**: Arial
- **Размер**: 12pt
- **Выравнивание**: по левому краю
- **Полужирный**: да
- **Нумерация**: 1.1, 1.2 … или (1), (2) …

### 2.6. Заголовок 3-го уровня
- **Шрифт**: Times New Roman
- **Размер**: 12pt
- **Выравнивание**: по левому краю
- **Полужирный**: да
- **Нумерация**: 1.1.1, 1.1.2 …

## 3. Формат основного текста

### 3.1. Шрифт и размер
- **Основной шрифт**: Times New Roman
- **Размер**: 12pt

### 3.2. Формат абзацев
- **Выравнивание**: по ширине
- **Отступ первой строки**: 2 знака
- **Междустрочный интервал**: 1,5
- **Интервалы до и после**: 0

## 4. Формат рисунков и таблиц

### 4.1. Рисунки
- **Номер**: Рис. 1, Рис. 2 … (по центру)
- **Подпись**: под рисунком, по центру
- **Шрифт**: Times New Roman, 10,5pt

### 4.2. Таблицы
- **Номер**: Таблица 1, Таблица 2 … (по центру)
- **Название**: над таблицей, по центру
- **Шрифт**: Times New Roman, 10,5pt
- **Шапка**: полужирная

## 5. Список литературы

### 5.1. Оформление (ГОСТ Р 7.0.5-2008)
- **Шрифт**: Times New Roman
- **Размер**: 10,5pt
- **Междустрочный интервал**: одинарный
- **Висячий отступ**: 2 знака

### 5.2. Примеры записей
- **Статья**: [1] Автор. Название статьи // Название журнала. — Год. — Т. X, № Y. — С. 00–00.
- **Книга**: [2] Автор. Название книги. — Город: Издательство, год. — С. 00–00.
- **Сборник**: [3] Автор. Название доклада // Название сборника. — Город: Издательство, год. — С. 00–00.
- **Диссертация**: [4] Автор. Название: дис. … канд. наук. — Город, год.
- **Электронный ресурс**: [5] Автор. Название. — URL: адрес (дата обращения: дд.мм.гггг).

## 6. Нумерация страниц

- **Позиция**: внизу по центру
- **Формат**: арабские цифры
- **Шрифт**: Times New Roman
- **Размер**: 10,5pt
- **Первая страница**: номер можно не показывать
"

        guide.Tags = New List(Of String) From {"научная", "статья", "журнал", "диплом", "学术", "论文", "期刊", "毕业论文"}

        Return guide
    End Function

    ''' <summary>创建商务报告排版规范</summary>
    Private Function CreateBusinessReportGuide() As StyleGuideResource
        Dim guide As New StyleGuideResource With {
            .Id = "preset-guide-business",
            .Name = "Стандарт оформления бизнес-отчёта",
            .Description = "Современный стандарт оформления бизнес-отчётов: проектные, аналитические и другие отчёты",
            .Category = "Бизнес",
            .TargetApp = "Word",
            .IsPreset = True,
            .SourceFileName = "business_report_guide.md",
            .SourceFileExtension = ".md",
            .FileEncoding = "UTF-8"
        }

        guide.GuideContent = "# Стандарт оформления бизнес-отчёта

## 1. Параметры страницы

### 1.1. Бумага
- **Размер**: A4
- **Ориентация**: книжная (в особых случаях допустима альбомная)

### 1.2. Поля
- **Верхнее**: 2,5cm
- **Нижнее**: 2,5cm
- **Левое**: 2,5cm
- **Правое**: 2,5cm

## 2. Шрифты

### 2.1. Главный заголовок (титул отчёта)
- **Шрифт**: Arial
- **Размер**: 24pt
- **Цвет**: тёмно-синий (#2E5090)
- **Выравнивание**: по центру
- **Полужирный**: да

### 2.2. Подзаголовок
- **Шрифт**: Arial
- **Размер**: 16pt
- **Цвет**: серый (#666666)
- **Выравнивание**: по центру

### 2.3. Заголовок раздела (1-й уровень)
- **Шрифт**: Arial
- **Размер**: 18pt
- **Цвет**: тёмно-синий (#2E5090)
- **Выравнивание**: по левому краю
- **Полужирный**: да
- **Интервал до**: 1 строка
- **Интервал после**: 0,5 строки

### 2.4. Заголовок подраздела (2-й уровень)
- **Шрифт**: Arial
- **Размер**: 14pt
- **Цвет**: чёрный
- **Выравнивание**: по левому краю
- **Полужирный**: да
- **Интервал до**: 0,5 строки
- **Интервал после**: 0,5 строки

### 2.5. Основной текст
- **Шрифт**: Arial
- **Размер**: 12pt
- **Цвет**: чёрный (#333333)
- **Выравнивание**: по ширине
- **Отступ первой строки**: 2 знака
- **Междустрочный интервал**: 1,5

## 3. Особые элементы

### 3.1. Маркированный список
- **Маркер**: закрашенный кружок (•) или короткое тире (–)
- **Отступ**: 1cm
- **Междустрочный интервал**: 1,2
- **Цвет**: для акцентов можно использовать фирменный цвет

### 3.2. Таблицы с данными
- **Фон шапки**: светло-синий (#E6F0FA)
- **Границы**: светло-серый (#CCCCCC), 1pt
- **Внутренние отступы ячеек**: сверху и снизу 5pt, слева и справа 8pt
- **Шрифт шапки**: Arial, полужирный
- **Шрифт данных**: Arial, обычный

### 3.3. Подписи к рисункам и таблицам
- **Позиция**: под рисунком или таблицей
- **Шрифт**: Arial
- **Размер**: 10,5pt
- **Цвет**: серый
- **Выравнивание**: по центру
- **Формат**: Рис. X: название

### 3.4. Врезка-цитата
- **Фон**: светло-серый (#F5F5F5)
- **Граница**: слева сплошная 4pt, фирменный цвет
- **Внутренние отступы**: 15pt
- **Шрифт**: Arial, курсив

## 4. Колонтитулы

### 4.1. Верхний колонтитул
- **Содержимое**: название отчёта или компании
- **Шрифт**: Arial
- **Размер**: 9pt
- **Цвет**: серый
- **Позиция**: по правому краю
- **Разделитель**: тонкая линия (0,5pt)

### 4.2. Нижний колонтитул
- **Содержимое**: номер страницы
- **Формат**: Страница X из X
- **Шрифт**: Arial
- **Размер**: 9pt
- **Позиция**: по центру или по правому краю

## 5. Цветовая схема

### Рекомендуемая палитра
- **Основной цвет**: #2E5090 (тёмно-синий)
- **Дополнительный**: #4A90D9 (ярко-синий)
- **Акцент**: #F5A623 (оранжевый)
- **Цвет текста**: #333333 (тёмно-серый)
- **Второстепенный текст**: #666666 (серый)
- **Фон**: #FFFFFF (белый)

## 6. Примечания

1. Соблюдай единый стиль; используй не более трёх основных цветов
2. Важные данные визуализируй диаграммами
3. Не перегружай страницы, оставляй поля
4. Рекомендуемое соотношение текста и графики — 6:4
5. Ключевые выводы выделяй полужирным или цветом
"

        guide.Tags = New List(Of String) From {"бизнес", "отчёт", "отчет", "компания", "проект", "商务", "报告", "企业", "项目"}

        Return guide
    End Function

#End Region

End Class
