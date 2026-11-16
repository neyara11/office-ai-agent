## Why

Пять пользовательских жалоб указывают на общий корень: агент получает неполный контекст Office, исполняет невалидированные команды и показывает служебный протокол в чате.

1. Модель постоянно выдаёт битые команды.
2. Модель «не видит» текст презентации.
3. Дизайн слайдов генерируется произвольно.
4. Чатом неудобно пользоваться, логика непонятна; настройки выглядят сломанными.
5. Вложения (docx/xlsx/doc/ppt) не читаются, CSV читается с трудом.

### Диагностика

| Проблема | Где корень |
|---|---|
| Битые команды | `ShareRibbon/Controls/Services/CodeExecutionService.vb:49-52` — «умный ремонт» JSON делает `Replace("\"", """)`, снимая легальное экранирование; `BalanceBraces` дописывает скобки вслепую. `PowerPointAi/ChatControl.vb:1334-1355` — fallback снимает слой экранирования со всей строки, портя VBA с кавычками. |
| Схема не проверяется в агенте | `PowerPointJsonCommandSchema.ValidateJsonStructure` вызывается только в легаси-ветке (`PowerPointAi/ChatControl.vb:1231`), `GetFormatCorrectionPrompt` (`PowerPointJsonCommandSchema.vb:120`) не вызывается; `LoopEngine` исполняет `toolCall.Parameters` сразу (`ShareRibbon/Agent/LoopEngine.vb:168-170`). |
| Два конкурирующих промпта | `ShareRibbon/Agent/PromptManager.vb` (агент) и `ShareRibbon/Config/PromptManager.vb` + `*JsonCommandSchema.vb` (легаси) описывают разные форматы. |
| Бедные схемы инструментов | `ShareRibbon/Tools/ppt/InsertTable.json:23-27` — `data` описано как «Данные таблицы», нет вложенной схемы и `slideIndex`; `CreateSlides.json` — проза вместо структуры. |
| Не видит текст презентации | `PowerPointAi/Context/PowerPointContextProvider.vb:96-99` отдаёт только «слайдов N, текущий M»; `ShareRibbon/Agent/PromptManager.vb:214-221` обрезает содержимое документа до 500 символов. |
| Произвольный дизайн | Два пути: профессиональный `CreateSlides` (`PowerPointAi/Design/*`) и сырые `InsertSlide/InsertText/InsertTable/BeautifySlides`, которые обходят дизайн-систему; в контексте нет темы/шрифтов/паттернов текущей колоды. |
| Вложения не читаются | `ShareRibbon/Resources/js/message-sender.js:91` отправляет `file.path`, которого в WebView2 нет → уходит только имя; бэкенд ищет файл только в папке активного документа (`ShareRibbon/Controls/BaseChatControl.vb:1301-1312`). Разбор docx/xlsx/pptx идёт через COM с новым экземпляром приложения (`WordAi/ChatControl.vb:2156`, `ExcelAi/ChatControl.vb:897`, `PowerPointAi/ChatControl.vb:435`) без managed-fallback. |
| CSV с трудом | `ShareRibbon/Controls/Services/FileParserService.vb:184-191` для `.csv` пробует GB18030 до определения кодировки по содержимому → русский UTF-8 без BOM определяется как GBK. `ParseCsvFile` режет по строкам и `Split(delimiter)` без учёта кавычек, отдаёт 25 строк. |
| Неудобный чат | В ленте виден сырой JSON/`\n`, счётчик «Потрачено токенов: 0», неясные состояния плана; настройки — ручная раскладка `ShareRibbon/Config/ConfigApiForm.vb` (~106 КБ) с обрезанными списками и чекбоксами. |

## What Changes

- **Единый шлюз валидации команд**: проверка параметров по JSON Schema/типизированным `ToolParam` до исполнения в `ToolRegistry`/`LoopEngine`; ошибка возвращается как `CodeOperationSchemaInvalid` с путём до поля и `recoverable=True`.
- **Удаление деструктивного «ремонта» JSON**: остаётся строгий `JToken.Parse` и один явный fallback снятия слоя экранирования, применяемый только если результат проходит полную валидацию.
- **Схемы инструментов из одного источника**: вложенные схемы для `slides[]`, `chart`, `table`, `data`, компактные примеры; секция промпта генерируется из `Tools/*.json`; легаси-дубли выводятся из эксплуатации.
- **Протокол скрыт из чата**: human-readable «Понял → План → Выполняю → Результат», JSON/thought — за «Показать протокол»; починить учёт токенов и статусы шагов; добавить «Отменить» и отчёт об изменениях.
- **Карта текста презентации**: новый `PowerPointDeckReader` (title/body/notes/таблицы/типы фигур/переполнения по слайдам) с токен-бюджетом и кэшем по `FullName + ModifiedTime`; прокидывается в `ContextPack.Document` и `currentContent`.
- **Дизайн через `CreateSlides`**: генерация слайдов маршрутизируется только в `CreateSlides`; `Insert*` — только точечные правки с `targetRef` и наследованием стиля; в `CreateSlides` передаётся дизайн-контекст колоды (палитра, шрифты, размер, примеры слайдов); обязательный preflight и визуальная верификация.
- **Реальные пути вложений**: host-диалог выбора файла через VSTO-bridge или копирование выбранного файла в temp-каталог приложения при drag&drop/`input type=file`.
- **Managed-разбор вложений**: OpenXml SDK для `.docx/.xlsx/.pptx`, COM — только fallback для `.doc/.xls/.ppt`; структурированное представление (листы/диапазоны/заголовки/примеры) и лимит по токенам; отчёт о разборе в UI и в `ToolResult`.
- **Корректный CSV**: UTF-8-валидность проверяется первой, codepage — только при невалидном UTF-8; парсер RFC 4180 (кавычки, экранирование, многострочные поля).
- **Редизайн настроек**: dock/anchor вместо ручной раскладки; вкладки «Модели / Рассуждения / Инструменты и MCP / Дополнительно»; «Проверить подключение», валидация URL/ключа, явные «Сохранить/Отмена».

## Capabilities

### New Capabilities

- `command-schema-gate`: валидация и нормализация команд до исполнения, структурированные schema-ошибки и repair-контракт.
- `presentation-deck-context`: чтение структуры и текста презентации с бюджетом и кэшем.
- `attachment-ingestion`: получение реальных путей, managed-разбор docx/xlsx/pptx/csv, структурированная выдача и отчёт.
- `chat-execution-ux`: человекочитаемые состояния выполнения, скрытие протокола, отмена и отчёт об изменениях.

### Modified Capabilities

- `command-generation`: схемы инструментов и промпты приводятся к единому источнику; протокол больше не показывается пользователю.
- `deck-design`: маршрутизация на `CreateSlides`, дизайн-контекст колоды, обязательная верификация.

## Impact

- **ShareRibbon**: `Controls/Services/CodeExecutionService.vb`, `Controls/Services/FileParserService.vb`, `Controls/BaseChatControl.vb`, `Agent/ToolRegistry.vb`, `Agent/ToolParameterValidator.vb` (new), `Agent/AgentRunJournal.vb` (new), `Agent/LoopEngine*.vb`, `Agent/PromptManager.vb`, `Agent/Context/ContextPack.vb`, `Tools/ppt/*.json`, `Resources/js/message-sender.js`, `Config/ConfigApiForm.vb`.
- **PowerPointAi**: `Context/PowerPointContextProvider.vb`, новый `Context/PowerPointDeckReader.vb`, `ChatControl.vb`, `Design/*` (контракт дизайн-контекста).
- **WordAi / ExcelAi**: `ChatControl.vb` — COM-разбор переводится в fallback, основной путь в `ShareRibbon`.
- **Проектные файлы**: новые `.vb`/ресурсы добавить в соответствующие `.vbproj`.
- **Совместимость**: легаси JSON-формат некоторое время поддерживается как fallback; BREAKING-изменений публичного API нет.
- **Вне области**: смена провайдера моделей, переписывание `AgentKernel`/`LoopEngine`, новая система памяти.
