# Handoff: макеты и UI-панели

Документ для агента, который продолжит работу. Сначала прочитать `CLAUDE.md`, `AGENTS.md`, `PowerPointAi/AGENTS.md`, `ShareRibbon/AGENTS.md`, `ShareRibbon/Controls/AGENTS.md`, `ShareRibbon/Config/AGENTS.md`.

## Текущее состояние

- Все четыре проекта собираются: `build-code.bat`.
- Уже реализованы и работают нативные PPT-инструменты:
  - `DuplicateSlide` — `slideIndex` с 1, `position: after|end`, сохраняет оформление/картинки/таблицы/диаграммы;
  - `ReplaceSlideText` — `slideIndex` с 1, `replacements:[{find,replaceWith}]`, обходит фигуры, группы, **ячейки таблиц**, гибко матчит переносы/пробелы;
  - `UpdateChartData` — `slideIndex` с 1, `chartIndex`, `categories`, `values`, `seriesName`; меняет данные первого ряда, сохраняя тип/оформление;
  - `InsertSlide` — `layout` + placeholder-заполнение;
  - `CreateSlides` — Scene/дизайн-система, **не наследует дизайн текущей колоды**.
- `ExecuteVBA` отключён: не в `GetAvailableTools`, не в legacy `SupportedCommands`, хост проверяет `AccessVBOM` до COM. Не возвращать VBA в типовые сценарии.
- Диагностика: `ShareRibbon/Agent/AgentRunJournal.vb` пишет JSONL в `%LOCALAPPDATA%\OfficeAiAppData\logs\agent-runs-YYYYMMDD.jsonl` (фазы run_start/plan/thought/tool_call/tool_result/repair). Отключается `OFFICE_AI_AGENT_DIAGNOSTICS=0`. В логах нет base64/скриншотов.
- UI-карточка Agent: план показывает только «Остановить», статус «Выполняется…»; JSON из `thought` вырезается. Не возвращать кнопки «Начать/Изменить/Отмена» без реального ожидания подтверждения в backend.

## Критичные грабли

### 1. Ресурсы UI не обновляются сами

- HTML/JS/CSS встраиваются через `ShareRibbon/My Project/Resources.resx` (ResXFileRef) и распаковываются `ShareRibbon/Resources/ResourceExtractor.vb` в `%LOCALAPPDATA%\OfficeAI\www`.
- Переизвлечение происходит только если изменён `_resourceVersion` **или** размер уже распакованного файла разошёлся с `.manifest`.
- Из-за этого правки JS месяцами не доезжали. **При любом изменении ресурсов поднимать `_resourceVersion`** и проверять `%LOCALAPPDATA%\OfficeAI\www\js\...`.
- Новые JS/CSS нужно добавить: в `Resources.resx` (ключ + ResXFileRef), в словарь `ResourceExtractor`, в `.vbproj` (`None Include`), и поднять версию.
- Для проверки содержимого сборки: `ResourceManager("ShareRibbon.Resources", asm).GetString("agent_card")`.

### 2. VSTO staging

- Debug-аддин зарегистрирован как `file:///D:/git/office-ai-agent/PowerPointAi/bin/Debug/PowerPointAi.vsto|vstolocal`.
- Реально загружается из `%LOCALAPPDATA%\assembly\dl3\...`; staging обновляется при сборке (проверять размеры/время), но **Office нужно перезапустить**, иначе работает старый код.
- `PowerPointAi\bin\Debug\ShareRibbon.dll` есть; при сомнениях чистить `%LOCALAPPDATA%\assembly\dl3` не нужно, сверить размеры с bin.

### 3. ID инструментов глобальны

- Инструменты грузятся из `ShareRibbon/Tools/**/*.json` (wildcard embedded).
- Одинаковый `id` в Word и PPT сливается: параметры первого остаются. Так было с `ReplaceText`.
- Новые PPT-инструменты называть уникально (`ReplaceSlideText`, `UpdateChartData`), не переиспользовать word/excel-имена.

### 4. Skill allow-list скрывает инструменты

- `GetVisibleTools` режет список по `allowed-tools` выбранного Skill.
- Если инструмента нет в `ShareRibbon/Skills/powerpoint-*/SKILL.md`, модель его не увидит и будет выдумывать/VBA.
- Добавляя инструмент, обновить оба PPT-skill.

### 5. Семантика slideIndex разная

- Большинство legacy-команд: 0-based, `-1` текущий.
- `DuplicateSlide`, `ReplaceSlideText`, `UpdateChartData`: **1-based**, `0/-1` текущий.
- Наблюдение (`ResolvePowerPointTargetSlideIndex`) учитывает это; при добавлении команд поддерживать консистентность и обновлять резолвер.

### 6. Legacy и Agent — два контура

- Agent: `ShareRibbon/Agent/PromptManager.vb` + `LoopEngine.Planning.vb` + `ToolRegistry`.
- Legacy: `ShareRibbon/Config/PromptManager.vb` + `PowerPointAi/PowerPointJsonCommandSchema.vb` + карточка «Выполнить план» из `code-handler.js`.
- Пользователь может иметь **сохранённый кастомный prompt**; `IsLegacyBuiltInJsonSchemaConstraint` признаёт старым только 22/16-командные/VBA-варианты. При изменении поведения обновлять оба контура и критерии устаревания.
- Последняя незавершённая правка: блок «Типовые задачи → команды» в промпты **не добавился** (replace не совпал из-за CRLF). Его нужно добавить в оба контура.

### 7. Наблюдаемость

- Ошибки хоста возвращаются как `ExceptionClassifier.CodeUnknown`; подробности кладутся в `_lastPptCommandError` и попадают в `userMessage/debugDetail/warnings`.
- Для диагностики смотреть JSONL по `toolId`, `parameters`, `observation.before/after.textHash`.

## Ключевые файлы

| Что | Где |
|---|---|
| Dispatch PPT-команд | `PowerPointAi/ChatControl.vb` (`ExecutePPTCommand`, `ExecuteJsonCommandWithToolResult`) |
| Замена текста/таблиц | `PowerPointAi/ChatControl.vb` (`ExecuteReplaceText`, `ReplaceTextInShapes`, `ReplaceTextInSingleShape`) |
| Диаграммы | `PowerPointAi/ChatControl.vb` (`ExecuteUpdateChartData`) |
| Дублирование | `PowerPointAi/ChatControl.vb` (`ExecuteDuplicateSlide`) |
| Наблюдение/снимок слайда | `PowerPointAi/ChatControl.vb` (`CapturePowerPointCommandSnapshot`, `ResolvePowerPointTargetSlideIndex`) |
| Валидация/видимость инструментов | `ShareRibbon/Agent/ToolRegistry.vb`, `ToolParameterValidator.vb` |
| Парсинг вызовов и план | `ShareRibbon/Agent/LoopEngine.Planning.vb` |
| Промпты агента | `ShareRibbon/Agent/PromptManager.vb` |
| Legacy-промпт | `ShareRibbon/Config/PromptManager.vb` |
| Legacy-схема | `PowerPointAi/PowerPointJsonCommandSchema.vb` |
| Skills | `ShareRibbon/Skills/powerpoint-deck-agent`, `powerpoint-professional-design-agent` |
| UI-карточки | `ShareRibbon/Resources/js/agent-card.js`, `code-handler.js` |
| Ресурсы UI | `ShareRibbon/Resources/ResourceExtractor.vb`, `ShareRibbon/My Project/Resources.resx` |
| Журнал | `ShareRibbon/Agent/AgentRunJournal.vb` |

## Сборка и проверка

```powershell
cmd /c build-code.bat
# при изменениях JS/CSS: поднять _resourceVersion, перезапустить PowerPoint
# при странном поведении: проверить %LOCALAPPDATA%\OfficeAI\www\js\<file> и dl3 staging
```

- Новые `.vb` добавлять в соответствующий `.vbproj`.
- Новые JSON-инструменты: `ShareRibbon/Tools/ppt/`, обязательно уникальный `id`, `appType:"ppt"`, `schema`, затем Skill allow-list.
- Проверка схем/парсера в этой сессии делалась ad-hoc через `csc.exe` и reflection; отдельного тест-проекта нет.
- Office COM снаружи не аттачится (ROT недоступен), поэтому runtime-диагностика — через JSONL.

## Инварианты

- Никакого второго Harness/Loop/Safety/AiGateway.
- Все COM-операции — в UI-потоке через существующий вход; освобождать COM-объекты.
- Не писать base64/скриншоты в Observation/History/логи.
- Не менять `.vdproj` без необходимости.
- Версии AssemblyInfo/.vbproj/.vdproj поднимать согласованно (`scripts/set-version.ps1`).
- Не добавлять VBA в типовые сценарии.
