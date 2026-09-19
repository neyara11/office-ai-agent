# План русификации Office AI Agent (форк `neyara11`)

| Параметр | Значение |
|---|---|
| Репозиторий | `https://github.com/neyara11/office-ai-agent` (fork of `it235/office-ai-agent`) |
| База | `1fd78d3` («merge: ai native cleanup and office compatibility fixes») |
| Лицензия | Apache-2.0 |
| Целевой контур | закрытый, без доступа в интернет |
| Статус | план, к реализации не приступали |

## 1. Цель

1. **Русскоязычный интерфейс** — риббон, диалоги, чат-панель, справка.
2. **Понимание русских команд** — распознавание намерений и маршрутизация инструментов на русском языке.
3. **Ответы на русском** — модель отвечает по-русски независимо от языка входных данных и внутренних промптов.
4. Сохранение работы в закрытом контуре (offline), включая локальную справку.

## 2. Критерии приёмки

- [ ] 100% видимых элементов UI на русском (риббон, все диалоги, чат-панель, справка).
- [ ] Набор из ≥30 русских команд выполняется без подсказок (intent + tool routing).
- [ ] 100% ответов на русском на тест-наборе, включая случаи, когда документ/контекст на другом языке.
- [ ] Китайские и английские команды продолжают работать (обратная совместимость матчинга).
- [ ] Справка открывается локально, без обращения к `officeso.cn`.
- [ ] При отключённой сети всё ядро работает против внутреннего OpenAI-совместимого endpoint.
- [ ] `platform`-идентификаторы, ID инструментов и ID контролов риббона не изменены.

## 3. Ключевая проблема: языковой дрейф модели

Слабые модели склонны отвечать на языке промпта и/или команды. Оставлять китайские промпты нельзя — при русском вводе ответ может уйти в китайский. Поэтому:

1. Все системные промпты переводятся на русский.
2. В **каждый** слой промптов добавляется «языковой контракт»:

   > Отвечай только на русском языке. Не переключай язык, даже если входные данные, документ, имена файлов или предыдущие сообщения на другом языке. Цитаты и код сохраняй как есть.

3. Описания инструментов (tools) и навыков (skills) переводятся на русский — они попадают в контекст модели и влияют на язык.
4. Опциональный guardrail (P4): пост-проверка ответа на преобладание кириллицы; при нарушении — повторный запрос с усиленной инструкцией.

## 4. Инвентарь правок

### 4.1 Промпты и описания инструментов (P0 — критично)

Файлы (копируются в вывод как `Content`, правятся без перекомпиляции DLL):

- `ShareRibbon/Prompts/system-base.json` — базовые роль/ограничения (Layer 1).
- `ShareRibbon/Prompts/word-context.json`, `excel-context.json`, `ppt-context.json`.
- `ShareRibbon/Prompts/planning-strategy.json`, `react-strategy.json`, `reflection-strategy.json`.
- `ShareRibbon/Prompts/Scenarios/academic-paper.json`, `business-report.json`, `official-document.json`.
- `ShareRibbon/Tools/word/*.json`, `Tools/excel/*.json`, `Tools/ppt/*.json` (~50 файлов, поля `description`).
- `ShareRibbon/Skills/*.json` и `Skills/**/SKILL.md` (5 навыков).

Inline-промпты в коде (требуют правки `.vb`):

- `ShareRibbon/Agent/PromptManager.vb` — сборка 6-слойного системного промпта, блоки зашиты в код (см. `:44-70`, `:66-70` «不可覆盖的执行协议», «通用约束»).
- `ShareRibbon/Controls/Services/HttpStreamService.vb:1184` — системное сообщение ReAct/tool-loop.
- `ShareRibbon/Services/Memory/LlmMemoryExtractor.vb:86` — извлечение долговременной памяти (иначе память будет на китайском).
- `ShareRibbon/Services/Proofread/ProofreadPromptBuilder.vb`, `Services/Reformat/SemanticPromptBuilder.vb`, `Services/Reformat/*`.
- `ShareRibbon/Config/PromptManager.vb` (~330 строк с китайским) — библиотека внутренних промптов.
- `ShareRibbon/Controls/Services/IntentRecognitionService.vb` (~398).
- `ShareRibbon/Translate/*` — доменные шаблоны перевода.

Действие: перевод на русский + языковой контракт. Имена инструментов и их параметры (`InsertText`, `FormatText`, …) не переводим — только `description`.

Риск: перевод описаний может снизить качество tool-selection. Митигация: оставлять англоязычные имена/схемы, при деградации — двуязычные `description` («RU — … / EN — …»).

### 4.2 Понимание русских команд (P0)

В коде **335 мест** с матчингом по китайским строкам (`Contains("…")`, `Equals`, `Case`). Их нельзя заменять — только дополнять русскими синонимами.

Ключевые точки:

- `WordAi/ChatControl.vb:1386-1401` — определение намерения форматирования/нумерации.
- `ShareRibbon/Controls/Services/IntentRecognitionService.vb`.
- `ShareRibbon/Services/Reformat/ReformatIntentRecognizer.vb`.
- `ShareRibbon/Agent/Execution/SafetyChecker.vb`, `ShareRibbon/Loop/Checkers/PreSendChecker.vb`, `Validators/PostFlushValidator.vb`.
- `WordAi/ChatControl.vb:2169,2765` — определение стилей заголовков (`标题` / `heading`).

Домены, для которых нужны русские синонимы (минимальный набор):

| Домен | Примеры русских ключей |
|---|---|
| Форматирование | отформатируй, формат, оформление, стиль, единый стиль |
| Нумерация/заголовки | нумерация, пронумеруй, заголовок, подзаголовок, уровни, оглавление |
| Правка текста | исправь, перепиши, перефразируй, сократи, расширь, дополни, продолжение |
| Орфография/пунктуация | проверь орфографию, опечатки, пунктуация, вычитай |
| Перевод | переведи, перевод, на русский/английский |
| Excel | формула, диаграмма, график, сводная, фильтр, сортировка, анализ, дубликаты |
| PowerPoint | слайд, презентация, макет, тема, заметки докладчика, анимация |
| Word-структура | таблица, содержание, сноска, комментарий, колонтитул |

### 4.3 UI: риббон и WinForms (P1)

`ShareRibbon/Ribbon/BaseOfficeRibbon.vb` + `.Designer.vb`; `WordAi|ExcelAi|PowerPointAi/Ribbon1.Designer.vb`; диалоги: `ConfigApiForm.vb`, `ConfigPromptForm.vb`, `MCPConfigForm.vb`, `StdioConfigForm.vb`, `ImportConfigForm.vb`, `AboutForm.vb`, `Translate*Form.vb`, `ReformatTemplate*Form.vb`, `SkillsConfigForm.vb`, `CommandPreviewForm.vb`, `JsonPreviewDialog.vb`, `BatchDataGenerationForm.vb`, `Spotlight.vb`, `GlobalStatusStrip*.vb`.

Особенности:
- Формы **не** `Localizable` — подписи лежат в `.Designer.vb`; правки синхронизировать с `.resx` (`ChatControl.resx`, `Ribbon1.resx`), иначе Visual Studio перезапишет.
- Русский текст длиннее китайского → править `Size`/`AutoSize`/`Width` контролов и колонок, иначе обрезка.

### 4.4 Чат-панель WebView2 (P1)

- `ShareRibbon/Resources/chat-template-refactored.html` (~110 строк китайского; хранится как строковый ресурс `chat_template_refactored` в `ShareRibbon/My Project/Resources.resx:174`, отдаётся через виртуальный хост `officeai.local`).
- `ShareRibbon/Resources/js/*.js` — 16 модулей: `core.html`-строки, `chat-manager`, `message-sender`, `code-handler`, `proofread-ui`, `autocomplete`, `reformat-*`, `revision-manager`, `agent-card`, `agent-protocol`, `history-manager`, `markdown-renderer`, `mcp-manager`.
- Приветственные подсказки и ссылки на `cloud.siliconflow.cn`, `jsj.top`, `officeso.cn` — заменить на нейтральные/локальные.

### 4.5 Справка локально (P1)

- Сейчас: `BaseOfficeRibbon.vb:114-134` → `Process.Start("https://www.officeso.cn/study/{word|excel|ppt}")`; в репо только иконка `Resources/help.png` (`Resources.resx:229`).
- План: добавить `ShareRibbon/Resources/help/ru/*.html` (Word/Excel/PowerPoint + общая), показывать в WebView2-панели или открывать локальный файл; включить в `Content`/установщик `OfficeAgent/OfficeAgent.vdproj`; ссылки `officeso.cn` из шаблона чата убрать.

### 4.6 Внешние панели и ссылки (P3, offline)

- Панели DeepSeek / Doubao (`BaseDeepseekChat.vb:547` → `chat.deepseek.com`, `BaseDoubaoChat.vb:32` → `doubao.com`) — скрыть/отключить кнопки для закрытого контура.
- Веб-краулер (`WebDataCapturePane.vb`, `BaseDataCapturePane.vb`) — зависит от интернета, скрыть или оставить с пометкой.
- `AboutForm.vb` (bilibili/gitee/github) — заменить на локальные/нейтральные.

## 5. Что НЕ трогаем

- `platform` в `PresetProviders.vb` — это ключ хранилища API-ключей (`ShareRibbon/Config/SecureConfigIntegration.vb:28`, `SecureConfigManager.SaveApiKey(item.key, item.platform)`).
- ID инструментов и команды (`InsertText`, `FormatText`, `inserttext`, …) и схемы параметров.
- ID контролов риббона и `.Designer.vb`-имена.
- Имена Excel-функций (`ALLM`, `CLLM`) и структуру БД.

## 6. Этапы

| Этап | Приоритет | Содержание | Результат |
|---|---|---|---|
| P0 | критично | Языковой контракт + перевод системных промптов, описаний tools/skills, inline-промптов; русские синонимы в матчинг намерений | Понимание русских команд и русские ответы |
| P1 | высокий | UI риббон/диалоги, чат-панель, локальная справка | Русский интерфейс |
| P2 | средний | Layout-фиксы, чистка внешних ссылок | Нет обрезки, нет внешних переходов |
| P3 | средний | Offline: скрыть DeepSeek/Doubao/краулер, проверить сборку/подпись | Готовность к закрытому контуру |
| P4 | опц. | Guardrail языка ответа, двуязычные описания tools, автотесты | Устойчивость к дрейфу |

## 7. QA и тестирование

- Скрипты репо: `scripts/run-golden-l0.ps1`, `scripts/smoke-*.ps1`, `scripts/run-code-checks.ps1`.
- Ручные сценарии: контрольный список ≥30 русских команд по доменам из §4.2, с проверкой языка ответа.
- Регресс: те же действия китайскими и английскими командами — должны по-прежнему срабатывать.
- Offline-тест: отключить сеть, оставить только внутренний endpoint; проверить чат, агент, память, справку.
- Метрики: доля ответов на русском = 100%; доля успешных RU-команд ≥ целевой (задать после первого прогона).

## 8. Риски и митигации

| Риск | Митигация |
|---|---|
| Дрейф языка ответа у слабых моделей | Языковой контракт во всех слоях + пост-проверка и retry |
| Деградация tool-selection после перевода описаний | Оставить EN-имена/схемы; при необходимости двуязычные описания |
| Обрезка UI из-за длины русского текста | Ревизия размеров контролов/колонок, `AutoEllipsis`, тултипы |
| `.Designer.vb` перезаписывается дизайнером VS | Править вместе с `.resx`, не открывать дизайнер без необходимости |
| Конфликты при обновлениях upstream | Минимизировать правки, вести список изменённых файлов, отдельная ветка |
| Потеря функциональности из-за перевода логических строк | Не заменять логические ключи, только добавлять русские |
| Кодировка | Все файлы UTF-8 (с BOM для `.vb`), проверять после массовых правок |

## 9. Инструменты автоматизации

1. Скрипт извлечения CJK-строк из `.vb`/`.js`/`.html` в CSV (строки, файл, номер) для перевода.
2. Скрипт обратной подстановки с чёрным списком logic-строк (§4.2) и защищённых идентификаторов (§5).
3. Проверка, что не тронуты `platform`, ID инструментов, ID контролов.
4. Журнал изменённых файлов для переноса на новые версии upstream.

## 10. Артефакты

- Ветка: `ru-localization` (при активной работе), база `1fd78d3`.
- Коммиты по этапам P0 → P3 (отдельно код и контент).
- Настоящий документ: `docs/ru-localization-plan.md`.

## 11. Сборка и установка в закрытом контуре (напоминание)

- MSI собирается на подключённой машине (Visual Studio + VSTO workload + *Microsoft Visual Studio Installer Projects*), NuGet — через локальный фид/предварительный restore.
- Офлайн-предпосылки: .NET Framework 4.7.2, VSTO Runtime (`vstor_redist.exe`), WebView2 Runtime (Fixed Version/offline), Office 2016+ или WPS.
- Подпись VSTO-манифестов своим сертификатом (`build/SignArtifacts.ps1`, `OFFICE_AI_SIGN_PFX` / `_CERT_THUMBPRINT`) + Trusted Publishers через GPO; XLL — подпись или Trusted Location.
