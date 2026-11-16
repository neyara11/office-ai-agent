## Why

Плагин умеет работать с существующей презентацией, но типовые задачи решаются неправильно:

- «сделай копию/такой же слайд» уходит в `CreateSlides`, который строит собственную дизайн-систему и не совпадает с колодой;
- «добавь слайд на основе этого» не использует похожие слайды колоды;
- у плагина нет панелей **создания, просмотра и выбора макетов**;
- модель периодически пытается использовать `ExecuteVBA` (запрещён и отключён), вместо нативных операций.

Нужен слой шаблонов/макетов: каталог, извлечение макета из слайда/колоды, подбор похожего макета, клонирование и заполнение контентом, плюс UI-панели каталога и создания.

## What Changes

- **Контракт макета и каталог**: `SlideTemplateManifest`, `SlideSpec`, `TemplateCatalog`, файловое хранилище в `%LOCALAPPDATA%\OfficeAiAppData\templates`.
- **Извлечение макетов** из активной презентации: один слайд или вся колода; архетип, слоты, вместимость, fingerprint, theme tokens, превью.
- **Подбор и заполнение**: детерминированный селектор по блокам контента; клонирование слайда и заполнение title/body/bullets/table/chart/image; верификация и rollback.
- **Инструменты**: `ListSlideTemplates`, `ExtractTemplateFromSlide`, `ExtractTemplateSetFromDeck`, `AddSlideFromTemplate`, `AddSlideFromNearestSlide`, `BuildDeckFromTemplateSet`.
- **UI-панели** в WebView2: каталог макетов (drawer), создание макета (sheet), выбор при вставке, inline-карточка в чате; Ribbon-кнопки «Макеты» и «Сохранить как макет».
- **Промпты**: типовые задачи → нативные команды; `CreateSlides` только для колод с нуля/fallback; `ExecuteVBA` не используется.
- **Никаких новых Harness/Loop/Design System**: только существующие `ToolRegistry`/`LoopEngine` и clone-and-fill.

## Capabilities

### New Capabilities

- `slide-template-catalog`: контракт, хранение, список, импорт/экспорт макетов.
- `template-extraction`: построение манифеста из слайда/колоды, превью, fingerprint.
- `template-selection-fill`: подбор макета, clone-and-fill, верификация, rollback.
- `template-ui-panels`: каталог/создание/выбор макетов в WebView2 и Ribbon.

### Modified Capabilities

- `deck-design`: `CreateSlides` больше не используется для задач «на основе существующей колоды».
- `command-generation`: типовые задачи покрываются нативными инструментами без VBA.

## Impact

- `ShareRibbon/Agent/DeckTemplates/` — новые контракты и storage.
- `ShareRibbon/Agent/ToolRegistry.vb`, `LoopEngine*.vb`, `PromptManager.vb` — регистрация инструментов и правила выбора.
- `ShareRibbon/Skills/powerpoint-*/SKILL.md` — `allowed-tools`.
- `ShareRibbon/Resources/js|css` + `Resources.resx` + `ResourceExtractor.vb` + `_resourceVersion`.
- `PowerPointAi/Design/DeckTemplate/` — COM-извлечение, клонирование, заполнение, верификация.
- `PowerPointAi/ChatControl.vb` — dispatch новых команд и observation.
- `PowerPointAi/PowerPointJsonCommandSchema.vb`, `ShareRibbon/Config/PromptManager.vb` — legacy-схема и промпты.
- `ShareRibbon/ShareRibbon.vbproj`, `PowerPointAi/PowerPointAi.vbproj` — новые файлы и ресурсы.
