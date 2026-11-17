## 1. P0 — сначала починить маршрутизацию типовых задач

- [x] 1.1 В `ShareRibbon/Agent/PromptManager.vb` (planning + react) и в `ShareRibbon/Config/PromptManager.vb` (PPT default) добавить блок «Типовые задачи → команды»:
  - «копия/такой же слайд, но на русском / с другими данными» → `DuplicateSlide` (`position: after|end`), затем `ReplaceSlideText` и/или `UpdateChartData`; `CreateSlides` не использовать;
  - «добавь последним» → `DuplicateSlide position=end`, правки копии через `slideIndex=-1`;
  - «переведи таблицу/текст» → `ReplaceSlideText` (обходит таблицы);
  - «новая колода с нуля» → `CreateSlides`; «вставить таблицу» → `InsertTable`; «удалить/переместить» → `DeleteSlide`/`MoveSlide`;
  - `ExecuteVBA` запрещён.
- [x] 1.2 То же добавить в `PowerPointAi/PowerPointJsonCommandSchema.vb` (legacy prompt) и в `SupportedCommands`/валидации, если появились новые команды.
- [x] 1.3 Усилить устаревание сохранённого legacy-промпта: для PPT считать legacy, если строка содержит `ExecuteVBA` **или** не содержит `DuplicateSlide`; проверить `IsLegacyBuiltInJsonSchemaConstraint`.
- [ ] 1.4 Проверка: запрос «сделай копию этого слайда, только на русском» и «добавь график на основе слайда 16 последним слайдом» не приводят к `CreateSlides`/`ExecuteVBA`; план — `DuplicateSlide` + правки. (Требует ручного прогона в PowerPoint + JSONL.)
- [x] 1.5 При изменении JS/CSS поднять `_resourceVersion` в `ResourceExtractor.vb`. (P0 JS/CSS не менял.)

## 2. P1 — контракт макета и каталог

- [x] 2.1 Добавить `ShareRibbon/Agent/DeckTemplates/`:
  - `SlideTemplateManifest`: `id`, `name`, `setId`, `archetype`, `slots[]`, `fingerprint`, `themeTokens`, `source`, `createdAt`, `previewPath`;
  - `TemplateSlot`: `role` (title/subtitle/body/bullets/table/chart/image/caption/footer), `shapeId`, `capacityChars`, `itemCount`, `contentKind`, `optional`;
  - `SlideSpec`/`SlideBlockSpec`: запрос контента от модели (title, bullets, table, chart, image, kpi);
  - `TemplateCatalog`: list/get/find by archetype/tags, версия схемы.
- [x] 2.2 `TemplateStorage`:
  - путь `%LOCALAPPDATA%\OfficeAiAppData\templates\<set>\`
  - `template.pptx` (один слайд-эталон), `manifest.json`, `preview.png` (опционально);
  - операции save/load/delete/import/export; атомарная запись; без секретов;
  - добавить файлы в `.vbproj`.
- [x] 2.3 Инструмент `ListSlideTemplates` (read-only): компактный каталог для модели/UI (id, name, archetype, slots summary), без превью в base64.
- [x] 2.4 Acceptance: каталог читается/пишется, переживает перезапуск Office, не ломает существующие AppData-пути. Проверено smoke-тестом storage: save → list → export → delete → import → load; служебные папки `.staging/.import/.trash` очищаются.

## 3. P2 — извлечение макетов (PowerPointAi)

- [x] 3.1 Создать `PowerPointAi/Design/DeckTemplate/DeckTemplateExtractor.vb`:
  - классификация `archetype` по составу фигур/placeholder (`cover`, `section`, `bullets`, `two-column`, `chart`, `table`, `kpi`, `image-text`, `quote`, `closing`);
  - `slots` по placeholder type и ролям фигур; `capacityChars` по геометрии и шрифту; `itemCount` для списков/таблиц;
  - `fingerprint` (типы + относительная геометрия фигур) для поиска похожих;
  - `themeTokens` из `SlideMaster.Theme` (accent/text/fonts) — только как метаданные.
- [x] 3.2 `ExtractTemplateFromSlide(slideIndex, name)`:
  - копирует слайд в отдельную презентацию/файл `template.pptx` (сохранить оформление, без связи с исходной колодой);
  - строит manifest; превью через `Slide.Export` во временный PNG, кладёт в хранилище;
  - не переносит заметки/анимации, если они не часть дизайна (решить и задокументировать).
- [x] 3.3 `ExtractTemplateSetFromDeck`: все слайды → набор макетов.
- [x] 3.4 Инструменты `ExtractTemplateFromSlide`, `ExtractTemplateSetFromDeck`; добавить в оба PPT-skill `allowed-tools`.
- [x] 3.5 Верификация: сохранённый `template.pptx` открывается, количество фигур совпадает, превью создано, COM освобождён.
- [x] 3.6 Acceptance: сохранить слайд 16 как макет, увидеть его в `ListSlideTemplates`, перезапустить Office — макет на месте. Проверено на реальном слайде 15 (chart): `template.pptx` + `preview.png` созданы, `archetype=chart`, слоты/тема/fingerprint заполнены; набор из 3 слайдов дал `image-text`, `table`, `kpi`; temp/snapshot и служебные папки удаляются.

## 4. P3 — подбор и заполнение

- [x] 4.1 `SlideTemplateSelector`:
  - score по покрытию блоков `SlideSpec`, вместимости, числу элементов, архетипу;
  - источники: явный каталог → implicit-каталог текущей колоды → `CreateSlides`/вопрос;
  - детерминированность и объяснимый выбор (`matchedBy`, `score`).
- [x] 4.2 `DeckTemplateRenderer` (PowerPointAi):
  - открыть `template.pptx`, скопировать слайд в активную колоду (`position: after|end`), закрыть источник;
  - заполнение: title/body/bullets → `ReplaceSlideText`; таблица → заполнение ячеек; график → `UpdateChartData`; изображение → вставка в picture-placeholder (позже);
  - удалить пустые необязательные слоты/плейсхолдеры, не ломая декор;
  - верификация: текст помещается, число строк/категорий совпало, нет пустых обязательных слотов;
  - rollback клона по маркеру при ошибке.
- [x] 4.3 Инструменты `AddSlideFromTemplate`, `AddSlideFromNearestSlide`.
- [x] 4.4 Observation: `templateId`, `sourceSlide`, `slotsFilled`, `warnings`, diff слайдов.
- [x] 4.5 Acceptance: «добавь слайд про X на основе похожего» клонирует ближайший макет и заполняет без переполнений; при отсутствии каталога — ближайший слайд текущей колоды. Проверено через COM: selector выбрал slide-15 (score=23) под chart-spec; `AddSlideFromTemplate` вставил слайд с заголовком и 7 значениями диаграммы; `InsertFromSlide` клонировал слайд 15; при ошибке клон удаляется по SlideID.

## 5. P4 — сборка колоды

- [x] 5.1 `BuildDeckFromTemplateSet(outline)`:
  - outline (title + blocks) → селектор → клоны → заполнение;
  - порядок, титул, финальный слайд; лимит слайдов.
- [x] 5.2 Preview: список выбранных макетов и outline до записи.
- [x] 5.3 Rollback всей пачки при ошибке (по маркерам).
- [x] 5.4 Acceptance: «сделай другую презентацию на основе этой» даёт колоду в стиле исходной, без `CreateSlides`. Проверено через COM: preview вернул план из 2 макетов без записи; `target=current` добавил 2 слайда; `target=new` создал презентацию из 2 слайдов (стартовый пустой удалён); при ошибке шага renderer откатывает всю пачку.

## 6. P5 — UI-панели (WebView2)

- [x] 6.1 UI-кит на существующих токенах: drawer/sheet, карточка макета, галерея, stepper, empty state, toast, confirm. Новые файлы: `Resources/js/templates-view.js`, `template-editor.js`, `Resources/css/templates.css`.
- [x] 6.2 Каталог: открыть из Ribbon «Макеты» и из чата; поиск/теги/архетип; карточки с превью; действия вставить/изменить/дублировать/удалить; импорт/экспорт набора. Сделано: кнопка «Макеты» в панели чата и Ribbon (видна только в PowerPoint), drawer, поиск/архетип, карточки с превью-картинкой, вставить/открыть/удалить/экспорт, импорт zip. Дублирование и редактирование (имя/архетип/теги) добавлены.
- [x] 6.3 Создание: «Сохранить как макет» из Ribbon/контекста; форма имени/архетипа, карта слотов, переопределение роли (декор/слот); сохранить слайд или всю колоду.
- [x] 6.4 Вставка: выбор макета → форма контента (title/bullets/table/chart) → «после текущего/в конец» → предпросмотр → применить.
- [x] 6.5 Inline-карточка в чате: 2–3 предложенных макета, «Вставить», «Другой макет», «Показать каталог». Карточка появляется после вызова `ListSlideTemplates` (host перехватывает ToolResult → `window.showTemplateSuggestions`), максимум 3 варианта; проверено в Playwright.
- [x] 6.6 Bridge: сообщения `templates:list`, `templates:saveFromSlide`, `templates:saveDeck`, `templates:insert`, `templates:preview`, `templates:delete`; обработчики в `BaseChatControl`/`ChatControl`; длинные операции — async, COM — на UI-потоке.
- [x] 6.7 Регистрация ресурсов: ResXFileRef, `ResourceExtractor`, `.vbproj`, `_resourceVersion` (поднят до `2026.11.17.5`; превью отдаются как `officeai.local/templates-preview/<id>.png`).
- [x] 6.8 Acceptance: в узкой панели 360–480 px каталог и создание макета работают без обрезания; вставка из UI даёт тот же результат, что и через чат. Проверено в Playwright на 380 px с mock-мостом: drawer открывается, карточки и формы рендерятся, insert-запрос уходит с title/chart и получает ответ; в PowerPoint нужен финальный прогон после перезапуска.

## 7. P6 — чистка и защита от регрессий

- [x] 7.1 Убрать `CreateSlides` из сценариев «на основе существующей колоды» на уровне промптов и (при необходимости) skill-правил.
- [x] 7.2 Проверить, что `ExecuteVBA` не появляется ни в одном prompt/плане; legacy `SupportedCommands` без VBA.
- [x] 7.3 Убрать дублирующиеся подсказки «выполнить план» в agent-контуре: backend выполняет сразу, карточка — только «Остановить».
- [x] 7.4 Добавить smoke-сценарии (ручные/скриптом):
  - сохранить макет из слайда 16 → список → вставить с новым текстом;
  - «копия слайда на русском» для таблицы: `DuplicateSlide` + `ReplaceSlideText`, таблица переведена;
  - «график на основе слайда»: `DuplicateSlide` + `UpdateChartData`, тип/цвета сохранены;
  - «другая презентация на основе этой»: `BuildDeckFromTemplateSet`.
- [x] 7.5 Проверка логов: нет base64/скриншотов, есть `templateId`/`slotsFilled`/`warnings`. В `agent-runs-20261117.jsonl` (45 записей) нет `base64`/`data:image`; шаблонные инструменты пишут templateId/slotsFilled.
- [x] 7.6 `build-code.bat` зелёный; новые `.vb`/ресурсы в `.vbproj`; `_resourceVersion` поднят.

## Definition of Done

- Пользователь может: сохранить слайд/колоду как макет, увидеть каталог, выбрать макет, вставить слайд с новым контентом; агент сам предлагает похожий макет.
- Типовые задачи не используют VBA и не создают произвольный дизайн через `CreateSlides`.
- Текст/таблицы/графики заполняются без переполнений; при ошибке — rollback и понятное сообщение.
- Каталог и настройки переживают перезапуск Office; UI не ломается в узкой панели.

## Не делать

- Не создавать второй Harness/Loop/Safety/AiGateway/дизайн-систему.
- Не хранить base64/скриншоты в Observation/History/логах; превью — только файлом в хранилище макетов.
- Не переиспользовать ID инструментов между Word/Excel/PowerPoint.
- Не добавлять новые инструменты без записи в `allowed-tools` PPT-skills.
- Не трогать `.vdproj` без необходимости; версии поднимать согласованно.
