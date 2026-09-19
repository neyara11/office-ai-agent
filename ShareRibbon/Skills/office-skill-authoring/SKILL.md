---
name: office-skill-authoring
description: Используйте при создании или проверке навыков Office AI. Обеспечивает структуру каталогов SKILL.md, краткие YAML-метаданные, прогрессивное раскрытие, references, scripts, область приложения и границы разрешённых инструментов.
application: Excel,Word,PowerPoint
tags: skill, skills, authoring, SKILL.md, progressive-disclosure, harness, agent-loop
allowed-tools: memory.search, memory.list_recent
intent_types: skill_authoring, architecture
---

# Создание навыков Office

Используйте этот навык при добавлении или проверке Skill для Office AI Agent.

## Требуемая структура

Каждый навык на основе каталога должен использовать:

```text
skill-name/
├── SKILL.md
├── references/        optional detailed docs, loaded only after the skill is selected
├── scripts/           optional executable helpers
└── assets/            optional templates or examples
```

`SKILL.md` должен начинаться с YAML front matter. Обязательные поля:

```yaml
---
name: skill-name
description: One concise sentence that explains when this skill should be used.
---
```

Рекомендуемые поля для этого проекта:

```yaml
application: Excel
tags: excel, formula, chart
allowed-tools: ApplyFormula, CreateChart
intent_types: data_analysis, formula
```

## Правила

1. Держите описание ориентированным на действие; оно используется для первичного выбора навыка.
2. Помещайте длинные примеры и предметные правила в `references/`, чтобы первый промпт оставался небольшим.
3. Ограничивайте `allowed-tools` минимальным полезным набором.
4. Чётко указывайте область приложения: Excel, Word, PowerPoint или общая.
5. Навык должен учить агента тому, как решать и выполнять; он не должен быть набором ключевых слов-маршрутов.
6. Скрипты должны находиться в `scripts/` и при необходимости объявлять аргументы в близлежащей документации.
