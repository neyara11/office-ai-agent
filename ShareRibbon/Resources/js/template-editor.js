/**
 * template-editor.js
 * Формы создания и вставки макета. Работает внутри drawer TemplatesView.
 */
(function (window, document) {
    'use strict';

    function byId(id) {
        return document.getElementById(id);
    }

    function post(type, payload) {
        var message = payload || {};
        message.type = type;
        if (window.officeAi && typeof window.officeAi.post === 'function') {
            return window.officeAi.post(message);
        }
        if (window.TemplatesView) window.TemplatesView.showToast('Мост Office недоступен', true);
        return false;
    }

    function toast(text, isError) {
        if (window.TemplatesView) window.TemplatesView.showToast(text, isError);
    }

    function field(labelText, control) {
        var label = document.createElement('label');
        label.appendChild(document.createTextNode(labelText));
        label.appendChild(control);
        return label;
    }

    function input(type, placeholder, value) {
        var element = document.createElement('input');
        element.type = type || 'text';
        if (placeholder) element.placeholder = placeholder;
        if (value !== undefined && value !== null) element.value = value;
        return element;
    }

    function textarea(placeholder, rows) {
        var element = document.createElement('textarea');
        element.placeholder = placeholder || '';
        element.rows = rows || 3;
        return element;
    }

    function select(options) {
        var element = document.createElement('select');
        options.forEach(function (item) {
            var option = document.createElement('option');
            option.value = item.value;
            option.textContent = item.label;
            element.appendChild(option);
        });
        return element;
    }

    function clearForm() {
        var host = byId('template-form');
        if (host) host.innerHTML = '';
    }

    function formHost() {
        var host = byId('template-form');
        if (host) host.style.display = 'block';
        return host;
    }

    function close() {
        clearForm();
        var host = byId('template-form');
        if (host) host.style.display = 'none';
    }

    function button(text, className, onClick) {
        var element = document.createElement('button');
        element.type = 'button';
        element.className = 'templates-btn ' + (className || '');
        element.textContent = text;
        element.addEventListener('click', onClick);
        return element;
    }

    function openCreate(mode) {
        var host = formHost();
        if (!host) return;
        host.innerHTML = '';

        var form = document.createElement('div');
        form.className = 'templates-form';

        var title = document.createElement('h3');
        title.textContent = mode === 'deck' ? 'Сохранить колоду как набор макетов' : 'Сохранить слайд как макет';
        form.appendChild(title);

        var nameInput = input('text', mode === 'deck' ? 'Имя набора' : 'Имя макета', '');
        form.appendChild(field('Имя', nameInput));

        var archetypeSelect = select([
            { value: '', label: 'Определить автоматически' },
            { value: 'cover', label: 'Титул' },
            { value: 'section', label: 'Раздел' },
            { value: 'bullets', label: 'Тезисы' },
            { value: 'two-column', label: 'Две колонки' },
            { value: 'chart', label: 'Диаграмма' },
            { value: 'table', label: 'Таблица' },
            { value: 'kpi', label: 'KPI' },
            { value: 'image-text', label: 'Изображение + текст' },
            { value: 'quote', label: 'Цитата' },
            { value: 'closing', label: 'Финальный' }
        ]);
        if (mode !== 'deck') {
            form.appendChild(field('Архетип', archetypeSelect));
        }

        var slideIndexInput = null;
        var maxSlidesInput = null;
        var skipEmptyInput = null;

        if (mode === 'deck') {
            var row = document.createElement('div');
            row.className = 'templates-form-row';
            maxSlidesInput = input('number', '20', '20');
            maxSlidesInput.min = '1';
            maxSlidesInput.max = '40';
            row.appendChild(field('Максимум слайдов', maxSlidesInput));

            var skipLabel = document.createElement('label');
            skipEmptyInput = input('checkbox', '', '');
            skipEmptyInput.checked = true;
            skipLabel.appendChild(skipEmptyInput);
            skipLabel.appendChild(document.createTextNode(' Пропускать пустые'));
            row.appendChild(skipLabel);
            form.appendChild(row);

            var deckHint = document.createElement('div');
            deckHint.className = 'templates-hint';
            deckHint.textContent = 'Каждый непустой слайд сохранится отдельным макетом с тем же setId.';
            form.appendChild(deckHint);
        } else {
            slideIndexInput = input('number', 'текущий', '');
            slideIndexInput.min = '1';
            form.appendChild(field('Номер слайда (1-based, пусто — текущий)', slideIndexInput));
        }

        var actions = document.createElement('div');
        actions.className = 'template-card-actions';
        actions.appendChild(button('Сохранить', 'primary', function () {
            if (mode === 'deck') {
                post('templates:saveDeck', {
                    setName: nameInput.value.trim(),
                    maxSlides: parseInt(maxSlidesInput.value, 10) || 20,
                    skipEmpty: !!skipEmptyInput.checked
                });
            } else {
                var index = parseInt(slideIndexInput.value, 10);
                post('templates:saveFromSlide', {
                    slideIndex: isNaN(index) ? 0 : index,
                    name: nameInput.value.trim(),
                    archetype: archetypeSelect.value
                });
            }
        }));
        actions.appendChild(button('Отмена', '', close));
        form.appendChild(actions);

        host.appendChild(form);
    }

    function parseTable(text) {
        var lines = (text || '').split(/\r?\n/).map(function (line) { return line.trim(); }).filter(Boolean);
        if (!lines.length) return null;
        var rows = lines.map(function (line) {
            return line.split(/\t|\|/).map(function (cell) { return cell.trim(); });
        });
        return { headers: rows[0], rows: rows.slice(1) };
    }

    function parseChart(valuesText, categoriesText) {
        var values = (valuesText || '').split(/[,;\s]+/).map(Number).filter(function (value) {
            return !isNaN(value);
        });
        if (!values.length) return null;
        var categories = (categoriesText || '').split(/[,;\n]+/).map(function (item) {
            return item.trim();
        }).filter(Boolean);
        return { values: values, categories: categories };
    }

    function openInsert(templateId) {
        var host = formHost();
        if (!host) return;
        host.innerHTML = '';

        var form = document.createElement('div');
        form.className = 'templates-form';

        var title = document.createElement('h3');
        title.textContent = 'Вставить макет: ' + templateId;
        form.appendChild(title);

        var titleInput = input('text', 'Заголовок слайда', '');
        form.appendChild(field('Заголовок', titleInput));

        var subtitleInput = input('text', 'Подзаголовок', '');
        form.appendChild(field('Подзаголовок', subtitleInput));

        var bulletsInput = textarea('Каждый пункт с новой строки', 3);
        form.appendChild(field('Тезисы', bulletsInput));

        var tableInput = textarea('Первая строка — заголовки; столбцы через Tab или |', 3);
        form.appendChild(field('Таблица (необязательно)', tableInput));

        var chartRow = document.createElement('div');
        chartRow.className = 'templates-form-row';
        var chartValuesInput = input('text', '30, 25, 10, 10, 5, 10, 10', '');
        var chartCategoriesInput = input('text', 'категории через запятую', '');
        chartRow.appendChild(field('Значения диаграммы (доли для процентов)', chartValuesInput));
        chartRow.appendChild(field('Категории (необязательно)', chartCategoriesInput));
        form.appendChild(chartRow);

        var positionSelect = select([
            { value: 'after', label: 'После текущего слайда' },
            { value: 'end', label: 'В конец колоды' }
        ]);
        form.appendChild(field('Позиция', positionSelect));

        var hint = document.createElement('div');
        hint.className = 'templates-hint';
        hint.textContent = 'Для процентных диаграмм значения — доли: 25% = 0.25. Если подписи не меняются, оставьте категории пустыми.';
        form.appendChild(hint);

        var actions = document.createElement('div');
        actions.className = 'template-card-actions';
        actions.appendChild(button('Вставить', 'primary', function () {
            var spec = {};
            var titleValue = titleInput.value.trim();
            var subtitleValue = subtitleInput.value.trim();
            var bullets = bulletsInput.value.split(/\r?\n/).map(function (line) {
                return line.trim();
            }).filter(Boolean);
            var table = parseTable(tableInput.value);
            var chart = parseChart(chartValuesInput.value, chartCategoriesInput.value);

            if (titleValue) spec.title = titleValue;
            if (subtitleValue) spec.subtitle = subtitleValue;
            if (bullets.length) spec.bullets = bullets;
            if (table) spec.table = table;
            if (chart) spec.chart = chart;

            if (!Object.keys(spec).length) {
                toast('Заполните хотя бы одно поле', true);
                return;
            }

            post('templates:insert', {
                templateId: templateId,
                spec: spec,
                position: positionSelect.value
            });
        }));
        actions.appendChild(button('Отмена', '', close));
        form.appendChild(actions);

        host.appendChild(form);
    }

    function showPreview(data) {
        var host = formHost();
        if (!host) return;
        host.innerHTML = '';

        var form = document.createElement('div');
        form.className = 'templates-form';

        var title = document.createElement('h3');
        title.textContent = 'Макет: ' + (data.name || data.id || '');
        form.appendChild(title);

        var details = document.createElement('div');
        details.className = 'templates-hint';
        details.textContent = 'Архетип: ' + (data.archetype || '—') +
            ' · Слоты: ' + ((data.slots || []).join(', ') || '—') +
            ' · Превью: ' + (data.hasPreview ? 'есть' : 'нет');
        form.appendChild(details);

        var actions = document.createElement('div');
        actions.className = 'template-card-actions';
        actions.appendChild(button('Вставить', 'primary', function () {
            openInsert(data.id);
        }));
        actions.appendChild(button('Открыть файл', '', function () {
            post('templates:open', { templateId: data.id });
        }));
        actions.appendChild(button('Закрыть', '', close));
        form.appendChild(actions);

        host.appendChild(form);
    }

    window.TemplateEditor = {
        openCreate: openCreate,
        openInsert: openInsert,
        showPreview: showPreview,
        close: close
    };
})(window, document);
