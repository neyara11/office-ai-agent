/**
 * templates-view.js
 * Каталог макетов слайдов: drawer, список, действия. Хост — PowerPoint; для Word/Excel
 * хост вернёт ошибку, и панель покажет понятное сообщение.
 */
(function (window, document) {
    'use strict';

    var state = {
        templates: [],
        open: false,
        loading: false,
        lastMessage: ''
    };

    function post(type, payload) {
        var message = payload || {};
        message.type = type;
        if (window.officeAi && typeof window.officeAi.post === 'function') {
            return window.officeAi.post(message);
        }
        showToast('Мост Office недоступен', true);
        return false;
    }

    function byId(id) {
        return document.getElementById(id);
    }

    function showToast(text, isError) {
        var toast = byId('templates-toast');
        if (!toast) return;
        toast.textContent = text || '';
        toast.classList.toggle('error', !!isError);
        toast.classList.add('show');
        window.clearTimeout(showToast._timer);
        showToast._timer = window.setTimeout(function () {
            toast.classList.remove('show');
        }, 3200);
    }

    function setOpen(open) {
        state.open = !!open;
        var drawer = byId('templates-drawer');
        var overlay = byId('templates-overlay');
        var button = byId('templates-button');
        if (drawer) drawer.classList.toggle('open', state.open);
        if (overlay) overlay.classList.toggle('open', state.open);
        if (button) button.classList.toggle('active', state.open);
        if (state.open) refresh();
    }

    function refresh() {
        if (state.loading) return;
        state.loading = true;
        renderStatus('Загружаю макеты…');
        post('templates:list', {});
    }

    function renderStatus(text) {
        var grid = byId('templates-grid');
        var empty = byId('templates-empty');
        if (grid) grid.innerHTML = '';
        if (empty) {
            empty.style.display = 'block';
            empty.textContent = text || '';
        }
    }

    function archetypeLabel(value) {
        var map = {
            cover: 'Титул',
            section: 'Раздел',
            bullets: 'Тезисы',
            'two-column': 'Две колонки',
            chart: 'Диаграмма',
            table: 'Таблица',
            kpi: 'KPI',
            'image-text': 'Изображение + текст',
            quote: 'Цитата',
            closing: 'Финальный',
            custom: 'Другой'
        };
        return map[value] || value || 'Другой';
    }

    function filterTemplates() {
        var query = (byId('templates-search') && byId('templates-search').value || '').trim().toLowerCase();
        var archetype = (byId('templates-archetype') && byId('templates-archetype').value) || '';
        return state.templates.filter(function (item) {
            if (archetype && (item.archetype || '') !== archetype) return false;
            if (!query) return true;
            var haystack = [item.name, item.id, item.setId, (item.tags || []).join(' ')].join(' ').toLowerCase();
            return haystack.indexOf(query) >= 0;
        });
    }

    function renderList() {
        state.loading = false;
        var grid = byId('templates-grid');
        var empty = byId('templates-empty');
        if (!grid || !empty) return;
        grid.innerHTML = '';

        var items = filterTemplates();
        if (!items.length) {
            empty.style.display = 'block';
            empty.textContent = state.templates.length
                ? 'Ничего не найдено по фильтру.'
                : 'Макетов пока нет. Сохраните слайд или колоду — они появятся здесь.';
            return;
        }

        empty.style.display = 'none';
        items.forEach(function (item) {
            var card = document.createElement('div');
            card.className = 'template-card';

            if (item.previewUrl) {
                var preview = document.createElement('img');
                preview.className = 'template-card-preview';
                preview.loading = 'lazy';
                preview.alt = item.name || item.id;
                preview.src = item.previewUrl;
                card.appendChild(preview);
            }

            var name = document.createElement('div');
            name.className = 'template-card-name';
            name.textContent = item.name || item.id;
            card.appendChild(name);

            var meta = document.createElement('div');
            meta.className = 'template-card-meta';
            var badge = document.createElement('span');
            badge.className = 'template-badge';
            badge.textContent = archetypeLabel(item.archetype);
            meta.appendChild(badge);
            if (item.hasPreview) {
                var previewBadge = document.createElement('span');
                previewBadge.className = 'template-badge';
                previewBadge.textContent = 'превью';
                meta.appendChild(previewBadge);
            }
            card.appendChild(meta);

            var details = document.createElement('div');
            details.className = 'template-card-meta';
            details.textContent = 'Слоты: ' + ((item.slotRoles || []).join(', ') || '—');
            card.appendChild(details);

            var actions = document.createElement('div');
            actions.className = 'template-card-actions';

            var insert = document.createElement('button');
            insert.className = 'templates-btn primary';
            insert.textContent = 'Вставить';
            insert.addEventListener('click', function () {
                if (window.TemplateEditor) window.TemplateEditor.openInsert(item.id, item.slotRoles || []);
            });
            actions.appendChild(insert);

            var open = document.createElement('button');
            open.className = 'templates-btn';
            open.textContent = 'Открыть';
            open.addEventListener('click', function () {
                post('templates:open', { templateId: item.id });
            });
            actions.appendChild(open);

            var editButton = document.createElement('button');
            editButton.className = 'templates-btn';
            editButton.textContent = 'Изменить';
            editButton.addEventListener('click', function () {
                if (window.TemplateEditor) window.TemplateEditor.openEdit(item);
            });
            actions.appendChild(editButton);

            var duplicateButton = document.createElement('button');
            duplicateButton.className = 'templates-btn';
            duplicateButton.textContent = 'Дублировать';
            duplicateButton.addEventListener('click', function () {
                post('templates:duplicate', { templateId: item.id });
            });
            actions.appendChild(duplicateButton);

            var exportButton = document.createElement('button');
            exportButton.className = 'templates-btn';
            exportButton.textContent = 'Экспорт';
            exportButton.addEventListener('click', function () {
                post('templates:export', { templateId: item.id });
            });
            actions.appendChild(exportButton);

            var remove = document.createElement('button');
            remove.className = 'templates-btn danger';
            remove.textContent = 'Удалить';
            remove.addEventListener('click', function () {
                if (window.confirm('Удалить макет «' + (item.name || item.id) + '»?')) {
                    post('templates:delete', { templateId: item.id });
                }
            });
            actions.appendChild(remove);

            card.appendChild(actions);
            grid.appendChild(card);
        });
    }

    function onHostResponse(response) {
        if (!response || typeof response !== 'object') return;
        var action = response.action || '';
        var ok = !!response.success;

        if (action === 'list') {
            state.loading = false;
            if (ok) {
                state.templates = response.data || [];
                renderList();
            } else {
                renderStatus(response.message || 'Не удалось загрузить макеты');
            }
            return;
        }

        if (!ok) {
            state.loading = false;
            showToast(response.message || 'Операция не выполнена', true);
            if (action === 'insert' || action === 'saveFromSlide' || action === 'saveDeck') {
                if (action === 'insert' && window.TemplateEditor) window.TemplateEditor.close();
            }
            return;
        }

        showToast(response.message || 'Готово');
        if (action === 'delete' || action === 'saveFromSlide' || action === 'saveDeck' || action === 'import' || action === 'duplicate' || action === 'update') {
            refresh();
        }
        if (action === 'insert') {
            if (window.TemplateEditor) window.TemplateEditor.close();
            refresh();
        }
        if (action === 'preview' && window.TemplateEditor) {
            window.TemplateEditor.showPreview(response.data || {});
        }
    }

    function saveCurrentSlide() {
        if (window.TemplateEditor) window.TemplateEditor.openCreate('slide');
    }

    function saveDeck() {
        if (window.TemplateEditor) window.TemplateEditor.openCreate('deck');
    }

    function bind() {
        var overlay = byId('templates-overlay');
        if (overlay) overlay.addEventListener('click', function () { setOpen(false); });

        var close = byId('templates-close');
        if (close) close.addEventListener('click', function () { setOpen(false); });

        var search = byId('templates-search');
        if (search) search.addEventListener('input', renderList);

        var archetype = byId('templates-archetype');
        if (archetype) archetype.addEventListener('change', renderList);

        var refreshBtn = byId('templates-refresh');
        if (refreshBtn) refreshBtn.addEventListener('click', refresh);

        var saveSlide = byId('templates-save-slide');
        if (saveSlide) saveSlide.addEventListener('click', saveCurrentSlide);

        var saveDeckBtn = byId('templates-save-deck');
        if (saveDeckBtn) saveDeckBtn.addEventListener('click', saveDeck);

        var importBtn = byId('templates-import');
        if (importBtn) importBtn.addEventListener('click', function () {
            post('templates:import', {});
        });

        var openBtn = byId('templates-button');
        if (openBtn) openBtn.addEventListener('click', function () { setOpen(!state.open); });

        document.addEventListener('keydown', function (event) {
            if (event.key === 'Escape' && state.open) setOpen(false);
        });
    }

    function showSuggestions(items) {
        if (!items || !items.length) return;
        var container = byId('chat-container');
        if (!container) return;

        var existing = container.querySelector('.template-suggestions');
        if (existing) existing.parentNode.removeChild(existing);

        var wrap = document.createElement('div');
        wrap.className = 'template-suggestions';

        var header = document.createElement('div');
        header.className = 'template-suggestions-header';
        header.textContent = 'Подходящие макеты';
        wrap.appendChild(header);

        var list = document.createElement('div');
        list.className = 'template-suggestions-list';
        items.slice(0, 3).forEach(function (item) {
            var card = document.createElement('div');
            card.className = 'template-suggestion';

            var name = document.createElement('div');
            name.className = 'template-suggestion-name';
            name.textContent = item.name || item.id;
            card.appendChild(name);

            var meta = document.createElement('div');
            meta.className = 'template-suggestion-meta';
            meta.textContent = archetypeLabel(item.archetype);
            card.appendChild(meta);

            var actions = document.createElement('div');
            actions.className = 'template-suggestion-actions';
            var insert = document.createElement('button');
            insert.className = 'templates-btn primary';
            insert.textContent = 'Вставить';
            insert.addEventListener('click', function () {
                if (window.TemplateEditor) window.TemplateEditor.openInsert(item.id, item.slotRoles || []);
            });
            actions.appendChild(insert);
            card.appendChild(actions);

            list.appendChild(card);
        });
        wrap.appendChild(list);

        var footer = document.createElement('div');
        footer.className = 'template-suggestions-actions';
        var catalog = document.createElement('button');
        catalog.className = 'templates-btn';
        catalog.textContent = 'Показать каталог';
        catalog.addEventListener('click', function () { setOpen(true); });
        footer.appendChild(catalog);
        var other = document.createElement('button');
        other.className = 'templates-btn';
        other.textContent = 'Другой макет';
        other.addEventListener('click', function () { setOpen(true); });
        footer.appendChild(other);
        wrap.appendChild(footer);

        container.appendChild(wrap);
        try {
            wrap.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        } catch (e) {
        }
    }

    window.TemplatesView = {
        open: function () { setOpen(true); },
        close: function () { setOpen(false); },
        refresh: refresh,
        onHostResponse: onHostResponse,
        showToast: showToast,
        showSuggestions: showSuggestions
    };
    window.showTemplateSuggestions = showSuggestions;

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bind);
    } else {
        bind();
    }
})(window, document);
