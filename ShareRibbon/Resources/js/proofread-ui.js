/**
 * proofread-ui.js - 校对UI交互模块
 * 提供WPS风格的校对体验：波浪线标注 + Hover Tooltip + 问题列表
 */

// ========== 校对专注模式 ==========

/**
 * 显示校对侧边面板
 */
function showProofreadSidePanel() {
    // 移除其他面板
    if (typeof hideTemplateEditorPane === 'function') {
        hideTemplateEditorPane();
    }

    // 检查是否已存在
    if (document.getElementById('proofread-side-panel')) return;

    // 标记校对模式激活（供 message-sender.js 检测）
    window.proofreadModeActive = true;

    // 创建侧边面板容器
    var panel = document.createElement('div');
    panel.id = 'proofread-side-panel';
    panel.className = 'proofread-side-panel';
    panel.innerHTML = '<div class="proofread-panel-header">' +
        '<div class="proofread-panel-title">' +
        '<span class="proofread-panel-icon">✓</span>' +
        '<span>AI-корректура</span>' +
        '</div>' +
        '<div class="proofread-panel-actions">' +
        '<button class="proofread-panel-action" title="Свернуть/развернуть" onclick="toggleProofreadPanelCollapse()">‹</button>' +
        '<button class="proofread-panel-action" title="Выйти из корректуры" onclick="proofreadExit()">×</button>' +
        '</div>' +
        '</div>' +
        '<div class="proofread-plan-summary" id="proofread-plan-summary"></div>' +
        '<div class="proofread-panel-content" id="proofread-panel-content"></div>';

    document.body.appendChild(panel);

    // 标记面板打开；不再强制挤压主体，避免在窄任务窗格中形成遮盖感。
    document.body.classList.add('proofread-panel-open');

    // 添加面板样式（如果尚未添加）
    injectProofreadStyles();
}

/**
 * 隐藏校对侧边面板
 */
function hideProofreadSidePanel() {
    var panel = document.getElementById('proofread-side-panel');
    if (panel) panel.remove();

    // 清除校对模式标记
    window.proofreadModeActive = false;
    window.proofreadSelectedText = '';
    window.proofreadIssueCount = 0;

    // 恢复主体布局
    document.body.classList.remove('proofread-panel-open');
}

/**
 * 显示校对列表
 */
function showProofreadList(html) {
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = html;
        bindProofreadListEvents();
    }
}

/**
 * 绑定校对列表事件
 */
function bindProofreadListEvents() {
    // 接受按钮
    var acceptBtns = document.querySelectorAll('.issue-btn.accept');
    acceptBtns.forEach(function(btn) {
        btn.addEventListener('click', function() {
            var issueId = this.getAttribute('data-issue-id');
            acceptProofreadIssue(issueId);
        });
    });
    
    // 忽略按钮
    var ignoreBtns = document.querySelectorAll('.issue-btn.ignore');
    ignoreBtns.forEach(function(btn) {
        btn.addEventListener('click', function() {
            var issueId = this.getAttribute('data-issue-id');
            ignoreProofreadIssue(issueId);
        });
    });
}

/**
 * 接受校对修正
 */
function acceptProofreadIssue(issueId) {
    var payload = {
        type: 'proofread',
        action: 'accept',
        issueId: issueId
    };
    sendProofreadAction(payload);
    
    // 移除该项
    var item = document.querySelector('.proofread-issue-item[data-issue-id="' + issueId + '"]');
    if (item) {
        item.style.opacity = '0.5';
        item.style.pointerEvents = 'none';
        setTimeout(function() { item.remove(); }, 300);
    }
}

/**
 * 忽略校对问题
 */
function ignoreProofreadIssue(issueId) {
    var payload = {
        type: 'proofread',
        action: 'ignore',
        issueId: issueId
    };
    sendProofreadAction(payload);
    
    // 标记该项为已忽略
    var item = document.querySelector('.proofread-issue-item[data-issue-id="' + issueId + '"]');
    if (item) {
        item.style.opacity = '0.4';
        var actions = item.querySelector('.issue-actions');
        if (actions) actions.remove();
    }
}

/**
 * 接受所有校对修正
 */
function proofreadAcceptAll() {
    var payload = {
        type: 'proofread',
        action: 'acceptAll',
        issueId: ''
    };
    sendProofreadAction(payload);
    
    // 隐藏列表
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = '<div class="proofread-success">' +
            '<span class="success-icon">🎉</span>' +
            '<span class="success-text">Все проблемы исправлены!</span>' +
            '</div>';
    }
}

/**
 * Режим выхода из корректуры
 */
function proofreadExit() {
    var payload = {
        type: 'proofread',
        action: 'exit',
        issueId: ''
    };
    sendProofreadAction(payload);

    // 清除校对模式标记和布局
    window.proofreadModeActive = false;
    window.proofreadSelectedText = '';
    window.proofreadIssueCount = 0;
    document.body.classList.remove('proofread-panel-open');

    hideProofreadSidePanel();
    hideProofreadModeIndicator();
}

/**
 * 发送校对操作到VB
 */
function sendProofreadAction(payload) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(payload);
    } else if (window.vsto && typeof window.vsto.postMessage === 'function') {
        window.vsto.postMessage(payload);
    } else {
        console.error('[ProofreadUI] 无法发送消息，WebView不可用');
    }
}

/**
 * 更新校对摘要
 */
function updateProofreadSummary(total, high, medium, low) {
    // 查找现有摘要元素
    var existingSummary = document.getElementById('proofread-summary');
    if (existingSummary) {
        existingSummary.innerHTML = 'Всего ' + total + ' проблем (' +
            '<span class="high">' + high + ' — к исправлению</span>，' +
            '<span class="medium">' + medium + ' — рекомендуется исправить</span>，' +
            '<span class="low">' + low + ' — необязательно</span>)';
    }
}

/**
 * 显示无问题消息
 */
function showProofreadNoIssues() {
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = '<div class="proofread-success">' +
            '<span class="success-icon">✅</span>' +
            '<span class="success-text">Проблем не найдено!</span>' +
            '<p class="success-hint">В документе нет изменений, требующих правки.</p>' +
            '<button class="proofread-btn secondary" onclick="proofreadExit()">Выйти из корректуры</button>' +
            '</div>';
    }
}

/**
 * 更新校对计划摘要
 */
function updateProofreadPlanSummary(planText) {
    var summary = document.getElementById('proofread-plan-summary');
    if (!summary) return;

    if (!planText) {
        summary.style.display = 'none';
        summary.innerHTML = '';
        return;
    }

    summary.style.display = 'block';
    summary.innerHTML = '<div class="proofread-plan-label">План этого шага</div>' +
        '<div class="proofread-plan-text">' + escapeProofreadHtml(planText).replace(/\n/g, '<br>') + '</div>';
}

/**
 * Свернуть/развернуть панель корректуры
 */
function toggleProofreadPanelCollapse() {
    var panel = document.getElementById('proofread-side-panel');
    if (!panel) return;

    panel.classList.toggle('collapsed');
    var btn = panel.querySelector('.proofread-panel-action');
    if (btn) {
        btn.textContent = panel.classList.contains('collapsed') ? '›' : '‹';
    }
}

/**
 * 显示校对解析失败消息
 */
function showProofreadParseError(payload) {
    var content = document.getElementById('proofread-panel-content');
    if (!content) return;

    payload = payload || {};
    var errorMessage = escapeProofreadHtml(payload.errorMessage || 'AI вернул неверный формат, не удалось построить список правок.');
    var rawPreview = escapeProofreadHtml(payload.rawPreview || '');
    var rawBlock = rawPreview
        ? '<pre class="proofread-error-raw">' + rawPreview + '</pre>'
        : '<p class="proofread-error-hint">Исходный ответ для показа отсутствует.</p>';

    content.innerHTML = '<div class="proofread-error">' +
        '<span class="error-icon">⚠️</span>' +
        '<span class="error-title">Не удалось разобрать результат корректуры</span>' +
        '<p class="error-message">' + errorMessage + '</p>' +
        '<p class="proofread-error-hint">Обычно это значит, что AI не вернул JSON по требованиям. Документ не изменён.</p>' +
        rawBlock +
        '<div class="proofread-list-actions">' +
        '<button class="proofread-btn secondary" onclick="proofreadExit()">Выйти из корректуры</button>' +
        '</div>' +
        '</div>';
}

function escapeProofreadHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * 显示全部修正完成消息
 */
function showProofreadAllCorrected() {
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = '<div class="proofread-success">' +
            '<span class="success-icon">🎉</span>' +
            '<span class="success-text">Все проблемы исправлены!</span>' +
            '<p class="success-hint">Документ полностью исправлен, панель корректуры можно закрыть.</p>' +
            '</div>';
    }
}

// ========== 注入校对样式 ==========

function injectProofreadStyles() {
    if (document.getElementById('proofread-styles')) return;
    
    var style = document.createElement('style');
    style.id = 'proofread-styles';
    style.textContent =
/* ========== 校对面板打开时主体布局避让 ========== */
'body.proofread-panel-open {' +
'    margin-right: 0 !important;' +
'    transition: margin-right 0.3s ease;' +
'}' +
'' +
/* ========== 校对面板样式 ========== */
'.proofread-side-panel {' +
'    position: fixed;' +
'    right: 8px;' +
'    top: 48px;' +
'    bottom: 8px;' +
'    width: min(360px, calc(100vw - 24px));' +
'    max-width: calc(100vw - 24px);' +
'    background: #fff;' +
'    box-shadow: 0 8px 28px rgba(15,23,42,0.18);' +
'    border: 1px solid #e5e7eb;' +
'    border-radius: 10px;' +
'    z-index: 1000;' +
'    display: flex;' +
'    flex-direction: column;' +
'    font-family: "Microsoft YaHei", "PingFang SC", sans-serif;' +
'    overflow: hidden;' +
'    transition: width 0.2s ease, max-width 0.2s ease;' +
'}' +
'.proofread-side-panel.collapsed {' +
'    width: 46px;' +
'    min-width: 46px;' +
'}' +
'.proofread-side-panel.collapsed .proofread-panel-title span:last-child,' +
'.proofread-side-panel.collapsed .proofread-plan-summary,' +
'.proofread-side-panel.collapsed .proofread-panel-content,' +
'.proofread-side-panel.collapsed .proofread-panel-actions button:last-child {' +
'    display: none;' +
'}' +
'.proofread-panel-header {' +
'    height: 44px;' +
'    flex: 0 0 auto;' +
'    display: flex;' +
'    align-items: center;' +
'    justify-content: space-between;' +
'    padding: 0 10px 0 12px;' +
'    background: #f8fafc;' +
'    border-bottom: 1px solid #e5e7eb;' +
'}' +
'.proofread-panel-title {' +
'    display: flex;' +
'    align-items: center;' +
'    gap: 8px;' +
'    min-width: 0;' +
'    font-size: 14px;' +
'    font-weight: 600;' +
'    color: #111827;' +
'}' +
'.proofread-panel-icon {' +
'    width: 22px;' +
'    height: 22px;' +
'    display: inline-flex;' +
'    align-items: center;' +
'    justify-content: center;' +
'    border-radius: 999px;' +
'    background: #2563eb;' +
'    color: white;' +
'    font-size: 13px;' +
'}' +
'.proofread-panel-actions {' +
'    display: flex;' +
'    gap: 4px;' +
'}' +
'.proofread-panel-action {' +
'    width: 26px;' +
'    height: 26px;' +
'    border: none;' +
'    border-radius: 6px;' +
'    background: transparent;' +
'    color: #475569;' +
'    cursor: pointer;' +
'    font-size: 18px;' +
'    line-height: 1;' +
'}' +
'.proofread-panel-action:hover {' +
'    background: #e5e7eb;' +
'}' +
'.proofread-plan-summary {' +
'    display: none;' +
'    flex: 0 0 auto;' +
'    padding: 10px 14px;' +
'    border-bottom: 1px solid #e5e7eb;' +
'    background: #f9fafb;' +
'}' +
'.proofread-plan-label {' +
'    font-size: 11px;' +
'    color: #64748b;' +
'    margin-bottom: 4px;' +
'}' +
'.proofread-plan-text {' +
'    font-size: 12px;' +
'    line-height: 1.5;' +
'    color: #334155;' +
'}' +
'.proofread-panel-content {' +
'    flex: 1;' +
'    overflow-y: auto;' +
'    padding: 16px;' +
'}' +
/* ========== 校对列表样式 ========== */
'.proofread-list {' +
'    width: 100%;' +
'}' +
'.proofread-list-header {' +
'    display: flex;' +
'    align-items: center;' +
'    gap: 8px;' +
'    padding: 12px 0;' +
'    border-bottom: 1px solid #e5e7eb;' +
'    margin-bottom: 16px;' +
'}' +
'.proofread-summary {' +
'    margin: -8px 0 14px;' +
'    padding: 8px 10px;' +
'    border-radius: 6px;' +
'    background: #f8fafc;' +
'    color: #475569;' +
'    font-size: 12px;' +
'    line-height: 1.5;' +
'}' +
'.proofread-summary .high { color: #dc2626; }' +
'.proofread-summary .medium { color: #d97706; }' +
'.proofread-summary .low { color: #2563eb; }' +
'.proofread-list-icon {' +
'    font-size: 20px;' +
'}' +
'.proofread-list-title {' +
'    font-size: 16px;' +
'    font-weight: 600;' +
'    color: #1f2937;' +
'}' +
'.proofread-severity-group {' +
'    margin-bottom: 16px;' +
'}' +
'.severity-header {' +
'    font-size: 13px;' +
'    font-weight: 600;' +
'    padding: 8px 12px;' +
'    border-radius: 6px;' +
'    margin-bottom: 8px;' +
'}' +
'.severity-header.high {' +
'    background: #fef2f2;' +
'    color: #dc2626;' +
'}' +
'.severity-header.medium {' +
'    background: #fef3c7;' +
'    color: #d97706;' +
'}' +
'.severity-header.low {' +
'    background: #f0fdf4;' +
'    color: #16a34a;' +
'}' +
'.proofread-issue-item {' +
'    background: #f9fafb;' +
'    border-radius: 8px;' +
'    padding: 12px;' +
'    margin-bottom: 8px;' +
'    border-left: 3px solid transparent;' +
'    transition: opacity 0.3s;' +
'}' +
'.proofread-issue-item.high {' +
'    border-left-color: #dc2626;' +
'}' +
'.proofread-issue-item.medium {' +
'    border-left-color: #d97706;' +
'}' +
'.proofread-issue-item.low {' +
'    border-left-color: #16a34a;' +
'}' +
'.issue-header {' +
'    display: flex;' +
'    justify-content: space-between;' +
'    margin-bottom: 8px;' +
'    font-size: 12px;' +
'}' +
'.issue-location {' +
'    color: #6b7280;' +
'}' +
'.issue-type {' +
'    background: #e5e7eb;' +
'    color: #4b5563;' +
'    padding: 2px 8px;' +
'    border-radius: 10px;' +
'}' +
'.issue-content {' +
'    margin-bottom: 8px;' +
'}' +
'.issue-original,' +
'.issue-suggestion {' +
'    font-size: 13px;' +
'    margin-bottom: 4px;' +
'}' +
'.issue-original .label,' +
'.issue-suggestion .label {' +
'    color: #6b7280;' +
'    margin-right: 6px;' +
'}' +
'.issue-original .text {' +
'    color: #dc2626;' +
'}' +
'.issue-suggestion .text {' +
'    color: #16a34a;' +
'    font-weight: 500;' +
'}' +
'.issue-explanation {' +
'    font-size: 12px;' +
'    color: #6b7280;' +
'    background: #f3f4f6;' +
'    padding: 6px 10px;' +
'    border-radius: 4px;' +
'    margin-bottom: 8px;' +
'}' +
'.issue-actions {' +
'    display: flex;' +
'    gap: 8px;' +
'}' +
'.issue-btn {' +
'    flex: 1;' +
'    padding: 6px 12px;' +
'    border: none;' +
'    border-radius: 6px;' +
'    font-size: 12px;' +
'    cursor: pointer;' +
'    transition: all 0.2s;' +
'}' +
'.issue-btn.accept {' +
'    background: #2563eb;' +
'    color: white;' +
'}' +
'.issue-btn.accept:hover {' +
'    background: #1d4ed8;' +
'}' +
'.issue-btn.ignore {' +
'    background: #f3f4f6;' +
'    color: #6b7280;' +
'}' +
'.issue-btn.ignore:hover {' +
'    background: #e5e7eb;' +
'}' +
'.proofread-list-actions {' +
'    display: flex;' +
'    gap: 8px;' +
'    margin-top: 16px;' +
'    padding-top: 16px;' +
'    border-top: 1px solid #e5e7eb;' +
'}' +
'.proofread-btn {' +
'    flex: 1;' +
'    padding: 10px 16px;' +
'    border: none;' +
'    border-radius: 8px;' +
'    font-size: 14px;' +
'    font-weight: 500;' +
'    cursor: pointer;' +
'    transition: all 0.2s;' +
'}' +
'.proofread-btn.primary {' +
'    background: #2563eb;' +
'    color: white;' +
'}' +
'.proofread-btn.primary:hover {' +
'    background: #1d4ed8;' +
'}' +
'.proofread-btn.secondary {' +
'    background: #f3f4f6;' +
'    color: #4b5563;' +
'}' +
'.proofread-btn.secondary:hover {' +
'    background: #e5e7eb;' +
'}' +
'.proofread-more {' +
'    text-align: center;' +
'    color: #6b7280;' +
'    font-size: 12px;' +
'    padding: 8px;' +
'}' +
/* ========== 校对成功样式 ========== */
'.proofread-success {' +
'    display: flex;' +
'    flex-direction: column;' +
'    align-items: center;' +
'    justify-content: center;' +
'    padding: 60px 20px;' +
'    text-align: center;' +
'}' +
'.proofread-success .success-icon {' +
'    font-size: 56px;' +
'    margin-bottom: 16px;' +
'}' +
'.proofread-success .success-text {' +
'    font-size: 18px;' +
'    color: #16a34a;' +
'    font-weight: 600;' +
'    margin-bottom: 8px;' +
'}' +
'.proofread-success .success-hint {' +
'    font-size: 14px;' +
'    color: #6b7280;' +
'    margin-top: 8px;' +
'}' +
'.proofread-error {' +
'    display: flex;' +
'    flex-direction: column;' +
'    align-items: stretch;' +
'    justify-content: center;' +
'    padding: 36px 16px;' +
'    text-align: left;' +
'}' +
'.proofread-error .error-icon {' +
'    font-size: 42px;' +
'    text-align: center;' +
'    margin-bottom: 12px;' +
'}' +
'.proofread-error .error-title {' +
'    font-size: 17px;' +
'    color: #b45309;' +
'    font-weight: 600;' +
'    text-align: center;' +
'    margin-bottom: 8px;' +
'}' +
'.proofread-error .error-message {' +
'    color: #92400e;' +
'    background: #fffbeb;' +
'    border: 1px solid #fde68a;' +
'    border-radius: 6px;' +
'    padding: 10px;' +
'    font-size: 13px;' +
'    line-height: 1.5;' +
'}' +
'.proofread-error-hint {' +
'    color: #6b7280;' +
'    font-size: 13px;' +
'    line-height: 1.5;' +
'}' +
'.proofread-error-raw {' +
'    max-height: 160px;' +
'    overflow: auto;' +
'    white-space: pre-wrap;' +
'    word-break: break-word;' +
'    background: #111827;' +
'    color: #e5e7eb;' +
'    border-radius: 6px;' +
'    padding: 10px;' +
'    font-size: 12px;' +
'    line-height: 1.45;' +
'}';
    
    document.head.appendChild(style);
}

// ========== 页面加载时初始化 ==========

// 等待DOM加载完成
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectProofreadStyles);
} else {
    injectProofreadStyles();
}
