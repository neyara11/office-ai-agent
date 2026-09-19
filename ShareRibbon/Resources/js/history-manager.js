/**
 * history-manager.js - History Sidebar Management
 * Handles chat session display and navigation
 */

window.historyManager = {
    isOpen: false,

    // Initialize history functionality
    init: function () {
        const toggleBtn = document.getElementById('history-toggle-btn');
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');
        const closeBtn = document.getElementById('close-sidebar-btn');
        const newSessionBtn = document.getElementById('new-session-btn');

        // Bind events
        toggleBtn.addEventListener('click', () => this.toggleSidebar());
        closeBtn.addEventListener('click', () => this.closeSidebar());
        overlay.addEventListener('click', () => this.closeSidebar());
        if (newSessionBtn) {
            newSessionBtn.addEventListener('click', () => {
                this.sendMessageToVB({ type: 'newSession' });
                this.closeSidebar();
            });
        }

        // Keyboard event
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this.isOpen) {
                this.closeSidebar();
            }
        });
    },

    // Toggle sidebar visibility
    toggleSidebar: function () {
        if (this.isOpen) {
            this.closeSidebar();
        } else {
            this.openSidebar();
        }
    },

    // Open sidebar
    openSidebar: function () {
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');

        sidebar.classList.remove('sidebar-hidden');
        sidebar.classList.add('sidebar-visible');
        overlay.classList.remove('overlay-hidden');
        overlay.classList.add('overlay-visible');

        this.isOpen = true;

        // Load history files
        this.loadHistoryFiles();
    },

    // Close sidebar
    closeSidebar: function () {
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');

        sidebar.classList.remove('sidebar-visible');
        sidebar.classList.add('sidebar-hidden');
        overlay.classList.remove('overlay-visible');
        overlay.classList.add('overlay-hidden');

        this.isOpen = false;
    },

    // Load history files list
    loadHistoryFiles: function () {
        const historyList = document.getElementById('history-list');

        // Show loading state
        historyList.innerHTML = '<div class="loading-state">Загрузка истории...</div>';

        // Request session list from backend (conversation/session_summary)
        this.sendMessageToVB({
            type: 'getSessionList'
        });
    },

    // Display session list from backend
    displayHistoryFiles: function (files) {
        const historyList = document.getElementById('history-list');

        if (!files || files.length === 0) {
            historyList.innerHTML = `
                <div class="empty-state">
                    <div class="empty-state-icon">📄</div>
                    <div class="empty-state-text">У вас пока нет истории сессий</div>
                </div>
            `;
            return;
        }

        files.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
        const itemsHtml = files.map(s => {
            const title = (s.title || 'Сессия').replace(/'/g, "\\'");
            const sid = (s.sessionId || '').replace(/'/g, "\\'");
            return `<div class="history-item" data-session-id="${s.sessionId}" onclick="historyManager.loadSession('${sid}')">
                <div class="history-item-title">${this.escapeHtml(title)}</div>
                <div class="history-item-date">${this.formatSessionDate(s.createdAt)}</div>
            </div>`;
        }).join('');
        historyList.innerHTML = itemsHtml;
    },

    escapeHtml: function (text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    formatSessionDate: function (createdAt) {
        if (!createdAt) return 'Неизвестное время';
        return String(createdAt).replace('T', ' ').substring(0, 19);
    },

    loadSession: function (sessionId) {
        this.sendMessageToVB({
            type: 'loadSession',
            sessionId: sessionId
        });
        this.closeSidebar();
    },

    // Send message to VB backend
    sendMessageToVB: function (message) {
        try {
            if (window.officeAi && typeof window.officeAi.post === 'function') {
                window.officeAi.post(message);
            } else if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage(message);
            } else if (window.vsto) {
                if (typeof window.vsto.sendMessage === 'function') {
                    window.vsto.sendMessage(JSON.stringify(message));
                } else if (typeof window.vsto.postMessage === 'function') {
                    window.vsto.postMessage(message);
                }
            } else {
                console.error('无法与后端通信');
            }
        } catch (error) {
            console.error('发送消息到VB后端失败:', error);
        }
    }
};

// Global function for VB backend to call
window.setHistoryFilesList = function (files) {
    historyManager.displayHistoryFiles(files);
};
