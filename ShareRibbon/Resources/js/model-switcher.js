/**
 * model-switcher.js - Model Switching Management
 * Handles model display and switching functionality
 */

/**
 * Open model configuration dialog
 * Sends message to VB.NET to open ConfigApiForm
 */
function openModelConfig() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'openApiConfigForm'
        });
    } else if (window.vsto) {
        window.vsto.openApiConfigForm();
    } else {
        alert('Не удалось открыть настройки: интерфейс связи не найден');
    }
}

/**
 * Update the current model display in the header bar
 * Called from VB.NET after model changes
 * @param {string} platform - The platform/provider name
 * @param {string} modelName - The model name
 */
function updateCurrentModelDisplay(platform, modelName) {
    var displayElement = document.getElementById('current-model-display');
    if (displayElement) {
        if (platform && modelName) {
            displayElement.textContent = platform + ' / ' + modelName;
        } else if (modelName) {
            displayElement.textContent = modelName;
        } else {
            displayElement.textContent = 'Модель не настроена';
        }
    }
}

/**
 * Request current model info from VB.NET
 * Used to initialize the display on page load
 */
function requestCurrentModelInfo() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'getCurrentModel'
        });
    }
}

/**
 * Initialize model switcher on page load
 */
(function initModelSwitcher() {
    // Request current model info when page loads
    // Small delay to ensure VB.NET communication is ready
    setTimeout(function() {
        requestCurrentModelInfo();
    }, 500);
})();
