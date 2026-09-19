/**
 * agent-card.js - 统一 Agent UI 组件
 * Единое отображение плана Agent, шагов выполнения, подтверждений иПояснение выполнения
 * 支持 ReAct 循环展示：Think → Action → Observation
 */

// 统一 Agent 状态
window.agentCardState = {
    active: false,
    session: null,
    locked: false
};

/**
 * 锁定聊天输入（Agent 执行期间）
 */
function lockChatInput() {
    window.agentCardState.locked = true;
    const smartInput = document.getElementById('smart-input');
    const sendBtn = document.getElementById('send-button');
    const chatInput = document.getElementById('chat-input');

    if (smartInput) {
        smartInput.contentEditable = 'false';
        smartInput.classList.add('input-locked');
        smartInput.dataset.placeholder = 'Agent выполняется, дождитесь завершения или нажмите «Стоп»...';
    }
    if (chatInput) chatInput.disabled = true;
    if (sendBtn) sendBtn.disabled = true;
}

/**
 * 解锁聊天输入
 */
function unlockChatInput() {
    window.agentCardState.locked = false;
    const smartInput = document.getElementById('smart-input');
    const sendBtn = document.getElementById('send-button');
    const chatInput = document.getElementById('chat-input');

    if (smartInput) {
        smartInput.contentEditable = 'true';
        smartInput.classList.remove('input-locked');
        smartInput.dataset.placeholder = 'Введите вопрос... Enter — отправить, Tab — принять подсказку';
    }
    if (chatInput) chatInput.disabled = false;
    if (sendBtn) sendBtn.disabled = false;
}

/**
 * 恢复一次 Agent 请求占用的全部输入 UI。
 * 输入锁和发送/终止按钮由不同模块管理，Agent 进入终态时必须同时复位。
 */
function restoreAgentRequestUi() {
    unlockChatInput();

    if (typeof changeSendButton === 'function') {
        changeSendButton();
        return;
    }

    // 脚本异常或加载顺序变化时的兜底，避免终止按钮永久停留。
    const sendButton = document.getElementById('send-button');
    const stopButton = document.getElementById('stop-button');
    if (sendButton) sendButton.style.setProperty('display', 'flex', 'important');
    if (stopButton) stopButton.style.setProperty('display', 'none', 'important');
    if (typeof hideLoadingIndicator === 'function') hideLoadingIndicator();
}

/**
 * 检查是否被 Agent 锁定
 */
function isAgentLocked() {
    return window.agentCardState.locked;
}

/**
 * 显示 Agent 规划卡片（嵌入聊天流）
 * @param {Object} planData - { sessionId, understanding, steps, summary, replaceThinkingUuid }
 */
function showAgentPlanCard(planData) {
    try {
        window.agentCardState.active = true;
        window.agentCardState.session = planData;

        lockChatInput();

        const uuid = planData.sessionId || generateUUID();
        const timestamp = formatDateTime(new Date());

        let chatContainer;
        if (planData.replaceThinkingUuid) {
            const thinkingDiv = document.getElementById('content-' + planData.replaceThinkingUuid);
            const parentContainer = thinkingDiv ? thinkingDiv.closest('.chat-container') : null;
            if (parentContainer) {
                chatContainer = parentContainer;
                chatContainer.className = 'chat-container agent-card-container';
                chatContainer.id = 'agent-plan-' + uuid;
                chatContainer.innerHTML = '';
            }
        }

        if (!chatContainer) {
            chatContainer = document.createElement('div');
            chatContainer.className = 'chat-container agent-card-container';
            chatContainer.id = 'agent-plan-' + uuid;
        }

        const stepsHtml = planData.steps ? planData.steps.map((step, idx) => `
            <div class="agent-step" id="agent-step-${uuid}-${idx}" data-status="pending">
                <div class="agent-step-header">
                    <span class="agent-step-icon" id="step-icon-${uuid}-${idx}">⏳</span>
                    <span class="agent-step-num">${idx + 1}</span>
                    <span class="agent-step-desc">${escapeHtml(step.description)}</span>
                </div>
                <div class="agent-step-detail" id="step-detail-${uuid}-${idx}" style="display:none;">
                    <div class="step-detail-text">${escapeHtml(step.detail || step.code || '')}</div>
                </div>
                <div class="agent-iteration-area" id="iteration-area-${uuid}-${idx}"></div>
            </div>
        `).join('') : '';

        chatContainer.innerHTML = `
            <div class="message-header">
                <div class="avatar-ai">AI</div>
                <div class="sender-info">
                    <div class="sender-name">Agent <span class="agent-badge">Авто-выполнение</span></div>
                    <div class="timestamp">${timestamp}</div>
                </div>
            </div>
            <div class="message-content agent-plan-content">
                <div class="agent-understanding">
                    <strong>📋 Понимание:</strong>${escapeHtml(planData.understanding || '')}
                </div>
                <div class="agent-steps-container">
                    <div class="agent-steps-header">
                        <span>📝 План выполнения</span>
                        <span class="agent-step-count">${planData.steps ? planData.steps.length : 0} шагов</span>
                    </div>
                    <div class="agent-steps-list" id="agent-steps-${uuid}">
                        ${stepsHtml}
                    </div>
                </div>
                <div class="agent-summary">
                    <strong>🎯 Ожидаемый результат:</strong>${escapeHtml(planData.summary || '')}
                </div>
                <div class="agent-actions" id="agent-actions-${uuid}">
                    <button class="agent-btn agent-btn-execute" onclick="confirmAgentExecution('${uuid}')">
                        ▶ Начать выполнение
                    </button>
                    <button class="agent-btn agent-btn-refine" onclick="refineAgentPlan('${uuid}')">🔄 Изменить план</button>
                    <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                        ✖ Отмена
                    </button>
                </div>
            </div>
            <div class="agent-status-bar" id="agent-status-${uuid}">
                <span class="status-icon">⏸</span>
                <span class="status-text">Ожидание подтверждения</span>
            </div>
        `;

        const isAlreadyInContainer = chatContainer.parentElement && chatContainer.parentElement.id === 'chat-container';
        if (!isAlreadyInContainer) {
            const chatHistoryContainer = document.getElementById('chat-container');
            if (chatHistoryContainer) {
                chatHistoryContainer.appendChild(chatContainer);
            }
        }

        chatContainer.scrollIntoView({ behavior: 'smooth', block: 'end' });

        window.agentCardState.session.uuid = uuid;

        // Если задача простая (авто-выполнение), нажать выполнение автоматически
        if (planData.autoExecute) {
            setTimeout(() => confirmAgentExecution(uuid), 500);
        }
    } catch (err) {
        console.error('showAgentPlanCard error:', err);
        unlockChatInput();
    }
}

/**
 * 更新 Agent 步骤状态
 * @param {string} sessionId - 会话 ID
 * @param {number} stepIndex - 步骤索引
 * @param {string} status - 状态：pending / running / completed / failed / skipped
 * @param {string} message - 可选消息
 */
function updateAgentStep(sessionId, stepIndex, status, message) {
    try {
        const stepEl = document.getElementById(`agent-step-${sessionId}-${stepIndex}`);
        if (!stepEl) return;

        stepEl.setAttribute('data-status', status);

        const iconEl = document.getElementById(`step-icon-${sessionId}-${stepIndex}`);
        if (iconEl) {
            const icons = {
                pending: '⏳',
                running: '🔄',
                completed: '✅',
                failed: '❌',
                skipped: '⏭'
            };
            iconEl.textContent = icons[status] || '⏳';
        }

        if (status === 'running') {
            stepEl.classList.add('step-running');
            const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
            if (detailEl) detailEl.style.display = 'block';
        } else if (status === 'completed') {
            stepEl.classList.remove('step-running');
            stepEl.classList.add('step-completed');
        } else if (status === 'failed') {
            stepEl.classList.remove('step-running');
            stepEl.classList.add('step-failed');
        }

        if (message) {
            const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
            if (detailEl) {
                const msgDiv = document.createElement('div');
                msgDiv.className = `step-message step-msg-${status}`;
                msgDiv.textContent = message;
                detailEl.appendChild(msgDiv);
            }
        }
    } catch (err) {
        console.error('updateAgentStep error:', err);
    }
}

/**
 * 更新 ReAct 迭代（Think → Action → Observation）
 * @param {string} sessionId - 会话 ID
 * @param {Object} iteration - { index, thought, action, observation }
 */
function updateAgentIteration(sessionId, iteration) {
    try {
        if (!iteration) return;

        const areaEl = document.getElementById(`iteration-area-${sessionId}-${iteration.index}`);
        if (!areaEl) {
            // 如果没有精确匹配的步骤区域，放到当前运行的步骤下
            const runningStep = document.querySelector(`#agent-steps-${sessionId} .step-running`);
            if (runningStep) {
                const idx = runningStep.querySelector('.agent-step-num')?.textContent;
                if (idx) {
                    const fallbackArea = document.getElementById(`iteration-area-${sessionId}-${parseInt(idx) - 1}`);
                    if (fallbackArea) fallbackArea.innerHTML = buildIterationHtml(iteration);
                }
            }
            return;
        }

        areaEl.innerHTML = buildIterationHtml(iteration);
    } catch (err) {
        console.error('updateAgentIteration error:', err);
    }
}

/**
 * 构建迭代 HTML
 */
function buildIterationHtml(iteration) {
    const thought = escapeHtml(iteration.thought || '');
    const action = escapeHtml(iteration.action || '');
    const observation = escapeHtml(iteration.observation || '');
    const explanation = iteration.explanation || null;
    const explanationText = explanation ? escapeHtml(explanation.ExplanationText || explanation.explanationText || '') : '';

    return `
        <div class="react-iteration">
            <div class="iteration-thought">
                <span class="iteration-label">💭 Размышление</span>
                <div class="iteration-content">${thought}</div>
            </div>
            ${action ? `
            <div class="iteration-action">
                <span class="iteration-label">🔧 Действие</span>
                <div class="iteration-content"><code>${action}</code></div>
            </div>` : ''}
            ${observation ? `
            <div class="iteration-observation">
                <span class="iteration-label">👁 Наблюдение</span>
                <div class="iteration-content">${observation}</div>
            </div>` : ''}
            ${explanationText ? `
            <div class="iteration-explanation">
                <span class="iteration-label">Пояснение выполнения</span>
                <div class="iteration-content">${explanationText}</div>
            </div>` : ''}
        </div>
    `;
}

/**
 * Показать пошаговое пояснение выполнения
 * @param {string} sessionId - 会话 ID
 * @param {Object} explanation - ExecutionExplanation
 */
function showAgentExecutionExplanation(sessionId, explanation) {
    try {
        if (!explanation) return;
        const stepIndex = explanation.StepIndex ?? explanation.stepIndex;
        const text = explanation.ExplanationText || explanation.explanationText || explanation.Message || explanation.message || '';
        if (stepIndex === undefined || !text) return;

        const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
        if (!detailEl) return;

        const explanationEl = document.createElement('details');
        explanationEl.className = 'step-explanation';
        explanationEl.open = false;

        const toolId = explanation.ToolId || explanation.toolId || '';
        const category = explanation.ToolCategory || explanation.toolCategory || '';
        const skillName = explanation.SkillName || explanation.skillName || '';
        const scriptFileName = explanation.ScriptFileName || explanation.scriptFileName || '';
        const mcpToolName = explanation.McpToolName || explanation.mcpToolName || '';
        const mcpStatus = explanation.McpStatus || explanation.mcpStatus || '';
        const failureReason = explanation.FailureReason || explanation.failureReason || '';
        const risk = explanation.RiskLevel || explanation.riskLevel || '';
        const paramsJson = explanation.ParametersJson || explanation.parametersJson || '';
        const fixed = explanation.FixAttempts || explanation.fixAttempts || 0;
        const elapsedMs = explanation.ElapsedMs ?? explanation.elapsedMs ?? 0;
        const beforeSummary = explanation.BeforeSummary || explanation.beforeSummary || '';
        const afterSummary = explanation.AfterSummary || explanation.afterSummary || '';
        const undoPointName = explanation.UndoPointName || explanation.undoPointName || '';
        const undoHint = explanation.UndoHint || explanation.undoHint || '';
        const canUndo = explanation.CanUndo ?? explanation.canUndo;
        const autoRepairSummary = explanation.AutoRepairSummary || explanation.autoRepairSummary || '';

        explanationEl.innerHTML = `
            <summary>${escapeHtml(text)}</summary>
            <div class="step-explanation-meta">
                ${toolId ? `<div><strong>Инструмент</strong> <code>${escapeHtml(toolId)}</code></div>` : ''}
                ${category ? `<div><strong>Категория</strong> ${escapeHtml(category)}</div>` : ''}
                ${elapsedMs ? `<div><strong>Время</strong> ${Number(elapsedMs).toLocaleString()} ms</div>` : ''}
                ${skillName ? `<div><strong>Skill</strong> ${escapeHtml(skillName)}</div>` : ''}
                ${scriptFileName ? `<div><strong>Скрипт</strong> <code>${escapeHtml(scriptFileName)}</code></div>` : ''}
                ${mcpToolName ? `<div><strong>MCP</strong> <code>${escapeHtml(mcpToolName)}</code>${mcpStatus ? ' ' + escapeHtml(mcpStatus) : ''}</div>` : ''}
                ${risk ? `<div><strong>Риск</strong> ${escapeHtml(risk)}</div>` : ''}
                ${beforeSummary ? `<div><strong>До</strong> ${escapeHtml(beforeSummary)}</div>` : ''}
                ${afterSummary ? `<div><strong>После</strong> ${escapeHtml(afterSummary)}</div>` : ''}
                ${fixed ? `<div><strong>Авто-исправление</strong> ${fixed} раз</div>` : ''}
                ${autoRepairSummary ? `<div><strong>Результат исправления</strong> ${escapeHtml(autoRepairSummary)}</div>` : ''}
                ${undoPointName ? `<div><strong>Точка отмены</strong> ${escapeHtml(undoPointName)}${canUndo === false ? ' <span>(отменить вручную)</span>' : ''}</div>` : ''}
                ${undoHint ? `<div><strong>Подсказка отмены</strong> ${escapeHtml(undoHint)}</div>` : ''}
                ${failureReason ? `<div><strong>Причина сбоя</strong> ${escapeHtml(failureReason)}</div>` : ''}
                ${paramsJson ? `<pre>${escapeHtml(paramsJson)}</pre>` : ''}
            </div>
        `;
        detailEl.appendChild(explanationEl);
    } catch (err) {
        console.error('showAgentExecutionExplanation error:', err);
    }
}

/**
 * 显示审批请求 UI
 * @param {string} sessionId - 会话 ID
 * @param {string} message - 审批提示消息
 */
function showAgentApproval(sessionId, message) {
    try {
        const actionsEl = document.getElementById(`agent-actions-${sessionId}`);
        if (actionsEl) {
            actionsEl.innerHTML = `
                <div class="agent-approval-request">
                    <span class="approval-msg">${escapeHtml(message)}</span>
                    <button class="agent-btn agent-btn-execute" onclick="agentApprove('${sessionId}')">✅ Подтвердить</button>
                    <button class="agent-btn agent-btn-abort" onclick="agentReject('${sessionId}')">❌ Пропустить</button>
                </div>
            `;
        }
        updateAgentStatus(sessionId, 'waitingApproval', 'Ожидание подтверждения пользователя...');
    } catch (err) {
        console.error('showAgentApproval error:', err);
    }
}

/**
 * 用户确认执行 Agent
 */
function confirmAgentExecution(uuid) {
    const executeBtn = document.querySelector(`[onclick="confirmAgentExecution('${uuid}')"]`);
    const abortBtn = document.querySelector(`[onclick="abortAgent('${uuid}')"]`);
    if (executeBtn) { executeBtn.disabled = true; executeBtn.textContent = '⏳ Выполняется...'; }
    if (abortBtn) { abortBtn.disabled = true; }

    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = `
            <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                ⏹ Остановить
            </button>
        `;
    }

    updateAgentStatus(uuid, 'running', 'Выполняется...');
    requestApprove(uuid);
}

/**
 * 用户批准当前审批项
 */
function agentApprove(sessionId) {
    requestApprove(sessionId);
}

/**
 * 用户拒绝当前审批项
 */
function agentReject(sessionId) {
    requestReject(sessionId);
}

/**
 * 修改 Agent 计划
 */
function refineAgentPlan(uuid) {
    const feedback = prompt('Введите замечания к плану:');
    if (feedback && feedback.trim()) {
        requestRefinePlan(uuid, feedback.trim());
    }
}

/**
 * 终止 Agent
 */
function abortAgent(uuid) {
    updateAgentStatus(uuid, 'aborted', 'Остановлено');
    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = '<span class="agent-terminated">Остановлено</span>';
    }
    requestAbortAgent();
    restoreAgentRequestUi();
}

/**
 * 更新 Agent 状态栏
 * @param {string} uuid - 会话 UUID
 * @param {string} status - 状态
 * @param {string} text - 显示文本
 */
function updateAgentStatus(uuid, status, text) {
    const statusBar = document.getElementById('agent-status-' + uuid);
    if (!statusBar) return;

    const icons = {
        running: '🔄',
        waitingApproval: '⏸',
        completed: '✅',
        failed: '❌',
        aborted: '⏹',
        paused: '⏸'
    };

    statusBar.innerHTML = `
        <span class="status-icon">${icons[status] || '⏳'}</span>
        <span class="status-text">${escapeHtml(text || '')}</span>
    `;
}

/**
 * 完成 Agent（成功或失败）
 * @param {string} uuid - 会话 UUID
 * @param {boolean} success - 是否成功
 * @param {string} message - 完成消息
 */
function completeAgent(uuid, success, message, thinkingUuid) {
    try {
        const statusBar = document.getElementById('agent-status-' + uuid);
        if (statusBar) {
            const icon = success ? '✅' : '❌';
            const text = success ? 'Задача выполнена' : 'Задача не выполнена';
            statusBar.innerHTML = `
                <span class="status-icon">${icon}</span>
                <span class="status-text">${text}${message ? ': ' + escapeHtml(message) : ''}</span>
            `;
        }

        const actions = document.getElementById('agent-actions-' + uuid);
        if (actions) {
            actions.innerHTML = `<span class="agent-finished">${success ? '✅ Готово' : '❌ Ошибка'}</span>`;
        }

        // 没有生成 plan card 时，状态仍显示在最初的 thinking 消息中。
        if (thinkingUuid) {
            const thinkingDiv = document.getElementById('content-' + thinkingUuid);
            if (thinkingDiv) {
                const icon = success ? '✅' : '❌';
                const text = success ? 'Задача выполнена' : 'Задача не выполнена';
                thinkingDiv.innerHTML = `<div class="agent-terminal-message ${success ? 'success' : 'failed'}">` +
                    `<span class="status-icon">${icon}</span>` +
                    `<span class="status-text">${text}${message ? ': ' + escapeHtml(message) : ''}</span>` +
                    `</div>`;
            }
        }

        window.agentCardState.active = false;
        window.agentCardState.session = null;
    } catch (err) {
        console.error('completeAgent error:', err);
    } finally {
        restoreAgentRequestUi();
    }
}

/**
 * 获取步骤状态图标
 */
function getStepIcon(status) {
    const icons = {
        pending: '⏳',
        running: '🔄',
        completed: '✅',
        failed: '❌',
        skipped: '⏭'
    };
    return icons[status] || '⏳';
}

// 导出到全局
window.agentCardState = window.agentCardState;
window.showAgentPlanCard = showAgentPlanCard;
window.updateAgentStep = updateAgentStep;
window.updateAgentIteration = updateAgentIteration;
window.showAgentExecutionExplanation = showAgentExecutionExplanation;
window.showAgentApproval = showAgentApproval;
window.confirmAgentExecution = confirmAgentExecution;
window.agentApprove = agentApprove;
window.agentReject = agentReject;
window.refineAgentPlan = refineAgentPlan;
window.abortAgent = abortAgent;
window.updateAgentStatus = updateAgentStatus;
window.completeAgent = completeAgent;
window.lockChatInput = lockChatInput;
window.unlockChatInput = unlockChatInput;
window.restoreAgentRequestUi = restoreAgentRequestUi;
window.isAgentLocked = isAgentLocked;
