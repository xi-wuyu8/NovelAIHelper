/* ./src/taskpane/taskpane.js */

// Global state for selected text
const state = {
    selectedText: "",
};

// Prevents re-initialization
let isOfficeInitialized = false;

// Stores the conversation history
let chatHistory = [];

// Pre-defined models for different AI providers
const models = {
    gemini: ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.5-flash-lite','gemini-2.0-flash', 'gemini-2.0-flash-lite',''],
    openai: ['gpt-4o', 'gpt-4o-mini', 'gpt-3.5-turbo'],
};


Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        if (isOfficeInitialized) {
            return;
        }

        // --- Event Listeners Setup ---
        document.getElementById("send-button").onclick = sendToAI;
        document.getElementById("add-file-button").onclick = () => document.getElementById('file-input').click();
        document.getElementById('file-input').onchange = handleFileSelect;
        
        const userInput = document.getElementById('user-input');
        userInput.onkeydown = (event) => {
            if (event.key === 'Enter' && !event.shiftKey) {
                event.preventDefault();
                sendToAI();
            }
        };
        // Auto-resize textarea
        userInput.addEventListener('input', () => {
            userInput.style.height = 'auto';
            userInput.style.height = (userInput.scrollHeight) + 'px';
        });

        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
        
        // --- API Modal Listeners ---
        document.getElementById('settings-button').onclick = () => toggleApiSettingsModal(true);
        document.getElementById('close-modal-button').onclick = () => toggleApiSettingsModal(false);
        document.getElementById('provider-select').onchange = updateModelOptions;
        document.getElementById('save-settings-button').onclick = saveApiSettings;

        // --- Initial Load Actions ---
        loadApiSettings();
        lucide.createIcons();
        updateKnowledgeBaseUI();
        onSelectionChanged();

        isOfficeInitialized = true;
    }
});

// ====================================================================================
// API Settings Modal Functions
// ====================================================================================

function toggleApiSettingsModal(show) {
    const modal = document.getElementById('settings-modal');
    modal.style.display = show ? 'flex' : 'none';
}

function updateModelOptions() {
    const provider = document.getElementById('provider-select').value;
    const modelSelect = document.getElementById('model-select');
    modelSelect.innerHTML = '';
    
    if (models[provider]) {
        models[provider].forEach(model => {
            const option = document.createElement('option');
            option.value = model;
            option.textContent = model;
            modelSelect.appendChild(option);
        });
    }
}

function loadApiSettings() {
    try {
        const settings = JSON.parse(localStorage.getItem('ai-settings')) || {};
        document.getElementById('provider-select').value = settings.provider || 'gemini';
        document.getElementById('api-key-input').value = settings.apiKey || '';
        
        updateModelOptions();
        if (settings.model) {
            document.getElementById('model-select').value = settings.model;
        }
    } catch (e) {
        console.error("Error loading API settings:", e);
    }
}

function saveApiSettings() {
    const provider = document.getElementById('provider-select').value;
    const apiKey = document.getElementById('api-key-input').value;
    const model = document.getElementById('model-select').value;

    if (!apiKey) {
        updateStatus("API 密钥不能为空！", 3000);
        return;
    }

    const settings = { provider, apiKey, model };
    localStorage.setItem('ai-settings', JSON.stringify(settings));
    updateStatus("API 设置已保存", 3000);
    toggleApiSettingsModal(false);
}

// ====================================================================================
// Core AI Communication Function
// ====================================================================================

async function sendToAI() {
    const userInputElement = document.getElementById("user-input");
    const userInput = userInputElement.value.trim();
    if (!userInput) return;

    addMessageToHistory(userInput, "user");
    chatHistory.push({ role: "user", content: userInput });
    userInputElement.value = "";
    userInputElement.style.height = 'auto'; // Reset height after sending

    const settings = JSON.parse(localStorage.getItem('ai-settings')) || {};
    if (!settings.apiKey) {
        const errorMsg = "请先在API设置中配置您的API密钥！";
        updateStatus(errorMsg, 5000);
        addMessageToHistory(errorMsg, "ai");
        chatHistory.pop(); // Remove user message if config is missing
        return;
    }

    try {
        updateStatus("正在获取主文档最新内容...");
        const mainDocContent = await getDocumentText();
        
        updateStatus(`正在请求 ${settings.provider} AI 服务...`);
        const response = await fetch("http://localhost:3001/api/chat", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                prompt: userInput,
                selectedText: state.selectedText,
                mainDocContent: mainDocContent,
                settings: settings,
                chatHistory: chatHistory,
            }),
        });

        if (!response.ok) {
            const errorBody = await response.json();
            throw new Error(`服务器错误 (${response.status}): ${errorBody.error || response.statusText}`);
        }

        const data = await response.json();
        const aiResponse = data.response;

        addMessageToHistory(aiResponse, "ai");
        chatHistory.push({ role: "assistant", content: aiResponse });

        if (state.selectedText) {
            await replaceSelectionWith(aiResponse);
            updateStatus("内容已智能修改", 3000);
        } else {
            askToInsertText(aiResponse);
            updateStatus("准备就绪");
        }
    } catch (error) {
        console.error("Error:", error);
        const errorMsg = `发生错误: ${error.message}`;
        updateStatus(errorMsg, 5000);
        addMessageToHistory(`抱歉，处理时发生错误: ${error.message}`, "ai");
        chatHistory.pop(); // Remove user's message from history on error to allow retry
    }
}

// ====================================================================================
// Knowledge Base Functions
// ====================================================================================

async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    updateStatus(`正在上传并处理 "${file.name}"...`);
    const formData = new FormData();
    formData.append('file', file);
    try {
        const response = await fetch("http://localhost:3001/api/upload", { method: 'POST', body: formData });
        if (!response.ok) throw new Error(`文件处理失败: ${await response.text()}`);
        const data = await response.json();
        updateStatus(`"${data.fileName}" 已成功添加!`, 3000);
        await updateKnowledgeBaseUI();
    } catch (error) {
        console.error("Upload error:", error);
        updateStatus(`上传错误: ${error.message}`, 5000);
    } finally {
        event.target.value = null;
    }
}

async function updateKnowledgeBaseUI() {
    try {
        const response = await fetch("http://localhost:3001/api/knowledge-base");
        if (!response.ok) throw new Error("无法连接到后端");
        const files = await response.json();
        renderFileList(files);
    } catch (error) {
        console.error("Failed to fetch knowledge base:", error);
        updateStatus("无法连接到后端服务", 5000);
    }
}

async function deleteFileFromKnowledgeBase(fileName) {
    try {
        const response = await fetch(`http://localhost:3001/api/knowledge-base/${encodeURIComponent(fileName)}`, { method: 'DELETE' });
        if (!response.ok) throw new Error('删除失败');
        updateStatus(`"${fileName}" 已删除`, 3000);
        await updateKnowledgeBaseUI();
    } catch (error) {
        console.error("Delete error:", error);
        updateStatus(`删除失败: ${error.message}`, 5000);
    }
}

function renderFileList(files) {
    const listElement = document.getElementById("file-list");
    listElement.innerHTML = "";
    const mainDocItem = document.createElement("li");
    mainDocItem.className = "file-item";
    mainDocItem.innerHTML = `<span><i data-lucide="file-text" class="button-icon"></i> 当前小说 (主文档)</span>`;
    listElement.appendChild(mainDocItem);
    files.filter(name => name !== "当前小说 (主文档)").forEach(fileName => {
        const listItem = document.createElement("li");
        listItem.className = "file-item";
        listItem.innerHTML = `<span><i data-lucide="file-text" class="button-icon"></i> ${fileName}</span>`;
        const deleteButton = document.createElement("button");
        deleteButton.className = "delete-file-btn";
        deleteButton.title = `删除 ${fileName}`;
        deleteButton.innerHTML = `<i data-lucide="x"></i>`;
        deleteButton.onclick = (e) => {
            e.stopPropagation();
            deleteFileFromKnowledgeBase(fileName);
        };
        listItem.appendChild(deleteButton);
        listElement.appendChild(listItem);
    });
    lucide.createIcons();
}

// ====================================================================================
// Word Document Interaction & UI Utility Functions
// ====================================================================================

async function onSelectionChanged() {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            context.load(range, "text");
            await context.sync();
            const selectedText = range.text.trim();
            state.selectedText = selectedText;
            const display = document.getElementById("selected-text-display");
            const input = document.getElementById("user-input");
            if (selectedText) {
                display.innerText = selectedText.length > 200 ? selectedText.substring(0, 200) + '...' : selectedText;
                input.placeholder = "对选中内容进行操作 (如: 润色)...";
            } else {
                display.innerHTML = `<p class="placeholder">请在文档中选择文本以进行编辑...</p>`;
                input.placeholder = "输入指令或问题...";
            }
        });
    } catch (error) {
        console.log("Selection change error (ignorable):", error);
    }
}

function replaceSelectionWith(newText) {
    return Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(newText, Word.InsertLocation.replace);
        await context.sync();
    });
}

function getDocumentText() {
    return Word.run(async (context) => {
        const body = context.document.body;
        context.load(body, "text");
        await context.sync();
        return body.text;
    });
}

function askToInsertText(textToInsert) {
    const history = document.getElementById("chat-history");
    const lastMessage = history.lastElementChild;
    if (!lastMessage || lastMessage.classList.contains('user-message')) return;
    const buttonContainer = document.createElement('div');
    buttonContainer.className = 'insert-button-container';
    const insertButton = document.createElement("button");
    insertButton.textContent = "插入到光标位置";
    insertButton.className = "ms-Button";
    insertButton.onclick = () => {
        replaceSelectionWith(textToInsert);
        insertButton.disabled = true;
        insertButton.textContent = "已插入";
    };
    buttonContainer.appendChild(insertButton);
    lastMessage.appendChild(buttonContainer);
}

function addMessageToHistory(message, sender) {
    const history = document.getElementById("chat-history");
    const messageDiv = document.createElement("div");
    messageDiv.className = sender === "user" ? "user-message" : "ai-message";
    const formattedMessage = document.createElement('div');
    formattedMessage.innerText = message;
    messageDiv.innerHTML = formattedMessage.innerHTML.replace(/\n/g, '<br>');
    history.appendChild(messageDiv);
    history.scrollTop = history.scrollHeight;
}

let statusTimeout;
function updateStatus(message, duration = 0) {
    const statusElement = document.getElementById("status-message");
    statusElement.textContent = message;
    if (statusTimeout) clearTimeout(statusTimeout);
    if (duration > 0) {
        statusTimeout = setTimeout(() => { statusElement.textContent = "准备就绪"; }, duration);
    }
}
