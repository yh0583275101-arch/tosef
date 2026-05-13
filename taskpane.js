// ============================================
// TORANI TEXT EDITOR - Office.js Add-in
// ============================================

// Default System Prompt
const DEFAULT_SYSTEM_PROMPT = `You are an expert Torani Editor. Your goal is to refine Hebrew text into a high-level Rabbinic/Yeshivish style.
- Improve flow using traditional scholarly vocabulary.
- Fix punctuation (Pissuk) according to Torah publication standards.
- If requested, add Nikkud (vowel marks).
- Format structured arguments clearly.
- Maintain the sanctity and accuracy of the source text while making it readable for scholars.
- Use academic Hebrew terminology and Yeshivish expressions like: חזינן, צ"ע, אתי שפיר, בודאי`;

// Global state
let currentPreviewChanges = [];
let currentFullDocPreview = null;

// ============================================
// Initialize
// ============================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js is ready for Word");
        initializeApp();
    }
});

function initializeApp() {
    // Set default system prompt
    document.getElementById("systemPrompt").value = DEFAULT_SYSTEM_PROMPT;
    
    // Set default yeshivish style instructions
    document.getElementById("yeshivishStyle").value = `סגנון יישיבה - הנחיות:
- השתמש בטרמינולוגיה נלמדנית בעברית אקדמית
- כלול ביטויים כמו: חזינן (is seen), צ"ע (difficult), אתי שפיר (makes sense)
- שנה משפטים פשוטים לטון נלמדני וחקיקתי
- שמור על העקביות הלשונית בכל הטקסט`;

    // Add event listeners
    setupEventListeners();
    
    // Load saved settings from localStorage
    loadSettings();
    
    addLog("✓ יישום נטען בהצלחה");
}

function setupEventListeners() {
    document.getElementById("processSelection").addEventListener("click", processSelection);
    document.getElementById("processDocument").addEventListener("click", processDocument);
    
    // Collapsible sections
    document.querySelectorAll(".collapsible-header").forEach(header => {
        header.addEventListener("click", function() {
            const content = this.nextElementSibling;
            if (content && content.classList.contains("collapsible-content")) {
                content.classList.toggle("active");
                const arrow = this.querySelector("span");
                arrow.style.transform = content.classList.contains("active") ? "rotate(180deg)" : "rotate(0deg)";
            }
        });
    });
}

// ============================================
// Settings Management
// ============================================
function loadSettings() {
    const saved = localStorage.getItem("torani-settings");
    if (saved) {
        const settings = JSON.parse(saved);
        Object.keys(settings).forEach(key => {
            const element = document.getElementById(key);
            if (element) {
                if (element.type === "checkbox") {
                    element.checked = settings[key];
                } else {
                    element.value = settings[key];
                }
            }
        });
    }
}

function saveSettings() {
    const settings = {
        apiKey: document.getElementById("apiKey").value,
        modelSelect: document.getElementById("modelSelect").value,
        checkYeshivish: document.getElementById("checkYeshivish").checked,
        checkPunctuation: document.getElementById("checkPunctuation").checked,
        checkFormatting: document.getElementById("checkFormatting").checked,
        checkVocalization: document.getElementById("checkVocalization").checked,
        previewMode: document.getElementById("previewMode").checked,
        alertEnabled: document.getElementById("alertEnabled").checked,
        systemPrompt: document.getElementById("systemPrompt").value,
        yeshivishStyle: document.getElementById("yeshivishStyle").value
    };
    localStorage.setItem("torani-settings", JSON.stringify(settings));
    addLog("✓ הגדרות שמורות");
}

// ============================================
// UI Toggles
// ============================================
function toggleCollapsible(header) {
    const content = header.nextElementSibling;
    if (content && content.classList.contains("collapsible-content")) {
        content.classList.toggle("active");
        const arrow = header.querySelector("span");
        arrow.style.transform = content.classList.contains("active") ? "rotate(180deg)" : "rotate(0deg)";
    }
}

function toggleEditSystemPrompt() {
    const systemPromptField = document.getElementById("systemPrompt");
    const button = event.target;
    
    if (systemPromptField.readOnly) {
        systemPromptField.readOnly = false;
        systemPromptField.style.backgroundColor = "#fffacd";
        button.textContent = "שמור את ה-System Prompt";
        addLog("⚠️ מצב עריכה - System Prompt");
    } else {
        systemPromptField.readOnly = true;
        systemPromptField.style.backgroundColor = "#fff";
        button.textContent = "ערוך את ה-System Prompt";
        saveSettings();
        addLog("✓ System Prompt נשמר");
    }
}

function previewFullDocument() {
    const previewEnabled = document.getElementById("previewMode").checked;
    if (!previewEnabled) {
        showAlert("⚠️ יש להפעיל את מצב התצוגה המקדימה", "warning");
        return;
    }
    
    Word.run(async (context) => {
        try {
            const body = context.document.body;
            body.load("text");
            await context.sync();
            
            // For now, show the current document text as preview
            const text = body.text;
            document.getElementById("fullDocPreview").innerHTML = `<p>${escapeHtml(text)}</p>`;
            document.getElementById("fullDocPreviewSection").style.display = "block";
            
            addLog("✓ תצוגה מלאה של המסמך מוכנה");
        } catch (error) {
            addLog(`❌ שגיאה בתצוגה: ${error.message}`);
        }
    });
}

function cancelFullDocPreview() {
    document.getElementById("fullDocPreviewSection").style.display = "none";
}

function applyFullDocChanges() {
    // This will be called after the user confirms
    cancelFullDocPreview();
    showAlert("✓ השינויים יחולו על כל המסמך", "success");
}

function cancelPreview() {
    document.getElementById("previewSection").style.display = "none";
}

function applyAllChanges() {
    if (currentPreviewChanges.length === 0) return;
    
    // Apply all changes
    Word.run(async (context) => {
        try {
            // Implement actual changes here
            await applyChangesToDocument(currentPreviewChanges, context);
            cancelPreview();
            showAlert("✓ כל השינויים הוחלו בהצלחה", "success");
            addLog(`✓ הוחלו ${currentPreviewChanges.length} שינויים`);
        } catch (error) {
            addLog(`❌ שגיאה בהחלת שינויים: ${error.message}`);
        }
    });
}

// ============================================
// Main Processing Functions
// ============================================
async function processSelection() {
    saveSettings();
    
    const apiKey = document.getElementById("apiKey").value;
    if (!apiKey) {
        showAlert("❌ יש להכניס API Key", "error");
        return;
    }
    
    const checkedActions = getSelectedActions();
    if (checkedActions.length === 0) {
        showAlert("❌ בחר לפחות פעולה אחת", "error");
        return;
    }
    
    Word.run(async (context) => {
        try {
            addLog("⏳ קורא טקסט נבחר...");
            
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            
            const selectedText = selection.text;
            if (!selectedText || selectedText.trim() === "") {
                showAlert("❌ בחר טקסט ראשון", "error");
                return;
            }
            
            addLog(`✓ נבחר טקסט (${selectedText.length} תווים)`);
            
            // Process the selection
            await processTextWithAI(selectedText, checkedActions, context);
            
        } catch (error) {
            addLog(`❌ שגיאה: ${error.message}`);
            showAlert(`❌ ${error.message}`, "error");
        }
    });
}

async function processDocument() {
    saveSettings();
    
    const apiKey = document.getElementById("apiKey").value;
    if (!apiKey) {
        showAlert("❌ יש להכניס API Key", "error");
        return;
    }
    
    const checkedActions = getSelectedActions();
    if (checkedActions.length === 0) {
        showAlert("❌ בחר לפחות פעולה אחת", "error");
        return;
    }
    
    Word.run(async (context) => {
        try {
            addLog("⏳ קורא את כל המסמך...");
            
            const body = context.document.body;
            body.load("text");
            await context.sync();
            
            const fullText = body.text;
            if (!fullText || fullText.trim() === "") {
                showAlert("❌ המסמך ריק", "error");
                return;
            }
            
            addLog(`✓ מסמך טעון (${fullText.length} תווים)`);
            
            // Process the entire document
            await processTextWithAI(fullText, checkedActions, context, true);
            
        } catch (error) {
            addLog(`❌ שגיאה: ${error.message}`);
            showAlert(`❌ ${error.message}`, "error");
        }
    });
}

// ============================================
// AI Processing
// ============================================
async function processTextWithAI(text, actions, context, isFullDocument = false) {
    const apiKey = document.getElementById("apiKey").value;
    const model = document.getElementById("modelSelect").value;
    let systemPrompt = document.getElementById("systemPrompt").value;
    const freeTextRequest = document.getElementById("freeTextRequest").value;
    const yeshivishStyle = document.getElementById("yeshivishStyle").value;
    
    // Build the system prompt based on selected actions
    let enhancedSystemPrompt = systemPrompt;
    
    if (actions.includes("yeshivish")) {
        enhancedSystemPrompt += "\n\n[YESHIVISH STYLE INSTRUCTIONS]:\n" + yeshivishStyle;
    }
    
    // Build user prompt
    let userPrompt = "Please process the following Hebrew text:\n\n";
    
    if (freeTextRequest) {
        userPrompt += `Special Request: ${freeTextRequest}\n\n`;
    }
    
    userPrompt += "Apply the following operations:\n";
    
    if (actions.includes("yeshivish")) {
        userPrompt += "- Convert to Yeshivish/academic style\n";
    }
    if (actions.includes("punctuation")) {
        userPrompt += "- Fix punctuation according to Torah publication standards\n";
    }
    if (actions.includes("formatting")) {
        userPrompt += "- Improve formatting and layout\n";
    }
    if (actions.includes("vocalization")) {
        userPrompt += "- Add Nikkud (vowel marks) where needed\n";
    } else {
        userPrompt += "- DO NOT add any Nikkud under any circumstances\n";
    }
    
    userPrompt += `\nText to process:\n"${text}"`;
    
    try {
        addLog("⏳ שולח בקשה ל-Gemini AI...");
        
        const response = await callGeminiAPI(apiKey, model, systemPrompt, userPrompt);
        const processedText = extractTextFromResponse(response);
        
        addLog("✓ קיבלנו תשובה מ-AI");
        
        if (document.getElementById("previewMode").checked) {
            // Show preview
            showPreviewTable(text, processedText, actions);
        } else {
            // Apply directly
            await applyChangesToDocument([{ original: text, proposed: processedText }], context);
            applyFormatting(context);
            showAlert("✓ השינויים הוחלו בהצלחה", "success");
            addLog("✓ שינויים הוחלו ישירות");
        }
        
    } catch (error) {
        addLog(`❌ שגיאה: ${error.message}`);
        showAlert(`❌ ${error.message}`, "error");
    }
}

async function callGeminiAPI(apiKey, model, systemPrompt, userPrompt) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    const payload = {
        contents: [
            {
                role: "user",
                parts: [
                    {
                        text: `${systemPrompt}\n\n${userPrompt}`
                    }
                ]
            }
        ]
    };
    
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify(payload)
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(`API Error: ${error.error?.message || response.statusText}`);
    }
    
    return await response.json();
}

function extractTextFromResponse(response) {
    if (response.candidates && response.candidates[0] && response.candidates[0].content) {
        const parts = response.candidates[0].content.parts;
        if (parts && parts[0]) {
            return parts[0].text;
        }
    }
    throw new Error("Invalid API response format");
}

// ============================================
// Preview & Application
// ============================================
function showPreviewTable(original, proposed, actions) {
    currentPreviewChanges = [{ original, proposed, actions }];
    
    const previewTable = document.getElementById("previewTable");
    previewTable.innerHTML = "";
    
    // Header row
    const headerRow = document.createElement("div");
    headerRow.className = "preview-row";
    headerRow.innerHTML = `
        <div class="preview-col preview-col-header">טקסט מקורי</div>
        <div class="preview-col preview-col-header">הטקסט המוצע</div>
    `;
    previewTable.appendChild(headerRow);
    
    // Data row
    const dataRow = document.createElement("div");
    dataRow.className = "preview-row";
    dataRow.innerHTML = `
        <div class="preview-col preview-col-original">${escapeHtml(original)}</div>
        <div class="preview-col preview-col-proposed">${escapeHtml(proposed)}</div>
    `;
    previewTable.appendChild(dataRow);
    
    document.getElementById("previewSection").style.display = "block";
    addLog("✓ תצוגה מקדימה מוכנה");
    
    // Show alerts if enabled
    checkAndShowAlerts(actions);
}

async function applyChangesToDocument(changes, context) {
    try {
        for (const change of changes) {
            const body = context.document.body;
            
            // Search for the original text and replace it
            const searchResults = body.getRange("Start").search(change.original, { matchCase: false });
            searchResults.load("items");
            await context.sync();
            
            if (searchResults.items.length > 0) {
                searchResults.items[0].insertText(change.proposed, Word.InsertLocation.replace);
            }
        }
        
        await context.sync();
        
    } catch (error) {
        console.error("Error applying changes:", error);
    }
}

function applyFormatting(context) {
    Word.run(async (context) => {
        try {
            const body = context.document.body;
            
            // Set font to Frank-Riehl or David
            body.font.name = "David";
            body.font.size = 11;
            
            // Set line spacing to 1.5
            body.lineSpacing = 1.5;
            
            // Set full justification
            body.alignment = Word.Alignment.justified;
            
            await context.sync();
            addLog("✓ עיצוב הוחל: David, 1.5 line spacing, justified");
            
        } catch (error) {
            addLog(`⚠️ עיצוב חלקי: ${error.message}`);
        }
    });
}

// ============================================
// Helper Functions
// ============================================
function getSelectedActions() {
    const actions = [];
    
    if (document.getElementById("checkYeshivish").checked) actions.push("yeshivish");
    if (document.getElementById("checkPunctuation").checked) actions.push("punctuation");
    if (document.getElementById("checkFormatting").checked) actions.push("formatting");
    if (document.getElementById("checkVocalization").checked) actions.push("vocalization");
    
    return actions;
}

function checkAndShowAlerts(actions) {
    if (!document.getElementById("alertEnabled").checked) return;
    
    const alerts = [];
    
    if (actions.includes("yeshivish") && document.getElementById("alertYeshivish").checked) {
        alerts.push("סגנון יישיבה");
    }
    if (actions.includes("punctuation") && document.getElementById("alertPunctuation").checked) {
        alerts.push("פיסוק");
    }
    if (actions.includes("formatting") && document.getElementById("alertFormatting").checked) {
        alerts.push("עיצוב");
    }
    if (actions.includes("vocalization") && document.getElementById("alertVocalization").checked) {
        alerts.push("ניקוד");
    }
    
    if (alerts.length > 0) {
        showAlert(`✓ שינויים בקטגוריות: ${alerts.join(", ")}`, "info");
    }
}

function showAlert(message, type = "info") {
    const status = document.getElementById("status");
    status.innerHTML = `<p>[${type.toUpperCase()}] ${message}</p>`;
    
    setTimeout(() => {
        status.innerHTML = "<p>מוכן לעריכה...</p>";
    }, 5000);
}

function addLog(message) {
    const logsBox = document.getElementById("logs");
    const entry = document.createElement("div");
    entry.className = "log-entry";
    const timestamp = new Date().toLocaleTimeString("he-IL");
    entry.textContent = `[${timestamp}] ${message}`;
    logsBox.insertBefore(entry, logsBox.firstChild);
    
    // Keep only last 20 logs
    while (logsBox.children.length > 20) {
        logsBox.removeChild(logsBox.lastChild);
    }
}

function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

// ============================================
// Footnote Conversion
// ============================================
function convertParenthesesAndToFootnotes() {
    Word.run(async (context) => {
        try {
            const body = context.document.body;
            body.load("text");
            await context.sync();
            
            const text = body.text;
            const footnotePattern = /\(([^)]+)\)/g;
            
            // This is a simplified implementation
            // In production, you would parse more carefully
            
            let match;
            while ((match = footnotePattern.exec(text)) !== null) {
                const footnoteText = match[1];
                // Create footnote logic here
                addLog(`📌 זוהה הערה שולית: "${footnoteText}"`);
            }
            
        } catch (error) {
            addLog(`⚠️ שגיאה בהמרת הערות שוליות: ${error.message}`);
        }
    });
}

// ============================================
// Export & Persistence
// ============================================
function exportSettings() {
    saveSettings();
    const settings = localStorage.getItem("torani-settings");
    const dataStr = JSON.stringify(JSON.parse(settings), null, 2);
    const dataBlob = new Blob([dataStr], { type: "application/json" });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "torani-settings.json";
    link.click();
    addLog("✓ הגדרות יוצאו");
}

// Make functions globally accessible
window.processSelection = processSelection;
window.processDocument = processDocument;
window.previewFullDocument = previewFullDocument;
window.applyAllChanges = applyAllChanges;
window.cancelPreview = cancelPreview;
window.applyFullDocChanges = applyFullDocChanges;
window.cancelFullDocPreview = cancelFullDocPreview;
window.toggleCollapsible = toggleCollapsible;
window.toggleEditSystemPrompt = toggleEditSystemPrompt;
