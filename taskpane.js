let geminiResult = "";
let chatgptResult = "";
let currentProvider = "gemini";
let isComposeMode = false;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        const body = document.getElementById("app-body");
        if (body) body.style.display = "flex";
        
        setupTabs();
        setupKeySaving();
        loadKeys();
        initApp();
    }
});

function setupTabs() {
    const gBtn = document.getElementById("tab-gemini");
    const cBtn = document.getElementById("tab-chatgpt");
    const gSet = document.getElementById("settings-gemini");
    const cSet = document.getElementById("settings-chatgpt");
    const gCont = document.getElementById("content-gemini");
    const cCont = document.getElementById("content-chatgpt");

    if (gBtn && cBtn) {
        gBtn.onclick = () => {
            currentProvider = "gemini";
            updateTabUI(gBtn, cBtn, gSet, cSet, gCont, cCont);
        };
        cBtn.onclick = () => {
            currentProvider = "chatgpt";
            updateTabUI(cBtn, gBtn, cSet, gSet, cCont, gCont);
        };
    }
}

function updateTabUI(activeBtn, inBtn, activeSet, inSet, activeCont, inCont) {
    activeBtn.classList.add("active");
    inBtn.classList.remove("active");
    activeSet.style.display = "block";
    inSet.style.display = "none";
    activeCont.style.display = "block";
    inCont.style.display = "none";
    document.getElementById("key-status").innerText = ""; // Clear status on tab switch
    updateInsertBtn();
}

function showKeySaved() {
    const status = document.getElementById("key-status");
    status.innerText = "Key Saved Locally!";
    setTimeout(() => { status.innerText = ""; }, 3000);
}

function setupKeySaving() {
    const gSave = document.getElementById("save-gemini-btn");
    const cSave = document.getElementById("save-chatgpt-btn");

    if (gSave) {
        gSave.onclick = () => {
            localStorage.setItem("gemini_key", document.getElementById("gemini-key-input").value);
            showKeySaved();
        };
    }
    if (cSave) {
        cSave.onclick = () => {
            localStorage.setItem("chatgpt_key", document.getElementById("chatgpt-key-input").value);
            showKeySaved();
        };
    }
}

function loadKeys() {
    const gIn = document.getElementById("gemini-key-input");
    const cIn = document.getElementById("chatgpt-key-input");
    if (gIn) gIn.value = localStorage.getItem("gemini_key") || "";
    if (cIn) cIn.value = localStorage.getItem("chatgpt_key") || "";
}

function initApp() {
    const item = Office.context.mailbox.item;
    isComposeMode = !!item.body.setSelectedDataAsync;
    // document.getElementById("status").innerText = isComposeMode ? "Compose Mode" : "Read Mode";
    
    document.getElementById("run").innerText = isComposeMode ? "Suggest Reply" : "Summarize & Translate";

    document.getElementById("run").onclick = () => {
        const prompt = isComposeMode ? "Follow the instructions to suggest the mail:" : "Summarize this and translate to Traditional Chinese:";
        startAI(prompt);
    };

    document.getElementById("insert-btn").onclick = () => {
        const text = (currentProvider === "gemini") ? geminiResult : chatgptResult;
        Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
    };
}

async function startAI(prompt) {
    const box = document.getElementById(`suggestion-${currentProvider}`);
    box.classList.remove("result-mode");
    box.innerText = "Thinking...";
    document.getElementById("insert-btn").style.display = "none";

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
        try {
            const response = (currentProvider === "gemini") ? 
                await callGemini(result.value, prompt) : await callChatGPT(result.value, prompt);
            
            if (currentProvider === "gemini") geminiResult = response;
            else chatgptResult = response;

            box.classList.add("result-mode");
            box.innerText = response;
            updateInsertBtn();
        } catch (e) {
            box.innerText = "Error: " + e.message;
        }
    });
}

function updateInsertBtn() {
    const res = (currentProvider === "gemini") ? geminiResult : chatgptResult;
    document.getElementById("insert-btn").style.display = (isComposeMode && res) ? "block" : "none";
}

async function callGemini(text, prompt) {
    const key = localStorage.getItem("gemini_key");
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`;
    const res = await fetch(url, { method: "POST", body: JSON.stringify({ contents: [{ parts: [{ text: prompt + "\n" + text }] }] }) });
    const data = await res.json();
    return data.candidates[0].content.parts[0].text;
}

async function callChatGPT(text, prompt) {
    const key = localStorage.getItem("chatgpt_key");
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${key}` },
        body: JSON.stringify({ model: "gpt-4o-mini", messages: [{ role: "user", content: prompt + "\n" + text }] })
    });
    const data = await res.json();
    return data.choices[0].message.content;
}