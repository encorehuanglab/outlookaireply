/* * Global variable to store the AI suggestion 
 * for insertion into the email body later.
 */
let currentAiSuggestion = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";

    // Set up Save Key button
    document.getElementById("save-key-btn").onclick = saveApiKey;
    
    // Pre-fill input if key exists
    const savedKey = localStorage.getItem("gemini_api_key");
    if (savedKey) {
        document.getElementById("api-key-input").value = savedKey;
    }


    initApp();
  }
});

function saveApiKey() {
    const key = document.getElementById("api-key-input").value;
    if (key) {
        localStorage.setItem("gemini_api_key", key);
        document.getElementById("status").innerText = "API Key saved locally!";
    } else {
        alert("Please enter a valid key.");
    }
}

/**
 * Detects whether the user is in Read mode or Compose (Reply) mode
 * and initializes the UI buttons and logic accordingly.
 */
function initApp() {
  const item = Office.context.mailbox.item;
  const runBtn = document.getElementById("run");
  const statusIndicator = document.getElementById("status");
  const insertBtn = document.getElementById("insert-btn");

  // Check if the current context allows setting data (Compose Mode)
  if (item.body.setSelectedDataAsync) {
    /* --- COMPOSE / REPLY MODE --- */
    statusIndicator.innerText = "Reply mode detected.";
    runBtn.innerText = "Generate AI Reply";
    runBtn.onclick = handleReplyMode;
    
    // Hide insert button until a suggestion is generated
    insertBtn.style.display = "none"; 
    insertBtn.onclick = insertToEmail;
  } else {
    /* --- READ MODE --- */
    statusIndicator.innerText = "Read mode detected.";
    runBtn.innerText = "Summarize & Translate";
    runBtn.onclick = handleReadMode;
    insertBtn.style.display = "none";
  }
}

/**
 * Handles Logic for Read Mode:
 * Prompt: Summarize and translate to Traditional Chinese.
 */
async function handleReadMode() {
  const statusIndicator = document.getElementById("status");
  const suggestionDiv = document.getElementById("suggestion");
  
  statusIndicator.innerText = "Summarizing email...";
  suggestionDiv.innerText = "Processing...";

  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      try {
        const prompt = "Please breif summary the email and translate to traditional Chinese";
        const responseText = await callGeminiAPI(result.value, prompt);
        
        suggestionDiv.innerText = responseText;
        statusIndicator.innerText = "Summary completed.";
      } catch (error) {
        // statusIndicator.innerText = "Failed to summarize.";
        statusIndicator.innerText = error.message;
        console.error(error);
      }
    }
  });
}

/**
 * Handles Logic for Reply Mode:
 * Prompt: Suggest a brief, polite reply.
 */
async function handleReplyMode() {
  const statusIndicator = document.getElementById("status");
  const suggestionDiv = document.getElementById("suggestion");
  const insertBtn = document.getElementById("insert-btn");

  statusIndicator.innerText = "Generating suggestion...";
  suggestionDiv.innerText = "Processing...";

  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      try {
        const prompt = "Please revise or suggest a brief, polite reply to this email";
        const responseText = await callGeminiAPI(result.value, prompt);
        
        currentAiSuggestion = responseText;
        suggestionDiv.innerText = currentAiSuggestion;
        statusIndicator.innerText = "Suggestion ready.";
        insertBtn.style.display = "block"; // Show the button to insert text
      } catch (error) {
        // statusIndicator.innerText = "Failed to generate reply.";
        statusIndicator.innerText = error.message;
        console.error(error);
      }
    }
  });
}

/**
 * Inserts the stored AI suggestion into the current email body.
 */
function insertToEmail() {
  if (!currentAiSuggestion) return;

  Office.context.mailbox.item.body.setSelectedDataAsync(
    currentAiSuggestion,
    { coercionType: Office.CoercionType.Text },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("status").innerText = "Successfully inserted!";
      } else {
        console.error("Insertion failed: " + result.error.message);
      }
    }
  );
}

/**
 * Core function to communicate with the Google Gemini API.
 * Uses gemini-2.0-flash for high speed and low latency.
 */
async function callGeminiAPI(emailText, systemPrompt) {
  // const API_KEY = "your_api_key_here"; // Replace with your actual API key or retrieve from secure storage
  // Retrieve the key from local storage instead of hardcoding it
  const API_KEY = localStorage.getItem("gemini_api_key");
  if (!API_KEY) {
      throw new Error("API Key not found. Please enter it in settings.");
  }
  // Updated to the latest stable flash model (2.0)
  const MODEL = "gemini-2.5-flash"; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${API_KEY}`;

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{
        parts: [{
          text: `${systemPrompt}\n\nEmail Content:\n${emailText}`
        }]
      }]
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error.message || "API Request Failed");
  }

  const data = await response.json();
  
  // Extract text from Gemini response structure
  return data.candidates[0].content.parts[0].text;
}