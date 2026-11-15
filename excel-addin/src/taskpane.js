/* global Office, Excel */

// State
let conversationId = null;
let currentSelectionData = null;
let isBusy = false;

// DOM Elements
const chatContainer = document.getElementById("chatContainer");
const welcomeBox = document.getElementById("welcomeBox");
const userInput = document.getElementById("userInput");
const btnSend = document.getElementById("btnSend");
const btnNewChat = document.getElementById("btnNewChat");
const btnRefreshSelection = document.getElementById("btnRefreshSelection");
const btnAnalyze = document.getElementById("btnAnalyze");
const btnExplainFormula = document.getElementById("btnExplainFormula");
const btnSuggestFormula = document.getElementById("btnSuggestFormula");
const currentSelectionEl = document.getElementById("currentSelection");
const notificationArea = document.getElementById("notificationArea");

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    setupEventListeners();
    readActiveSelection(false);
  } else {
    // Non-Excel fallback / browser preview
    setupEventListeners();
    showNotification("Running in standalone preview mode (Office.js host not detected).", "info");
  }
});

function setupEventListeners() {
  userInput.addEventListener("input", () => {
    btnSend.disabled = !userInput.value.trim() || isBusy;
  });

  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  btnSend.addEventListener("click", () => sendMessage());
  btnNewChat.addEventListener("click", () => startNewChat());
  btnRefreshSelection.addEventListener("click", () => readActiveSelection(true));
  btnAnalyze.addEventListener("click", () => handleQuickAction("analyze"));
  btnExplainFormula.addEventListener("click", () => handleQuickAction("explain_formula"));
  btnSuggestFormula.addEventListener("click", () => handleQuickAction("generate_formula"));

  document.querySelectorAll(".prompt-chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const prompt = chip.getAttribute("data-prompt");
      if (prompt) {
        userInput.value = prompt;
        btnSend.disabled = false;
        sendMessage();
      }
    });
  });
}

function showNotification(text, type = "error") {
  notificationArea.innerHTML = `<div class="notification ${type}">${escapeHtml(text)}</div>`;
  setTimeout(() => {
    notificationArea.innerHTML = "";
  }, 6000);
}

function escapeHtml(str) {
  if (typeof str !== "string") return "";
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function startNewChat() {
  conversationId = null;
  chatContainer.innerHTML = "";
  if (welcomeBox) chatContainer.appendChild(welcomeBox);
  userInput.value = "";
  btnSend.disabled = true;
  showNotification("New conversation started.", "success");
}

// 1. Read Excel Selection using Excel JavaScript API
async function readActiveSelection(notify = false) {
  if (typeof Excel === "undefined") {
    currentSelectionEl.textContent = "Excel API not loaded";
    return null;
  }

  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();

      sheet.load("name");
      range.load(["address", "rowCount", "columnCount", "values", "formulas"]);

      await context.sync();

      // Guard against massive payloads (>15,000 cells)
      const totalCells = range.rowCount * range.columnCount;
      if (totalCells > 15000) {
        showNotification(`Selected range has ${totalCells} cells. Limiting preview for performance.`, "error");
      }

      // Extract headers from first row if multi-row selection
      let headers = [];
      if (range.rowCount > 1 && range.values && range.values.length > 0) {
        headers = range.values[0].map((v) => (v != null ? String(v) : ""));
      }

      currentSelectionData = {
        worksheet: sheet.name,
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        headers: headers,
        preview: range.values.slice(0, 100).map((row) => row.slice(0, 50)),
        formulas: range.formulas.slice(0, 50).map((row) => row.slice(0, 50)),
      };

      currentSelectionEl.textContent = `${sheet.name}!${range.address} (${range.rowCount}x${range.columnCount})`;
      if (notify) {
        showNotification(`Read ${range.address} successfully.`, "success");
      }
      return currentSelectionData;
    });
  } catch (err) {
    currentSelectionEl.textContent = "No range selected";
    if (notify) {
      showNotification(`Failed to read selection: ${err.message || err}`, "error");
    }
    return null;
  }
}

// Quick Actions
async function handleQuickAction(actionType) {
  await readActiveSelection(false);

  if (!currentSelectionData) {
    showNotification("Please select a cell or range in Excel first.", "error");
    return;
  }

  if (actionType === "analyze") {
    userInput.value = `Analyze the data in ${currentSelectionData.address} and provide key insights, statistics, and any anomalies.`;
    btnSend.disabled = false;
    sendMessage("analyze");
  } else if (actionType === "explain_formula") {
    // Check if cell has a formula
    let formula = null;
    if (currentSelectionData.formulas && currentSelectionData.formulas[0] && currentSelectionData.formulas[0][0]) {
      const f = currentSelectionData.formulas[0][0];
      if (typeof f === "string" && f.startsWith("=")) {
        formula = f;
      }
    }

    if (formula) {
      userInput.value = `Explain this Excel formula step by step: ${formula}`;
    } else {
      userInput.value = `Inspect the formula in ${currentSelectionData.address}. If it's a value, explain how to compute it dynamically.`;
    }
    btnSend.disabled = false;
    sendMessage("explain_formula");
  } else if (actionType === "generate_formula") {
    userInput.value = `Write an optimized Excel formula for ${currentSelectionData.address}.`;
    userInput.focus();
    btnSend.disabled = false;
  }
}

// Insert formula into active Excel cell
async function insertFormulaIntoCell(formulaText) {
  if (typeof Excel === "undefined") {
    showNotification("Excel JavaScript API is not available.", "error");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values", "formulas"]);
      await context.sync();

      // Check if target already has content
      const existing = range.values[0][0];
      if (existing !== "" && existing != null) {
        const confirmOverwrite = confirm(
          `Cell ${range.address} already contains "${existing}". Overwrite with formula: ${formulaText}?`
        );
        if (!confirmOverwrite) return;
      }

      // Ensure starts with =
      let cleanFormula = formulaText.trim();
      if (!cleanFormula.startsWith("=")) cleanFormula = "=" + cleanFormula;

      range.formulas = [[cleanFormula]];
      await context.sync();

      showNotification(`Successfully inserted formula ${cleanFormula} into ${range.address}!`, "success");
    });
  } catch (err) {
    showNotification(`Error inserting formula: ${err.message || err}`, "error");
  }
}

// Append Chat Message Bubble
function appendMessage(role, text, formula = null) {
  if (welcomeBox && welcomeBox.parentNode === chatContainer) {
    welcomeBox.remove();
  }

  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${role}`;

  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.textContent = text;
  msgDiv.appendChild(bubble);

  // If a formula was detected/returned, render formula insertion action card
  if (formula) {
    const formulaCard = document.createElement("div");
    formulaCard.className = "formula-card";
    formulaCard.innerHTML = `
      <div class="formula-card-title">
        <span>DETECTED EXCEL FORMULA</span>
        <span>✨ Gridex</span>
      </div>
      <div class="formula-code">${escapeHtml(formula)}</div>
      <div class="formula-actions">
        <button class="btn-insert-formula">Insert into Active Cell</button>
      </div>
    `;

    formulaCard.querySelector(".btn-insert-formula").addEventListener("click", () => {
      insertFormulaIntoCell(formula);
    });

    msgDiv.appendChild(formulaCard);
  }

  chatContainer.appendChild(msgDiv);
  chatContainer.scrollTop = chatContainer.scrollHeight;
  return msgDiv;
}

// Send Message to Next.js API
async function sendMessage(action = "general") {
  const text = userInput.value.trim();
  if (!text || isBusy) return;

  userInput.value = "";
  btnSend.disabled = true;
  isBusy = true;

  // Append user message to UI
  appendMessage("user", text);

  // Append thinking placeholder
  const thinkingBubble = appendMessage("assistant", "Thinking...");

  try {
    // Read current selection if not already cached
    if (!currentSelectionData) {
      await readActiveSelection(false);
    }

    const payload = {
      message: text,
      conversationId: conversationId,
      excelContext: currentSelectionData
        ? {
            ...currentSelectionData,
            action,
          }
        : undefined,
    };

    const backendUrl = "/api/excel";

    const response = await fetch(backendUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    const data = await response.json().catch(() => ({ error: "Server response was unreadable." }));

    if (!response.ok) {
      throw new Error(data.error || `Server returned status ${response.status}`);
    }

    // Remove thinking message
    thinkingBubble.remove();

    if (data.conversationId) {
      conversationId = data.conversationId;
    }

    // Append AI response
    appendMessage("assistant", data.content || "No response received.", data.detectedFormula);
  } catch (err) {
    thinkingBubble.remove();
    appendMessage("assistant", `⚠️ Error: ${err.message || "Failed to reach ASIF AI backend."}`);
    showNotification(err.message || "Request failed", "error");
  } finally {
    isBusy = false;
    btnSend.disabled = !userInput.value.trim();
    userInput.focus();
  }
}
