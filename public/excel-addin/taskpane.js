/* global Office, Excel */

// State
let conversationId = null;
let currentSelectionData = null;
let currentWorkbookData = null;
let isBusy = false;
let currentMode = "agent"; // "agent" | "chat"

// DOM Elements
const chatContainer = document.getElementById("chatContainer");
const welcomeBox = document.getElementById("welcomeBox");
const userInput = document.getElementById("userInput");
const btnSend = document.getElementById("btnSend");
const btnNewChat = document.getElementById("btnNewChat");
const btnScanWorkbook = document.getElementById("btnScanWorkbook");
const btnMgmtSummary = document.getElementById("btnMgmtSummary");
const btnVarianceReport = document.getElementById("btnVarianceReport");
const btnFindDuplicates = document.getElementById("btnFindDuplicates");
const currentSelectionEl = document.getElementById("currentSelection");
const notificationArea = document.getElementById("notificationArea");
const modeChat = document.getElementById("modeChat");
const modeAgent = document.getElementById("modeAgent");
const agentModeIndicator = document.getElementById("agentModeIndicator");

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    setupEventListeners();
    scanWorkbookContext(false);
  } else {
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
  btnScanWorkbook.addEventListener("click", () => scanWorkbookContext(true));

  // Quick Agent Actions
  btnMgmtSummary.addEventListener("click", () => {
    userInput.value = "Analyze all data in this workbook. Create a new sheet named Management Summary. Show total production, order-wise performance, monthly trends, top 10 orders, exceptions and charts.";
    btnSend.disabled = false;
    sendMessage();
  });

  btnVarianceReport.addEventListener("click", () => {
    userInput.value = "Analyze the active worksheet. Highlight negative variances, compute efficiency %, and generate a variance summary table.";
    btnSend.disabled = false;
    sendMessage();
  });

  btnFindDuplicates.addEventListener("click", () => {
    userInput.value = "Scan for duplicate orders or duplicate rows in this sheet and create an Audit Report.";
    btnSend.disabled = false;
    sendMessage();
  });

  // Mode Toggles
  modeChat.addEventListener("click", () => setAppMode("chat"));
  modeAgent.addEventListener("click", () => setAppMode("agent"));

  // Prompt Chips
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

function setAppMode(mode) {
  currentMode = mode;
  if (mode === "agent") {
    modeAgent.classList.add("active");
    modeChat.classList.remove("active");
    agentModeIndicator.style.display = "inline-block";
    userInput.placeholder = "Instruct ASIF AI (e.g. 'Create Management Summary with charts')...";
  } else {
    modeChat.classList.add("active");
    modeAgent.classList.remove("active");
    agentModeIndicator.style.display = "none";
    userInput.placeholder = "Ask ASIF AI about Excel formulas or analysis...";
  }
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
  showNotification("Session reset.", "success");
}

// Custom Modal Dialog for Excel Task Pane (replaces unsupported window.confirm)
function showCustomConfirmModal({ title, message, details, confirmLabel = "Confirm", cancelLabel = "Cancel" }) {
  return new Promise((resolve) => {
    const existingModal = document.getElementById("asifConfirmModal");
    if (existingModal) existingModal.remove();

    const overlay = document.createElement("div");
    overlay.id = "asifConfirmModal";
    overlay.className = "modal-overlay";

    overlay.innerHTML = `
      <div class="modal-content">
        <div class="modal-header">
          <span class="modal-icon">⚠️</span>
          <span class="modal-title">${escapeHtml(title || "Confirmation Required")}</span>
        </div>
        <div class="modal-body">
          <p>${escapeHtml(message)}</p>
          ${details ? `<div class="modal-target-cell">${details}</div>` : ""}
        </div>
        <div class="modal-actions">
          <button id="modalBtnCancel" class="modal-btn modal-btn-cancel">${escapeHtml(cancelLabel)}</button>
          <button id="modalBtnConfirm" class="modal-btn modal-btn-confirm">${escapeHtml(confirmLabel)}</button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    const btnConfirm = overlay.querySelector("#modalBtnConfirm");
    const btnCancel = overlay.querySelector("#modalBtnCancel");

    const cleanup = (result) => {
      overlay.remove();
      resolve(result);
    };

    btnConfirm.addEventListener("click", () => cleanup(true));
    btnCancel.addEventListener("click", () => cleanup(false));
  });
}

// 1. Workbook Intelligence Engine (Scan Entire Workbook & Worksheets)
async function scanWorkbookContext(notify = false) {
  if (typeof Excel === "undefined") {
    currentSelectionEl.textContent = "Excel API not loaded";
    return null;
  }

  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      sheets.load("items/name");
      activeSheet.load("name");

      await context.sync();

      const worksheetSummaries = [];

      for (const sheet of sheets.items) {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["address", "rowCount", "columnCount", "values", "formulas"]);
        await context.sync();

        if (usedRange.isNullObject || usedRange.rowCount === 0) {
          worksheetSummaries.push({
            name: sheet.name,
            rowCount: 0,
            columnCount: 0,
            address: "Empty",
            headers: [],
            inferredTypes: {},
            sampleRows: [],
          });
          continue;
        }

        // Infer headers
        let headers = [];
        let inferredTypes = {};
        if (usedRange.rowCount > 0 && usedRange.values && usedRange.values.length > 0) {
          headers = usedRange.values[0].map((v) => (v != null ? String(v) : ""));

          // Sample data rows to infer column types (Numeric, Date, Text)
          if (usedRange.rowCount > 1) {
            headers.forEach((hdr, colIdx) => {
              for (let r = 1; r < Math.min(usedRange.rowCount, 10); r++) {
                const val = usedRange.values[r][colIdx];
                if (val !== null && val !== "") {
                  if (typeof val === "number") {
                    inferredTypes[hdr] = "number";
                    break;
                  } else if (typeof val === "string" && !isNaN(Date.parse(val)) && (val.includes("-") || val.includes("/"))) {
                    inferredTypes[hdr] = "date";
                    break;
                  } else {
                    inferredTypes[hdr] = "text";
                  }
                }
              }
            });
          }
        }

        // Chunked sampling: take up to 25 sample rows and 30 columns to protect memory
        const sampleRows = usedRange.values
          .slice(0, 25)
          .map((row) => row.slice(0, 30));

        worksheetSummaries.push({
          name: sheet.name,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          address: usedRange.address,
          headers: headers,
          inferredTypes: inferredTypes,
          sampleRows: sampleRows,
        });
      }

      currentWorkbookData = {
        activeWorksheet: activeSheet.name,
        worksheets: worksheetSummaries,
      };

      currentSelectionEl.textContent = `${sheets.items.length} Sheet(s) Scanned (Active: ${activeSheet.name})`;
      if (notify) {
        showNotification(`Scanned ${sheets.items.length} worksheet(s) successfully!`, "success");
      }
      return currentWorkbookData;
    });
  } catch (err) {
    currentSelectionEl.textContent = "Scan failed";
    if (notify) {
      showNotification(`Failed to scan workbook: ${err.message || err}`, "error");
    }
    return null;
  }
}

// 2. Action Engine: Execute Plan through Office.js sequentially
async function executeActionPlan(actionPlan, planCardEl) {
  if (!actionPlan || !actionPlan.actions || !Array.isArray(actionPlan.actions)) {
    showNotification("No actionable steps found in plan.", "error");
    return;
  }

  if (typeof Excel === "undefined") {
    showNotification("Excel JavaScript API is not available.", "error");
    return;
  }

  isBusy = true;
  btnSend.disabled = true;

  const stepsListEl = planCardEl.querySelector(".plan-steps");
  const stepItems = stepsListEl ? stepsListEl.querySelectorAll(".plan-step") : [];

  const updateStepStatus = (index, status, text) => {
    if (stepItems[index]) {
      stepItems[index].className = `plan-step ${status}`;
      const icon = status === "completed" ? "✅" : status === "in-progress" ? "⏳" : status === "failed" ? "❌" : "•";
      stepItems[index].querySelector(".step-icon").textContent = icon;
      if (text) {
        stepItems[index].querySelector(".step-desc").textContent = text;
      }
    }
  };

  try {
    for (let i = 0; i < actionPlan.actions.length; i++) {
      const act = actionPlan.actions[i];
      updateStepStatus(i, "in-progress");

      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;

        if (act.type === "create_worksheet") {
          const sheetName = act.worksheet || "Summary";
          // Check if worksheet exists
          const existingSheet = sheets.getItemOrNullObject(sheetName);
          await context.sync();

          if (!existingSheet.isNullObject) {
            // Sheet already exists: activate it or prompt
            existingSheet.activate();
          } else {
            const newSheet = sheets.add(sheetName);
            newSheet.activate();
          }
          await context.sync();
        } else if (act.type === "write_range") {
          const sheet = act.worksheet ? sheets.getItem(act.worksheet) : sheets.getActiveWorksheet();
          const range = sheet.getRange(act.address);

          if (act.values && act.values.length > 0) {
            range.values = act.values;
          }
          if (act.formulas && act.formulas.length > 0) {
            range.formulas = act.formulas;
          }

          // Apply cell styling if specified
          if (act.format) {
            if (act.format.bold !== undefined) range.format.font.bold = act.format.bold;
            if (act.format.fontSize) range.format.font.size = act.format.fontSize;
            if (act.format.fontColor) range.format.font.color = act.format.fontColor;
            if (act.format.fillColor) range.format.fill.color = act.format.fillColor;
            if (act.format.horizontalAlignment) range.format.horizontalAlignment = act.format.horizontalAlignment;
            if (act.format.numberFormat) range.numberFormat = [[act.format.numberFormat]];
          }

          await context.sync();
        } else if (act.type === "format_range") {
          const sheet = act.worksheet ? sheets.getItem(act.worksheet) : sheets.getActiveWorksheet();
          const range = sheet.getRange(act.address);

          if (act.format) {
            if (act.format.bold !== undefined) range.format.font.bold = act.format.bold;
            if (act.format.fontSize) range.format.font.size = act.format.fontSize;
            if (act.format.fontColor) range.format.font.color = act.format.fontColor;
            if (act.format.fillColor) range.format.fill.color = act.format.fillColor;
            if (act.format.horizontalAlignment) range.format.horizontalAlignment = act.format.horizontalAlignment;
            if (act.format.numberFormat) range.numberFormat = [[act.format.numberFormat]];
          }
          await context.sync();
        } else if (act.type === "create_table") {
          const sheet = act.worksheet ? sheets.getItem(act.worksheet) : sheets.getActiveWorksheet();
          const range = sheet.getRange(act.address);
          sheet.tables.add(range, true /* hasHeaders */);
          await context.sync();
        } else if (act.type === "autofit") {
          const sheet = act.worksheet ? sheets.getItem(act.worksheet) : sheets.getActiveWorksheet();
          if (act.address) {
            sheet.getRange(act.address).format.autofitColumns();
          } else {
            sheet.getUsedRange().format.autofitColumns();
          }
          await context.sync();
        } else if (act.type === "create_chart" && act.chart) {
          const sheet = act.worksheet ? sheets.getItem(act.worksheet) : sheets.getActiveWorksheet();
          const dataRange = sheet.getRange(act.chart.dataRange);
          const chartType = act.chart.chartType || "ColumnClustered";

          const chart = sheet.charts.add(chartType, dataRange, "Auto");
          if (act.chart.title) chart.title.text = act.chart.title;
          if (act.chart.position) {
            chart.setPosition(act.chart.position, null, null);
          }
          if (act.chart.width) chart.width = act.chart.width;
          if (act.chart.height) chart.height = act.chart.height;

          await context.sync();
        } else if (act.type === "highlight_variance") {
          const sheet = act.worksheet ? sheets.getItem(act.worksheet) : sheets.getActiveWorksheet();
          const range = sheet.getRange(act.address);
          range.load(["values", "rowCount", "columnCount"]);
          await context.sync();

          // Highlight negative values in red
          for (let r = 0; r < range.rowCount; r++) {
            for (let c = 0; c < range.columnCount; c++) {
              const val = range.values[r][c];
              if (typeof val === "number" && val < 0) {
                const cell = range.getCell(r, c);
                cell.format.font.color = "#dc2626";
                cell.format.fill.color = "#fee2e2";
              }
            }
          }
          await context.sync();
        }
      });

      updateStepStatus(i, "completed");
    }

    showNotification("All planned Excel operations completed successfully!", "success");

    // Replace Execute button with Completed status
    const actionArea = planCardEl.querySelector(".plan-actions");
    if (actionArea) {
      actionArea.innerHTML = `<div style="color: var(--text-neon); font-size: 11px; font-weight: 700; display: flex; align-items: center; gap: 4px;">✅ Actions successfully executed in Excel.</div>`;
    }

    await scanWorkbookContext(false);
  } catch (err) {
    showNotification(`Execution stopped: ${err.message || err}`, "error");
  } finally {
    isBusy = false;
    btnSend.disabled = !userInput.value.trim();
  }
}

// Render Action Plan Card inside chat
function renderActionPlan(actionPlan) {
  const card = document.createElement("div");
  card.className = "plan-card";

  const actionsHtml = actionPlan.actions
    .map(
      (act, idx) => `
      <div class="plan-step">
        <span class="step-icon">○</span>
        <span class="step-desc">
          <strong>${escapeHtml(formatActionTypeName(act.type))}:</strong> 
          ${escapeHtml(act.worksheet ? `[${act.worksheet}] ` : "")}${escapeHtml(act.address || "")} 
          ${act.chart ? `(${act.chart.title || act.chart.chartType})` : ""}
        </span>
      </div>
    `
    )
    .join("");

  card.innerHTML = `
    <div class="plan-header">
      <div class="plan-title">⚡ AUTONOMOUS EXECUTION PLAN</div>
      <div class="plan-tag">${actionPlan.actions.length} ACTIONS</div>
    </div>
    <div class="plan-summary">${escapeHtml(actionPlan.summary || "Ready to execute planned Excel operations.")}</div>
    <div class="plan-steps">${actionsHtml}</div>
    <div class="plan-actions">
      <button class="btn-execute-plan">
        <span>⚡</span> Approve &amp; Execute in Excel
      </button>
      <button class="btn-cancel-plan">Dismiss</button>
    </div>
  `;

  const btnExecute = card.querySelector(".btn-execute-plan");
  const btnCancel = card.querySelector(".btn-cancel-plan");

  btnExecute.addEventListener("click", async () => {
    if (actionPlan.requiresConfirmation) {
      const confirmed = await showCustomConfirmModal({
        title: "Confirm Execution",
        message: "This plan involves creating or modifying worksheet content. Proceed?",
        confirmLabel: "Yes, Execute",
        cancelLabel: "Cancel",
      });
      if (!confirmed) {
        showNotification("Execution cancelled by user.", "info");
        return;
      }
    }
    executeActionPlan(actionPlan, card);
  });

  btnCancel.addEventListener("click", () => {
    card.remove();
  });

  return card;
}

function formatActionTypeName(type) {
  switch (type) {
    case "create_worksheet":
      return "Create Worksheet";
    case "write_range":
      return "Write Data & Formulas";
    case "format_range":
      return "Format Range";
    case "create_table":
      return "Create Excel Table";
    case "create_chart":
      return "Generate Chart";
    case "autofit":
      return "Autofit Columns";
    case "highlight_variance":
      return "Highlight Variance";
    default:
      return type;
  }
}

// Append Chat Message Bubble
function appendMessage(role, text, formula = null, actionPlan = null) {
  if (welcomeBox && welcomeBox.parentNode === chatContainer) {
    welcomeBox.remove();
  }

  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${role}`;

  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.textContent = text;
  msgDiv.appendChild(bubble);

  // If action plan exists, render execution plan card
  if (actionPlan) {
    const planCard = renderActionPlan(actionPlan);
    msgDiv.appendChild(planCard);
  }

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

// Insert formula into active Excel cell
async function insertFormulaIntoCell(formulaText) {
  if (typeof Excel === "undefined") {
    showNotification("Excel JavaScript API is not available.", "error");
    return;
  }

  try {
    let cleanFormula = formulaText.trim();
    if (!cleanFormula.startsWith("=")) cleanFormula = "=" + cleanFormula;

    let cellInfo = null;
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values", "formulas"]);
      await context.sync();

      const existingVal = range.values[0] ? range.values[0][0] : null;
      const existingFormula = range.formulas[0] ? range.formulas[0][0] : null;
      const isOccupied =
        (existingFormula && typeof existingFormula === "string" && existingFormula.trim() !== "") ||
        (existingVal !== "" && existingVal != null);

      cellInfo = {
        address: range.address,
        isOccupied: isOccupied,
        displayValue: existingFormula || existingVal,
      };
    });

    if (!cellInfo) {
      showNotification("Could not inspect active cell.", "error");
      return;
    }

    if (cellInfo.isOccupied) {
      const userApproved = await showCustomConfirmModal({
        title: "Overwrite Existing Content?",
        message: `Cell ${cellInfo.address} already contains data or a formula. Overwrite?`,
        details: `<div style="color: #f87171;">Current: ${escapeHtml(String(cellInfo.displayValue))}</div><div style="color: #4ade80;">New: ${escapeHtml(cleanFormula)}</div>`,
        confirmLabel: "Insert Formula",
        cancelLabel: "Cancel",
      });

      if (!userApproved) {
        showNotification("Formula insertion cancelled.", "info");
        return;
      }
    }

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.formulas = [[cleanFormula]];
      range.select();
      await context.sync();
    });

    showNotification(`Inserted formula ${cleanFormula} into ${cellInfo.address}!`, "success");
  } catch (err) {
    showNotification(`Error inserting formula: ${err.message || err}`, "error");
  }
}

// Send Message to Next.js API
async function sendMessage() {
  const text = userInput.value.trim();
  if (!text || isBusy) return;

  userInput.value = "";
  btnSend.disabled = true;
  isBusy = true;

  appendMessage("user", text);
  const thinkingBubble = appendMessage("assistant", currentMode === "agent" ? "Analyzing workbook & formulating execution plan..." : "Thinking...");

  try {
    // Scan workbook if not yet available
    if (!currentWorkbookData) {
      await scanWorkbookContext(false);
    }

    const payload = {
      message: text,
      conversationId: conversationId,
      mode: currentMode,
      workbookContext: currentWorkbookData,
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

    thinkingBubble.remove();

    if (data.conversationId) {
      conversationId = data.conversationId;
    }

    appendMessage("assistant", data.content || "Plan formulated.", data.detectedFormula, data.actionPlan);
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
