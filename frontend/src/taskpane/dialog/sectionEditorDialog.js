/* global Office */

function $(id) {
  return document.getElementById(id);
}

function safeJsonParse(s) {
  try {
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

function setStatus(text) {
  const el = $("status");
  if (el) el.textContent = text || "";
}

function setContext({ protocolSectionLabel = "", protocolText = "", crfFormName = "", crfText = "" } = {}) {
  const row = $("context-row");
  const protocolName = $("protocol-section-name");
  const protocolBox = $("protocol-context");
  const crfBox = $("crf-context");
  const crfName = $("crf-form-name");
  const isLoa = currentSection === "__LOA__";
  const isTitlePage = currentSection === "__TITLE_PAGE__";

  if (row) row.style.display = !isLoa && !isTitlePage && (protocolText || crfText || crfFormName) ? "block" : "none";
  if (protocolName) {
    protocolName.textContent = protocolSectionLabel ? `Protocol: ${protocolSectionLabel}` : "Protocol context";
  }
  if (protocolBox) protocolBox.value = protocolText || "";
  if (crfBox) crfBox.value = crfText || "";
  if (crfName) crfName.textContent = crfFormName ? `CRF: ${crfFormName}` : "CRF context";
}

function setButtonsEnabled({ canInsertClear, canGenerate, canApply }) {
  const btnInsert = $("btn-insert");
  const btnClear = $("btn-clear");
  const btnModify = $("btn-modify");
  const btnGenerate = $("btn-generate");
  const btnRegenerate = $("btn-regenerate");
  const btnApply = $("btn-apply-abbrev");

  if (btnInsert) btnInsert.disabled = !canInsertClear;
  if (btnClear) btnClear.disabled = !canInsertClear;
  if (btnModify) btnModify.disabled = !canInsertClear;
  if (btnGenerate) btnGenerate.disabled = !canGenerate;
  if (btnRegenerate) btnRegenerate.disabled = !canGenerate;
  if (btnApply) btnApply.disabled = !canApply;
}

function toggleModifyArea(show) {
  const area = $("modify-area");
  if (area) area.style.display = show ? "block" : "none";
}

function messageParent(payload) {
  Office.context.ui.messageParent(JSON.stringify(payload));
}

let currentStudy = "";
let currentSection = "";
let currentLabel = "";

function isTitlePageMode() {
  return currentSection === "__TITLE_PAGE__";
}

function setTitlePageValues(values = {}) {
  const keys = ["protocolTitle", "protocolNumber", "protocolVersionDate", "documentVersionDate"];
  keys.forEach((k) => {
    const el = $(`tp-${k}`);
    if (el) el.value = values?.[k] || "";
  });
}

function getTitlePageValues() {
  return {
    protocolTitle: $("tp-protocolTitle")?.value || "",
    protocolNumber: $("tp-protocolNumber")?.value || "",
    protocolVersionDate: $("tp-protocolVersionDate")?.value || "",
    documentVersionDate: $("tp-documentVersionDate")?.value || ""
  };
}

function clearTitlePageValues() {
  setTitlePageValues({});
}

function hasAnyTitlePageValue() {
  const vals = getTitlePageValues();
  return Object.values(vals).some((v) => String(v || "").trim().length > 0);
}

function toggleModeUi() {
  const titlePage = isTitlePageMode();
  const editor = $("editor");
  const tpEditor = $("titlepage-editor");
  const modifyBtn = $("btn-modify");
  const modifyArea = $("modify-area");
  const applyBtn = $("btn-apply-abbrev");
  if (editor) editor.style.display = titlePage ? "none" : "block";
  if (tpEditor) tpEditor.style.display = titlePage ? "block" : "none";
  if (modifyBtn) modifyBtn.style.display = titlePage ? "none" : "inline-block";
  if (modifyArea && titlePage) modifyArea.style.display = "none";
  if (applyBtn) applyBtn.style.display = currentSection === "__LOA__" ? "inline-block" : "none";
}

function updateTitle() {
  const el = $("section-title");
  if (!el) return;
  el.textContent = currentLabel || "Section";
}

function updateEditorState() {
  const text = ($("editor")?.value || "").trim();
  const isLoa = currentSection === "__LOA__";
  const isTitlePage = isTitlePageMode();
  setButtonsEnabled({
    canInsertClear: isTitlePage ? hasAnyTitlePageValue() : text.length > 0,
    canGenerate: true,
    canApply: !isTitlePage && isLoa && text.length > 0
  });
}

function handleInit(msg) {
  currentStudy = msg.studyNumber || "";
  currentSection = msg.sectionNumber || "";
  currentLabel = msg.label || `${currentSection}`.trim();
  updateTitle();

  toggleModeUi();

  if (isTitlePageMode()) {
    setTitlePageValues(msg.titlePageValues || {});
  } else if ($("editor") && typeof msg.initialText === "string") {
    $("editor").value = msg.initialText;
  }
  setContext({ protocolText: "", crfFormName: "", crfText: "" });
  updateEditorState();

  if (msg.autoGenerate) {
    setStatus("Generating...");
    setButtonsEnabled({ canInsertClear: false, canGenerate: false });
    messageParent({ type: "generate", studyNumber: currentStudy, sectionNumber: currentSection });
  }
}

function handleGenerated(msg) {
  if ($("editor") && typeof msg.text === "string") $("editor").value = msg.text;
  setStatus("AI generation complete.");
  updateEditorState();
}

function handleGeneratedContext(msg) {
  setContext({
    protocolSectionLabel: msg?.protocolSectionLabel || "",
    protocolText: msg?.protocolText || "",
    crfFormName: msg?.crfFormName || "",
    crfText: msg?.crfText || ""
  });
}

function handleGeneratedTitlePage(msg) {
  setTitlePageValues(msg?.values || {});
  setStatus("AI generation complete.");
  updateEditorState();
}

function handleInserted() {
  setStatus("Inserted.");
  updateEditorState();
}

function handleCleared() {
  if (isTitlePageMode()) {
    clearTitlePageValues();
  } else if ($("editor")) {
    $("editor").value = "";
  }
  setContext({ protocolText: "", crfFormName: "", crfText: "" });
  setStatus("Cleared.");
  updateEditorState();
}

function handleError(msg) {
  setStatus(msg?.message ? `Error: ${msg.message}` : "Error");
  updateEditorState();
}

function handleApplyProgress(msg) {
  if (msg?.message) setStatus(msg.message);
}

function handleApplyDone(msg) {
  const replacements = msg?.replacementsMade ?? 0;
  const sectionsChanged = msg?.sectionsChanged ?? 0;
  setStatus(`Applied abbreviations. Replacements: ${replacements}. Sections updated: ${sectionsChanged}.`);
  updateEditorState();
}

Office.onReady(() => {
  // Notify parent that dialog is ready to receive init data.
  messageParent({ type: "dialogReady" });

  $("btn-generate")?.addEventListener("click", () => {
    if (!currentStudy || !currentSection) return;
    setStatus("Generating...");
    setButtonsEnabled({ canInsertClear: false, canGenerate: false });
    messageParent({ type: "generate", studyNumber: currentStudy, sectionNumber: currentSection });
  });

  $("btn-insert")?.addEventListener("click", () => {
    if (!currentStudy || !currentSection) return;
    setStatus("Inserting...");
    setButtonsEnabled({ canInsertClear: false, canGenerate: false });
    if (isTitlePageMode()) {
      messageParent({
        type: "insert",
        studyNumber: currentStudy,
        sectionNumber: currentSection,
        values: getTitlePageValues()
      });
    } else {
      const text = $("editor")?.value || "";
      messageParent({ type: "insert", studyNumber: currentStudy, sectionNumber: currentSection, text });
    }
  });

  $("btn-clear")?.addEventListener("click", () => {
    if (!currentStudy || !currentSection) return;
    setStatus("Clearing...");
    setButtonsEnabled({ canInsertClear: false, canGenerate: false });
    messageParent({ type: "clear", studyNumber: currentStudy, sectionNumber: currentSection });
  });

  $("btn-modify")?.addEventListener("click", () => {
    toggleModifyArea(true);
  });

  $("btn-cancel-modify")?.addEventListener("click", () => {
    toggleModifyArea(false);
  });

  $("btn-regenerate")?.addEventListener("click", () => {
    if (!currentStudy || !currentSection) return;
    const modifyPrompt = $("modify-prompt")?.value || "";
    if (!modifyPrompt.trim()) {
      setStatus("Please enter instructions for modification.");
      return;
    }
    setStatus("Regenerating...");
    setButtonsEnabled({ canInsertClear: false, canGenerate: false });
    messageParent({
      type: "regenerate",
      studyNumber: currentStudy,
      sectionNumber: currentSection,
      modifyPrompt
    });
    toggleModifyArea(false);
  });

  $("btn-apply-abbrev")?.addEventListener("click", () => {
    if (!currentStudy || !currentSection) return;
    const text = $("editor")?.value || "";
    if (!text.trim()) {
      setStatus("Please provide abbreviations first.");
      return;
    }
    setStatus("Applying abbreviations to body sections...");
    setButtonsEnabled({ canInsertClear: false, canGenerate: false, canApply: false });
    messageParent({
      type: "applyAbbrevToBody",
      studyNumber: currentStudy,
      sectionNumber: currentSection,
      abbrevText: text
    });
  });

  $("editor")?.addEventListener("input", updateEditorState);
  ["protocolTitle", "protocolNumber", "protocolVersionDate", "documentVersionDate"].forEach((k) => {
    $(`tp-${k}`)?.addEventListener("input", updateEditorState);
  });

  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    const msg = safeJsonParse(arg?.message);
    if (!msg?.type) return;
    if (msg.type === "init") return handleInit(msg);
    if (msg.type === "generatedContext") return handleGeneratedContext(msg);
    if (msg.type === "generatedTitlePage") return handleGeneratedTitlePage(msg);
    if (msg.type === "generated") return handleGenerated(msg);
    if (msg.type === "inserted") return handleInserted();
    if (msg.type === "cleared") return handleCleared();
    if (msg.type === "applyProgress") return handleApplyProgress(msg);
    if (msg.type === "applyDone") return handleApplyDone(msg);
    if (msg.type === "error") return handleError(msg);
  });
});

