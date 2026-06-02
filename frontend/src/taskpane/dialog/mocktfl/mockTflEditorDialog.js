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

function setTitle(text) {
  const el = $("workitem-title");
  if (el) el.textContent = text || "MockTFL Item";
}

function messageParent(payload) {
  Office.context.ui.messageParent(JSON.stringify(payload));
}

function setEditorText(text) {
  const el = $("editor");
  if (el) el.value = text || "";
}

function getEditorText() {
  return $("editor")?.value || "";
}

function setFigureRCode(text) {
  const el = $("figure-r-code");
  if (el) el.value = String(text || "");
}

function getFigureRCode() {
  return $("figure-r-code")?.value || "";
}

function setFigurePreviewImage(base64Image) {
  const image = String(base64Image || "").trim();
  const row = $("figure-preview-row");
  const img = $("figure-preview-image");
  if (!row || !img) return;
  if (!image) {
    row.style.display = "none";
    img.removeAttribute("src");
    return;
  }
  img.src = `data:image/png;base64,${image}`;
  row.style.display = "block";
}

function renderFigureMockDataTable(text) {
  const row = $("figure-mock-data-table-row");
  const table = $("figure-mock-data-table");
  if (!row || !table) return;
  table.innerHTML = "";

  const payload = safeJsonParse(text);
  const mockData = payload?.mockData || payload;

  // New flat-records shape: mockData.records is an array of flat row objects.
  // Fall back to the older mockData.data array for backward compatibility.
  const tidySource = Array.isArray(mockData?.records)
    ? mockData.records
    : Array.isArray(mockData?.data)
    ? mockData.data
    : [];
  const tidyRows = tidySource.filter((r) => r && typeof r === "object" && !Array.isArray(r));
  if (tidyRows.length > 0) {
    const columns = [];
    tidyRows.forEach((rowObj) => {
      Object.keys(rowObj).forEach((key) => {
        if (!columns.includes(key)) columns.push(key);
      });
    });

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    columns.forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);

    const tbody = document.createElement("tbody");
    tidyRows.forEach((rowObj) => {
      const tr = document.createElement("tr");
      columns.forEach((col) => {
        const td = document.createElement("td");
        const val = rowObj[col];
        td.textContent = val == null ? "" : String(val);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    row.style.display = "block";
    return;
  }

  // Backward-compat: legacy timePoints/groupSeries shape.
  const timePoints = Array.isArray(mockData?.timePoints) ? mockData.timePoints : [];
  const groupSeries = Array.isArray(mockData?.groupSeries) ? mockData.groupSeries : [];
  if (timePoints.length === 0 || groupSeries.length === 0) {
    row.style.display = "none";
    return;
  }

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  ["Time", ...groupSeries.map((series) => String(series?.group || "Group").trim() || "Group")].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);

  const tbody = document.createElement("tbody");
  timePoints.forEach((timePoint, idx) => {
    const tr = document.createElement("tr");
    const timeCell = document.createElement("td");
    timeCell.textContent = String(timePoint);
    tr.appendChild(timeCell);

    groupSeries.forEach((series) => {
      const td = document.createElement("td");
      const values = Array.isArray(series?.values) ? series.values : [];
      td.textContent = values[idx] == null ? "" : String(values[idx]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  row.style.display = "block";
}

function resetCorrespondingTableUi() {
  const row = $("corresponding-table-row");
  const box = $("corresponding-table-box");
  if (box) box.textContent = "";
  if (row) row.style.display = "none";
  setFigureTypeText("");
  showElement("figure-type-row", false);
  showElement("figure-auto-generate-row", false);
  showElement("figure-find-table-row", true);
  setTableInfoStatus(null);
}

/**
 * Show whether the corresponding table's info was successfully fetched from the
 * database during auto generate.
 * - state="pending": neutral "fetching" message
 * - state="success": green, includes table label + the column groups applied
 * - state="fail": orange fallback message
 * - null: hide the status line
 */
function setTableInfoStatus(state, payload) {
  const el = $("figure-table-info-status");
  if (!el) return;

  if (!state) {
    el.style.display = "none";
    el.textContent = "";
    return;
  }

  if (state === "pending") {
    el.style.display = "block";
    el.style.background = "#f3f4f6";
    el.style.color = "#374151";
    el.style.border = "1px solid #e5e7eb";
    el.textContent = "Fetching corresponding table info and generating mock data...";
    return;
  }

  const info = payload || {};
  if (state === "success") {
    const groups = Array.isArray(info.usedColumnGroups) ? info.usedColumnGroups.filter(Boolean) : [];
    const groupNote = groups.length >= 2 ? ` (curve groups applied: ${groups.join(" / ")})` : "";
    el.style.display = "block";
    el.style.background = "#ecfdf5";
    el.style.color = "#065f46";
    el.style.border = "1px solid #a7f3d0";
    el.textContent = `Corresponding table info fetched successfully: ${String(info.label || "").trim()}${groupNote}`;
    return;
  }

  // fail
  el.style.display = "block";
  el.style.background = "#fff7ed";
  el.style.color = "#9a3412";
  el.style.border = "1px solid #fed7aa";
  el.textContent = "Corresponding table info not found; generated with default groups (run analysis and Save to Database on the matching TABLE first if you need consistency).";
}

function setCorrespondingTableUi(tableInfo) {
  const info = tableInfo || null;
  const row = $("corresponding-table-row");
  const box = $("corresponding-table-box");
  const label = String(info?.label || `${info?.type || "TABLE"} ${info?.number || ""}`).trim();
  if (!info || !label) {
    resetCorrespondingTableUi();
    return;
  }
  const score = typeof info.score === "number" ? ` (similarity ${info.score.toFixed(3)})` : "";
  if (box) box.textContent = `${label}${score}`;
  if (row) row.style.display = "block";
  showElement("figure-find-table-row", true);
  showElement("figure-type-row", true);
  showElement("figure-auto-generate-row", true);
}

function getFigureTypeText() {
  return $("figure-type-text")?.value || "";
}

function setFigureTypeText(text) {
  const el = $("figure-type-text");
  if (el) el.value = String(text || "");
}

function setFigureMockDataText(text) {
  const value = String(text || "");
  const row = $("figure-mock-data-row");
  const box = $("figure-mock-data-text");
  if (!row || !box) return;
  if (!value.trim()) {
    box.value = "";
    row.style.display = "none";
    renderFigureMockDataTable("");
    return;
  }
  box.value = value;
  row.style.display = "block";
  renderFigureMockDataTable(value);
}

function normalizeColumnOrientation(value, allowEmpty = false) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "horizontal") return "horizontal";
  if (normalized === "vertical") return "vertical";
  return allowEmpty ? "" : "vertical";
}

function normalizeColumnGroups(availableColumns, savedColumnGroups = []) {
  const defaults = Array.isArray(availableColumns)
    ? availableColumns.map((value) => String(value || "").trim()).filter(Boolean)
    : [];
  const saved = Array.isArray(savedColumnGroups)
    ? savedColumnGroups.map((value) => String(value || "").trim()).filter(Boolean)
    : [];
  if (saved.length === 0) return defaults;
  return defaults.map((value, idx) => saved[idx] || value).filter(Boolean);
}

function renderColumnGroupInputs(availableColumns, savedColumnGroups = []) {
  const container = $("columns-checkbox-list");
  if (!container) return;

  container.innerHTML = ""; // clear old

  if (!Array.isArray(availableColumns) || availableColumns.length === 0) {
    showElement("columns-selection-row", false);
    return;
  }

  showElement("columns-selection-row", true);
  const normalizedColumnGroups = normalizeColumnGroups(availableColumns, savedColumnGroups);

  availableColumns.forEach((colName, idx) => {
    const row = document.createElement("div");
    row.className = "column-group-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "column-group-checkbox";
    checkbox.checked = idx < normalizedColumnGroups.length;

    const input = document.createElement("input");
    input.type = "text";
    input.className = "column-group-input";
    input.value = normalizedColumnGroups[idx] || String(colName || "").trim();
    input.placeholder = `Column group ${idx + 1}`;

    row.appendChild(checkbox);
    row.appendChild(input);
    container.appendChild(row);
  });
}

function getSelectedColumnGroups() {
  const container = $("columns-checkbox-list");
  if (!container) return [];

  return Array.from(container.querySelectorAll(".column-group-row"))
    .filter((row) => row.querySelector(".column-group-checkbox")?.checked)
    .map((row) => String(row.querySelector(".column-group-input")?.value || "").trim())
    .filter(Boolean);
}

function hasEmptySelectedColumnGroup() {
  const container = $("columns-checkbox-list");
  if (!container) return false;

  return Array.from(container.querySelectorAll(".column-group-row")).some((row) => {
    const checked = row.querySelector(".column-group-checkbox")?.checked;
    const value = String(row.querySelector(".column-group-input")?.value || "").trim();
    return checked && !value;
  });
}

function setColumnGroupSubtitle(text) {
  const input = $("column-group-subtitle");
  if (input) input.value = String(text || "").trim();
}

function getColumnGroupSubtitle() {
  return String($("column-group-subtitle")?.value || "").trim();
}

function setColumnOrientation(value) {
  const select = $("column-orientation-select");
  if (select) select.value = normalizeColumnOrientation(value, true);
}

function getColumnOrientation() {
  return normalizeColumnOrientation($("column-orientation-select")?.value, true);
}

function updateWorkflowVisibilityByOrientation() {
  if (currentWorkItemType === "FIGURE") {
    showElement("vertical-workflow-root", false);
    showElement("horizontal-workflow-root", false);
    return;
  }
  const orientation = getColumnOrientation();
  const verticalRoot = $("vertical-workflow-root");
  const horizontalRoot = $("horizontal-workflow-root");
  if (verticalRoot) verticalRoot.style.display = orientation === "vertical" ? "block" : "none";
  if (horizontalRoot) horizontalRoot.style.display = orientation === "horizontal" ? "block" : "none";
}

function normalizeSublevelHeaders(headers) {
  const values = Array.isArray(headers) ? headers : [];
  const next = values.map((value) => String(value || "").trim());
  return next.length > 0 ? next : ["", ""];
}

function getSublevelConfigListEl() {
  return $("sublevel-config-list");
}

function updateSublevelInputPlaceholders(blockEl) {
  const inputs = Array.from(blockEl.querySelectorAll(".sublevel-input"));
  inputs.forEach((input, idx) => {
    input.placeholder = `Header ${idx + 1}`;
  });
}

function updateSublevelInputButtons(blockEl) {
  const rows = Array.from(blockEl.querySelectorAll(".sublevel-input-row"));
  rows.forEach((row) => {
    const removeBtn = row.querySelector(".sublevel-input-remove-btn");
    if (removeBtn) removeBtn.disabled = rows.length <= 1;
  });
}

function createSublevelInput(value = "") {
  const input = document.createElement("input");
  input.type = "text";
  input.className = "sublevel-input";
  input.value = String(value || "").trim();
  return input;
}

function updateSublevelBlockLevels() {
  const container = getSublevelConfigListEl();
  if (!container) return;
  Array.from(container.querySelectorAll(".sublevel-block")).forEach((block, idx) => {
    const level = idx + 1;
    block.dataset.level = String(level);
    const title = block.querySelector(".sublevel-title");
    if (title) title.textContent = `Header layer ${level}`;
  });
}

function createSublevelInputRow(blockEl, value = "") {
  const row = document.createElement("div");
  row.className = "sublevel-input-row";

  const input = createSublevelInput(value);
  const actions = document.createElement("div");
  actions.className = "sublevel-input-actions";

  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.className = "icon-btn sublevel-input-add-btn";
  addBtn.textContent = "+";

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "icon-btn sublevel-input-remove-btn";
  removeBtn.textContent = "-";

  addBtn.addEventListener("click", () => {
    row.insertAdjacentElement("afterend", createSublevelInputRow(blockEl, ""));
    updateSublevelInputPlaceholders(blockEl);
    updateSublevelInputButtons(blockEl);
  });

  removeBtn.addEventListener("click", () => {
    const inputRows = blockEl.querySelectorAll(".sublevel-input-row");
    if (inputRows.length <= 1) return;
    row.remove();
    updateSublevelInputPlaceholders(blockEl);
    updateSublevelInputButtons(blockEl);
  });

  actions.appendChild(addBtn);
  actions.appendChild(removeBtn);
  row.appendChild(input);
  row.appendChild(actions);
  return row;
}

function createSublevelBlock(level, headers = ["", ""]) {
  const block = document.createElement("div");
  block.className = "sublevel-block";
  block.dataset.level = String(level);

  const headerRow = document.createElement("div");
  headerRow.className = "sublevel-header-row";

  const title = document.createElement("span");
  title.className = "label sublevel-title";
  title.textContent = `Header layer ${level}`;

  const actions = document.createElement("div");
  actions.className = "sublevel-actions";

  const removeLevelBtn = document.createElement("button");
  removeLevelBtn.type = "button";
  removeLevelBtn.className = "sublevel-remove-level-btn";
  removeLevelBtn.textContent = "Remove this header layer";

  const inputList = document.createElement("div");
  inputList.className = "sublevel-input-list";

  // New sublevels still start with two inputs, but each row can be edited independently.
  normalizeSublevelHeaders(headers).forEach((headerText) => {
    inputList.appendChild(createSublevelInputRow(block, headerText));
  });

  removeLevelBtn.addEventListener("click", () => {
    block.remove();
    updateSublevelBlockLevels();
  });

  actions.appendChild(removeLevelBtn);
  headerRow.appendChild(title);
  headerRow.appendChild(actions);

  block.appendChild(headerRow);
  block.appendChild(inputList);

  updateSublevelInputPlaceholders(block);
  updateSublevelInputButtons(block);
  return block;
}

function clearSublevelBlocks() {
  const container = getSublevelConfigListEl();
  if (container) container.innerHTML = "";
}

function addSublevelBlock(headers = ["", ""]) {
  const container = getSublevelConfigListEl();
  if (!container) return;
  const level = container.querySelectorAll(".sublevel-block").length + 1;
  container.appendChild(createSublevelBlock(level, headers));
  updateSublevelBlockLevels();
}

function renderSublevelBlocks(subLevels) {
  clearSublevelBlocks();
  const arr = Array.isArray(subLevels) ? subLevels : [];
  arr.forEach((subLevel) => {
    addSublevelBlock(normalizeSublevelHeaders(subLevel?.headers));
  });
}

function collectColumnHeaderConfig() {
  const sublevelContainer = getSublevelConfigListEl();
  const headerLayers = !sublevelContainer
    ? []
    : Array.from(sublevelContainer.querySelectorAll(".sublevel-block")).map((block, idx) => ({
        // Header layers are positional in the UI below the column groups row.
        level: idx + 1,
        headers: Array.from(block.querySelectorAll(".sublevel-input")).map((input) =>
          String(input.value || "").trim()
        )
      }));

  return {
    orientation: getColumnOrientation(),
    columnGroups: getSelectedColumnGroups(),
    columnGroupSubtitle: getColumnGroupSubtitle(),
    headerLayers
  };
}

function validateColumnHeaderConfig(columnHeaderConfig) {
  const orientation = normalizeColumnOrientation(columnHeaderConfig?.orientation, true);
  if (!orientation) {
    return "Please select a column header direction.";
  }
  // Horizontal can already be chosen/saved even though its detailed config is not implemented yet.
  if (orientation === "horizontal") {
    return "";
  }

  const columnGroups = Array.isArray(columnHeaderConfig?.columnGroups)
    ? columnHeaderConfig.columnGroups.filter(Boolean)
    : [];
  if (columnGroups.length === 0) {
    return "Please select at least one column group.";
  }
  if (hasEmptySelectedColumnGroup()) {
    return "Please fill in every selected column group name.";
  }

  const headerLayers = Array.isArray(columnHeaderConfig?.headerLayers) ? columnHeaderConfig.headerLayers : [];
  for (let idx = 0; idx < headerLayers.length; idx++) {
    const level = idx + 1;
    const headers = Array.isArray(headerLayers[idx]?.headers) ? headerLayers[idx].headers : [];
    if (headers.length < 1) {
      return `Header layer ${level} must contain at least one header.`;
    }
    if (headers.some((headerText) => !String(headerText || "").trim())) {
      return `Please fill in every header text for header layer ${level}.`;
    }
  }

  return "";
}

function renderColumnHeaderConfig(availableColumns, savedColumnHeaderConfig) {
  const savedConfig = savedColumnHeaderConfig || {};
  setColumnOrientation(savedConfig.orientation || "");
  // Column groups come from study-level defaults but remain editable in the dialog.
  renderColumnGroupInputs(availableColumns, savedConfig.columnGroups);
  setColumnGroupSubtitle(savedConfig.columnGroupSubtitle || "");
  renderSublevelBlocks(savedConfig.headerLayers);
  updateWorkflowVisibilityByOrientation();
}

function setSapMatchedSection({ label, text }) {
  const labelEl = $("sap-matched-label");
  const textEl = $("sap-section-text");
  if (labelEl) {
    const safeLabel = String(label || "").trim();
    labelEl.textContent = safeLabel ? `Matched SAP section: ${safeLabel}` : "Matched SAP section text";
  }
  if (textEl) textEl.value = String(text || "");
}

function setSapStatisticalVariablesText(text) {
  const el = $("sap-statistical-variables-text");
  if (el) el.value = String(text || "");
}

function getSapStatisticalVariablesText() {
  return $("sap-statistical-variables-text")?.value || "";
}

function setMappedCrfQaText(text) {
  const el = $("crf-mapped-qa-text");
  if (el) el.value = String(text || "");
}

function showElement(id, show) {
  const el = $(id);
  if (!el) return;
  if (!show) {
    el.style.display = "none";
    return;
  }
  el.style.display = id === "confirm-variables-row" ? "flex" : "block";
}

function setButtonEnabled(id, enabled) {
  const el = $(id);
  if (!el) return;
  el.disabled = !enabled;
}

function clearSapAnalysisPanels() {
  setSapMatchedSection({ label: "", text: "" });
  setSapStatisticalVariablesText("");
  setMappedCrfQaText("");
  setEditorText("");
  showElement("sap-section-row", false);
  showElement("variables-row", false);
  showElement("confirm-variables-row", false);
  showElement("mapping-row", false);
  showElement("editor-row", false);
  showElement("editor-actions-row", false);
}

function showStartAnalysisButton(show) {
  const btn = $("btn-start-analysis");
  if (!btn) return;
  btn.style.display = show ? "inline-block" : "none";
  btn.disabled = !show;
}

function setStartAnalysisButtonText(text) {
  const btn = $("btn-start-analysis");
  if (btn) btn.textContent = text || "Start analysis";
}

function showMockTflWorkflowMode(type) {
  currentWorkItemType = String(type || "").trim().toUpperCase();
  const isFigure = currentWorkItemType === "FIGURE";
  showElement("column-orientation-row", !isFigure);
  showElement("figure-workflow-root", isFigure);

  if (isFigure) {
    showElement("vertical-workflow-root", false);
    showElement("horizontal-workflow-root", false);
  } else {
    updateWorkflowVisibilityByOrientation();
  }
}

let currentWorkItemType = "";
let analysisEnabledForCurrentItem = false;
let mappingCompleted = false;
let generatedFigureBase64 = "";
let generatedFigureMockData = "";

function bindUi() {
  const btnInsert = $("btn-insert");
  const btnClear = $("btn-clear");
  const btnClose = $("btn-close");
  const btnStartAnalysis = $("btn-start-analysis");
  const btnConfirmVariables = $("btn-confirm-variables");
  const btnAddSublevelHeader = $("btn-add-sublevel-header");
  const btnGenerateFigure = $("btn-generate-figure");
  const btnAutoGenerateFigure = $("btn-auto-generate-figure");
  const btnFindCorrespondingTable = $("btn-find-corresponding-table");
  const btnInsertFigure = $("btn-insert-figure");
  const btnClearFigure = $("btn-clear-figure");
  const btnCloseFigure = $("btn-close-figure");
  const columnOrientationSelect = $("column-orientation-select");
  const variablesInput = $("sap-statistical-variables-text");

  if (btnInsert) {
    btnInsert.addEventListener("click", () => {
      const columnHeaderConfig = collectColumnHeaderConfig();
      const validationError = validateColumnHeaderConfig(columnHeaderConfig);
      if (validationError) {
        setStatus(validationError);
        return;
      }
      messageParent({ 
        type: "mocktfl:insert", 
        text: getEditorText(),
        selectedColumns: columnHeaderConfig.columnGroups,
        columnHeaderConfig
      });
    });
  }
  if (btnClear) {
    btnClear.addEventListener("click", () => {
      setEditorText("");
      messageParent({ type: "mocktfl:clear" });
    });
  }
  if (btnClose) {
    btnClose.addEventListener("click", () => {
      messageParent({ type: "mocktfl:close" });
      try {
        Office.context.ui.closeContainer();
      } catch (e) {}
    });
  }

  if (btnFindCorrespondingTable) {
    btnFindCorrespondingTable.addEventListener("click", () => {
      setButtonEnabled("btn-find-corresponding-table", false);
      setStatus("Finding corresponding table...");
      messageParent({ type: "mocktfl:findCorrespondingTable" });
    });
  }

  if (btnAutoGenerateFigure) {
    btnAutoGenerateFigure.addEventListener("click", () => {
      generatedFigureBase64 = "";
      generatedFigureMockData = "";
      setFigurePreviewImage("");
      setFigureMockDataText("");
      setButtonEnabled("btn-insert-figure", false);
      setButtonEnabled("btn-auto-generate-figure", false);
      setButtonEnabled("btn-generate-figure", false);
      setStatus("Auto-generating mock data and figure...");
      messageParent({ type: "mocktfl:autoGenerateFigure", figureType: getFigureTypeText() });
    });
  }

  const btnDefineFigureType = $("btn-define-figure-type");
  if (btnDefineFigureType) {
    btnDefineFigureType.addEventListener("click", () => {
      setButtonEnabled("btn-define-figure-type", false);
      setStatus("Defining figure type with GPT...");
      messageParent({ type: "mocktfl:defineFigureType" });
    });
  }

  if (btnGenerateFigure) {
    btnGenerateFigure.addEventListener("click", () => {
      const rCodeText = getFigureRCode();
      if (!rCodeText.trim()) {
        setStatus("Please paste R code before generating the figure.");
        return;
      }
      generatedFigureBase64 = "";
      generatedFigureMockData = "";
      setFigurePreviewImage("");
      setFigureMockDataText("");
      setButtonEnabled("btn-insert-figure", false);
      setButtonEnabled("btn-auto-generate-figure", false);
      setButtonEnabled("btn-generate-figure", false);
      setStatus("Generating figure with R...");
      messageParent({ type: "mocktfl:generateFigure", rCodeText });
    });
  }

  if (btnInsertFigure) {
    btnInsertFigure.addEventListener("click", () => {
      messageParent({
        type: "mocktfl:insert",
        rCodeText: getFigureRCode(),
        figureBase64: generatedFigureBase64,
        figureMockDataText: generatedFigureMockData
      });
    });
  }

  if (btnClearFigure) {
    btnClearFigure.addEventListener("click", () => {
      generatedFigureBase64 = "";
      generatedFigureMockData = "";
      setFigureRCode("");
      setFigurePreviewImage("");
      setFigureMockDataText("");
      setButtonEnabled("btn-insert-figure", false);
      messageParent({ type: "mocktfl:clear" });
    });
  }

  if (btnCloseFigure) {
    btnCloseFigure.addEventListener("click", () => {
      messageParent({ type: "mocktfl:close" });
      try {
        Office.context.ui.closeContainer();
      } catch (e) {}
    });
  }

  if (btnStartAnalysis) {
    btnStartAnalysis.addEventListener("click", () => {
      // The parent controller will enforce a strict feature gate as well.
      setStatus("Analyzing...");
      clearSapAnalysisPanels();
      mappingCompleted = false;
      if (analysisEnabledForCurrentItem) setButtonEnabled("btn-insert", false);
      setButtonEnabled("btn-start-analysis", false);
      messageParent({ type: "mocktfl:startAnalysis" });
    });
  }

  if (btnConfirmVariables) {
    btnConfirmVariables.addEventListener("click", () => {
      // Send the current textarea contents so the backend maps exactly what the
      // user has reviewed and, if needed, manually edited.
      setStatus("Mapping confirmed variables to CRF...");
      setButtonEnabled("btn-confirm-variables", false);
      messageParent({
        type: "mocktfl:confirmVariables",
        variablesText: getSapStatisticalVariablesText()
      });
    });
  }

  if (btnAddSublevelHeader) {
    btnAddSublevelHeader.addEventListener("click", () => {
      addSublevelBlock(["", ""]);
    });
  }

  if (columnOrientationSelect) {
    columnOrientationSelect.addEventListener("change", () => {
      updateWorkflowVisibilityByOrientation();
      const orientation = getColumnOrientation();
      if (orientation !== "vertical") {
        setStatus("");
      } else if (!String($("status")?.textContent || "").trim()) {
        setStatus("Ready.");
      }
      if (!orientation) return;
      messageParent({
        type: "mocktfl:updateColumnOrientation",
        columnHeaderConfig: collectColumnHeaderConfig()
      });
    });
  }

  if (variablesInput) {
    variablesInput.addEventListener("input", () => {
      const canConfirm = !!getSapStatisticalVariablesText().trim();
      setButtonEnabled("btn-confirm-variables", canConfirm);
    });
  }

  const editorInput = $("editor");
  if (editorInput) {
    editorInput.addEventListener("input", () => {
      const canInsert = !analysisEnabledForCurrentItem || (mappingCompleted && !!getEditorText().trim());
      setButtonEnabled("btn-insert", canInsert);
    });
  }

  // Parent -> dialog messages
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    const msg = safeJsonParse(arg?.message);
    const msgType = String(msg?.type || "");
    if (!msgType.startsWith("mocktfl:")) return;

    if (msgType === "mocktfl:init") {
      const label = msg?.workItem?.label || msg?.workItem?.number || "MockTFL Item";
      setTitle(label);
      showMockTflWorkflowMode(msg?.workItem?.type);
      analysisEnabledForCurrentItem = !!msg?.workItem?.enableAnalysis;
      mappingCompleted = false;
      generatedFigureBase64 = "";
      generatedFigureMockData = "";
      clearSapAnalysisPanels();
      setFigureRCode(String(msg?.savedSectionData?.rCodeText || ""));
      setFigurePreviewImage(String(msg?.savedSectionData?.figureBase64 || ""));
      setFigureMockDataText(String(msg?.savedSectionData?.figureMockDataText || ""));
      if (String(msg?.savedSectionData?.figureBase64 || "").trim()) {
        generatedFigureBase64 = String(msg.savedSectionData.figureBase64).trim();
      }
      if (String(msg?.savedSectionData?.figureMockDataText || "").trim()) {
        generatedFigureMockData = String(msg.savedSectionData.figureMockDataText);
      }
      // The figure workflow starts at "Find Corresponding Table" and only reveals
      // "Auto Generate Data & Figure" once a corresponding table is known.
      resetCorrespondingTableUi();
      if (msg?.savedSectionData?.correspondingTable) {
        setCorrespondingTableUi(msg.savedSectionData.correspondingTable);
        setFigureTypeText(String(msg?.savedSectionData?.figureType || ""));
      }

      if (Array.isArray(msg?.availableColumns)) {
        renderColumnHeaderConfig(msg.availableColumns, msg?.savedColumnHeaderConfig);
      } else {
        renderColumnHeaderConfig([], msg?.savedColumnHeaderConfig);
      }

      showStartAnalysisButton(analysisEnabledForCurrentItem);
      setStartAnalysisButtonText("Start analysis");
      setButtonEnabled("btn-start-analysis", analysisEnabledForCurrentItem);
      setButtonEnabled("btn-confirm-variables", false);
      setButtonEnabled("btn-insert", !analysisEnabledForCurrentItem);
      setButtonEnabled("btn-find-corresponding-table", true);
      setButtonEnabled("btn-auto-generate-figure", true);
      setButtonEnabled("btn-generate-figure", true);
      setButtonEnabled("btn-insert-figure", !!generatedFigureBase64);
      setStatus(getColumnOrientation() === "vertical" ? "Ready." : "");
      return;
    }

    if (msgType === "mocktfl:loadSavedSectionData") {
      const label = msg?.workItem?.label || msg?.workItem?.number || "MockTFL Item";
      const saved = msg?.savedSectionData || {};
      setTitle(label);
      showMockTflWorkflowMode(msg?.workItem?.type);
      analysisEnabledForCurrentItem = !!msg?.workItem?.enableAnalysis;
      mappingCompleted = true;
      generatedFigureBase64 = String(saved?.figureBase64 || "").trim();
      generatedFigureMockData = String(saved?.figureMockDataText || "");
      clearSapAnalysisPanels();
      setFigureRCode(String(saved?.rCodeText || ""));
      setFigurePreviewImage(generatedFigureBase64);
      setFigureMockDataText(generatedFigureMockData);
      resetCorrespondingTableUi();
      if (saved?.correspondingTable) {
        setCorrespondingTableUi(saved.correspondingTable);
        setFigureTypeText(String(saved?.figureType || ""));
      }

      if (Array.isArray(msg?.availableColumns)) {
        renderColumnHeaderConfig(msg.availableColumns, saved?.columnHeaderConfig);
      } else {
        renderColumnHeaderConfig([], saved?.columnHeaderConfig);
      }

      showStartAnalysisButton(analysisEnabledForCurrentItem);
      setStartAnalysisButtonText("Re-analysis");
      setSapMatchedSection({
        label: String(saved?.matchedSapLabel || ""),
        text: String(saved?.matchedSapSectionText || "")
      });
      setSapStatisticalVariablesText(String(saved?.suggestedVariablesText || ""));
      setMappedCrfQaText(String(saved?.mappedCrfText || ""));
      setEditorText(String(saved?.notesText || ""));
      showElement("sap-section-row", true);
      showElement("variables-row", true);
      showElement("confirm-variables-row", true);
      showElement("mapping-row", true);
      showElement("editor-row", true);
      showElement("editor-actions-row", true);
      setButtonEnabled("btn-confirm-variables", !!getSapStatisticalVariablesText().trim());
      setButtonEnabled("btn-insert", !!getEditorText().trim());
      setButtonEnabled("btn-find-corresponding-table", true);
      setButtonEnabled("btn-auto-generate-figure", true);
      setButtonEnabled("btn-generate-figure", true);
      setButtonEnabled("btn-insert-figure", !!generatedFigureBase64);
      setStatus(getColumnOrientation() === "vertical" ? "Loaded saved section data from database." : "");
      return;
    }

    if (msgType === "mocktfl:analysisRunning") {
      setStatus("Analyzing SAP and generating suggested variables...");
      return;
    }

    if (msgType === "mocktfl:analysisStep1Result") {
      // Step 1 only reveals the SAP evidence and editable variable suggestions.
      setSapMatchedSection({
        label: String(msg?.sapMatchedLabel || ""),
        text: String(msg?.sapSectionText || "")
      });
      setSapStatisticalVariablesText(String(msg?.sapStatisticalVariablesText || ""));
      setMappedCrfQaText("");
      showElement("sap-section-row", true);
      showElement("variables-row", true);
      showElement("confirm-variables-row", true);
      showElement("mapping-row", false);
      setButtonEnabled("btn-start-analysis", true);
      setStartAnalysisButtonText("Re-analysis");
      setButtonEnabled("btn-confirm-variables", true);
      setStatus("Review or edit the suggested variables, then confirm to map them to CRF.");
      return;
    }

    if (msgType === "mocktfl:analysisStep2Running") {
      setStatus("Mapping confirmed variables to CRF...");
      setButtonEnabled("btn-confirm-variables", false);
      return;
    }

    if (msgType === "mocktfl:analysisStep2Result") {
      // Final Notes text is now built from confirmed variables + mapping result,
      // replacing the old full-question-list source.
      const finalNotesText = String(msg?.finalNotesText || "").trim();
      if (finalNotesText) {
        setEditorText(`${finalNotesText}\n`);
      }
      mappingCompleted = true;
      setMappedCrfQaText(String(msg?.mappedEvidenceText || ""));
      showElement("mapping-row", true);
      showElement("editor-row", true);
      showElement("editor-actions-row", true);
      setButtonEnabled("btn-start-analysis", true);
      setStartAnalysisButtonText("Re-analysis");
      setButtonEnabled("btn-confirm-variables", true);
      if (analysisEnabledForCurrentItem && !!getEditorText().trim()) {
        setButtonEnabled("btn-insert", true);
      }
      setStatus("Analysis ready.");
      return;
    }

    if (msgType === "mocktfl:findingCorrespondingTable") {
      setButtonEnabled("btn-find-corresponding-table", false);
      setStatus("Finding corresponding table...");
      return;
    }

    if (msgType === "mocktfl:correspondingTableFound") {
      setCorrespondingTableUi(msg?.matchedTable || null);
      setButtonEnabled("btn-find-corresponding-table", true);
      setStatus("Corresponding table found. Define the figure type, then auto generate data and figure.");
      return;
    }

    if (msgType === "mocktfl:definingFigureType") {
      setButtonEnabled("btn-define-figure-type", false);
      setStatus("Defining figure type with GPT...");
      return;
    }

    if (msgType === "mocktfl:figureTypeDefined") {
      setFigureTypeText(String(msg?.figureType || ""));
      setButtonEnabled("btn-define-figure-type", true);
      setStatus("Figure type defined. You can edit it, then auto generate data and figure.");
      return;
    }

    if (msgType === "mocktfl:autoFigureGenerating") {
      setStatus("Auto-generating mock data and figure...");
      setTableInfoStatus("pending");
      setButtonEnabled("btn-auto-generate-figure", false);
      setButtonEnabled("btn-generate-figure", false);
      setButtonEnabled("btn-insert-figure", false);
      return;
    }

    if (msgType === "mocktfl:autoFigureGenerated") {
      generatedFigureBase64 = String(msg?.imageBase64 || "").trim();
      generatedFigureMockData = String(msg?.figureMockDataText || "");
      if (String(msg?.generatedRCode || "").trim()) {
        setFigureRCode(String(msg.generatedRCode));
      }
      setFigureMockDataText(generatedFigureMockData);
      setFigurePreviewImage(generatedFigureBase64);
      const tableInfo = msg?.tableInfo || { fetched: false };
      setTableInfoStatus(tableInfo.fetched ? "success" : "fail", tableInfo);
      setButtonEnabled("btn-auto-generate-figure", true);
      setButtonEnabled("btn-generate-figure", true);
      setButtonEnabled("btn-insert-figure", !!generatedFigureBase64);
      setStatus("Auto-generated mock data and figure are ready. Review then click Insert.");
      return;
    }

    if (msgType === "mocktfl:figureGenerating") {
      setStatus("Generating figure with R...");
      setButtonEnabled("btn-auto-generate-figure", false);
      setButtonEnabled("btn-generate-figure", false);
      setButtonEnabled("btn-insert-figure", false);
      return;
    }

    if (msgType === "mocktfl:figureGenerated") {
      generatedFigureBase64 = String(msg?.imageBase64 || "").trim();
      setFigurePreviewImage(generatedFigureBase64);
      setButtonEnabled("btn-auto-generate-figure", true);
      setButtonEnabled("btn-generate-figure", true);
      setButtonEnabled("btn-insert-figure", !!generatedFigureBase64);
      setStatus("Figure generated. Review the preview, then click Insert.");
      return;
    }

    if (msgType === "mocktfl:inserted") {
      setStatus("Inserted.");
      return;
    }

    if (msgType === "mocktfl:cleared") {
      setStatus("Cleared.");
      mappingCompleted = false;
      clearSapAnalysisPanels();
      setButtonEnabled("btn-start-analysis", analysisEnabledForCurrentItem);
      setStartAnalysisButtonText("Start analysis");
      setButtonEnabled("btn-confirm-variables", false);
      setButtonEnabled("btn-insert", !analysisEnabledForCurrentItem);
      generatedFigureMockData = "";
      setFigureMockDataText("");
      resetCorrespondingTableUi();
      setButtonEnabled("btn-find-corresponding-table", true);
      setButtonEnabled("btn-auto-generate-figure", true);
      setButtonEnabled("btn-generate-figure", true);
      setButtonEnabled("btn-insert-figure", false);
      return;
    }

    if (msgType === "mocktfl:error") {
      setTableInfoStatus(null);
      setButtonEnabled("btn-start-analysis", analysisEnabledForCurrentItem);
      setButtonEnabled("btn-confirm-variables", !!getSapStatisticalVariablesText().trim());
      setButtonEnabled("btn-find-corresponding-table", true);
      setButtonEnabled("btn-define-figure-type", true);
      setButtonEnabled("btn-auto-generate-figure", true);
      setButtonEnabled("btn-generate-figure", true);
      setButtonEnabled("btn-insert-figure", !!generatedFigureBase64);
      if (!analysisEnabledForCurrentItem || (mappingCompleted && !!getEditorText().trim())) {
        setButtonEnabled("btn-insert", true);
      }
      setStatus(`Error: ${msg?.message || "Unknown"}`);
    }
  });
}

Office.initialize = () => {
  bindUi();
  setStatus("Loading...");
  messageParent({ type: "mocktfl:dialogReady" });
};

