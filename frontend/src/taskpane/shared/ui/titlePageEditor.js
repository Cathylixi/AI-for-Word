// Title Page editor UI (protocol metadata fields) rendered inside taskpane.

const FIELD_DEFS = [
  { key: "protocolTitle", label: "Protocol Title" },
  { key: "protocolNumber", label: "Protocol Number" },
  { key: "protocolVersionDate", label: "Protocol Version, Date" },
  { key: "documentVersionDate", label: "Document Version, Date" }
];

function renderTitlePageEditor(rootEl) {
  if (!rootEl) return;
  rootEl.innerHTML = "";

  FIELD_DEFS.forEach((f) => {
    const row = document.createElement("div");
    row.style.marginBottom = "8px";

    const label = document.createElement("div");
    label.textContent = f.label;
    label.style.fontWeight = "600";
    label.style.marginBottom = "4px";

    const input = document.createElement("textarea");
    input.id = `titlepage-${f.key}`;
    input.rows = 2;
    input.style.width = "100%";
    input.style.padding = "8px";
    input.style.border = "1px solid #ddd";
    input.style.borderRadius = "6px";
    input.style.fontFamily = "inherit";
    input.style.fontSize = "13px";

    row.appendChild(label);
    row.appendChild(input);
    rootEl.appendChild(row);
  });
}

function showTitlePageEditor(rowEl, show) {
  if (!rowEl) return;
  rowEl.style.display = show ? "block" : "none";
}

function getTitlePageValues(rootEl) {
  const values = {};
  FIELD_DEFS.forEach((f) => {
    const el = rootEl?.querySelector(`#titlepage-${f.key}`);
    values[f.key] = el?.value || "";
  });
  return values;
}

function setTitlePageValues(rootEl, values) {
  FIELD_DEFS.forEach((f) => {
    const el = rootEl?.querySelector(`#titlepage-${f.key}`);
    if (el) el.value = values?.[f.key] || "";
  });
}

function clearTitlePageValues(rootEl) {
  FIELD_DEFS.forEach((f) => {
    const el = rootEl?.querySelector(`#titlepage-${f.key}`);
    if (el) el.value = "";
  });
}

function hasAnyTitlePageValue(rootEl) {
  return FIELD_DEFS.some((f) => {
    const el = rootEl?.querySelector(`#titlepage-${f.key}`);
    return !!(el && el.value && el.value.trim());
  });
}

function bindTitlePageInput(rootEl, handler) {
  FIELD_DEFS.forEach((f) => {
    const el = rootEl?.querySelector(`#titlepage-${f.key}`);
    if (el) el.addEventListener("input", handler);
  });
}

export {
  renderTitlePageEditor,
  showTitlePageEditor,
  getTitlePageValues,
  setTitlePageValues,
  clearTitlePageValues,
  hasAnyTitlePageValue,
  bindTitlePageInput
};
