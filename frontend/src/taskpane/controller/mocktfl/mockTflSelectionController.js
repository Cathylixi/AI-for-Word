export * from "../../domains/mocktfl/controller/mockTflSelectionController";

import { getMockTflExampleEntries } from "../../api/mocktfl/references";
import { autoSelectMockTflSections } from "../../api/mocktfl/autoSelect";

function renderMockTflChecklist({ containerEl, entries, onSelectionChange }) {
  if (!containerEl) return;
  containerEl.innerHTML = "";
  (entries || []).forEach((e) => {
    const type = String(e.type || "").trim().toUpperCase();
    const number = String(e.number || "").trim();
    const title = String(e.title || "").trim();
    const order = typeof e.order === "number" ? e.order : null;
    if (!type || !number || !title) return;

    const id = `mocktfl-section-${type}-${number.replace(/[^a-zA-Z0-9_-]/g, "_")}`;
    const wrapper = document.createElement("div");
    wrapper.style.display = "flex";
    wrapper.style.gap = "8px";
    wrapper.style.alignItems = "center";
    wrapper.style.marginBottom = "6px";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.id = id;
    checkbox.value = number;
    checkbox.dataset.type = type;
    checkbox.dataset.title = title;
    checkbox.dataset.order = order === null ? "" : String(order);
    checkbox.addEventListener("change", () => {
      if (onSelectionChange) onSelectionChange();
    });

    const label = document.createElement("label");
    label.setAttribute("for", id);
    label.textContent = `${type} ${number} ${title}`.trim();

    wrapper.appendChild(checkbox);
    wrapper.appendChild(label);
    containerEl.appendChild(wrapper);
  });
}

function getSelectedMockTflEntries(containerEl) {
  if (!containerEl) return [];
  const inputs = Array.from(containerEl.querySelectorAll("input[type='checkbox']"));
  return inputs
    .filter((i) => i.checked)
    .map((i) => ({
      type: String(i.dataset.type || "").trim().toUpperCase(),
      number: String(i.value || "").trim(),
      title: String(i.dataset.title || "").trim(),
      order: Number.isFinite(Number(i.dataset.order)) ? Number(i.dataset.order) : null
    }))
    .filter((e) => e.type && e.number && e.title)
    .sort((a, b) => {
      const ao = typeof a.order === "number" ? a.order : Number.MAX_SAFE_INTEGER;
      const bo = typeof b.order === "number" ? b.order : Number.MAX_SAFE_INTEGER;
      if (ao !== bo) return ao - bo;
      if (a.type !== b.type) return a.type.localeCompare(b.type, "en");
      return a.number.localeCompare(b.number, "en");
    });
}

function updateInsertTemplateEnabled(containerEl, buttonEl) {
  if (!buttonEl) return;
  const selected = getSelectedMockTflEntries(containerEl);
  buttonEl.disabled = selected.length === 0;
}

function applyAutoSelection(containerEl, recommended) {
  if (!containerEl) return;
  const recMap = new Map(
    (recommended || []).map((r) => [
      `${String(r?.type || "").trim().toUpperCase()}::${String(r?.number || "").trim()}`,
      !!r?.recommended
    ])
  );
  const inputs = Array.from(containerEl.querySelectorAll("input[type='checkbox']"));
  inputs.forEach((input) => {
    const type = String(input.dataset.type || "").trim().toUpperCase();
    const number = String(input.value || "").trim();
    input.checked = !!recMap.get(`${type}::${number}`);
  });
}

function setMockTflWorkItemOptions(selectEl, entries) {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Select a section...</option>`;
  (entries || []).forEach((e) => {
    const type = String(e.type || "").trim().toUpperCase();
    const number = String(e.number || "").trim();
    const title = String(e.title || "").trim();
    if (!type || !number || !title) return;
    const opt = document.createElement("option");
    opt.value = `${type}:${number}`;
    opt.textContent = `${type} ${number} ${title}`;
    selectEl.appendChild(opt);
  });
}

async function loadMockTflEntries() {
  return await getMockTflExampleEntries();
}

async function runMockTflAutoSelection({ studyNumber, entries, threshold }) {
  return await autoSelectMockTflSections({
    studyNumber,
    tflSections: entries,
    threshold
  });
}

export {
  renderMockTflChecklist,
  getSelectedMockTflEntries,
  updateInsertTemplateEnabled,
  applyAutoSelection,
  setMockTflWorkItemOptions,
  loadMockTflEntries,
  runMockTflAutoSelection
};
