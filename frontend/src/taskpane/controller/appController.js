import { getExistingStudyNumbers } from "../api/studies";

import {
  getSapExampleEntries,
  generateSapSection,
  generateTitlePageMetadata,
  generateAbbreviations,
  extractAbbreviationsFromText,
  autoSelectSapSections,
  clearSapTemplate,
  generateSapFrontMatter,
  generateSapAfterToc,
  updateTableOfContents,
  generateOrUpdateTocAtPlaceholder,
  generateAbbreviationsAtPlaceholder,
  readAbbreviationsFromPlaceholder,
  selectSectionBody,
  readSectionBody,
  writeSectionBody,
  readAllSapBodySections,
  writeTitlePageFields,
  clearTitlePageFields,
  readTitlePageFields
} from "../domains/sap";

import { debounce } from "../shared/utils/debounce";
import { getDom } from "../shared/ui/dom";
import { setStatus } from "../shared/ui/status";
import { showRow, showActions } from "../shared/ui/visibility";
import { setStudyOptions, setSectionOptions, setWorkItemOptions, bindChange } from "../shared/ui/selectors";
import { setAiEnabled, bindAiClick } from "../shared/ui/aiButton";
import {
  showEditor,
  setEditorText,
  getEditorText,
  clearEditor,
  hasEditorText,
  bindEditorInput
} from "../shared/ui/sectionEditor";
import { setInsertEnabled, setClearEnabled, bindInsert, bindClear } from "../shared/ui/actionsBar";
import {
  renderTitlePageEditor,
  showTitlePageEditor,
  getTitlePageValues,
  setTitlePageValues,
  clearTitlePageValues,
  hasAnyTitlePageValue,
  bindTitlePageInput
} from "../shared/ui/titlePageEditor";
import {
  handleSaveSap,
  checkAndPromptRestore,
  setSapSectionMatchMeta,
  clearSapSectionMatchMeta
} from "./sapStateController";
import {
  setupMockTflDocument,
  resetPageOrientationToPortrait,
  clearMockTflTemplate,
  generateMockTflTemplate,
  renderMockTflChecklist,
  getSelectedMockTflEntries,
  updateInsertTemplateEnabled as updateMockTflInsertTemplateEnabled,
  applyAutoSelection as applyMockTflAutoSelection,
  setMockTflWorkItemOptions,
  parseMockTflWorkItemValue,
  buildMockTflLabel,
  handleMockTflWorkItemChange,
  openMockTflEditorDialogForWorkItem,
  closeMockTflEditorDialog,
  loadMockTflEntries,
  runMockTflAutoSelection,
  handleSaveMockTfl
} from "../domains/mocktfl";
import { extractDrugComparison } from "../domains/mocktfl/api/drugComparison";
import { loadMockTflState } from "../domains/mocktfl/api/state";
import { saveMockTflDrugComparison } from "../domains/mocktfl/controller/mockTflStateController";

// Format date as "DDMonYYYY" (e.g., "29Jan2026") to match typical SAP conventions.
function formatSapDate(date) {
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, "0");
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mon = months[d.getMonth()] || "Jan";
  const year = String(d.getFullYear());
  return `${day}${mon}${year}`;
}

// Default document version/date suggestion for the Title Page.
function buildDefaultDocumentVersionDate() {
  return `Version 1.0, ${formatSapDate(new Date())}`;
}

async function initApp() {
  const dom = getDom();

    if (dom.sideload) dom.sideload.style.display = "none";
    if (dom.appBody) dom.appBody.style.display = "block";
    if (dom.hostInfo) dom.hostInfo.textContent = "";

    let currentStudy = "";
    let currentTask = "";
    let currentWorkItemType = "section";
    let currentSectionNumber = "";
    let sapEntries = [];
    let insertedSapEntries = [];
    let mockTflEntries = [];
    let insertedMockTflEntries = [];
    let sectionEditorDialog = null;

    function renderSectionChecklist(entries) {
      if (!dom.sectionCheckboxList) return;
      dom.sectionCheckboxList.innerHTML = "";
      (entries || []).forEach((e) => {
        const number = String(e.number || "").trim();
        const title = String(e.title || "").trim();
        const id = `sap-section-${number.replace(/[^a-zA-Z0-9_-]/g, "_")}`;
        const wrapper = document.createElement("div");
        wrapper.style.display = "flex";
        wrapper.style.gap = "8px";
        wrapper.style.alignItems = "center";
        wrapper.style.marginBottom = "6px";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.id = id;
        checkbox.value = number;
        checkbox.dataset.title = title;
        checkbox.addEventListener("change", updateInsertTemplateEnabled);

        const label = document.createElement("label");
        label.setAttribute("for", id);
        label.textContent = `${number} ${title}`.trim();

        wrapper.appendChild(checkbox);
        wrapper.appendChild(label);
        dom.sectionCheckboxList.appendChild(wrapper);
      });
      updateInsertTemplateEnabled();
    }

    function getSelectedSapEntries() {
      if (!dom.sectionCheckboxList) return [];
      const inputs = Array.from(dom.sectionCheckboxList.querySelectorAll("input[type='checkbox']"));
      return inputs
        .filter((i) => i.checked)
        .map((i) => ({
          number: String(i.value || "").trim(),
          title: String(i.dataset.title || "").trim()
        }))
        .filter((e) => e.number && e.title);
    }

    function updateInsertTemplateEnabled() {
      if (!dom.btnInsertTemplate) return;
      if (currentTask === "MockTFL") {
        updateMockTflInsertTemplateEnabled(dom.sectionCheckboxList, dom.btnInsertTemplate);
      } else {
      const selected = getSelectedSapEntries();
      dom.btnInsertTemplate.disabled = selected.length === 0;
      }
    }

    function getSectionLabel(sectionNumber) {
      if (sectionNumber === "__TITLE_PAGE__") return "Title Page";
      if (sectionNumber === "__TOC__") return "Table of Contents";
      if (sectionNumber === "__LOA__") return "List of Abbreviations";
      const inserted = (insertedSapEntries || []).find((e) => e.number === sectionNumber);
      if (inserted) {
        const displayNumber = inserted.displayNumber || inserted.number || "";
        return `${displayNumber} ${inserted.title || ""}`.trim();
      }
      const found = (sapEntries || []).find((e) => e.number === sectionNumber);
      if (!found) return String(sectionNumber || "");
      return `${found.number} ${found.title}`.trim();
    }

    function getSectionLevel(number) {
      const n = String(number || "").trim();
      if (!n) return 0;
      return n.split(".").length;
    }

    function getParentNumber(number) {
      const n = String(number || "").trim();
      const idx = n.lastIndexOf(".");
      return idx > 0 ? n.slice(0, idx) : "";
    }

    function buildRenumberedEntries(selectedEntries) {
      const allByNumber = new Map(
        (sapEntries || [])
          .map((e) => ({
            ...e,
            number: String(e?.number || "").trim(),
            title: String(e?.title || "").trim()
          }))
          .filter((e) => e.number && e.title)
          .map((e) => [e.number, e])
      );

      // Keep original SAP order as the canonical ordering source.
      const orderedAll = (sapEntries || [])
        .map((e) => ({
          ...e,
          number: String(e?.number || "").trim(),
          title: String(e?.title || "").trim()
        }))
        .filter((e) => e.number && e.title)
        .sort((a, b) => {
          const ao = typeof a.order === "number" ? a.order : Number.MAX_SAFE_INTEGER;
          const bo = typeof b.order === "number" ? b.order : Number.MAX_SAFE_INTEGER;
          if (ao !== bo) return ao - bo;
          return a.number.localeCompare(b.number, "en");
        });

      const included = new Set();
      (selectedEntries || []).forEach((e) => {
        const n = String(e?.number || "").trim();
        if (n) included.add(n);
      });

      // Auto-include missing parents so child sections never become orphans.
      for (const n of Array.from(included)) {
        let p = getParentNumber(n);
        while (p) {
          if (allByNumber.has(p)) included.add(p);
          p = getParentNumber(p);
        }
      }

      const picked = orderedAll.filter((e) => included.has(e.number));
      if (picked.length === 0) return [];

      // Renumber by hierarchy depth while preserving selected order from SAP template.
      const counters = [];
      const renumbered = picked.map((e) => {
        const level = getSectionLevel(e.number);
        if (level <= 0) return { ...e };
        if (counters.length < level) {
          while (counters.length < level) counters.push(0);
        }
        if (counters.length > level) counters.length = level;
        counters[level - 1] = (counters[level - 1] || 0) + 1;
        for (let i = level; i < counters.length; i++) counters[i] = 0;
        const displayNumber = counters.slice(0, level).join(".");
        return {
          ...e,
          displayNumber
        };
      });

      return renumbered;
    }

    function formatAbbrevItems(items) {
      return (items || [])
        .map((it) => {
          const term = String(it?.term || "").trim();
          const def = String(it?.definition || "").trim();
          return term ? `${term}\t${def}` : "";
        })
        .filter(Boolean)
        .join("\n");
    }

    function parseAbbrevItems(text) {
      const lines = String(text || "")
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
      return lines.map((line) => {
        const [term, ...rest] = line.split(/\t+/);
        return {
          term: (term || "").trim(),
          definition: rest.join(" ").trim()
        };
      }).filter((it) => it.term);
    }

    function escapeRegExp(text) {
      return String(text || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    function buildAbbrevPairs(items) {
      const map = new Map();
      (items || []).forEach((it) => {
        const abbr = String(it?.term || "").trim();
        const full = String(it?.definition || "").trim();
        if (!abbr || !full) return;
        const key = full.toLowerCase();
        if (map.has(key)) return;
        map.set(key, { abbr, full });
      });
      const pairs = Array.from(map.values());
      pairs.sort((a, b) => b.full.length - a.full.length);
      return pairs;
    }

    function applyAbbrevPairsToText(text, pairs, seenCounts) {
      let replacements = 0;
      let result = String(text || "");
      pairs.forEach(({ abbr, full }) => {
        const key = full.toLowerCase();
        const pattern = `\\b${escapeRegExp(full)}\\b`;
        const regex = new RegExp(pattern, "gi");
        result = result.replace(regex, (match) => {
          const next = (seenCounts.get(key) || 0) + 1;
          seenCounts.set(key, next);
          if (next === 1) return match;
          replacements += 1;
          return abbr;
        });
      });
      return { text: result, replacements };
    }

    async function generateLoaFromProtocolAndSap({ studyNumber, refineInstruction }) {
      // 1. Extract from Protocol (DISABLED per user request)
      /*
      const baseItems = await generateAbbreviations({ studyNumber });

      const existingMap = new Map();
      baseItems.forEach((it) => {
        const key = String(it?.term || "").trim().toLowerCase();
        if (key) existingMap.set(key, it);
      });
      */
      const existingMap = new Map(); // Start empty
      const existingDefSet = new Set(); // To track definitions (explanations)

      // 2. Scan SAP sections
      const sections = await readAllSapBodySections({ studyNumber });

      const MAX_BATCH_CHARS = 2500;
      let batchText = "";

      const flushBatch = async () => {
        if (!batchText.trim()) return;
        const existingTerms = Array.from(existingMap.keys());
        const newItems = await extractAbbreviationsFromText({
          text: batchText,
          existingTerms
        });
        newItems.forEach((it) => {
          const key = String(it?.term || "").trim().toLowerCase();
          const defKey = String(it?.definition || "").trim().toLowerCase();
          
          // Dedup by Abbreviation (key) AND Definition (defKey)
          if (!key || existingMap.has(key)) return;
          if (!defKey || existingDefSet.has(defKey)) return;
          
          existingMap.set(key, it);
          existingDefSet.add(defKey);
        });
        batchText = "";
      };

      for (const sec of sections) {
        const text = String(sec?.text || "").trim();
        if (!text) continue;
        const labeled = `[${sec.sectionNumber}]\n${text}\n\n`;
        if ((batchText + labeled).length > MAX_BATCH_CHARS) {
          await flushBatch();
        }
        batchText += labeled;
      }
      await flushBatch();

      return Array.from(existingMap.values());
    }

    function closeSectionEditorDialog() {
      try {
        if (sectionEditorDialog) sectionEditorDialog.close();
      } catch (e) {}
      sectionEditorDialog = null;
    }

    function sendToDialog(payload) {
      try {
        if (!sectionEditorDialog) return;
        sectionEditorDialog.messageChild(JSON.stringify(payload));
      } catch (e) {}
    }

    async function openSectionEditorDialogForSection({ studyNumber, sectionNumber, autoGenerate = false }) {
      if (!studyNumber || !sectionNumber) return;
      if (!Office?.context?.ui?.displayDialogAsync) {
        setStatus(dom.status, "Dialog API is not available in this host.");
        return;
      }

      closeSectionEditorDialog();

      const dialogUrl = `${window.location.origin}/sectionEditorDialog.html`;
      const label = getSectionLabel(sectionNumber);

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 50, displayInIframe: false },
        async (asyncResult) => {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            setStatus(dom.status, `Failed to open editor dialog: ${asyncResult.error?.message || asyncResult.error}`);
            return;
          }

          sectionEditorDialog = asyncResult.value;

          sectionEditorDialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
            let msg = null;
            try {
              msg = JSON.parse(arg?.message || "{}");
            } catch (e) {
              msg = null;
            }
            if (!msg?.type) return;

            if (msg.type === "dialogReady") {
              let initialText = "";
              let titlePageValues = null;
              if (sectionNumber === "__TITLE_PAGE__") {
                titlePageValues = await readTitlePageFields({ studyNumber });
              } else if (sectionNumber === "__LOA__") {
                const items = await readAbbreviationsFromPlaceholder();
                initialText = formatAbbrevItems(items);
              } else {
                initialText = await readSectionBody({ studyNumber, sectionNumber });
              }

              sendToDialog({
                type: "init",
                studyNumber,
                sectionNumber,
                label,
                initialText,
                titlePageValues,
                autoGenerate: !!autoGenerate
              });
              return;
            }

            if (msg.type === "generate" || msg.type === "regenerate") {
              try {
                let generatedText = "";
                if (sectionNumber === "__TITLE_PAGE__") {
                  const metadata = await generateTitlePageMetadata({ studyNumber });
                  metadata.documentVersionDate = buildDefaultDocumentVersionDate();
                  sendToDialog({ type: "generatedTitlePage", values: metadata });
                } else if (sectionNumber === "__LOA__") {
                   const mergedItems = await generateLoaFromProtocolAndSap({
                     studyNumber,
                     refineInstruction: msg.modifyPrompt
                   });
                   generatedText = formatAbbrevItems(mergedItems);
                } else {
                   const result = await generateSapSection({
                    studyNumber,
                    sapSectionNumber: sectionNumber,
                    refineInstruction: msg.modifyPrompt
                  });
                   generatedText = result?.generatedText || "";

                   // Persist matched context metadata in-memory so it can be saved to DB later.
                   // We store this under the SAP sectionNumber being edited.
                   // Example:
                   // - matched protocol section: "6.3 Patient Withdrawal from Treatment"
                   // - matched CRF form: "END OF TREATMENT PERIOD B"
                   try {
                     setSapSectionMatchMeta({
                       sectionNumber,
                       protocolSectionNumber: result?.matchedProtocolSection?.number || "",
                       protocolSectionTitle: result?.matchedProtocolSection?.title || "",
                       crfFormName: result?.crfContext?.formName || ""
                     });
                   } catch (e) {}

                   sendToDialog({
                     type: "generatedContext",
                     protocolSectionLabel: [
                       result?.matchedProtocolSection?.number || "",
                       result?.matchedProtocolSection?.title || ""
                     ]
                       .join(" ")
                       .trim(),
                     protocolText: result?.protocolContextText || "",
                     crfFormName: result?.crfContext?.formName || "",
                     crfText: result?.crfContext?.fullText || ""
                   });
                }

                if (sectionNumber === "__TITLE_PAGE__") return;
                if (!generatedText && sectionNumber !== "__LOA__") throw new Error("AI generation failed");
                sendToDialog({ type: "generated", text: generatedText });
              } catch (e) {
                sendToDialog({ type: "error", message: e?.message || String(e) });
              }
              return;
            }

            if (msg.type === "applyAbbrevToBody") {
              try {
                sendToDialog({ type: "applyProgress", message: "Reading abbreviations list..." });
                const items = parseAbbrevItems(msg.abbrevText || "");
                const pairs = buildAbbrevPairs(items);
                if (pairs.length === 0) {
                  sendToDialog({ type: "error", message: "No valid abbreviations found." });
                  return;
                }

                sendToDialog({ type: "applyProgress", message: "Reading SAP body sections..." });
                const sections = await readAllSapBodySections({ studyNumber });

                // Preserve SAP entry order when possible
                const map = new Map(sections.map((s) => [s.sectionNumber, s.text]));
                const ordered = [];
                (sapEntries || []).forEach((e) => {
                  if (map.has(e.number)) {
                    ordered.push({ sectionNumber: e.number, text: map.get(e.number) });
                    map.delete(e.number);
                  }
                });
                map.forEach((text, sectionNumber) => ordered.push({ sectionNumber, text }));

                const seenCounts = new Map();
                let totalReplacements = 0;
                let sectionsChanged = 0;

                for (const sec of ordered) {
                  const original = String(sec.text || "");
                  const { text: updated, replacements } = applyAbbrevPairsToText(original, pairs, seenCounts);
                  if (updated !== original) {
                    await writeSectionBody({
                      studyNumber,
                      sectionNumber: sec.sectionNumber,
                      text: updated
                    });
                    sectionsChanged += 1;
                    totalReplacements += replacements;
                  }
                }

                sendToDialog({
                  type: "applyDone",
                  replacementsMade: totalReplacements,
                  sectionsChanged
                });
              } catch (e) {
                sendToDialog({ type: "error", message: e?.message || String(e) });
              }
              return;
            }

            if (msg.type === "insert") {
              try {
                if (sectionNumber === "__TITLE_PAGE__") {
                  await writeTitlePageFields({
                    studyNumber,
                    values: msg.values || {}
                  });
                } else if (sectionNumber === "__LOA__") {
                  const items = parseAbbrevItems(msg.text || "");
                  await generateAbbreviationsAtPlaceholder(items);
                } else {
                  await writeSectionBody({
                    studyNumber,
                    sectionNumber,
                    text: msg.text || ""
                  });
                }
                sendToDialog({ type: "inserted" });
              } catch (e) {
                sendToDialog({ type: "error", message: e?.message || String(e) });
              }
              return;
            }

            if (msg.type === "clear") {
              try {
                if (sectionNumber === "__TITLE_PAGE__") {
                  await clearTitlePageFields({ studyNumber });
                } else if (sectionNumber === "__LOA__") {
                  await generateAbbreviationsAtPlaceholder([]);
                } else {
                  await writeSectionBody({ studyNumber, sectionNumber, text: "" });
                }
                sendToDialog({ type: "cleared" });
              } catch (e) {
                sendToDialog({ type: "error", message: e?.message || String(e) });
              }
            }
          });

          sectionEditorDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            // User closed dialog.
            closeSectionEditorDialog();
          });
        }
      );
    }

    function resetTaskAndBelow() {
      currentTask = "";
      currentWorkItemType = "section";
      currentSectionNumber = "";
      if (dom.taskSelect) dom.taskSelect.value = "";
      if (dom.selectMode) dom.selectMode.value = "";
      if (dom.sectionCheckboxList) dom.sectionCheckboxList.innerHTML = "";
      if (dom.btnInsertTemplate) dom.btnInsertTemplate.disabled = true;
      if (dom.sectionSelect) {
        dom.sectionSelect.innerHTML = `<option value="">Select a section...</option>`;
      }
      clearEditor(dom.editor);
      clearTitlePageValues(dom.titlePageRoot);
      showRow(dom.taskRow, false);
      showRow(dom.selectModeRow, false);
      showRow(dom.sectionSelectionRow, false);
      showRow(dom.sectionRow, false);
      showRow(dom.aiRow, false);
      showRow(dom.openEditorRow, false);
      showRow(dom.finishRow, false);
      showEditor(dom.editorRow, false);
      showTitlePageEditor(dom.titlePageRow, false);
      showActions(dom.actionRow, false);
      setAiEnabled(dom.btnAi, false);
      if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
      if (dom.btnFinishSap) dom.btnFinishSap.disabled = true;
      setInsertEnabled(dom.btnInsert, false);
      setClearEnabled(dom.btnClear, false);
      showRow(dom.saveRow, false);
      if (dom.btnSaveDb) dom.btnSaveDb.disabled = true;
      closeSectionEditorDialog();
      closeMockTflEditorDialog();
      insertedSapEntries = [];
      insertedMockTflEntries = [];
      mockTflEntries = [];

      // Clear cached SAP matched-context metadata when resetting task/workspace state.
      try {
        clearSapSectionMatchMeta();
      } catch (e) {}
    }

    function resetSectionAndEditor() {
      currentWorkItemType = "section";
      currentSectionNumber = "";
      if (dom.selectMode) dom.selectMode.value = "";
      if (dom.sectionCheckboxList) dom.sectionCheckboxList.innerHTML = "";
      if (dom.btnInsertTemplate) dom.btnInsertTemplate.disabled = true;
      if (dom.sectionSelect) {
        dom.sectionSelect.value = "";
        dom.sectionSelect.innerHTML = `<option value="">Select a section...</option>`;
      }
      clearEditor(dom.editor);
      clearTitlePageValues(dom.titlePageRoot);
      showRow(dom.selectModeRow, false);
      showRow(dom.sectionSelectionRow, false);
      showRow(dom.sectionRow, false);
      showRow(dom.aiRow, false);
      showRow(dom.openEditorRow, false);
      showRow(dom.finishRow, false);
      showEditor(dom.editorRow, false);
      showTitlePageEditor(dom.titlePageRow, false);
      showActions(dom.actionRow, false);
      setAiEnabled(dom.btnAi, false);
      if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
      if (dom.btnFinishSap) dom.btnFinishSap.disabled = true;
      setInsertEnabled(dom.btnInsert, false);
      setClearEnabled(dom.btnClear, false);
      showRow(dom.saveRow, false);
      if (dom.btnSaveDb) dom.btnSaveDb.disabled = true;
      closeSectionEditorDialog();
      closeMockTflEditorDialog();
      insertedSapEntries = [];
      insertedMockTflEntries = [];
      mockTflEntries = [];
    }

    async function loadStudies() {
      try {
        setStatus(dom.status, "Loading study numbers...");
        const studies = await getExistingStudyNumbers();
        setStudyOptions(dom.studySelect, studies);
        setStatus(dom.status, "");
      } catch (e) {
        setStatus(dom.status, `Failed to load studies: ${e.message || e}`);
      }
    }

    async function insertSapTemplate(selectedEntries) {
      if (!currentStudy || !Array.isArray(selectedEntries) || selectedEntries.length === 0) return;
      try {
        const finalEntries = buildRenumberedEntries(selectedEntries);
        if (finalEntries.length === 0) {
          setStatus(dom.status, "No valid SAP sections selected.");
          return;
        }

        setStatus(dom.status, "Clearing existing SAP template...");
        await clearSapTemplate();

        // Strict page order:
        // 1) Insert Title/Signature -> Revision History -> TOC
        setStatus(dom.status, "Inserting SAP template (front matter)...");
        await generateSapFrontMatter({ studyNumber: currentStudy, entries: finalEntries });

        // 2) Insert List of Abbreviations (placeholder) -> main sections
        setStatus(dom.status, "Inserting SAP template (after TOC)...");
        await generateSapAfterToc({
          studyNumber: currentStudy,
          entries: finalEntries
        });

        setWorkItemOptions(dom.sectionSelect, finalEntries);
        insertedSapEntries = finalEntries;

        showRow(dom.selectModeRow, false);
        showRow(dom.sectionSelectionRow, false);
        showRow(dom.sectionRow, true);
        showRow(dom.aiRow, true);
        showEditor(dom.editorRow, false);
        showTitlePageEditor(dom.titlePageRow, false);
        showActions(dom.actionRow, false);
        setAiEnabled(dom.btnAi, false);
        showRow(dom.finishRow, true);
        if (dom.btnFinishSap) dom.btnFinishSap.disabled = false;
        setInsertEnabled(dom.btnInsert, false);
        setClearEnabled(dom.btnClear, false);
        // Placeholder mode: updateTableOfContents is a no-op.
        setStatus(dom.status, "Finalizing...");
        try {
          await updateTableOfContents();
        } catch (e) {}
        
        // Show Save button
        showRow(dom.saveRow, true);
        if (dom.btnSaveDb) dom.btnSaveDb.disabled = false;

        setStatus(dom.status, "SAP template ready.");
      } catch (e) {
        const msg = e?.message || String(e);
        setStatus(dom.status, `Failed to generate SAP template: ${msg}`);
        // eslint-disable-next-line no-console
        console.error("SAP template generation error:", e);
      }
    }

    async function insertMockTflTemplate(selectedEntries) {
      if (!currentStudy || !Array.isArray(selectedEntries) || selectedEntries.length === 0) return;
      try {
        setStatus(dom.status, "Clearing existing Mock TFL template...");
        await clearMockTflTemplate();

        setStatus(dom.status, "Inserting Mock TFL placeholders...");
        await generateMockTflTemplate({
          studyNumber: currentStudy,
          entries: selectedEntries
        });

        insertedMockTflEntries = selectedEntries;
        setMockTflWorkItemOptions(dom.sectionSelect, selectedEntries);
        if (dom.sectionSelect) dom.sectionSelect.value = "";

        showRow(dom.selectModeRow, false);
        showRow(dom.sectionSelectionRow, false);
        showRow(dom.drugComparisonRow, true);
        showRow(dom.sectionRow, true);
        showRow(dom.aiRow, false);
        showRow(dom.openEditorRow, true);
        if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
        showRow(dom.finishRow, false);
        showEditor(dom.editorRow, false);
        showTitlePageEditor(dom.titlePageRow, false);
        showActions(dom.actionRow, false);
        setAiEnabled(dom.btnAi, false);
        setInsertEnabled(dom.btnInsert, false);
        setClearEnabled(dom.btnClear, false);
        showRow(dom.saveRow, true);
        if (dom.btnSaveDb) dom.btnSaveDb.disabled = false;

        setStatus(dom.status, "Mock TFL template ready.");
      } catch (e) {
        setStatus(dom.status, `Failed to generate Mock TFL template: ${e.message || e}`);
      }
    }

    bindChange(dom.studySelect, async () => {
      const newStudy = dom.studySelect?.value || "";
      if (newStudy !== currentStudy) {
        currentStudy = newStudy;
        resetTaskAndBelow();
        if (currentStudy) {
          showRow(dom.taskRow, true);
        }
        try {
          await clearSapTemplate();
          await clearMockTflTemplate();
        } catch (e) {}
      }
    });

    bindChange(dom.taskSelect, async () => {
      currentTask = dom.taskSelect?.value || "";
      resetSectionAndEditor();
      if (currentTask === "MockTFL") {
        if (!currentStudy) return;
        try {
          setStatus(dom.status, "Loading Mock TFL entries...");
          mockTflEntries = await loadMockTflEntries();
          if (mockTflEntries.length === 0) {
            setStatus(dom.status, "No Mock TFL entries found.");
            return;
          }
          renderMockTflChecklist({
            containerEl: dom.sectionCheckboxList,
            entries: mockTflEntries,
            onSelectionChange: updateInsertTemplateEnabled
          });

          await setupMockTflDocument();
          
          // Try to load any existing drug comparison text
          const stateResp = await loadMockTflState({ studyNumber: currentStudy });
          if (stateResp?.success && stateResp?.data) {
            if (dom.drug1Input && stateResp.data.drug1) dom.drug1Input.value = stateResp.data.drug1;
            if (dom.drug2Input && stateResp.data.drug2) dom.drug2Input.value = stateResp.data.drug2;
          }

          showRow(dom.selectModeRow, true);
          showRow(dom.sectionSelectionRow, false);
          showRow(dom.sectionRow, false);
          showRow(dom.aiRow, false);
          showRow(dom.openEditorRow, false);
          showRow(dom.finishRow, false);
          showRow(dom.saveRow, false);
          setStatus(dom.status, "Select 'Start selecting' to review Mock TFL items.");
        } catch (e) {
          setStatus(dom.status, `Failed to setup Mock TFL: ${e.message || e}`);
        }
        setAiEnabled(dom.btnAi, false);
        return;
      }

      if (currentTask === "SAP") {
        if (!currentStudy) return;
        try {
          // Ensure we are back to Portrait when switching to SAP
          await resetPageOrientationToPortrait();
          
          setStatus(dom.status, "Loading SAP example entries...");
          sapEntries = await getSapExampleEntries();
          if (sapEntries.length === 0) {
            setStatus(dom.status, "No SAP example entries found.");
            return;
          }
          renderSectionChecklist(sapEntries);
          showRow(dom.selectModeRow, true);
          showRow(dom.sectionSelectionRow, false);
          setStatus(dom.status, "Select 'Start selecting' to review SAP sections.");
          
          // Check for saved state and prompt restore
          await checkAndPromptRestore({
              studyNumber: currentStudy,
              onRestoreConfirmed: (restoredEntries) => {
                  // After restore, we need to sync UI state
                  insertedSapEntries = restoredEntries;

                  // Restore mode: skip "select sections" UI and go directly to working mode.
                  if (dom.selectMode) dom.selectMode.value = "";
                  showRow(dom.selectModeRow, false);
                  showRow(dom.sectionSelectionRow, false);
                  showRow(dom.sectionRow, true);
                  showRow(dom.aiRow, false);
                  showRow(dom.openEditorRow, true);
                  if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = false;
                  showRow(dom.finishRow, true);
                  if (dom.btnFinishSap) dom.btnFinishSap.disabled = false;
                  showRow(dom.saveRow, true);
                  if (dom.btnSaveDb) dom.btnSaveDb.disabled = false;
              },
              onCancel: () => {
                  // Do nothing, stay in "Start selecting" state
              }
          });

        } catch (e) {
          setStatus(dom.status, `Failed to load SAP entries: ${e.message || e}`);
        }
        setAiEnabled(dom.btnAi, false);
      } else {
        setAiEnabled(dom.btnAi, false);
      }
    });

    bindChange(dom.selectMode, async () => {
      const mode = dom.selectMode?.value || "";
      if (mode === "start") {
        showRow(dom.sectionSelectionRow, true);
        updateInsertTemplateEnabled();
      } else {
        showRow(dom.sectionSelectionRow, false);
      }
    });

    if (dom.btnAutoSelect) {
      dom.btnAutoSelect.addEventListener("click", async () => {
        if (!currentStudy) return;
        try {
          setStatus(dom.status, "Running auto selection...");
          if (currentTask === "MockTFL") {
            if (mockTflEntries.length === 0) return;
            const recommended = await runMockTflAutoSelection({
              studyNumber: currentStudy,
              entries: mockTflEntries
            });
            applyMockTflAutoSelection(dom.sectionCheckboxList, recommended);
          } else {
            if (sapEntries.length === 0) return;
          const recommended = await autoSelectSapSections({
            studyNumber: currentStudy,
            sapSections: sapEntries
          });
          const recMap = new Map(recommended.map((r) => [String(r.number), r.recommended]));
          const inputs = Array.from(dom.sectionCheckboxList.querySelectorAll("input[type='checkbox']"));
          inputs.forEach((input) => {
            const key = String(input.value || "");
            input.checked = !!recMap.get(key);
          });
          }
          updateInsertTemplateEnabled();
          setStatus(dom.status, "Auto selection complete. Please review and click Insert Template.");
        } catch (e) {
          setStatus(dom.status, `Auto selection failed: ${e.message || e}`);
        }
      });
    }

    if (dom.btnInsertTemplate) {
      dom.btnInsertTemplate.addEventListener("click", async () => {
        if (currentTask === "MockTFL") {
          const selected = getSelectedMockTflEntries(dom.sectionCheckboxList);
          if (selected.length === 0) return;
          await insertMockTflTemplate(selected);
        } else {
        const selected = getSelectedSapEntries();
        if (selected.length === 0) return;
        await insertSapTemplate(selected);
        }
      });
    }

    if (dom.btnSaveDb) {
      dom.btnSaveDb.addEventListener("click", async () => {
        if (currentTask === "MockTFL") {
          await handleSaveMockTfl({
            studyNumber: currentStudy,
            insertedEntries: insertedMockTflEntries
          });
        } else {
        await handleSaveSap({
          studyNumber: currentStudy,
          insertedSapEntries
        });
        }
      });
    }

    bindAiClick(dom.btnAi, async () => {
      const selection = dom.sectionSelect?.value || "";
      if (!currentStudy || !selection) return;

      if (selection === "__TITLE_PAGE__") {
        try {
          setStatus(dom.status, "Generating title page metadata...");
          const metadata = await generateTitlePageMetadata({ studyNumber: currentStudy });
          // Always prefill "Document Version, Date" with Version 1.0 + today's date for review.
          metadata.documentVersionDate = buildDefaultDocumentVersionDate();
          setTitlePageValues(dom.titlePageRoot, metadata);
          const hasText = hasAnyTitlePageValue(dom.titlePageRoot);
          setInsertEnabled(dom.btnInsert, hasText);
          setClearEnabled(dom.btnClear, hasText);
          setStatus(dom.status, "Title page suggestions ready.");
        } catch (e) {
          setStatus(dom.status, `Title page generation failed: ${e.message || e}`);
        }
        return;
      }

      if (selection === "__TOC__") {
        try {
          setStatus(dom.status, "Locating TOC placeholder...");
          const result = await generateOrUpdateTocAtPlaceholder();
          if (result === "updated") {
            setStatus(dom.status, "Table of Contents updated.");
          } else {
            setStatus(dom.status, "Placeholder selected. Please insert TOC via References -> Table of Contents.");
          }
        } catch (e) {
          setStatus(dom.status, `TOC operation failed: ${e.message || e}`);
        }
        return;
      }

      // For normal sections and LOA, we do not use this button anymore (Open Editor is used instead).
      // Keep existing AI behaviors for special work items only.
      if (selection && selection !== "__TITLE_PAGE__" && selection !== "__TOC__") {
        setStatus(dom.status, "Click 'Open Editor' to work on this section.");
      }
    });

    // Open Editor button: opens dialog for the currently selected normal section.
    if (dom.btnOpenEditor) {
      dom.btnOpenEditor.addEventListener("click", async () => {
        if (!currentStudy) return;
        const selection = dom.sectionSelect?.value || "";
        if (!selection) return;
        if (selection === "__TOC__") return;

        if (currentTask === "MockTFL") {
          const parsed = parseMockTflWorkItemValue(selection);
          if (!parsed) return;
          const entry = (insertedMockTflEntries || []).find(
            (e) =>
              String(e?.type || "").trim().toUpperCase() === parsed.type &&
              String(e?.number || "").trim() === parsed.number
          );
          const label = buildMockTflLabel({
            type: parsed.type,
            number: parsed.number,
            title: entry?.title || ""
          });
          try {
            await openMockTflEditorDialogForWorkItem({
              studyNumber: currentStudy,
              type: parsed.type,
              number: parsed.number,
              label
            });
          } catch (e) {
            setStatus(dom.status, `Failed to open MockTFL editor: ${e?.message || e}`);
          }
          return;
        }
        
        await openSectionEditorDialogForSection({
          studyNumber: currentStudy,
          sectionNumber: selection,
          autoGenerate: false
        });
      });
    }

    if (dom.btnExtractDrugComparison) {
      dom.btnExtractDrugComparison.addEventListener("click", async () => {
        if (!currentStudy) return;
        try {
          setStatus(dom.status, "Analyzing SAP introduction for drug comparison...");
          if (dom.btnExtractDrugComparison) dom.btnExtractDrugComparison.disabled = true;
          
          const result = await extractDrugComparison({ studyNumber: currentStudy });
          
          if (dom.drug1Input) dom.drug1Input.value = result.drug1 || "Drug 1";
          if (dom.drug2Input) dom.drug2Input.value = result.drug2 || "Drug 2";
          
          setStatus(dom.status, "Drug comparison extraction complete. Please review and save.");
        } catch (e) {
          setStatus(dom.status, `Failed to extract drug comparison: ${e.message || e}`);
        } finally {
          if (dom.btnExtractDrugComparison) dom.btnExtractDrugComparison.disabled = false;
        }
      });
    }

    if (dom.btnSaveDrugComparison) {
      dom.btnSaveDrugComparison.addEventListener("click", async () => {
        if (!currentStudy) return;
        try {
          setStatus(dom.status, "Saving drug comparison to database...");
          if (dom.btnSaveDrugComparison) dom.btnSaveDrugComparison.disabled = true;
          
          const drug1 = dom.drug1Input ? dom.drug1Input.value : "";
          const drug2 = dom.drug2Input ? dom.drug2Input.value : "";
          await saveMockTflDrugComparison({
            studyNumber: currentStudy,
            drug1,
            drug2
          });
          
          setStatus(dom.status, "Drug comparison saved successfully.");
        } catch (e) {
          setStatus(dom.status, `Failed to save drug comparison: ${e.message || e}`);
        } finally {
          if (dom.btnSaveDrugComparison) dom.btnSaveDrugComparison.disabled = false;
        }
      });
    }

    if (dom.btnFinishSap) {
      dom.btnFinishSap.addEventListener("click", async () => {
        if (!currentStudy) return;
        await openSectionEditorDialogForSection({
          studyNumber: currentStudy,
          sectionNumber: "__LOA__",
          autoGenerate: false
        });
      });
    }

    bindChange(dom.sectionSelect, async () => {
      const selection = dom.sectionSelect?.value || "";
      if (currentTask === "MockTFL") {
        // MockTFL work items are edited via MockTFL dialog and always map to a body Content Control.
        showEditor(dom.editorRow, false);
        showActions(dom.actionRow, false);
        showTitlePageEditor(dom.titlePageRow, false);
        showRow(dom.aiRow, false);
        showRow(dom.openEditorRow, true);

        if (!selection) {
          if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
          setAiEnabled(dom.btnAi, false);
          return;
        }

        try {
          await handleMockTflWorkItemChange({ studyNumber: currentStudy, value: selection });
          if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = false;
        } catch (e) {
          if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
          setStatus(dom.status, `Failed to select MockTFL item: ${e?.message || e}`);
        }
        setAiEnabled(dom.btnAi, false);
        return;
      }

      if (!selection) {
        setAiEnabled(dom.btnAi, false);
        showRow(dom.openEditorRow, false);
        if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
        setInsertEnabled(dom.btnInsert, false);
        setClearEnabled(dom.btnClear, false);
        return;
      }

      if (selection === "__TITLE_PAGE__") {
        currentWorkItemType = "titlePage";
        currentSectionNumber = "__TITLE_PAGE__";
        showEditor(dom.editorRow, false);
        showTitlePageEditor(dom.titlePageRow, false);
        showActions(dom.actionRow, false);
        showRow(dom.aiRow, false);
        showRow(dom.openEditorRow, true);
        if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = false;
        setAiEnabled(dom.btnAi, false);
        return;
      }

      if (selection === "__TOC__") {
        currentWorkItemType = "toc";
        currentSectionNumber = "";
        showEditor(dom.editorRow, false);
        showTitlePageEditor(dom.titlePageRow, false);
        showActions(dom.actionRow, false); // No insert/clear needed for TOC
        showRow(dom.openEditorRow, false);
        if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = true;
        setAiEnabled(dom.btnAi, true);
        dom.btnAi.textContent = "Select TOC Placeholder";
        return;
      }

      // Reset button text for normal sections
      dom.btnAi.textContent = "Generate with AI";
      showRow(dom.aiRow, true);

      currentWorkItemType = "section";
      currentSectionNumber = selection;
      showTitlePageEditor(dom.titlePageRow, false);
      await selectSectionBody({ studyNumber: currentStudy, sectionNumber: currentSectionNumber });
      setAiEnabled(dom.btnAi, true);
      // Hide inline editor/actions; section editing is done in dialog.
      showEditor(dom.editorRow, false);
      showActions(dom.actionRow, false);
      // Show Open Editor button for normal sections.
      showRow(dom.aiRow, false);
      showRow(dom.openEditorRow, true);
      if (dom.btnOpenEditor) dom.btnOpenEditor.disabled = false;
    });

    const onEditorInput = debounce(async () => {
      if (currentWorkItemType !== "section") return;
      const hasText = hasEditorText(dom.editor);
      setInsertEnabled(dom.btnInsert, hasText);
      setClearEnabled(dom.btnClear, hasText);
    }, 200);

    bindEditorInput(dom.editor, onEditorInput);
    renderTitlePageEditor(dom.titlePageRoot);
    bindTitlePageInput(dom.titlePageRoot, debounce(() => {
      if (currentWorkItemType !== "titlePage") return;
      const hasText = hasAnyTitlePageValue(dom.titlePageRoot);
      setInsertEnabled(dom.btnInsert, hasText);
      setClearEnabled(dom.btnClear, hasText);
    }, 200));

    bindInsert(dom.btnInsert, async () => {
      if (!currentStudy) return;
      try {
        setStatus(dom.status, "Inserting into document...");
        if (currentWorkItemType === "titlePage") {
          await writeTitlePageFields({
            studyNumber: currentStudy,
            values: getTitlePageValues(dom.titlePageRoot)
          });
        } else {
          await writeSectionBody({
            studyNumber: currentStudy,
            sectionNumber: currentSectionNumber,
            text: getEditorText(dom.editor)
          });
        }
        setStatus(dom.status, "Inserted.");
      } catch (e) {
        setStatus(dom.status, `Insert failed: ${e.message || e}`);
      }
    });

    bindClear(dom.btnClear, async () => {
      if (!currentStudy) return;
      try {
        if (currentWorkItemType === "titlePage") {
          clearTitlePageValues(dom.titlePageRoot);
          await clearTitlePageFields({ studyNumber: currentStudy });
        } else {
          clearEditor(dom.editor);
          await writeSectionBody({
            studyNumber: currentStudy,
            sectionNumber: currentSectionNumber,
            text: ""
          });
        }
        setInsertEnabled(dom.btnInsert, false);
        setClearEnabled(dom.btnClear, false);
        setStatus(dom.status, "Cleared.");
      } catch (e) {
        setStatus(dom.status, `Clear failed: ${e.message || e}`);
      }
    });

  await loadStudies();
}

export { initApp };
