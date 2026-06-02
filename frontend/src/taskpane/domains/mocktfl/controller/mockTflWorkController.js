/* global Office */

import {
  selectMockTflBody,
  readMockTflTitle,
  extractPureTitleFromMockTflTitleCcText
} from "../word/sections";
import { startMockTflAnalysis } from "../api/startAnalysis";
import { mapMockTflVariablesToCrf } from "../api/mapVariablesToCrf";
import { generateMockTflFigure } from "../api/generateFigure";
import { autoGenerateMockTflFigure } from "../api/autoGenerateFigure";
import { findCorrespondingTable } from "../api/findCorrespondingTable";
import { defineMockTflFigureType } from "../api/defineFigureType";
import { loadMockTflState } from "../api/state";
import {
  findMockTflSectionData,
  saveMockTflSectionDraft,
  saveMockTflSectionSnapshot
} from "./mockTflStateController";
import { insertTable14_1_1_2_PatientDisposition } from "../word/tables/table14_1_1_2_patientDisposition";
import { insertTable14_1_3_1_SummaryOfDemographics } from "../word/tables/table14_1_3_1_summaryOfDemographics";
import { insertTable14_2_1_1_PrimaryEfficacyEndpoint } from "../word/tables/table14_2_1_1_primaryEfficacyEndpoint";
import { insertMockTflFigure } from "../word/figures/insertFigure";

function parseMockTflWorkItemValue(value) {
  const raw = String(value || "").trim();
  const idx = raw.indexOf(":");
  if (idx <= 0) return null;
  const type = raw.slice(0, idx).trim().toUpperCase();
  const number = raw.slice(idx + 1).trim();
  if (!type || !number) return null;
  return { type, number };
}

function buildMockTflLabel({ type, number, title }) {
  const t = String(type || "").trim().toUpperCase();
  const n = String(number || "").trim();
  const ti = String(title || "").trim();
  return `${t} ${n} ${ti}`.trim();
}

let mockTflEditorDialog = null;

// Feature gate:
// We only enable the analysis prototype for a small allowlist of work items for now.
// IMPORTANT: Keep this strict to avoid enabling analysis for all items unintentionally.
const ANALYSIS_ALLOWLIST = new Set(["TABLE:14.1.1.2", "TABLE:14.1.3.1", "TABLE:14.2.1.1"]);

// Table insertion gate:
// For these work items, clicking "Insert" generates a structured Word table (not plain text).
const TABLE_INSERT_ALLOWLIST = new Set(["TABLE:14.1.1.2", "TABLE:14.1.3.1", "TABLE:14.2.1.1"]);

// Figure generation gate:
// Keep this prototype scoped to the first requested Figure item.
const FIGURE_ALLOWLIST = new Set(["FIGURE:14.2.1.1"]);

function closeMockTflEditorDialog() {
  try {
    if (mockTflEditorDialog) mockTflEditorDialog.close();
  } catch (e) {}
  mockTflEditorDialog = null;
}

function safeJsonParse(s) {
  try {
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

function sendToDialog(payload) {
  try {
    if (!mockTflEditorDialog) return;
    mockTflEditorDialog.messageChild(JSON.stringify(payload));
  } catch (e) {}
}

function isAnalysisTarget({ type, number }) {
  const t = String(type || "").trim().toUpperCase();
  const n = String(number || "").trim();
  return ANALYSIS_ALLOWLIST.has(`${t}:${n}`);
}

function isTableInsertTarget({ type, number }) {
  const t = String(type || "").trim().toUpperCase();
  const n = String(number || "").trim();
  return TABLE_INSERT_ALLOWLIST.has(`${t}:${n}`);
}

function isFigureTarget({ type, number }) {
  const t = String(type || "").trim().toUpperCase();
  const n = String(number || "").trim();
  return FIGURE_ALLOWLIST.has(`${t}:${n}`);
}

/**
 * Group CRF mappings by the final matched CRF question.
 *
 * Why:
 * - Several suggested variables may map to the same CRF question.
 * - We want both the evidence box and the final Notes box to use the same
 *   grouped source of truth.
 */
function groupCrfMappingsByQuestion(mappings) {
  const arr = Array.isArray(mappings) ? mappings : [];
  const groups = new Map();

  arr.forEach((m, idx) => {
    const bulletIndex = Number(m?.bulletIndex || idx + 1);
    const bulletText = String(m?.bulletText || "").trim() || "(Empty bullet text)";
    const originalQuestionText = String(m?.matchedQuestion?.questionText || "").trim() || "(No matched question)";
    const rewrittenQuestionText = String(m?.matchedQuestion?.rewrittenQuestionText || originalQuestionText).trim();
    const questionScore = typeof m?.matchedQuestion?.score === "number" ? m.matchedQuestion.score : null;
    const answers = Array.isArray(m?.answers) ? m.answers : [];

    if (!groups.has(originalQuestionText)) {
      groups.set(originalQuestionText, {
        originalQuestionText,
        rewrittenQuestionText,
        score: questionScore,
        variables: [],
        answers: []
      });
    }

    const group = groups.get(originalQuestionText);
    group.variables.push({ index: bulletIndex, text: bulletText });

    if (typeof questionScore === "number" && (!Number.isFinite(group.score) || questionScore > group.score)) {
      group.score = questionScore;
    }

    const seenAnswerTexts = new Set(group.answers.map((a) => String(a?.text || "").trim()));
    answers.forEach((answer) => {
      const text = String(answer?.text || "").trim();
      if (!text || seenAnswerTexts.has(text)) return;
      seenAnswerTexts.add(text);
      group.answers.push(answer);
    });
  });

  return Array.from(groups.values());
}

/**
 * Build detailed evidence text for the read-only mapping box.
 * This uses grouped mappings so repeated CRF questions are shown once.
 */
function formatGroupedMappingsForEvidence(groupedMappings) {
  const groups = Array.isArray(groupedMappings) ? groupedMappings : [];
  if (groups.length === 0) return "";

  const blocks = groups.map((group) => {
    const variableIndexes = group.variables.map((v) => String(v.index));
    const lines = [];
    lines.push(`Suggested variables ${variableIndexes.join(", ")}:`);
    group.variables.forEach((v) => {
      lines.push(`- ${v.text}`);
    });
    lines.push("");
    lines.push(
      `Matched CRF Question${typeof group.score === "number" ? ` (score=${group.score.toFixed(3)})` : ""}:`
    );
    lines.push(`[Original]: ${group.originalQuestionText}`);
    lines.push(`[Rewritten]: ${group.rewrittenQuestionText}`);
    lines.push("");
    lines.push("Answers/Options:");
    if (!Array.isArray(group.answers) || group.answers.length === 0) {
      lines.push("- (No answers/options found)");
    } else {
      group.answers.forEach((a) => {
        const t = String(a?.text || "").trim();
        if (!t) return;
        lines.push(`- ${t}`);
      });
    }
    return lines.join("\n");
  });

  return blocks.join("\n\n------------------------------\n\n");
}

/**
 * Build final Notes text from grouped CRF mappings.
 *
 * IMPORTANT:
 * - This intentionally follows the same grouped source used by the evidence box.
 * - The Notes box keeps a simplified format: question + options only.
 */
function formatGroupedMappingsForNotes(groupedMappings) {
  const lines = [];
  const groups = Array.isArray(groupedMappings) ? groupedMappings : [];
  if (groups.length === 0) {
    lines.push("- (No mapped CRF question found)");
  } else {
    groups.forEach((group) => {
      lines.push(String(group.rewrittenQuestionText || "").trim() || "(No matched question)");
      const answerTexts = (Array.isArray(group.answers) ? group.answers : [])
        .map((a) => String(a?.text || "").trim())
        .filter(Boolean);
      if (answerTexts.length > 0) {
        answerTexts.forEach((t) => lines.push(`- ${t}`));
  }
  lines.push("");
    });
  }
  return lines.join("\n");
}

// Removed parseDrugComparisonToColumns

async function handleMockTflWorkItemChange({ studyNumber, value }) {
  const parsed = parseMockTflWorkItemValue(value);
  if (!parsed) return { ok: false };
  await selectMockTflBody({ studyNumber, type: parsed.type, number: parsed.number });
  return { ok: true, ...parsed };
}

async function openMockTflEditorDialogForWorkItem({ studyNumber, type, number, label }) {
  if (!Office?.context?.ui?.displayDialogAsync) {
    throw new Error("Dialog API is not available in this host.");
  }
  closeMockTflEditorDialog();
  const dialogUrl = `${window.location.origin}/mockTflEditorDialog.html`;

  return await new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 70, width: 40 },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(asyncResult.error || new Error("Failed to open MockTFL editor dialog."));
          return;
        }

        mockTflEditorDialog = asyncResult.value;
        let currentAnalysisContext = {
          pureTitle: "",
          savedCrfFormName: ""
        };
        let currentFigureContext = {
          pureTitle: "",
          drug1: "",
          drug2: ""
        };
        let currentSavedSection = null;

        mockTflEditorDialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
          const msg = safeJsonParse(arg?.message);
          const msgType = String(msg?.type || "");
          if (!msgType.startsWith("mocktfl:")) return;

          try {
            if (msgType === "mocktfl:dialogReady") {
              const titleCcText = await readMockTflTitle({ studyNumber, type, number });
              const pureTitle = extractPureTitleFromMockTflTitleCcText(titleCcText);
              currentAnalysisContext = { pureTitle, savedCrfFormName: "" };
              const stateResp = await loadMockTflState({ studyNumber });
              const state = stateResp?.success ? stateResp.data || {} : {};
              const savedDrug1 = String(state?.drug1 || "").trim();
              const savedDrug2 = String(state?.drug2 || "").trim();
              currentFigureContext = {
                pureTitle,
                drug1: savedDrug1,
                drug2: savedDrug2
              };
              const savedSectionData = findMockTflSectionData(state, { type, number });
              currentSavedSection = savedSectionData || null;
              
              const availableColumns = [
                String(state?.drug1 || "").trim(),
                String(state?.drug2 || "").trim(),
                "Overall"
              ].filter(Boolean);

              if (savedSectionData?.saved) {
                // Rehydrate the dialog with the last saved header config for this section.
                sendToDialog({
                  type: "mocktfl:loadSavedSectionData",
                  studyNumber,
                  workItem: {
                    type,
                    number,
                    label: label || `${type} ${number}`.trim(),
                    enableAnalysis: isAnalysisTarget({ type, number })
                  },
                  savedSectionData,
                  availableColumns,
                  figureContext: currentFigureContext
                });
                return;
              }
              sendToDialog({
                type: "mocktfl:init",
                studyNumber,
                workItem: {
                  type,
                  number,
                  label: label || `${type} ${number}`.trim(),
                  // Let the dialog decide whether to show "Start analysis"
                  enableAnalysis: isAnalysisTarget({ type, number })
                },
                availableColumns,
                savedSectionData: savedSectionData || null,
                figureContext: currentFigureContext,
                // Rehydrate draft-only config such as orientation without forcing the "saved section" workflow.
                savedColumnHeaderConfig: savedSectionData?.columnHeaderConfig || null
              });
              return;
            }

            if (msgType === "mocktfl:updateColumnOrientation") {
              const columnHeaderConfig = msg?.columnHeaderConfig || null;
              currentSavedSection = {
                ...(currentSavedSection || {
                  key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                  type,
                  number,
                  saved: false,
                  matchedSapLabel: "",
                  matchedSapSectionText: "",
                  suggestedVariablesText: "",
                  mappedCrfText: "",
                  notesText: "",
                  rCodeText: "",
                  figureBase64: "",
                  figureMockDataText: ""
                }),
                columnHeaderConfig
              };
              await saveMockTflSectionDraft({
                studyNumber,
                snapshot: {
                  ...(currentSavedSection || {}),
                  type,
                  number,
                  label: label || `${type} ${number}`.trim(),
                  columnHeaderConfig
                }
              });
              return;
            }

            if (msgType === "mocktfl:findCorrespondingTable") {
              if (!isFigureTarget({ type, number })) {
                sendToDialog({
                  type: "mocktfl:error",
                  message: "Find corresponding table is only enabled for FIGURE 14.2.1.1 in the current prototype."
                });
                return;
              }

              sendToDialog({ type: "mocktfl:findingCorrespondingTable" });
              const resp = await findCorrespondingTable({
                studyNumber,
                type,
                number,
                pureTitle: currentFigureContext.pureTitle
              });
              const matchedTable = resp?.matchedTable || null;
              if (!matchedTable) {
                throw new Error("Backend did not return a corresponding table.");
              }

              currentSavedSection = {
                ...(currentSavedSection || {
                  key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                  type,
                  number,
                  saved: false,
                  matchedSapLabel: "",
                  matchedSapSectionText: "",
                  suggestedVariablesText: "",
                  mappedCrfText: "",
                  notesText: "",
                  columnHeaderConfig: null
                }),
                saved: false,
                correspondingTable: matchedTable
              };

              await saveMockTflSectionDraft({
                studyNumber,
                snapshot: {
                  ...currentSavedSection,
                  type,
                  number,
                  label: label || `${type} ${number}`.trim()
                }
              });

              sendToDialog({
                type: "mocktfl:correspondingTableFound",
                matchedTable,
                candidates: Array.isArray(resp?.candidates) ? resp.candidates : []
              });
              return;
            }

            if (msgType === "mocktfl:defineFigureType") {
              if (!isFigureTarget({ type, number })) {
                sendToDialog({
                  type: "mocktfl:error",
                  message: "Defining figure type is only enabled for FIGURE 14.2.1.1 in the current prototype."
                });
                return;
              }

              sendToDialog({ type: "mocktfl:definingFigureType" });
              const resp = await defineMockTflFigureType({
                studyNumber,
                type,
                number,
                pureTitle: currentFigureContext.pureTitle,
                correspondingTableType: currentSavedSection?.correspondingTable?.type || "",
                correspondingTableNumber: currentSavedSection?.correspondingTable?.number || ""
              });
              const figureType = String(resp?.figureType || "").trim();

              currentSavedSection = {
                ...(currentSavedSection || {
                  key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                  type,
                  number,
                  saved: false,
                  matchedSapLabel: "",
                  matchedSapSectionText: "",
                  suggestedVariablesText: "",
                  mappedCrfText: "",
                  notesText: "",
                  columnHeaderConfig: null
                }),
                saved: false,
                figureType
              };

              await saveMockTflSectionDraft({
                studyNumber,
                snapshot: {
                  ...currentSavedSection,
                  type,
                  number,
                  label: label || `${type} ${number}`.trim()
                }
              });

              sendToDialog({ type: "mocktfl:figureTypeDefined", figureType });
              return;
            }

            if (msgType === "mocktfl:autoGenerateFigure") {
              if (!isFigureTarget({ type, number })) {
                sendToDialog({
                  type: "mocktfl:error",
                  message: "Auto figure generation is only enabled for FIGURE 14.2.1.1 in the current prototype."
                });
                return;
              }

              const confirmedFigureType = String(msg?.figureType || currentSavedSection?.figureType || "").trim();
              sendToDialog({ type: "mocktfl:autoFigureGenerating" });
              const resp = await autoGenerateMockTflFigure({
                studyNumber,
                type,
                number,
                pureTitle: currentFigureContext.pureTitle,
                drug1: currentFigureContext.drug1,
                drug2: currentFigureContext.drug2,
                figureType: confirmedFigureType,
                correspondingTableType: currentSavedSection?.correspondingTable?.type || "",
                correspondingTableNumber: currentSavedSection?.correspondingTable?.number || ""
              });
              const figureBase64 = String(resp?.imageBase64 || "").trim();
              if (!figureBase64) {
                throw new Error("Backend did not return a generated figure image.");
              }
              const generatedRCode = String(resp?.generatedRCode || "");
              const figureMockDataText = JSON.stringify(
                {
                  figureKind: String(resp?.figureKind || ""),
                  chartTitle: String(resp?.chartTitle || ""),
                  mockData: resp?.mockData || null
                },
                null,
                2
              );

              currentSavedSection = {
                ...(currentSavedSection || {
                  key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                  type,
                  number,
                  saved: false,
                  matchedSapLabel: "",
                  matchedSapSectionText: "",
                  suggestedVariablesText: "",
                  mappedCrfText: "",
                  notesText: "",
                  columnHeaderConfig: null
                }),
                saved: false,
                figureType: confirmedFigureType,
                rCodeText: generatedRCode,
                figureBase64,
                figureMockDataText
              };

              await saveMockTflSectionDraft({
                studyNumber,
                snapshot: {
                  ...currentSavedSection,
                  type,
                  number,
                  label: label || `${type} ${number}`.trim()
                }
              });

              sendToDialog({
                type: "mocktfl:autoFigureGenerated",
                imageBase64: figureBase64,
                generatedRCode,
                figureMockDataText,
                tableInfo: resp?.tableInfo || { fetched: false, label: "", usedColumnGroups: [] }
              });
              return;
            }

            if (msgType === "mocktfl:generateFigure") {
              if (!isFigureTarget({ type, number })) {
                sendToDialog({
                  type: "mocktfl:error",
                  message: "Figure generation is only enabled for FIGURE 14.2.1.1 in the current prototype."
                });
                return;
              }

              const rCodeText = String(msg?.rCodeText || "");
              if (!rCodeText.trim()) {
                throw new Error("Please paste R code before generating the figure.");
              }

              sendToDialog({ type: "mocktfl:figureGenerating" });
              const resp = await generateMockTflFigure({
                studyNumber,
                type,
                number,
                rCodeText
              });
              const figureBase64 = String(resp?.imageBase64 || "").trim();
              if (!figureBase64) {
                throw new Error("Backend did not return a generated figure image.");
              }

              currentSavedSection = {
                ...(currentSavedSection || {
                  key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                  type,
                  number,
                  saved: false,
                  matchedSapLabel: "",
                  matchedSapSectionText: "",
                  suggestedVariablesText: "",
                  mappedCrfText: "",
                  notesText: "",
                  columnHeaderConfig: null,
                  figureMockDataText: ""
                }),
                saved: false,
                rCodeText,
                figureBase64,
                figureMockDataText: ""
              };

              // Save a draft immediately so closing the dialog does not lose a generated preview.
              await saveMockTflSectionDraft({
                studyNumber,
                snapshot: {
                  ...currentSavedSection,
                  type,
                  number,
                  label: label || `${type} ${number}`.trim()
                }
              });

              sendToDialog({
                type: "mocktfl:figureGenerated",
                imageBase64: figureBase64
              });
              return;
            }

            if (msgType === "mocktfl:startAnalysis") {
              // Safety gate: only allow this feature for allowlisted prototype items.
              if (!isAnalysisTarget({ type, number })) {
                sendToDialog({
                  type: "mocktfl:error",
                  message: "Start analysis is only enabled for TABLE 14.1.1.2, TABLE 14.1.3.1 and TABLE 14.2.1.1 in the current prototype."
                });
                return;
              }

              sendToDialog({ type: "mocktfl:analysisRunning" });

              // 1) Read title from Word Title Content Control (source of truth)
              const titleCcText = await readMockTflTitle({ studyNumber, type, number });
              const pureTitle = extractPureTitleFromMockTflTitleCcText(titleCcText);
              if (!pureTitle) {
                throw new Error("Failed to extract pure title from Word title Content Control.");
              }

              // 2) Step 1: run SAP matching + GPT variable suggestion only.
              const resp = await startMockTflAnalysis({
                studyNumber,
                type,
                number,
                pureTitle
              });

              // Keep only the minimal context required for the confirmation step.
              currentAnalysisContext = {
                pureTitle,
                savedCrfFormName: String(resp?.context?.savedCrfFormName || "").trim()
              };

              const sapMatchedLabel = String(resp?.sap?.matched?.label || "").trim();
              const sapSectionText = String(resp?.sap?.sectionText || "").trim();
              const sapStatisticalVariablesText = String(resp?.sap?.statisticalVariablesText || "").trim();
              currentSavedSection = {
                key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                type,
                number,
                saved: false,
                matchedSapLabel: sapMatchedLabel,
                matchedSapSectionText: sapSectionText,
                suggestedVariablesText: sapStatisticalVariablesText,
                mappedCrfText: "",
                notesText: "",
                rCodeText: currentSavedSection?.rCodeText || "",
                figureBase64: currentSavedSection?.figureBase64 || "",
                columnHeaderConfig: currentSavedSection?.columnHeaderConfig || null
              };

              sendToDialog({
                type: "mocktfl:analysisStep1Result",
                sapMatchedLabel,
                sapSectionText,
                sapStatisticalVariablesText
              });
              return;
            }

            if (msgType === "mocktfl:confirmVariables") {
              if (!isAnalysisTarget({ type, number })) {
                sendToDialog({
                  type: "mocktfl:error",
                  message: "Start analysis is only enabled for TABLE 14.1.1.2, TABLE 14.1.3.1 and TABLE 14.2.1.1 in the current prototype."
                });
                return;
              }

              if (!currentAnalysisContext.pureTitle) {
                throw new Error("Please click 'Start analysis' before confirming variables.");
              }

              sendToDialog({ type: "mocktfl:analysisStep2Running" });

              const resp = await mapMockTflVariablesToCrf({
                studyNumber,
                type,
                number,
                pureTitle: currentAnalysisContext.pureTitle,
                variablesText: String(msg?.variablesText || ""),
                savedCrfFormName: currentAnalysisContext.savedCrfFormName
              });

              const crfMappings = Array.isArray(resp?.crf?.mappings) ? resp.crf.mappings : [];
              const groupedMappings = groupCrfMappingsByQuestion(crfMappings);
              const mappedEvidenceText = formatGroupedMappingsForEvidence(groupedMappings);
              const finalNotesText = formatGroupedMappingsForNotes(groupedMappings);
              currentSavedSection = {
                ...(currentSavedSection || {
                  key: `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`,
                  type,
                  number,
                  saved: false
                }),
                saved: false,
                matchedSapLabel: String(currentSavedSection?.matchedSapLabel || "").trim(),
                matchedSapSectionText: String(currentSavedSection?.matchedSapSectionText || ""),
                suggestedVariablesText: String(msg?.variablesText || ""),
                mappedCrfText: mappedEvidenceText,
                notesText: finalNotesText,
                columnHeaderConfig: currentSavedSection?.columnHeaderConfig || null
              };

              sendToDialog({
                type: "mocktfl:analysisStep2Result",
                finalNotesText,
                mappedEvidenceText
              });
              return;
            }

            if (msgType === "mocktfl:insert") {
              if (isFigureTarget({ type, number })) {
                const rCodeText = String(msg?.rCodeText || currentSavedSection?.rCodeText || "");
                const figureBase64 = String(msg?.figureBase64 || currentSavedSection?.figureBase64 || "").trim();
                const figureMockDataText = String(
                  msg?.figureMockDataText || currentSavedSection?.figureMockDataText || ""
                );
                if (!figureBase64) {
                  throw new Error("No generated figure is available. Please click Generate Figure first.");
                }

                await insertMockTflFigure({
                  studyNumber,
                  type,
                  number,
                  base64Image: figureBase64
                });

                await saveMockTflSectionSnapshot({
                  studyNumber,
                  snapshot: {
                    ...(currentSavedSection || {}),
                    type,
                    number,
                    label: label || `${type} ${number}`.trim(),
                    rCodeText,
                    figureBase64,
                    figureMockDataText,
                    notesText: String(currentSavedSection?.notesText || ""),
                    columnHeaderConfig: currentSavedSection?.columnHeaderConfig || null
                  }
                });
                sendToDialog({ type: "mocktfl:inserted" });
                return;
              }

              // Keep existing text-insert behavior for all work items EXCEPT the special table target.
              // For table-insert targets, we insert a structured Word table where:
              // - column 1 = parsed editor text lines
              // - option lines are indented
              // - blank lines are preserved as blank table rows
              if (isTableInsertTarget({ type, number })) {
                const editorText = String(msg?.text || "");
                const selectedColumns = Array.isArray(msg?.selectedColumns) ? msg.selectedColumns : [];
                const columnHeaderConfig = msg?.columnHeaderConfig || null;
                const columnOrientation = String(columnHeaderConfig?.orientation || "vertical").trim().toLowerCase();
                
                if (!editorText.trim()) {
                  throw new Error("No content found in Notes. Please confirm variables before Insert.");
                }
                if (columnOrientation === "horizontal") {
                  // Save the user's choice first, but block insertion until horizontal rendering exists.
                  await saveMockTflSectionSnapshot({
                    studyNumber,
                    snapshot: {
                      ...(currentSavedSection || {}),
                      type,
                      number,
                      label: label || `${type} ${number}`.trim(),
                      notesText: editorText,
                      columnHeaderConfig
                    }
                  });
                  sendToDialog({
                    type: "mocktfl:error",
                    message: "Horizontal column structure is saved to the database, but insertion is not implemented yet."
                  });
                  return;
                }
                // Structured table targets now receive the full multi-level header config.
                // Route to the correct table skeleton inserter based on table number.
                if (String(number).trim() === "14.1.1.2") {
                  await insertTable14_1_1_2_PatientDisposition({
                    studyNumber,
                    text: editorText,
                    selectedColumns,
                    columnHeaderConfig
                  });
                } else if (String(number).trim() === "14.1.3.1") {
                  await insertTable14_1_3_1_SummaryOfDemographics({
                    studyNumber,
                    text: editorText,
                    selectedColumns,
                    columnHeaderConfig
                  });
                } else if (String(number).trim() === "14.2.1.1") {
                  await insertTable14_2_1_1_PrimaryEfficacyEndpoint({
                    studyNumber,
                    text: editorText,
                    selectedColumns,
                    columnHeaderConfig
                  });
                } else {
                  // This should never happen because of the allowlist, but keep a safety net.
                  throw new Error(`Unsupported table insert target: ${type} ${number}`);
                }
              }
              await saveMockTflSectionSnapshot({
                studyNumber,
                snapshot: {
                  ...(currentSavedSection || {}),
                  type,
                  number,
                  label: label || `${type} ${number}`.trim(),
                  notesText: String(msg?.text || ""),
                  columnHeaderConfig: msg?.columnHeaderConfig || currentSavedSection?.columnHeaderConfig || null
                }
              });
              sendToDialog({ type: "mocktfl:inserted" });
              return;
            }

            if (msgType === "mocktfl:clear") {
              currentAnalysisContext = { pureTitle: "", savedCrfFormName: "" };
              currentSavedSection = null;
              sendToDialog({ type: "mocktfl:cleared" });
              return;
            }

            if (msgType === "mocktfl:close") {
              closeMockTflEditorDialog();
              return;
            }
          } catch (e) {
            sendToDialog({ type: "mocktfl:error", message: e?.message || String(e) });
          }
        });

        mockTflEditorDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          // user closed
          closeMockTflEditorDialog();
        });

        resolve(true);
      }
    );
  });
}

export {
  parseMockTflWorkItemValue,
  buildMockTflLabel,
  handleMockTflWorkItemChange,
  openMockTflEditorDialogForWorkItem,
  closeMockTflEditorDialog
};

