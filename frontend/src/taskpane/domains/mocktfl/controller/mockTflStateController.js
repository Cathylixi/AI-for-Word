import { loadMockTflState, saveMockTflState } from "../api/state";
import { getDom } from "../../../shared/ui/dom";
import { setStatus } from "../../../shared/ui/status";

function buildMockTflKey({ type, number }) {
  return `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`;
}

function extractTitleFromLabel({ type, number, label }) {
  const rawLabel = String(label || "").trim();
  const prefix = `${String(type || "").trim().toUpperCase()} ${String(number || "").trim()}`.trim();
  if (!rawLabel) return "";
  if (!prefix) return rawLabel;
  if (rawLabel.toUpperCase().startsWith(prefix.toUpperCase())) {
    return rawLabel.slice(prefix.length).trim();
  }
  return rawLabel;
}

function normalizeEntries(entries) {
  const arr = Array.isArray(entries) ? entries : [];
  const seen = new Set();
  const out = [];
  arr.forEach((entry) => {
    const type = String(entry?.type || "").trim().toUpperCase();
    const number = String(entry?.number || "").trim();
    if (!type || !number) return;
    const key = buildMockTflKey({ type, number });
    if (seen.has(key)) return;
    seen.add(key);
    out.push({
      type,
      number,
      title: extractTitleFromLabel({ type, number, label: entry?.label || entry?.title || "" }) || String(entry?.title || "").trim(),
      label: String(entry?.label || `${type} ${number}`).trim()
    });
  });
  return out;
}

function uniqueNonEmptyStrings(values) {
  const arr = Array.isArray(values) ? values : [];
  const seen = new Set();
  const out = [];
  arr.forEach((value) => {
    const text = String(value || "").trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    out.push(text);
  });
  return out;
}

function normalizeOptionalString(value) {
  return String(value || "").trim();
}

function normalizeColumnOrientation(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "horizontal") return "horizontal";
  if (normalized === "vertical") return "vertical";
  return "";
}

function normalizeColumnHeaderConfig(columnHeaderConfig) {
  // Keep a single persisted shape for every section snapshot.
  const orientation = normalizeColumnOrientation(columnHeaderConfig?.orientation);
  const columnGroups = uniqueNonEmptyStrings(columnHeaderConfig?.columnGroups);
  const rawHeaderLayers = Array.isArray(columnHeaderConfig?.headerLayers) ? columnHeaderConfig.headerLayers : [];
  // Persist header layers in a predictable shape so reopen/restore stays stable.
  const headerLayers = rawHeaderLayers
    .map((headerLayer, idx) => ({
      level: Number.isFinite(Number(headerLayer?.level)) ? Number(headerLayer.level) : idx + 1,
      headers: uniqueNonEmptyStrings(headerLayer?.headers)
    }))
    .filter((headerLayer) => headerLayer.headers.length > 0)
    .sort((a, b) => a.level - b.level);
  const columnGroupSubtitle = normalizeOptionalString(columnHeaderConfig?.columnGroupSubtitle);

  return {
    orientation,
    columnGroups,
    columnGroupSubtitle,
    headerLayers
  };
}

function normalizeCorrespondingTable(correspondingTable) {
  if (!correspondingTable || typeof correspondingTable !== "object") return null;
  const type = String(correspondingTable?.type || "").trim().toUpperCase();
  const number = String(correspondingTable?.number || "").trim();
  const title = String(correspondingTable?.title || "").trim();
  const label = String(correspondingTable?.label || `${type} ${number} ${title}`).trim();
  if (!type && !number && !title && !label) return null;
  const rawScore = Number(correspondingTable?.score);
  return {
    type,
    number,
    title,
    label,
    score: Number.isFinite(rawScore) ? rawScore : null,
    matchedBy: String(correspondingTable?.matchedBy || "").trim(),
    embeddingModel: String(correspondingTable?.embeddingModel || "").trim(),
    matchedAt: correspondingTable?.matchedAt || new Date().toISOString()
  };
}

function createEmptySectionData({ type, number }) {
  return {
    key: buildMockTflKey({ type, number }),
    type: String(type || "").trim().toUpperCase(),
    number: String(number || "").trim(),
    saved: false,
    matchedSapLabel: "",
    matchedSapSectionText: "",
    suggestedVariablesText: "",
    mappedCrfText: "",
    notesText: "",
    rCodeText: "",
    figureBase64: "",
    figureMockDataText: "",
    figureType: "",
    correspondingTable: null,
    columnHeaderConfig: {
      orientation: "",
      columnGroups: [],
      columnGroupSubtitle: "",
      headerLayers: []
    },
    updatedAt: new Date().toISOString()
  };
}

function normalizeSectionsData(sectionsData) {
  const arr = Array.isArray(sectionsData) ? sectionsData : [];
  const seen = new Set();
  const out = [];
  arr.forEach((sectionData) => {
    const type = String(sectionData?.type || "").trim().toUpperCase();
    const number = String(sectionData?.number || "").trim();
    const key = String(sectionData?.key || buildMockTflKey({ type, number })).trim();
    if (!type || !number || !key || seen.has(key)) return;
    seen.add(key);
    out.push({
      key,
      type,
      number,
      saved: !!sectionData?.saved,
      matchedSapLabel: String(sectionData?.matchedSapLabel || "").trim(),
      matchedSapSectionText: String(sectionData?.matchedSapSectionText || ""),
      suggestedVariablesText: String(sectionData?.suggestedVariablesText || ""),
      mappedCrfText: String(sectionData?.mappedCrfText || ""),
      notesText: String(sectionData?.notesText || ""),
      rCodeText: String(sectionData?.rCodeText || ""),
      figureBase64: String(sectionData?.figureBase64 || "").trim(),
      figureMockDataText: String(sectionData?.figureMockDataText || ""),
      figureType: String(sectionData?.figureType || ""),
      correspondingTable: normalizeCorrespondingTable(sectionData?.correspondingTable),
      columnHeaderConfig: normalizeColumnHeaderConfig(sectionData?.columnHeaderConfig),
      updatedAt: sectionData?.updatedAt || new Date().toISOString()
    });
  });
  return out;
}

function findMockTflSectionData(state, { type, number }) {
  const key = buildMockTflKey({ type, number });
  const sectionsData = Array.isArray(state?.sectionsData) ? state.sectionsData : [];
  const sectionData = sectionsData.find((item) => String(item?.key || "").trim() === key);
  if (!sectionData) return null;
  return {
    ...sectionData,
    columnHeaderConfig: normalizeColumnHeaderConfig(sectionData?.columnHeaderConfig)
  };
}

function upsertMockTflSectionData(sectionsData, sectionData) {
  const normalizedSectionsData = normalizeSectionsData(sectionsData);
  const type = String(sectionData?.type || "").trim().toUpperCase();
  const number = String(sectionData?.number || "").trim();
  const key = buildMockTflKey({ type, number });
  const next = normalizedSectionsData.filter((item) => item.key !== key);
  next.push({
    ...createEmptySectionData({ type, number }),
    ...sectionData,
    key,
    type,
    number,
    saved: !!sectionData?.saved,
    updatedAt: new Date().toISOString()
  });
  return next;
}

function ensureSectionsDataForEntries(entries, existingSectionsData) {
  const next = normalizeSectionsData(existingSectionsData);
  normalizeEntries(entries).forEach((entry) => {
    const existing = next.find((item) => item.key === buildMockTflKey(entry));
    if (existing) return;
    next.push(createEmptySectionData(entry));
  });
  return next;
}

function mergeSectionSnapshot(sectionsData, snapshot) {
  const normalizedSectionsData = normalizeSectionsData(sectionsData);
  const type = String(snapshot?.type || "").trim().toUpperCase();
  const number = String(snapshot?.number || "").trim();
  return upsertMockTflSectionData(normalizedSectionsData, {
    type,
    number,
    saved: !!snapshot?.saved,
    matchedSapLabel: String(snapshot?.matchedSapLabel || "").trim(),
    matchedSapSectionText: String(snapshot?.matchedSapSectionText || ""),
    suggestedVariablesText: String(snapshot?.suggestedVariablesText || ""),
    mappedCrfText: String(snapshot?.mappedCrfText || ""),
    notesText: String(snapshot?.notesText || ""),
    rCodeText: String(snapshot?.rCodeText || ""),
    figureBase64: String(snapshot?.figureBase64 || "").trim(),
    figureMockDataText: String(snapshot?.figureMockDataText || ""),
    figureType: String(snapshot?.figureType || ""),
    correspondingTable: normalizeCorrespondingTable(snapshot?.correspondingTable),
    columnHeaderConfig: normalizeColumnHeaderConfig(snapshot?.columnHeaderConfig)
  });
}

async function buildMockTflSaveState({ insertedEntries, existingSectionsData, drug1, drug2 }) {
  return {
    saved: true,
    lastSavedAt: new Date().toISOString(),
    entries: normalizeEntries(insertedEntries),
    sectionsData: normalizeSectionsData(existingSectionsData),
    schemaVersion: 5,
    drug1: String(drug1 || "").trim(),
    drug2: String(drug2 || "").trim()
  };
}

async function handleSaveMockTfl({ studyNumber, insertedEntries }) {
  const dom = getDom();
  if (!studyNumber) return;
  try {
    setStatus(dom.status, "Saving MockTFL to database...");
    const existingResp = await loadMockTflState({ studyNumber });
    const existingState = existingResp?.success ? existingResp.data || {} : {};
    const normalizedEntries = normalizeEntries(insertedEntries);
    const state = await buildMockTflSaveState({
      insertedEntries: normalizedEntries,
      existingSectionsData: ensureSectionsDataForEntries(normalizedEntries, existingState?.sectionsData || []),
      drug1: existingState?.drug1,
      drug2: existingState?.drug2
    });
    await saveMockTflState({ studyNumber, state });
    setStatus(dom.status, "Saved successfully.");
  } catch (e) {
    setStatus(dom.status, `Save failed: ${e.message || e}`);
  }
}

async function persistMockTflSectionState({ studyNumber, insertedEntries, snapshot, saved = true }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  const existingResp = await loadMockTflState({ studyNumber });
  const existingState = existingResp?.success ? existingResp.data || {} : {};
  const nextEntries = normalizeEntries([
    ...(Array.isArray(existingState?.entries) ? existingState.entries : []),
    ...(Array.isArray(insertedEntries) ? insertedEntries : []),
    {
      type: snapshot?.type,
      number: snapshot?.number,
      label: snapshot?.label
    }
  ]);
  const nextSectionsData = mergeSectionSnapshot(existingState?.sectionsData || [], {
    ...snapshot,
    saved
  });
  const state = await buildMockTflSaveState({
    insertedEntries: nextEntries,
    existingSectionsData: ensureSectionsDataForEntries(nextEntries, nextSectionsData),
    drug1: existingState?.drug1,
    drug2: existingState?.drug2
  });
  await saveMockTflState({ studyNumber, state });
}

async function saveMockTflSectionSnapshot({ studyNumber, insertedEntries, snapshot }) {
  await persistMockTflSectionState({
    studyNumber,
    insertedEntries,
    snapshot,
    saved: true
  });
}

async function saveMockTflSectionDraft({ studyNumber, insertedEntries, snapshot }) {
  await persistMockTflSectionState({
    studyNumber,
    insertedEntries,
    snapshot,
    saved: !!snapshot?.saved
  });
}

async function saveMockTflDrugComparison({ studyNumber, drug1, drug2 }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  const existingResp = await loadMockTflState({ studyNumber });
  const existingState = existingResp?.success ? existingResp.data || {} : {};
  
  const state = await buildMockTflSaveState({
    insertedEntries: Array.isArray(existingState?.entries) ? existingState.entries : [],
    existingSectionsData: Array.isArray(existingState?.sectionsData) ? existingState.sectionsData : [],
    drug1: String(drug1 || "").trim(),
    drug2: String(drug2 || "").trim()
  });
  await saveMockTflState({ studyNumber, state });
}

export {
  handleSaveMockTfl,
  saveMockTflSectionSnapshot,
  saveMockTflSectionDraft,
  findMockTflSectionData,
  saveMockTflDrugComparison
};
