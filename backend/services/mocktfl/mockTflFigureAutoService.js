const OpenAI = require("openai");

const { generateMockTflFigureFromRCode } = require("./mockTflFigureService");
const {
  KNOWN_FIGURE_FAMILIES,
  buildFigureRCode,
  getMappingSchemaHint
} = require("./figureRTemplates");

const SUPPORTED_FIGURE_KEYS = new Set(["FIGURE:14.2.1.1"]);

function normalizeFigureKey(type, number) {
  return `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`;
}

/**
 * Treatment-group column names that are summary/aggregate columns and should NOT
 * be used as a plotted curve group (e.g., a cumulative incidence curve per arm).
 */
const SUMMARY_COLUMN_GROUPS = new Set(["overall", "total", "all", "combined"]);

function extractTreatmentGroups(columnGroups) {
  const arr = Array.isArray(columnGroups) ? columnGroups : [];
  return arr
    .map((g) => String(g || "").trim())
    .filter(Boolean)
    .filter((g) => !SUMMARY_COLUMN_GROUPS.has(g.toLowerCase()));
}

/**
 * Load the FULL saved MockTFL section data of the corresponding TABLE for a figure.
 *
 * Data source:
 * - llxdocument.studies { studyNumber, chunkType: "MockTFL" }
 * - data.sectionsData[] keyed by `${type}:${number}` (e.g., "TABLE:14.2.1.1")
 *
 * Returns { found:false } when the study state or the table section is missing,
 * so figure generation can gracefully fall back to drug1/drug2 placeholders.
 */
async function loadCorrespondingTableInfo({ studyNumber, tableType, tableNumber, mongoose }) {
  const type = String(tableType || "").trim().toUpperCase();
  const number = String(tableNumber || "").trim();
  if (!studyNumber || !type || !number) return { found: false };
  if (!mongoose?.connection?.client) return { found: false };

  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");
  const stateDoc = await studiesCol.findOne(
    { studyNumber: String(studyNumber).trim(), chunkType: "MockTFL" },
    { projection: { data: 1, _id: 0 } }
  );

  const sectionsData = Array.isArray(stateDoc?.data?.sectionsData) ? stateDoc.data.sectionsData : [];
  const key = `${type}:${number}`;
  const section =
    sectionsData.find((s) => String(s?.key || "").trim() === key) ||
    sectionsData.find(
      (s) =>
        String(s?.type || "").trim().toUpperCase() === type &&
        String(s?.number || "").trim() === number
    ) ||
    null;

  if (!section) return { found: false };

  const columnHeaderConfig =
    section.columnHeaderConfig && typeof section.columnHeaderConfig === "object"
      ? section.columnHeaderConfig
      : {};
  const columnGroups = Array.isArray(columnHeaderConfig.columnGroups)
    ? columnHeaderConfig.columnGroups.map((g) => String(g || "").trim()).filter(Boolean)
    : [];

  return {
    found: true,
    key,
    type,
    number,
    label: [type, number, String(section?.title || "").trim()].filter(Boolean).join(" ").trim(),
    columnHeaderConfig: {
      orientation: String(columnHeaderConfig.orientation || "").trim(),
      columnGroups,
      columnGroupSubtitle: String(columnHeaderConfig.columnGroupSubtitle || "").trim(),
      headerLayers: Array.isArray(columnHeaderConfig.headerLayers) ? columnHeaderConfig.headerLayers : []
    },
    mappedCrfText: String(section.mappedCrfText || ""),
    suggestedVariablesText: String(section.suggestedVariablesText || ""),
    matchedSapLabel: String(section.matchedSapLabel || ""),
    matchedSapSectionText: String(section.matchedSapSectionText || ""),
    notesText: String(section.notesText || "")
  };
}

function escapeRString(value) {
  return String(value == null ? "" : value)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"');
}

function parseJsonObjectFromText(text) {
  const raw = String(text || "").trim();
  if (!raw) return null;
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? String(fenced[1] || "").trim() : raw;
  try {
    return JSON.parse(candidate);
  } catch (e) {
    return null;
  }
}

function buildTableContextForPrompt(tableInfo) {
  if (!tableInfo || !tableInfo.found) return null;
  const cfg = tableInfo.columnHeaderConfig || {};
  const lines = [
    "Corresponding Table Context (use this to make the figure consistent with its table):",
    `- Table: ${tableInfo.label || ""}`.trim(),
    `- Column groups (treatment arms): ${(cfg.columnGroups || []).join(", ") || "(none)"}`,
    cfg.columnGroupSubtitle ? `- Column group subtitle: ${cfg.columnGroupSubtitle}` : "",
    cfg.orientation ? `- Column header orientation: ${cfg.orientation}` : "",
    tableInfo.mappedCrfText ? `- Table row headers / mapped CRF:\n${tableInfo.mappedCrfText}` : "",
    tableInfo.suggestedVariablesText ? `- Suggested statistical variables:\n${tableInfo.suggestedVariablesText}` : "",
    tableInfo.matchedSapLabel ? `- Matched SAP section: ${tableInfo.matchedSapLabel}` : "",
    tableInfo.matchedSapSectionText ? `- SAP section text:\n${tableInfo.matchedSapSectionText}` : "",
    tableInfo.notesText ? `- Notes:\n${tableInfo.notesText}` : ""
  ].filter(Boolean);
  return lines.join("\n");
}

/**
 * Step 1 of mock-data generation:
 * Ask GPT to list ALL variables that could plausibly appear for this figure,
 * given the figure title, the corresponding table info, and the figure type.
 * Returns a de-duplicated array of variable name strings (may be empty).
 */
async function listAllFigureVariablesByAi({ pureTitle, tableInfo, figureType, client, model }) {
  const tableContext = buildTableContextForPrompt(tableInfo);
  const figureTypeText = String(figureType || "").trim();
  const prompt = [
    "You are a senior clinical biostatistician.",
    "Task: list ALL variables that could plausibly appear in the data underlying the figure described below.",
    "Be exhaustive: include time/x-axis variables, the analysis/event variables, censoring variables,",
    "grouping/treatment variables, at-risk counts, derived summary measures, and any others relevant to the figure type.",
    "",
    "Return ONLY valid JSON of the form:",
    '{ "variables": ["variable 1", "variable 2", "..."] }',
    "",
    `Figure title: ${String(pureTitle || "").trim()}`,
    figureTypeText ? `Figure type: ${figureTypeText}` : "",
    "",
    tableContext || "Corresponding table information: (not available)"
  ]
    .filter(Boolean)
    .join("\n");

  try {
    const completion = await client.chat.completions.create({
      model,
      messages: [
        { role: "system", content: "You output strict JSON listing statistical variables." },
        { role: "user", content: prompt }
      ],
      temperature: 0.1,
      response_format: { type: "json_object" }
    });
    const parsed = parseJsonObjectFromText(completion?.choices?.[0]?.message?.content || "");
    const raw = Array.isArray(parsed?.variables) ? parsed.variables : [];
    const seen = new Set();
    const out = [];
    for (const v of raw) {
      const name = String(v || "").trim();
      if (!name || seen.has(name.toLowerCase())) continue;
      seen.add(name.toLowerCase());
      out.push(name);
    }
    return out.slice(0, 40);
  } catch (e) {
    return [];
  }
}

/**
 * Validate that the AI output uses a single FLAT `records` array.
 *
 * Throws a clear error if:
 * - `records` is missing or not a non-empty array
 * - any record is not a plain object
 * - any record value is a nested object or array
 *
 * Returns the validated records array.
 */
function validateFlatRecordsJson(parsed) {
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("AI mock data must be a JSON object containing a `records` array.");
  }
  const records = parsed.records;
  if (!Array.isArray(records) || records.length < 2) {
    throw new Error("AI mock data must contain a `records` array with at least 2 rows.");
  }
  records.forEach((row, idx) => {
    if (!row || typeof row !== "object" || Array.isArray(row)) {
      throw new Error(`AI mock data record #${idx + 1} must be a flat object.`);
    }
    for (const [key, value] of Object.entries(row)) {
      if (value !== null && typeof value === "object") {
        throw new Error(
          `AI mock data must use flat records JSON. Record #${idx + 1} field "${key}" is a nested object/array, which is not allowed.`
        );
      }
    }
  });
  return records;
}

/**
 * Step 1+2 of the pipeline: let GPT decide which variables the figure needs and
 * then design a FLAT-records JSON mock dataset for it.
 *
 * Returns:
 *   { chartTitle, variables: string[], dataJson: { variables, records, notes } }
 */
async function createSpecByAi({ pureTitle, drug1, drug2, tableInfo, figureType, client, model }) {
  const treatmentGroups = extractTreatmentGroups(tableInfo?.columnHeaderConfig?.columnGroups);
  const figureTypeText = String(figureType || "").trim();
  const chartTitle = String(pureTitle || "").trim() || "Generated Figure";

  if (!client) {
    throw new Error("Missing OpenAI API key for AI-driven figure data generation.");
  }

  // STEP 1: let GPT decide which variables this figure should contain.
  const variables = await listAllFigureVariablesByAi({ pureTitle, tableInfo, figureType, client, model });
  const variablesBlock = variables.length
    ? `Variables you have identified as relevant (use them to inform the data you design):\n${variables.map((v, i) => `${i + 1}. ${v}`).join("\n")}`
    : "";

  const tableContext = buildTableContextForPrompt(tableInfo);
  const groupRule =
    treatmentGroups.length >= 2
      ? `If your data includes a treatment-group / treatment-arm variable, its values MUST be exactly these names (in order): ${treatmentGroups.map((g) => `"${g}"`).join(", ")}.`
      : `If your data includes a treatment-group / treatment-arm variable, use these names: "${String(drug1 || "").trim() || "Drug 1"}", "${String(drug2 || "").trim() || "Drug 2"}".`;

  // STEP 2: ask GPT to design the mock data, but constrained to FLAT records so
  // the downstream R code can parse it reliably (no nested objects/arrays).
  const prompt = [
    "You are a senior clinical biostatistician.",
    "A figure needs to be drawn. Based on the figure title, figure type and the corresponding table information below:",
    "1. Decide which variables (columns) this figure needs. You choose the variables yourself; there is no fixed set.",
    "2. Generate realistic synthetic (fake/demo) mock data for those variables.",
    "",
    "Return STRICT JSON only (a single JSON object) with EXACTLY this shape:",
    "{",
    '  "statistical_reasoning": "First, conduct a diagnostic reasoning by answering these 2 questions: 1. Raw vs Summary Judgment: Should I generate raw subject-level records for R to calculate the probabilities, or direct X/Y summary points? 2. Baseline/Origin Judgment: Does this statistical figure have a theoretical starting point over time? (If generating raw data, simply ensure the time variable is relative to an origin).",',
    '  "variables": ["variable1", "variable2", "variable3"],',
    '  "records": [',
    '    { "variable1": <value>, "variable2": <value>, "variable3": <value> },',
    '    { "variable1": <value>, "variable2": <value>, "variable3": <value> }',
    "  ],",
    '  "notes": ["..."]',
    "}",
    "",
    "The names variable1/variable2/variable3 above are ONLY placeholders that show the structure.",
    "Replace them with the real, meaningful column names you decided in step 1.",
    "",
    "Hard rules (MUST follow, or the figure cannot be drawn):",
    "- The JSON MUST contain a single data array named `records`.",
    "- `records` MUST be tidy data: one observation per row.",
    "- Every record MUST be a FLAT object. Each value is ONLY a string, number, boolean, or null.",
    "- NO nested objects and NO arrays inside a record (e.g. do NOT put {\"baseline\": {...}} or {\"tags\": [...]}).",
    "- Every record must use the SAME keys, and those keys must be the ones listed in `variables`.",
    "- The records MUST include enough columns to draw or compute the figure. If generating raw survival/time-to-event data, include at least a time column, a status column (e.g., 1=Event, 0=Censored), and a grouping column if there are treatment arms.",
    "- Use your own meaningful column names.",
    "- No markdown, no comments, no explanation. JSON only.",
    "",
    "Data quality principles (MUST follow):",
    "- Raw Data Over Computed Data: If the figure involves complex statistical summaries (e.g., survival probabilities, cumulative incidence, Kaplan-Meier), NEVER generate the pre-calculated Y-axis probabilities/rates. Instead, you MUST generate the raw subject-level data needed to compute them (e.g., `time_to_event` and `event_status` where 1=Event, 0=Censored).",
    "- Action Enforcement: Your generated `records` MUST strictly reflect your statistical reasoning.",
    "- Clinical Realism: Introduce realistic clinical/biological variability and noise into the raw data.",
    "- Include enough rows to make the figure look complete and realistic (e.g., 20-50 raw subjects per treatment/group for time-to-event figures, so that R can compute a smooth or multi-step curve).",
    `- ${groupRule}`,
    "",
    `Figure title: ${String(pureTitle || "").trim()}`,
    figureTypeText ? `Figure type: ${figureTypeText}` : "",
    variablesBlock,
    "",
    tableContext || "Corresponding table information: (not available)"
  ]
    .filter(Boolean)
    .join("\n");

  try {
    const completion = await client.chat.completions.create({
      model,
      messages: [
        { role: "system", content: "You output strict JSON describing synthetic data for a statistical figure." },
        { role: "user", content: prompt }
      ],
      temperature: 0.3,
      response_format: { type: "json_object" }
    });
    const parsed = parseJsonObjectFromText(completion?.choices?.[0]?.message?.content || "");
    if (!parsed || typeof parsed !== "object") {
      throw new Error("AI did not return valid JSON mock data.");
    }
    validateFlatRecordsJson(parsed);
    return { chartTitle, variables, dataJson: parsed };
  } catch (e) {
    throw new Error(`AI mock data generation failed: ${e?.message || String(e)}`);
  }
}

/**
 * Step 3 of the pipeline (Path A): the AI no longer writes R code. It only picks
 * a figure family and maps the existing data columns to that family's roles. The
 * deterministic R code is then assembled locally by `figureRTemplates`.
 *
 * Returns: { figureFamily, specificType, mapping, labels }
 */
async function generateFigureMappingByAi({ dataJson, figureType, client, model }) {
  const figureTypeText = String(figureType || "").trim();
  const records = Array.isArray(dataJson?.records) ? dataJson.records : [];
  const availableColumns = records.length
    ? Array.from(new Set(records.flatMap((r) => (r && typeof r === "object" ? Object.keys(r) : []))))
    : Array.isArray(dataJson?.variables)
    ? dataJson.variables.map((v) => String(v || "").trim()).filter(Boolean)
    : [];
  const sampleRecords = records.slice(0, 5);

  const prompt = [
    "You are a senior clinical biostatistician choosing how to plot an existing dataset.",
    "You will NOT write any R code. Instead, classify the figure and map data columns to plotting roles.",
    "",
    getMappingSchemaHint(),
    "",
    "Return STRICT JSON only with EXACTLY this shape:",
    "{",
    `  "figureFamily": one of [${KNOWN_FIGURE_FAMILIES.map((f) => `"${f}"`).join(", ")}],`,
    '  "specificType": "short label, e.g. cumulative_incidence / overall_survival / mean_change",',
    '  "mapping": { "<role>": "<exact column name>" },',
    '  "labels": { "x": "X axis label", "y": "Y axis label" }',
    "}",
    "",
    "Hard rules:",
    "- Every column name you put in `mapping` MUST be one of the EXISTING columns listed below. Never invent names.",
    "- Choose the SINGLE family whose computation matches the figure's intent.",
    "- For survival / cumulative-incidence / Kaplan-Meier figures, ALWAYS use family `time_to_event` and map `timeColumn` + `statusColumn` (and `groupColumn` if there are treatment arms). The status column must be the one holding 1=event / 0=censored.",
    "- Set `specificType` to `cumulative_incidence` for incidence/recurrence curves, or `overall_survival`/`survival` for survival curves. This decides whether the y-axis is S(t) or 1 - S(t).",
    "- Only include mapping keys relevant to the chosen family. Omit optional keys you don't need.",
    "- No markdown, no comments. JSON only.",
    "",
    figureTypeText ? `Figure type hint: ${figureTypeText}` : "",
    `Existing data columns: ${availableColumns.join(", ") || "(none)"}`,
    "Sample records (first rows):",
    JSON.stringify(sampleRecords)
  ]
    .filter(Boolean)
    .join("\n");

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: "system", content: "You output strict JSON mapping data columns to a plotting family." },
      { role: "user", content: prompt }
    ],
    temperature: 0.1,
    response_format: { type: "json_object" }
  });

  const parsed = parseJsonObjectFromText(completion?.choices?.[0]?.message?.content || "");
  if (!parsed || typeof parsed !== "object") {
    throw new Error("AI did not return a valid figure mapping JSON.");
  }
  const figureFamily = String(parsed.figureFamily || "").trim();
  if (!KNOWN_FIGURE_FAMILIES.includes(figureFamily)) {
    throw new Error(
      `AI returned unknown figureFamily "${figureFamily}". Must be one of: ${KNOWN_FIGURE_FAMILIES.join(", ")}.`
    );
  }

  return {
    figureFamily,
    specificType: String(parsed.specificType || "").trim(),
    mapping: parsed.mapping && typeof parsed.mapping === "object" ? parsed.mapping : {},
    labels: parsed.labels && typeof parsed.labels === "object" ? parsed.labels : {},
    availableColumns
  };
}

async function autoGenerateMockTflFigure({
  studyNumber,
  type,
  number,
  pureTitle,
  drug1,
  drug2,
  figureType,
  correspondingTableType,
  correspondingTableNumber,
  mongoose,
  openaiApiKey
}) {
  const key = normalizeFigureKey(type, number);
  if (!SUPPORTED_FIGURE_KEYS.has(key)) {
    throw new Error("Auto figure generation is only enabled for FIGURE 14.2.1.1 in the current prototype.");
  }
  if (!studyNumber) {
    throw new Error("Missing studyNumber");
  }

  // Read the FULL saved info of the corresponding table (if any) so the figure
  // can be made consistent with the table (treatment arms, rows, SAP context).
  const tableInfo = await loadCorrespondingTableInfo({
    studyNumber,
    tableType: correspondingTableType,
    tableNumber: correspondingTableNumber,
    mongoose
  });

  const usedColumnGroups = extractTreatmentGroups(tableInfo?.columnHeaderConfig?.columnGroups).slice(0, 2);

  const client = openaiApiKey ? new OpenAI({ apiKey: openaiApiKey }) : null;
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";

  // Step 1+2: AI decides which variables matter and freely designs the JSON data.
  const spec = await createSpecByAi({ pureTitle, drug1, drug2, tableInfo, figureType, client, model });

  // Step 3: AI only classifies the figure family and maps the existing data
  // columns to plotting roles. It does NOT write R code.
  const figureMapping = await generateFigureMappingByAi({
    dataJson: spec.dataJson,
    figureType,
    client,
    model
  });

  // Step 3.5: assemble deterministic, pre-tested R code from the template
  // library. This is what eliminates AI-induced R syntax/statistics crashes.
  const rBody = buildFigureRCode({
    figureFamily: figureMapping.figureFamily,
    specificType: figureMapping.specificType,
    mapping: figureMapping.mapping,
    labels: figureMapping.labels,
    availableColumns: figureMapping.availableColumns
  });

  // Inject the JSON ourselves (correctly escaped) so the template code can rely
  // on `json_data` being present and valid.
  const singleLineJson = JSON.stringify(spec.dataJson);
  const jsonPrefix = `library(jsonlite)\njson_data <- "${escapeRString(singleLineJson)}"\n`;
  const generatedRCode = `${jsonPrefix}\n${rBody}`;

  // Step 4: render on the remote R runner. The R code is template-built, so a
  // failure here is rare; we still log it for diagnostics.
  let image;
  try {
    image = await generateMockTflFigureFromRCode({ rCodeText: generatedRCode });
  } catch (err) {
    console.error("\n[MOCKTFL FATAL] R Runner failed! The template generated the following R code which caused the crash:\n");
    console.error("====================================================");
    console.error(generatedRCode);
    console.error("====================================================\n");
    throw err;
  }

  const dataJsonObj =
    spec.dataJson && typeof spec.dataJson === "object" && !Array.isArray(spec.dataJson)
      ? spec.dataJson
      : { data: spec.dataJson };

  return {
    figureKind: String(figureType || figureMapping.specificType || dataJsonObj.plotType || "").trim(),
    chartTitle: spec.chartTitle,
    mockData: {
      variables: Array.isArray(spec.variables) ? spec.variables : [],
      ...dataJsonObj
    },
    figureMapping: {
      figureFamily: figureMapping.figureFamily,
      specificType: figureMapping.specificType,
      mapping: figureMapping.mapping,
      labels: figureMapping.labels
    },
    tableInfo: {
      fetched: !!tableInfo?.found,
      label: tableInfo?.found ? tableInfo.label : "",
      usedColumnGroups: usedColumnGroups.length >= 2 ? usedColumnGroups : []
    },
    generatedRCode,
    imageBase64: image.imageBase64,
    mimeType: image.mimeType,
    stdout: image.stdout,
    stderr: image.stderr
  };
}

module.exports = {
  autoGenerateMockTflFigure,
  loadCorrespondingTableInfo,
  __test: { createSpecByAi, generateFigureMappingByAi }
};

