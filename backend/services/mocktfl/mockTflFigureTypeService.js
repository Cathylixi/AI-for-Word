const OpenAI = require("openai");

const { loadCorrespondingTableInfo } = require("./mockTflFigureAutoService");

/**
 * Allowed figure types the model is encouraged to choose from. The answer is
 * stored as free text (and remains user-editable in the dialog), but we steer
 * the model toward this controlled vocabulary for consistency.
 */
const FIGURE_TYPE_OPTIONS = [
  "Cumulative incidence plot",
  "Kaplan-Meier survival curve",
  "Line plot (mean over time)",
  "Forest plot",
  "Waterfall plot",
  "Bar chart",
  "Scatter plot",
  "Other"
];

function buildPrompt({ figureTitle, tableInfo }) {
  const lines = [
    "You are a senior clinical biostatistician.",
    "Task: determine the most appropriate statistical figure/plot type for the figure described below,",
    "using the figure title and its corresponding table information.",
    "",
    "Choose the best match from this list when possible:",
    ...FIGURE_TYPE_OPTIONS.map((o) => `- ${o}`),
    "",
    "Answer with ONE short line: the figure type, optionally followed by a brief (one sentence) reason.",
    "Do not return JSON. Do not add extra commentary.",
    "",
    `Figure title: ${String(figureTitle || "").trim()}`
  ];

  if (tableInfo && tableInfo.found) {
    const cfg = tableInfo.columnHeaderConfig || {};
    lines.push(
      "",
      "Corresponding table information:",
      `- Table: ${tableInfo.label || ""}`.trim(),
      `- Column groups (treatment arms): ${(cfg.columnGroups || []).join(", ") || "(none)"}`,
      tableInfo.mappedCrfText ? `- Table row headers / mapped CRF:\n${tableInfo.mappedCrfText}` : "",
      tableInfo.suggestedVariablesText ? `- Suggested statistical variables:\n${tableInfo.suggestedVariablesText}` : "",
      tableInfo.matchedSapLabel ? `- Matched SAP section: ${tableInfo.matchedSapLabel}` : "",
      tableInfo.matchedSapSectionText ? `- SAP section text:\n${tableInfo.matchedSapSectionText}` : ""
    );
  } else {
    lines.push("", "Corresponding table information: (not available)");
  }

  return lines.filter(Boolean).join("\n");
}

/**
 * Ask GPT to define the figure type from the figure title + corresponding table content.
 *
 * @returns {Promise<{ figureType: string, tableInfoFetched: boolean }>}
 */
async function defineMockTflFigureType({
  studyNumber,
  pureTitle,
  tableType,
  tableNumber,
  mongoose,
  openaiApiKey
}) {
  if (!pureTitle || !String(pureTitle).trim()) throw new Error("Missing figure title");
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY");

  const tableInfo = await loadCorrespondingTableInfo({
    studyNumber,
    tableType,
    tableNumber,
    mongoose
  });

  const client = new OpenAI({ apiKey: openaiApiKey });
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";
  const prompt = buildPrompt({ figureTitle: pureTitle, tableInfo });

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: "system", content: "You are concise and only output the requested figure type." },
      { role: "user", content: prompt }
    ],
    temperature: 0.1
  });

  const figureType = String(completion?.choices?.[0]?.message?.content || "").trim();
  if (!figureType) throw new Error("Model returned an empty figure type.");

  return {
    figureType,
    tableInfoFetched: !!tableInfo?.found
  };
}

module.exports = {
  defineMockTflFigureType,
  FIGURE_TYPE_OPTIONS
};
