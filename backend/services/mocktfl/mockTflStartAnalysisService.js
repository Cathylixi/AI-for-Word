const OpenAI = require("openai");

/**
 * Compute cosine similarity for two vectors.
 * Returns null if vectors are invalid or have mismatched lengths.
 *
 * NOTE:
 * - We keep this implementation local to MockTFL to avoid mixing domains.
 */
function cosineSimilarity(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return null;
  let dot = 0;
  let normA = 0;
  let normB = 0;
  for (let i = 0; i < a.length; i++) {
    const x = a[i];
    const y = b[i];
    dot += x * y;
    normA += x * x;
    normB += y * y;
  }
  if (normA === 0 || normB === 0) return null;
  return dot / (Math.sqrt(normA) * Math.sqrt(normB));
}

function isValidEmbeddingVector(v) {
  if (!Array.isArray(v) || v.length === 0) return false;
  for (let i = 0; i < Math.min(v.length, 5); i++) {
    if (typeof v[i] !== "number" || !Number.isFinite(v[i])) return false;
  }
  return true;
}

/**
 * Normalize a human label / key for reliable matching across data sources.
 *
 * Why:
 * - SAP saves `matchedCrfFormName` as a human-facing label (e.g. "END OF TREATMENT PERIOD B").
 * - CRF chunks may store the same form name in `data.title` or in `formKey`.
 * - We want exact (case/space-insensitive) matching before falling back to embeddings.
 */
function normalizeLabel(s) {
  return String(s || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

/**
 * Extract question strings from CRF LabelForm list.
 * Expected CRF structure:
 * - doc.data.LabelForm: Array
 * - doc.data.LabelForm[i].content.question_part.text: string
 */
function extractQuestionsFromLabelForm(labelFormArray) {
  const arr = Array.isArray(labelFormArray) ? labelFormArray : [];
  const seen = new Set();
  const out = [];
  for (const item of arr) {
    const q = String(item?.content?.question_part?.text || "").trim();
    if (!q) continue;
    if (seen.has(q)) continue;
    seen.add(q);
    out.push(q);
  }
  return out;
}

/**
 * Parse numbered bullet text into structured items.
 * Expected input format:
 * 1. ...
 * 2. ...
 * 3. ...
 *
 * Fallback:
 * - If numbered bullets are not detected, treat non-empty lines as items.
 */
function parseNumberedBullets(text) {
  const raw = String(text || "").trim();
  if (!raw) return [];

  const lines = raw
    .split(/\r\n|\n/g)
    .map((l) => l.trim())
    .filter(Boolean);

  const items = [];
  for (const line of lines) {
    const m = line.match(/^(\d+)\.\s*(.+)$/);
    if (m) {
      items.push({ index: Number(m[1]), text: String(m[2] || "").trim() });
    }
  }

  if (items.length > 0) return items;
  // Fallback path for imperfect model outputs.
  return lines.map((line, i) => ({ index: i + 1, text: line }));
}

/**
 * Extract a JSON object from a model response.
 *
 * Why:
 * - The model may return plain JSON or wrap it in a fenced code block.
 * - We keep the parser local so the analysis flow can stay resilient.
 */
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

/**
 * Normalize a model-provided array of strings into clean, unique lines.
 */
function normalizeStringList(arr) {
  const input = Array.isArray(arr) ? arr : [];
  const seen = new Set();
  const out = [];

  input.forEach((item) => {
    const text = String(item || "").trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    out.push(text);
  });

  return out;
}

/**
 * Convert an array of items into the numbered bullet format already used by the UI.
 */
function formatNumberedListText(items) {
  const arr = Array.isArray(items) ? items : [];
  if (arr.length === 0) return "";
  return arr.map((item, idx) => `${idx + 1}. ${String(item || "").trim()}`).join("\n");
}

/**
 * Normalize user-confirmed variable inputs into the numbered-bullet structure
 * already used by the CRF mapping pipeline.
 *
 * Why:
 * - The dialog lets users freely edit the suggested variables before mapping.
 * - We accept either a structured array or free-form textarea text.
 */
function normalizeConfirmedAnalysisItems({ variables, variablesText }) {
  const normalizedVariables = normalizeStringList(variables);
  if (normalizedVariables.length > 0) {
    return normalizedVariables.map((text, idx) => ({ index: idx + 1, text }));
  }

  const parsedItems = parseNumberedBullets(variablesText).filter((item) => String(item?.text || "").trim());
  return parsedItems.map((item, idx) => ({
    index: Number(item?.index || idx + 1),
    text: String(item?.text || "").trim()
  }));
}

/**
 * Convert value_part into normalized answer options.
 * Priority:
 * 1) value_part.options[] (already structured)
 * 2) value_part.text_multiline (split by lines)
 * 3) value_part.text (single fallback)
 */
function extractAnswerOptionsFromValuePart(valuePart) {
  const vp = valuePart || {};

  if (Array.isArray(vp.options) && vp.options.length > 0) {
    return vp.options
      .map((opt, idx) => ({
        optionIndex: Number(opt?.option_index || idx + 1),
        text: String(opt?.text || "").trim()
      }))
      .filter((opt) => !!opt.text);
  }

  const multiline = String(vp.text_multiline || "").trim();
  if (multiline) {
    const lines = multiline
      .split(/\r\n|\n/g)
      .map((l) => l.trim())
      .filter(Boolean);
    return lines.map((text, idx) => ({ optionIndex: idx + 1, text }));
  }

  const single = String(vp.text || "").trim();
  if (!single) return [];
  return [{ optionIndex: 1, text: single }];
}

/**
 * Build CRF question candidates from LabelForm rows.
 * Each candidate contains:
 * - matchIndex
 * - questionText
 * - answer options (from value_part)
 */
function buildCrfQuestionCandidates(labelFormArray) {
  const arr = Array.isArray(labelFormArray) ? labelFormArray : [];
  const out = [];

  for (const item of arr) {
    const matchIndex = Number(item?.match_index);
    const questionText = String(
      item?.content?.question_part?.text ||
        item?.content?.full_text_without_number?.text ||
        item?.content?.full_text ||
        ""
    ).trim();
    if (!questionText) continue;

    const answers = extractAnswerOptionsFromValuePart(item?.content?.value_part);
    out.push({
      matchIndex: Number.isFinite(matchIndex) ? matchIndex : null,
      questionText,
      answers,
      rawValueText: String(item?.content?.value_part?.text || "").trim(),
      rawValueMultilineText: String(item?.content?.value_part?.text_multiline || "").trim()
    });
  }

  return out;
}

/**
 * Match CRF forms by title embedding and return best match + extracted questions.
 * We only use `data.title_embedding` for matching.
 */
async function matchCrfQuestionsByTitleEmbedding({ studyNumber, titleEmbedding, mongoose }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!isValidEmbeddingVector(titleEmbedding)) throw new Error("Invalid titleEmbedding");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");

  const crfDocs = await studiesCol
    .find(
      { studyNumber: String(studyNumber).trim(), chunkType: "crf_crfFormList" },
      { projection: { formKey: 1, data: 1, _id: 0 } }
    )
    .toArray();

  if (!Array.isArray(crfDocs) || crfDocs.length === 0) {
    throw new Error("No CRF forms found for this study (chunkType=crf_crfFormList)");
  }

  let best = null; // { score, formKey, title, doc }
  for (const doc of crfDocs) {
    const data = doc?.data || {};
    const score = cosineSimilarity(titleEmbedding, data.title_embedding);
    if (!Number.isFinite(score)) continue;
    if (!best || score > best.score) {
      best = {
        score,
        formKey: doc?.formKey || null,
        title: data?.title || null,
        doc
      };
    }
  }

  if (!best) {
    throw new Error("No comparable CRF title embeddings found (title_embedding missing or dimension mismatch)");
  }

  const questions = extractQuestionsFromLabelForm(best?.doc?.data?.LabelForm);
  return {
    matched: {
      formKey: best.formKey,
      title: best.title,
      score: best.score,
      matchedBy: "title_embedding"
    },
    questions,
    labelForm: Array.isArray(best?.doc?.data?.LabelForm) ? best.doc.data.LabelForm : []
  };
}

/**
 * Match a CRF form by the saved SAP matched CRF form name.
 *
 * Why:
 * - In SAP generation, we already matched a CRF form and saved the chosen form name
 *   into `llxdocument.studies(chunkType="SAP").data.sections[].matchedCrfFormName`.
 * - For MockTFL Start Analysis, we should re-use that saved decision to avoid re-matching
 *   against CRF embeddings (which can be noisy and pick the wrong form).
 *
 * Matching strategy:
 * - Try exact match (case/space-insensitive) against:
 *   - doc.data.title
 *   - doc.formKey
 * - If not found, throw with a clear message (caller can fallback to embeddings).
 */
async function matchCrfQuestionsBySapMatchedFormName({ studyNumber, matchedCrfFormName, mongoose }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  const target = normalizeLabel(matchedCrfFormName);
  if (!target) throw new Error("Missing matchedCrfFormName");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");

  const crfDocs = await studiesCol
    .find(
      { studyNumber: String(studyNumber).trim(), chunkType: "crf_crfFormList" },
      { projection: { formKey: 1, data: 1, _id: 0 } }
    )
    .toArray();

  if (!Array.isArray(crfDocs) || crfDocs.length === 0) {
    throw new Error("No CRF forms found for this study (chunkType=crf_crfFormList)");
  }

  let found = null;
  for (const doc of crfDocs) {
    const title = normalizeLabel(doc?.data?.title);
    const formKey = normalizeLabel(doc?.formKey);
    if (title && title === target) {
      found = doc;
      break;
    }
    if (formKey && formKey === target) {
      found = doc;
      break;
    }
  }

  if (!found) {
    throw new Error(`CRF form not found by saved SAP matchedCrfFormName: "${matchedCrfFormName}"`);
  }

  const questions = extractQuestionsFromLabelForm(found?.data?.LabelForm);
  return {
    matched: {
      formKey: found?.formKey || null,
      title: found?.data?.title || null,
      score: null,
      matchedBy: "sap_saved_matchedCrfFormName"
    },
    questions,
    labelForm: Array.isArray(found?.data?.LabelForm) ? found.data.LabelForm : []
  };
}

/**
 * Resolve the CRF form that should be used for mapping.
 *
 * Strategy:
 * 1) Prefer the CRF form name that was already saved during SAP generation.
 * 2) Fallback to title-embedding matching if the saved form cannot be reused.
 */
async function resolveCrfQuestionsContext({
  studyNumber,
  titleEmbedding,
  savedCrfFormName,
  mongoose
}) {
  const savedName = String(savedCrfFormName || "").trim();
  if (savedName) {
    try {
      return await matchCrfQuestionsBySapMatchedFormName({
        studyNumber,
        matchedCrfFormName: savedName,
        mongoose
      });
    } catch (e) {
      // Fallback: keep the system functional even if the saved form name
      // cannot be found in the latest CRF chunks.
    }
  }

  return await matchCrfQuestionsByTitleEmbedding({ studyNumber, titleEmbedding, mongoose });
}

/**
 * Match a MockTFL title embedding against saved SAP section title embeddings.
 *
 * Data source:
 * - llxdocument.studies { studyNumber, chunkType: "SAP" }
 * - data.entries[] contains number/title/title_embedding/title_embedded
 * - data.sections[] contains sectionNumber/text
 */
async function matchSavedSapSectionByTitleEmbedding({ studyNumber, titleEmbedding, mongoose }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!isValidEmbeddingVector(titleEmbedding)) throw new Error("Invalid titleEmbedding");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");

  const sapDoc = await studiesCol.findOne(
    { studyNumber: String(studyNumber).trim(), chunkType: "SAP" },
    { projection: { data: 1, _id: 0 } }
  );
  const sapState = sapDoc?.data || null;
  if (!sapState) {
    throw new Error("No saved SAP state found for this study. Please 'Save to Database' in SAP first.");
  }

  const entries = Array.isArray(sapState.entries) ? sapState.entries : [];
  const sections = Array.isArray(sapState.sections) ? sapState.sections : [];
  if (entries.length === 0) {
    throw new Error("Saved SAP state has no entries. Please re-save SAP to database.");
  }

  let best = null; // { score, number, title }
  for (const e of entries) {
    if (!e?.title_embedded) continue;
    if (!isValidEmbeddingVector(e?.title_embedding)) continue;
    const score = cosineSimilarity(titleEmbedding, e.title_embedding);
    if (!Number.isFinite(score)) continue;
    if (!best || score > best.score) {
      best = {
        score,
        number: String(e?.number || "").trim(),
        title: String(e?.title || "").trim()
      };
    }
  }

  if (!best?.number) {
    throw new Error(
      "No SAP title embeddings available for matching. Please 'Save to Database' in SAP to generate title embeddings."
    );
  }

  const matchedSection = sections.find((s) => String(s?.sectionNumber || "").trim() === best.number) || null;
  const sectionText = String(matchedSection?.text || "").trim();
  if (!sectionText) {
    throw new Error(`Matched SAP section text is empty (sectionNumber=${best.number}).`);
  }

  return {
    matched: {
      number: best.number,
      title: best.title || null,
      label: [best.number, best.title].filter(Boolean).join(" ").trim(),
      score: best.score
    },
    sectionText,
    // Return full matched section so callers can reuse SAP-saved metadata
    // (e.g., matchedCrfFormName) without re-querying.
    matchedSection: matchedSection || null
  };
}

/**
 * Ask the chat model for MockTFL statistical variables:
 * - which statistical variables, measures, populations, groupings, or denominators are needed.
 *
 * We return both structured arrays and preformatted numbered text so the caller
 * can reuse the same payload for UI rendering and downstream matching.
 */
async function analyzeSapSectionForMockTfl({
  pureTitle,
  matchedSapSectionLabel,
  matchedSapSectionText,
  openaiApiKey
}) {
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY");

  const client = new OpenAI({ apiKey: openaiApiKey });
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";

  const prompt = [
    "You are a senior biostatistician assistant helping prepare a Mock TFL.",
    "Task:",
    "Read the SAP section and return the statistical variables needed for the current MockTFL title.",
    "These should be key statistical variables, measures, populations, groupings, or denominators.",
    "",
    "Output requirements:",
    '- Return ONLY valid JSON with this exact shape: {"statisticalVariables":["..."]}',
    "- Each array item must be a short, standalone noun phrase.",
    "- Do not write full sentences.",
    "- Do not include markdown fences, headings, or explanatory prose outside JSON.",
    "- Keep the output specific to the current MockTFL title and the SAP section.",
    "",
    `MockTFL title context: ${String(pureTitle || "").trim()}`,
    `Matched SAP section: ${String(matchedSapSectionLabel || "").trim()}`,
    "",
    "SAP section text:",
    String(matchedSapSectionText || "")
  ].join("\n");

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: "system", content: "You produce concise, structured outputs for statistical programming." },
      { role: "user", content: prompt }
    ],
    temperature: 0.2
  });

  const text = completion?.choices?.[0]?.message?.content?.trim() || "";
  if (!text) throw new Error("OpenAI returned empty SAP MockTFL analysis.");

  const parsed = parseJsonObjectFromText(text);
  if (!parsed || typeof parsed !== "object") {
    throw new Error("OpenAI returned invalid JSON for SAP MockTFL analysis.");
  }

  const statisticalVariableItems = normalizeStringList(parsed?.statisticalVariables);

  return {
    statisticalVariableItems,
    statisticalVariablesText: formatNumberedListText(statisticalVariableItems)
  };
}

/**
 * For each SAP analysis bullet, find the best matched CRF question using cosine similarity.
 * Matching space:
 * - bullet embedding vs each candidate question embedding (inside matched CRF form)
 */
async function matchSapBulletsToCrfQuestions({
  analysisItems,
  labelForm,
  client,
  embeddingModel
}) {
  const bullets = Array.isArray(analysisItems)
    ? analysisItems.filter((it) => String(it?.text || "").trim())
    : [];
  const candidates = buildCrfQuestionCandidates(labelForm);

  if (bullets.length === 0 || candidates.length === 0) return [];

  // Build unique question text list and batch-embed once.
  const uniqueQuestionTexts = Array.from(new Set(candidates.map((c) => c.questionText)));
  const qEmbResp = await client.embeddings.create({
    model: embeddingModel,
    input: uniqueQuestionTexts
  });
  const qEmbByIndex = new Map((qEmbResp?.data || []).map((it) => [it.index, it.embedding]));
  const qEmbByText = new Map();
  uniqueQuestionTexts.forEach((text, idx) => {
    const emb = qEmbByIndex.get(idx);
    if (isValidEmbeddingVector(emb)) qEmbByText.set(text, emb);
  });

  // Batch-embed bullet texts.
  const bulletTexts = bullets.map((b) => String(b.text || "").trim());
  const bEmbResp = await client.embeddings.create({
    model: embeddingModel,
    input: bulletTexts
  });
  const bEmbByIndex = new Map((bEmbResp?.data || []).map((it) => [it.index, it.embedding]));

  const mappings = [];
  bullets.forEach((b, i) => {
    const bulletText = String(b?.text || "").trim();
    const bulletEmbedding = bEmbByIndex.get(i);
    if (!isValidEmbeddingVector(bulletEmbedding)) return;

    let best = null; // { score, candidate }
    for (const c of candidates) {
      const qEmb = qEmbByText.get(c.questionText);
      if (!isValidEmbeddingVector(qEmb)) continue;
      const score = cosineSimilarity(bulletEmbedding, qEmb);
      if (!Number.isFinite(score)) continue;
      if (!best || score > best.score) {
        best = { score, candidate: c };
      }
    }

    if (!best?.candidate) return;
    mappings.push({
      bulletIndex: Number(b?.index || i + 1),
      bulletText,
      matchedQuestion: {
        matchIndex: best.candidate.matchIndex,
        questionText: best.candidate.questionText,
        score: best.score
      },
      answers: best.candidate.answers,
      rawValueText: best.candidate.rawValueText,
      rawValueMultilineText: best.candidate.rawValueMultilineText
    });
  });

  return mappings;
}

/**
 * Call GPT to rewrite matched CRF questions into formal statistical noun phrases (table headers).
 *
 * Rules:
 * 1. Must be a noun phrase.
 * 2. Remove questions (Was, Did, Is, ?).
 * 3. Remove instructions (If Yes, please specify...).
 * 4. Keep it professional.
 */
async function rewriteCrfQuestionsForMockTfl({ questionTexts, pureTitle, client, model }) {
  const uniqueQuestions = Array.from(new Set(questionTexts.map(q => String(q).trim()).filter(Boolean)));
  if (uniqueQuestions.length === 0) return new Map();

  // We use a JSON object to map original text to rewritten text.
  // We send a numbered list of questions to GPT to save tokens, and ask for a JSON object back mapping index to rewritten text.
  const payloadForPrompt = uniqueQuestions.map((q, idx) => `${idx}: ${q}`).join("\n");

  const prompt = [
    "You are a senior biostatistician assistant.",
    "Task:",
    "Rewrite the provided CRF (Case Report Form) questions into formal statistical table row headers.",
    "",
    "Rules:",
    "1. OUTPUT MUST BE A NOUN PHRASE (e.g., 'Number of patients...', 'Reasons for...', 'Age', 'Sex').",
    "2. DO NOT output complete sentences with a subject and verb. DO NOT use question marks.",
    "3. REMOVE all question words ('Was', 'Did', 'Is', 'Are').",
    "4. REMOVE all form-filling instructions (e.g., 'If Yes, please specify...', 'Select from list...', '(If Death then please complete...)').",
    "5. Keep the statistical meaning intact and concise.",
    "",
    `Context (Current MockTFL Title): ${pureTitle}`,
    "",
    "Few-shot Examples:",
    "Input: 'Was treatment completed per protocol for period B?' -> Output: 'Number of Patients Completed Study Treatment'",
    "Input: 'Primary reason for treatment discontinuation ( If Death then please complete the Death Form.)' -> Output: 'Reasons for Early Study Treatment Discontinuation'",
    "Input: 'Did the patient experience any Major Protocol Deviations?' -> Output: 'Number of Patients with at least one major protocol deviation'",
    "Input: 'Age at time of consent' -> Output: 'Age (years)'",
    "",
    "Format Requirement:",
    "Return ONLY valid JSON. The JSON must be an object where the keys are the input indices, and the values are the rewritten noun phrases.",
    'Example: {"0": "Rewritten Phrase 0", "1": "Rewritten Phrase 1"}',
    "",
    "Input Questions to Rewrite:",
    payloadForPrompt
  ].join("\n");

  const completion = await client.chat.completions.create({
    model: model || "gpt-4o-mini",
    messages: [
      { role: "system", content: "You produce concise, structured outputs for statistical programming." },
      { role: "user", content: prompt }
    ],
    temperature: 0.1
  });

  const text = completion?.choices?.[0]?.message?.content?.trim() || "";
  const parsed = parseJsonObjectFromText(text) || {};

  const resultMap = new Map();
  uniqueQuestions.forEach((q, idx) => {
    const rewritten = parsed[String(idx)];
    if (rewritten && typeof rewritten === "string" && rewritten.trim()) {
      resultMap.set(q, rewritten.trim());
    } else {
      resultMap.set(q, q); // fallback to original if missing
    }
  });

  return resultMap;
}

/**
 * Step 1 for MockTFL Start Analysis:
 * - Embed pureTitle once
 * - SAP: match saved SAP entry + fetch section text
 * - SAP: run AI extraction on matched SAP section text
 * - Return only the editable variable suggestions and minimal context for step 2
 */
async function startMockTflAnalysis({ studyNumber, pureTitle, mongoose, openaiApiKey }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!pureTitle || !String(pureTitle).trim()) throw new Error("Missing pureTitle");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY");

  const client = new OpenAI({ apiKey: openaiApiKey });
  const embeddingModel = process.env.OPENAI_EMBEDDING_MODEL || process.env.EMBEDDING_MODEL || "text-embedding-3-small";

  // 1) Embed pure title ONCE and reuse the vector for both CRF and SAP matching.
  const embResp = await client.embeddings.create({
    model: embeddingModel,
    input: String(pureTitle).trim()
  });
  const titleEmbedding = embResp?.data?.[0]?.embedding || null;
  if (!isValidEmbeddingVector(titleEmbedding)) {
    throw new Error("Failed to create embedding for title");
  }

  // 2) SAP: match saved SAP section and get raw text (+ saved meta like matchedCrfFormName)
  const sapMatch = await matchSavedSapSectionByTitleEmbedding({ studyNumber, titleEmbedding, mongoose });
  const savedCrfFormName = String(sapMatch?.matchedSection?.matchedCrfFormName || "").trim();

  // 3) SAP: derive variable suggestions from the matched SAP text.
  const sapAnalysis = await analyzeSapSectionForMockTfl({
    pureTitle,
    matchedSapSectionLabel: sapMatch?.matched?.label || "",
    matchedSapSectionText: sapMatch?.sectionText || "",
    openaiApiKey
  });

  return {
    embeddingModel,
    context: {
      pureTitle: String(pureTitle).trim(),
      savedCrfFormName
    },
    sap: {
      matched: sapMatch.matched,
      sectionText: sapMatch.sectionText,
      statisticalVariablesText: sapAnalysis?.statisticalVariablesText || "",
      statisticalVariableItems: sapAnalysis?.statisticalVariableItems || [],
    }
  };
}

/**
 * Step 2 for MockTFL Start Analysis:
 * - Accept the user-confirmed variables from the dialog
 * - Resolve the CRF form context
 * - Map each confirmed variable to the best matching CRF question
 */
async function mapMockTflVariablesToCrf({
  studyNumber,
  pureTitle,
  variables,
  variablesText,
  savedCrfFormName,
  mongoose,
  openaiApiKey
}) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!pureTitle || !String(pureTitle).trim()) throw new Error("Missing pureTitle");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY");

  const analysisItems = normalizeConfirmedAnalysisItems({ variables, variablesText });
  if (analysisItems.length === 0) {
    throw new Error("No confirmed statistical variables were provided.");
  }

  const client = new OpenAI({ apiKey: openaiApiKey });
  const embeddingModel = process.env.OPENAI_EMBEDDING_MODEL || process.env.EMBEDDING_MODEL || "text-embedding-3-small";

  // Recreate the title embedding in step 2 so we can still fall back to title-based
  // CRF matching after the user has edited the variables.
  const embResp = await client.embeddings.create({
    model: embeddingModel,
    input: String(pureTitle).trim()
  });
  const titleEmbedding = embResp?.data?.[0]?.embedding || null;
  if (!isValidEmbeddingVector(titleEmbedding)) {
    throw new Error("Failed to create embedding for title");
  }

  const crf = await resolveCrfQuestionsContext({
    studyNumber,
    titleEmbedding,
    savedCrfFormName,
    mongoose
  });

  const mappings = await matchSapBulletsToCrfQuestions({
    analysisItems,
    labelForm: crf?.labelForm || [],
    client,
    embeddingModel
  });

  // Call GPT to rewrite the matched CRF questions
  const matchedQuestionTexts = mappings.map(m => m.matchedQuestion?.questionText).filter(Boolean);
  const rewrittenMap = await rewriteCrfQuestionsForMockTfl({
    questionTexts: matchedQuestionTexts,
    pureTitle,
    client,
    model: process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini"
  });

  // Inject rewritten text into mappings
  mappings.forEach(m => {
    if (m.matchedQuestion && m.matchedQuestion.questionText) {
      m.matchedQuestion.rewrittenQuestionText = rewrittenMap.get(m.matchedQuestion.questionText) || m.matchedQuestion.questionText;
    }
  });

  return {
    embeddingModel,
    crf: {
      matched: crf?.matched || null,
      questions: Array.isArray(crf?.questions) ? crf.questions : [],
      mappings
    },
    confirmedVariables: {
      text: formatNumberedListText(analysisItems.map((item) => item.text)),
      items: analysisItems
    }
  };
}

/**
 * Extract drug comparison from SAP Introduction
 */
async function extractDrugComparisonFromSap({ studyNumber, mongoose, openaiApiKey }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY");

  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");

  const sapDoc = await studiesCol.findOne(
    { studyNumber: String(studyNumber).trim(), chunkType: "SAP" },
    { projection: { data: 1, _id: 0 } }
  );
  const sapState = sapDoc?.data || null;
  if (!sapState) {
    throw new Error("No saved SAP state found for this study. Please 'Save to Database' in SAP first.");
  }

  const sections = Array.isArray(sapState.sections) ? sapState.sections : [];
  if (sections.length === 0) {
    throw new Error("Saved SAP state has no sections. Please re-save SAP to database.");
  }

  // Look for "Introduction" or "Background" or just take the first section
  let introSection = sections.find((s) => {
    const title = String(s?.title || "").toUpperCase();
    return title.includes("INTRODUCTION") || title.includes("BACKGROUND");
  });

  if (!introSection) {
    introSection = sections[0]; // Fallback to the first section
  }

  const textToAnalyze = String(introSection?.text || introSection?.content || "").trim();
  if (!textToAnalyze) {
    throw new Error("Found SAP Introduction section, but the text is empty.");
  }

  const client = new OpenAI({ apiKey: openaiApiKey });
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";

  const prompt = [
    "You are a clinical trial biostatistician.",
    "Task: Read the following text from the SAP (Statistical Analysis Plan) Introduction / Background.",
    "Identify what drugs or treatment groups are being compared in this trial.",
    "",
    "Instructions:",
    "- Find the two main drugs or treatment groups being compared.",
    "- If only one drug is tested against a placebo, use 'Placebo' as the second drug.",
    "- Return ONLY valid JSON with exactly this shape: {\"drug1\":\"...\",\"drug2\":\"...\"}",
    "- Output ONLY the JSON string, without any extra text, markdown formatting, or code fences.",
    "",
    "SAP Text:",
    textToAnalyze
  ].join("\n");

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: "system", content: "You extract key trial design information from SAP texts into structured JSON." },
      { role: "user", content: prompt }
    ],
    temperature: 0.1,
    response_format: { type: "json_object" }
  });

  const text = completion?.choices?.[0]?.message?.content?.trim() || "";
  const parsed = parseJsonObjectFromText(text) || {};

  return {
    drug1: String(parsed.drug1 || "Drug 1").trim(),
    drug2: String(parsed.drug2 || "Drug 2").trim(),
    matchedSapSectionNumber: introSection.sectionNumber,
    matchedSapSectionTitle: introSection.title
  };
}

module.exports = {
  startMockTflAnalysis,
  mapMockTflVariablesToCrf,
  extractDrugComparisonFromSap
};

