const OpenAI = require("openai");

/**
 * Validate whether a value looks like an embedding vector.
 * We keep this intentionally simple:
 * - must be an array
 * - must have at least 1 number
 */
function isValidEmbeddingVector(v) {
  if (!Array.isArray(v) || v.length === 0) return false;
  // We only check the first few items for performance.
  for (let i = 0; i < Math.min(v.length, 5); i++) {
    if (typeof v[i] !== "number" || !Number.isFinite(v[i])) return false;
  }
  return true;
}

/**
 * Normalize title for embedding input to reduce accidental variance.
 * IMPORTANT:
 * - User requirement: "title will not change"
 * - Still, we normalize whitespace to avoid saving redundant embeddings.
 */
function normalizeTitleForEmbedding(title) {
  return String(title || "")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Ensure SAP entries contain cached or newly generated title embeddings.
 *
 * Behavior:
 * - If an entry already exists in the saved SAP state with `title_embedded=true`
 *   and a valid `title_embedding`, we reuse it and do NOT call OpenAI again.
 * - Otherwise, we generate embeddings (batch) for missing entries and store:
 *   - entry.title_embedded = true
 *   - entry.title_embedding = [ ...vector ]
 *
 * Design notes:
 * - This function runs on the BACKEND so OpenAI keys never touch the frontend.
 * - We only embed the section title (not "number + title") per user requirement.
 * - We key caching by `entry.number` because it is stable and unique in SAP TOC.
 *
 * @param {object} params
 * @param {string} params.studyNumber
 * @param {object} params.incomingState - The state payload being saved (must include entries array).
 * @param {object|null} params.existingState - Existing saved SAP state from DB (may be null).
 * @param {object} params.mongoose
 * @param {string} params.openaiApiKey
 * @returns {Promise<object>} enrichedState
 */
async function ensureSapEntryTitleEmbeddings({
  studyNumber,
  incomingState,
  existingState,
  mongoose,
  openaiApiKey
}) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!incomingState || typeof incomingState !== "object") throw new Error("Missing incomingState");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const incomingEntries = Array.isArray(incomingState.entries) ? incomingState.entries : [];
  if (incomingEntries.length === 0) {
    // Nothing to do; keep payload as-is.
    return incomingState;
  }

  // Build cache map from existing state: number -> { title, embedded, embedding }
  const existingEntries = Array.isArray(existingState?.entries) ? existingState.entries : [];
  const existingByNumber = new Map();
  for (const e of existingEntries) {
    const number = String(e?.number || "").trim();
    if (!number) continue;
    existingByNumber.set(number, {
      title: String(e?.title || "").trim(),
      title_embedded: !!e?.title_embedded,
      title_embedding: e?.title_embedding
    });
  }

  // First pass: decide which entries need embedding.
  const toEmbed = []; // { idx, number, titleNorm }
  const enrichedEntries = incomingEntries.map((e, idx) => {
    const number = String(e?.number || "").trim();
    const title = String(e?.title || "").trim();

    // Keep original entry fields, but we may add/overwrite embedding fields.
    const out = { ...(e || {}) };

    if (!number || !title) {
      // If entry is malformed, do not attempt embedding.
      return out;
    }

    const cached = existingByNumber.get(number);
    const titleNorm = normalizeTitleForEmbedding(title);

    // Reuse cached embedding if:
    // - title is identical (safety belt)
    // - cached says embedded=true
    // - cached embedding vector is valid
    if (
      cached &&
      cached.title === title &&
      cached.title_embedded === true &&
      isValidEmbeddingVector(cached.title_embedding)
    ) {
      out.title_embedded = true;
      out.title_embedding = cached.title_embedding;
      return out;
    }

    // Otherwise we will embed this title.
    out.title_embedded = false;
    // Do NOT keep stale title_embedding if any, unless it is valid + embedded=true
    // (we handle the reuse case above).
    delete out.title_embedding;
    toEmbed.push({ idx, number, titleNorm });
    return out;
  });

  if (toEmbed.length === 0) {
    return { ...incomingState, entries: enrichedEntries };
  }

  if (!openaiApiKey) {
    // We need embeddings but cannot create them.
    // Fail fast so the user knows why title_embedded remains false.
    throw new Error(
      "OPENAI_API_KEY is missing. Cannot generate SAP entry title embeddings for new/unsaved entries."
    );
  }

  // Deduplicate titles to minimize embedding calls/cost.
  const uniqueTitles = [];
  const titleToUniqueIndex = new Map(); // titleNorm -> uniqueTitles index
  for (const item of toEmbed) {
    if (titleToUniqueIndex.has(item.titleNorm)) continue;
    titleToUniqueIndex.set(item.titleNorm, uniqueTitles.length);
    uniqueTitles.push(item.titleNorm);
  }

  const client = new OpenAI({ apiKey: openaiApiKey });
  const embeddingModel = process.env.OPENAI_EMBEDDING_MODEL || process.env.EMBEDDING_MODEL || "text-embedding-3-small";

  // Batch embedding request.
  const resp = await client.embeddings.create({
    model: embeddingModel,
    input: uniqueTitles
  });

  // Map embedding outputs back by returned index.
  const indexToEmbedding = new Map();
  for (const item of resp?.data || []) {
    indexToEmbedding.set(item.index, item.embedding);
  }
  if (indexToEmbedding.size !== uniqueTitles.length) {
    throw new Error(
      `Embedding count mismatch: expected ${uniqueTitles.length}, got ${indexToEmbedding.size}`
    );
  }

  // Fill embeddings back into entries.
  for (const item of toEmbed) {
    const uniqueIdx = titleToUniqueIndex.get(item.titleNorm);
    const emb = indexToEmbedding.get(uniqueIdx);
    if (!isValidEmbeddingVector(emb)) continue;
    enrichedEntries[item.idx].title_embedded = true;
    enrichedEntries[item.idx].title_embedding = emb;
    enrichedEntries[item.idx].title_embedding_model = embeddingModel; // helpful for debugging/migration
  }

  return { ...incomingState, entries: enrichedEntries };
}

module.exports = {
  ensureSapEntryTitleEmbeddings
};

