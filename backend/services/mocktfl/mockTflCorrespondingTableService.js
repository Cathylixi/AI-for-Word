const OpenAI = require("openai");

const SUPPORTED_FIGURE_KEYS = new Set(["FIGURE:14.2.1.1"]);

/**
 * Compute cosine similarity for two numeric vectors.
 * Returns null if vectors are invalid or have mismatched lengths.
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

function normalizeFigureKey(type, number) {
  return `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`;
}

function buildTableLabel(entry) {
  const type = String(entry?.type || "TABLE").trim().toUpperCase();
  const number = String(entry?.number || "").trim();
  const title = String(entry?.title || "").trim();
  return [type, number, title].filter(Boolean).join(" ").trim();
}

/**
 * Find the MockTFL TABLE entry whose name is most semantically similar to the
 * current FIGURE name.
 *
 * Why:
 * - The figure mock-data workflow needs to know which table it corresponds to.
 * - We reuse the title embeddings already stored in References.MockTFL_Example
 *   (built from `${type} ${number} ${title}`) and compare them against a fresh
 *   embedding of the figure title using cosine similarity.
 */
async function findCorrespondingTable({ studyNumber, type, number, pureTitle, mongoose, openaiApiKey }) {
  const key = normalizeFigureKey(type, number);
  if (!SUPPORTED_FIGURE_KEYS.has(key)) {
    throw new Error("Find corresponding table is only enabled for FIGURE 14.2.1.1 in the current prototype.");
  }
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const figureTitle = String(pureTitle || "").trim();
  if (!figureTitle) throw new Error("Missing figure title for corresponding table matching.");
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY for embedding generation.");

  const refDb = mongoose.connection.client.db("References");
  const tflCol = refDb.collection("MockTFL_Example");
  const tableDocs = await tflCol
    .find(
      { type: "TABLE" },
      { projection: { type: 1, number: 1, title: 1, embedding: 1, order: 1, _id: 0 } }
    )
    .toArray();

  const tablesWithEmbedding = tableDocs.filter((d) => isValidEmbeddingVector(d?.embedding));
  if (tablesWithEmbedding.length === 0) {
    throw new Error("No TABLE embeddings available in References.MockTFL_Example for matching.");
  }

  const client = new OpenAI({ apiKey: openaiApiKey });
  const embeddingModel = process.env.OPENAI_EMBEDDING_MODEL || process.env.EMBEDDING_MODEL || "text-embedding-3-small";
  const embResp = await client.embeddings.create({
    model: embeddingModel,
    input: figureTitle
  });
  const figureEmbedding = embResp?.data?.[0]?.embedding || null;
  if (!isValidEmbeddingVector(figureEmbedding)) {
    throw new Error("Failed to create embedding for figure title.");
  }

  const scored = [];
  for (const entry of tablesWithEmbedding) {
    const score = cosineSimilarity(figureEmbedding, entry.embedding);
    if (typeof score !== "number" || !Number.isFinite(score)) continue;
    scored.push({
      type: String(entry?.type || "TABLE").trim().toUpperCase(),
      number: String(entry?.number || "").trim(),
      title: String(entry?.title || "").trim(),
      label: buildTableLabel(entry),
      score
    });
  }

  if (scored.length === 0) {
    throw new Error("No comparable TABLE embeddings found (dimension mismatch).");
  }

  scored.sort((a, b) => b.score - a.score);
  const best = scored[0];

  return {
    matchedTable: {
      type: best.type,
      number: best.number,
      title: best.title,
      label: best.label,
      score: best.score,
      matchedBy: "mocktfl_table_title_embedding",
      embeddingModel,
      matchedAt: new Date().toISOString()
    },
    candidates: scored.slice(0, 3)
  };
}

module.exports = {
  findCorrespondingTable
};
