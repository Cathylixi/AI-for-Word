const OpenAI = require("openai");

/**
 * Compute cosine similarity for two vectors.
 * Returns null if vectors are invalid or have mismatched lengths.
 *
 * NOTE:
 * - We intentionally keep this implementation local (not importing SAP code)
 *   to avoid mixing domains, while still using the same mathematical logic.
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

/**
 * Extract the question text from a CRF LabelForm object.
 *
 * Expected CRF structure (as per current DB):
 * - doc.data.LabelForm: Array
 * - doc.data.LabelForm[i].content.question_part.text: string
 */
function extractQuestionsFromLabelForm(labelFormArray) {
  const arr = Array.isArray(labelFormArray) ? labelFormArray : [];

  // Preserve order but deduplicate identical question strings.
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
 * Analyze a specific MockTFL item by:
 * 1) Embedding the (pure) title text
 * 2) Matching against CRF form title embeddings for the given study
 * 3) Returning the best-matched CRF form and all questions from its LabelForm list
 *
 * IMPORTANT:
 * - This is designed for a specific use-case (TABLE 14.1.1.2) now.
 * - The route layer will enforce the specific target gating.
 */
async function analyzeMockTflCrfQuestions({ studyNumber, pureTitle, mongoose, openaiApiKey }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!pureTitle || !String(pureTitle).trim()) throw new Error("Missing pureTitle");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");
  if (!openaiApiKey) throw new Error("Missing OPENAI_API_KEY");

  const client = new OpenAI({ apiKey: openaiApiKey });
  const embeddingModel = process.env.OPENAI_EMBEDDING_MODEL || "text-embedding-3-small";

  // 1) Create embedding vector for the table title (without the TABLE 14.x prefix)
  const embResp = await client.embeddings.create({
    model: embeddingModel,
    input: String(pureTitle).trim()
  });
  const titleEmbedding = embResp?.data?.[0]?.embedding || null;
  if (!Array.isArray(titleEmbedding) || titleEmbedding.length === 0) {
    throw new Error("Failed to create embedding for title");
  }

  // 2) Load CRF forms for this study (chunkType = crf_crfFormList)
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

  // 3) Find best match using cosine similarity against data.title_embedding
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

  // 4) Extract questions from the best matched CRF form
  const questions = extractQuestionsFromLabelForm(best?.doc?.data?.LabelForm);

  return {
    matched: {
      formKey: best.formKey,
      title: best.title,
      score: best.score
    },
    questions
  };
}

module.exports = {
  analyzeMockTflCrfQuestions
};

