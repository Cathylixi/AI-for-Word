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

function bestContentSimilarity(tflEmbedding, contentEmbeddings) {
  if (!Array.isArray(contentEmbeddings) || contentEmbeddings.length === 0) return null;
  let best = null;
  for (const ch of contentEmbeddings) {
    const score = cosineSimilarity(tflEmbedding, ch?.embedding);
    if (score === null) continue;
    if (best === null || score > best) best = score;
  }
  return best;
}

async function autoSelectMockTflSections({ studyNumber, tflSections, threshold, mongoose }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");
  if (!Array.isArray(tflSections) || tflSections.length === 0) {
    throw new Error("Missing tflSections");
  }

  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");
  const embChunk = await studiesCol.findOne(
    { studyNumber: String(studyNumber), chunkType: "protocol_extraction_embeddings" },
    { projection: { data: 1, _id: 0 } }
  );
  if (!embChunk?.data?.sectionedText?.length) {
    throw new Error("Protocol extraction embeddings not found for this study");
  }

  const refDb = mongoose.connection.client.db("References");
  const tflCol = refDb.collection("MockTFL_Example");

  const keys = tflSections
    .map((s) => ({
      type: String(s.type || "").trim().toUpperCase(),
      number: String(s.number || "").trim()
    }))
    .filter((k) => k.type && k.number);

  const keyOrConditions = keys.map((k) => ({ type: k.type, number: k.number }));
  const tflDocs = keyOrConditions.length
    ? await tflCol
      .find({ $or: keyOrConditions }, { projection: { type: 1, number: 1, title: 1, embedding: 1, _id: 0 } })
      .toArray()
    : [];

  const tflMap = new Map(
    tflDocs.map((d) => [`${String(d.type || "").trim().toUpperCase()}::${String(d.number || "").trim()}`, d])
  );
  const scoreThreshold = typeof threshold === "number" ? threshold : 0.5;

  const recommended = tflSections.map((s) => {
    const type = String(s.type || "").trim().toUpperCase();
    const number = String(s.number || "").trim();
    const title = String(s.title || "").trim();
    const tflEntry = tflMap.get(`${type}::${number}`);
    if (!tflEntry?.embedding) {
      return { type, number, title, recommended: false, score: null, matchedProtocolTitle: null };
    }

    let best = null;
    for (const sec of embChunk.data.sectionedText) {
      const titleScore = cosineSimilarity(tflEntry.embedding, sec.title_embedding);
      const contentScore = bestContentSimilarity(tflEntry.embedding, sec.content_embeddings);
      const score = Math.max(
        typeof titleScore === "number" ? titleScore : -Infinity,
        typeof contentScore === "number" ? contentScore : -Infinity
      );
      if (!Number.isFinite(score)) continue;
      if (!best || score > best.score) {
        best = {
          score,
          titleScore: typeof titleScore === "number" ? titleScore : null,
          contentScore: typeof contentScore === "number" ? contentScore : null,
          title: sec.title || null,
          number: sec.number || null
        };
      }
    }

    const score = best?.score ?? null;
    return {
      type,
      number,
      title,
      recommended: typeof score === "number" ? score >= scoreThreshold : false,
      score,
      titleScore: best?.titleScore ?? null,
      contentScore: best?.contentScore ?? null,
      matchedBy:
        best?.titleScore !== null && best?.contentScore !== null
          ? best.titleScore >= best.contentScore
            ? "title"
            : "content"
          : best?.titleScore !== null
            ? "title"
            : best?.contentScore !== null
              ? "content"
              : null,
      matchedProtocolTitle: best?.title || null,
      matchedProtocolNumber: best?.number || null
    };
  });

  return recommended;
}

module.exports = {
  autoSelectMockTflSections
};
