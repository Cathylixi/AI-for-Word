const OpenAI = require("openai");

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

function bestContentSimilarity(sapEmbedding, contentEmbeddings) {
  if (!Array.isArray(contentEmbeddings) || contentEmbeddings.length === 0) return null;
  let best = null;
  for (const ch of contentEmbeddings) {
    const score = cosineSimilarity(sapEmbedding, ch?.embedding);
    if (score === null) continue;
    if (best === null || score > best) best = score;
  }
  return best;
}

function buildContextFromRanges(fullText, ranges) {
  const sorted = ranges
    .filter(r => typeof r.startChar === "number" && typeof r.endChar === "number")
    .sort((a, b) => a.startChar - b.startChar);

  let out = "";
  let cursor = 0;
  for (const r of sorted) {
    const start = Math.max(r.startChar, cursor);
    const end = Math.max(start, r.endChar);
    out += fullText.slice(start, end);
    cursor = end;
  }
  return out.trim();
}

function truncateText(text, maxChars) {
  const s = String(text || "");
  if (s.length <= maxChars) return s;
  return s.slice(0, maxChars);
}

function normalizeTitleForEmbedding(title) {
  return String(title || "")
    .replace(/\s+/g, " ")
    .trim();
}

async function embedSapTitleForExampleMatch(client, title) {
  const normalized = normalizeTitleForEmbedding(title);
  if (!normalized) return null;

  const embeddingModel = process.env.OPENAI_EMBEDDING_MODEL || process.env.EMBEDDING_MODEL || "text-embedding-3-small";
  const response = await client.embeddings.create({
    model: embeddingModel,
    input: normalized
  });

  return {
    embedding: response?.data?.[0]?.embedding || null,
    embeddingModel
  };
}

// Match the current SAP section title against the uploaded SPI611 example titles.
// We only use SPI611 as a writing-style example, never as the source of study facts.
async function matchSpi611ExampleSectionByTitleEmbedding({ titleEmbedding, mongoose }) {
  if (!Array.isArray(titleEmbedding) || titleEmbedding.length === 0) return null;
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const refDb = mongoose.connection.client.db("References");
  const spiDoc = await refDb.collection("SAP_SPI611").findOne(
    { docType: "state", docKey: "CMAA868D12302" },
    { projection: { data: 1, _id: 0 } }
  );

  const entries = Array.isArray(spiDoc?.data?.entries) ? spiDoc.data.entries : [];
  const sections = Array.isArray(spiDoc?.data?.sections) ? spiDoc.data.sections : [];
  if (entries.length === 0 || sections.length === 0) return null;

  let best = null;
  for (const entry of entries) {
    const score = cosineSimilarity(titleEmbedding, entry?.title_embedding);
    if (score === null) continue;
    if (!best || score > best.score) {
      best = {
        number: entry?.number || null,
        title: entry?.title || null,
        score,
        titleEmbeddingModel: entry?.title_embedding_model || null
      };
    }
  }

  if (!best?.number) return null;

  const matchedSection = sections.find((s) => String(s?.sectionNumber || "") === String(best.number));
  const sectionText = String(matchedSection?.text || "").trim();
  if (!sectionText) return null;

  return {
    matched: best,
    sectionText
  };
}

async function generateSapSection({
  studyNumber,
  sapSectionNumber,
  refineInstruction,
  mongoose,
  openaiApiKey
}) {
  if (!studyNumber || !sapSectionNumber) {
    throw new Error("Missing studyNumber or sapSectionNumber");
  }

  if (!mongoose?.connection?.client) {
    throw new Error("Database connection is not ready");
  }

  const client = new OpenAI({ apiKey: openaiApiKey });

  // 1) Read SAP example entry (with embedding)
  const refDb = mongoose.connection.client.db("References");
  const sapCol = refDb.collection("SAP_Example");
  const sapEntry = await sapCol.findOne(
    { number: String(sapSectionNumber) },
    { projection: { number: 1, title: 1, embedding: 1, text: 1, _id: 0 } }
  );
  if (!sapEntry?.embedding) {
    throw new Error("SAP example entry or embedding not found");
  }

  // Build a title-only embedding for SPI611 style-example matching.
  // This keeps the example lookup aligned with the saved title vectors in References.SAP_SPI611.
  const titleEmbeddingResult = await embedSapTitleForExampleMatch(client, sapEntry.title);
  const sectionTitleEmbedding = titleEmbeddingResult?.embedding || null;

  // 2) Read protocol_extraction_embeddings for the study
  const docDb = mongoose.connection.client.db("llxdocument");
  const studiesCol = docDb.collection("studies");

  const embChunk = await studiesCol.findOne(
    { studyNumber: String(studyNumber), chunkType: "protocol_extraction_embeddings" },
    { projection: { data: 1, _id: 0 } }
  );
  if (!embChunk?.data?.sectionedText?.length) {
    throw new Error("Protocol extraction embeddings not found for this study");
  }

  // 3) Find best-matching protocol section by title embedding
  let best = null;
  for (const sec of embChunk.data.sectionedText) {
    const titleScore = cosineSimilarity(sapEntry.embedding, sec.title_embedding);
    const contentScore = bestContentSimilarity(sapEntry.embedding, sec.content_embeddings);
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
        sectionIndex: sec.sectionIndex,
        title: sec.title,
        number: sec.number,
        content_embeddings: sec.content_embeddings
      };
    }
  }
  if (!best) {
    throw new Error("No comparable protocol section embeddings found");
  }

  // 4) Load raw protocol content to reconstruct text
  const extractionChunk = await studiesCol.findOne(
    { studyNumber: String(studyNumber), chunkType: "protocol_extraction" },
    { projection: { data: 1, _id: 0 } }
  );
  const protoSection = extractionChunk?.data?.sectionedText?.[best.sectionIndex];
  const fullContent = protoSection?.content || "";
  if (!fullContent) {
    throw new Error("Matched protocol section content is empty");
  }

  const contextText = buildContextFromRanges(fullContent, best.content_embeddings || []);

  // 5) Find best-matching CRF form by embeddings (title + full_text)
  const crfDocs = await studiesCol
    .find(
      { studyNumber: String(studyNumber), chunkType: "crf_crfFormList" },
      { projection: { formKey: 1, data: 1, _id: 0 } }
    )
    .toArray();
  let bestCrf = null; // { score, titleScore, fullScore, formKey, title, fullText }
  for (const doc of crfDocs) {
    const data = doc?.data || {};
    const titleScore = cosineSimilarity(sapEntry.embedding, data.title_embedding);
    const fullScore = cosineSimilarity(sapEntry.embedding, data.full_text_embedding);
    const score = Math.max(
      typeof titleScore === "number" ? titleScore : -Infinity,
      typeof fullScore === "number" ? fullScore : -Infinity
    );
    if (!Number.isFinite(score)) continue;
    if (!bestCrf || score > bestCrf.score) {
      bestCrf = {
        score,
        titleScore: typeof titleScore === "number" ? titleScore : null,
        fullScore: typeof fullScore === "number" ? fullScore : null,
        formKey: doc?.formKey || null,
        title: data.title || null,
        fullText: data.full_text || ""
      };
    }
  }

  const crfContextText = bestCrf?.fullText ? truncateText(bestCrf.fullText, 6000) : "";
  const crfTitleLine = bestCrf?.title || bestCrf?.formKey || "";

  // 6) Find the closest SPI611 section by title embedding.
  const spi611Match = sectionTitleEmbedding
    ? await matchSpi611ExampleSectionByTitleEmbedding({
        titleEmbedding: sectionTitleEmbedding,
        mongoose
      })
    : null;
  const spi611ExampleText = String(spi611Match?.sectionText || "").trim();

  // 7) Generate SAP section text
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";
  const promptParts = [
    `Write the SAP section for: ${sapEntry.number} ${sapEntry.title}.`,
    "Use formal SAP style, concise, and structured in paragraphs.",
    "Do not include the section number/title as a heading.",
    "Use the following protocol context:",
    contextText
  ];

  if (crfContextText || crfTitleLine) {
    promptParts.push(
      "Use the following CRF context to supplement operational details (e.g., data collection fields):",
      [crfTitleLine ? `CRF Form: ${crfTitleLine}` : "", crfContextText].filter(Boolean).join("\n")
    );
  }

  if (spi611ExampleText) {
    promptParts.push(
      "Use the following SPI611 SAP section only as a writing example.",
      [
        "Learn the paragraph structure, tone, and level of detail from it.",
        "Do not copy wording, study-specific facts, numbers, endpoints, or disease-specific content.",
        "Use it only for format/style guidance while writing from the current study context.",
        `SPI611 Example Section: ${[spi611Match?.matched?.number || "", spi611Match?.matched?.title || ""].join(" ").trim()}`,
        "SPI611 Example Text:",
        spi611ExampleText
      ].join("\n")
    );
  }

  if (refineInstruction) {
    promptParts.push(
      `\nIMPORTANT - USER FEEDBACK:\n${refineInstruction}\n\nPlease regenerate the section content strictly following this feedback while maintaining formal SAP style.`
    );
  }

  const prompt = promptParts.join("\n\n");

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: "system", content: "You are a senior medical writer specialized in Clinical SAP documents." },
      { role: "user", content: prompt }
    ],
    temperature: 0.2
  });

  const generatedText = completion?.choices?.[0]?.message?.content?.trim() || "";
  if (!generatedText) {
    throw new Error("OpenAI returned empty content");
  }

  return {
    generatedText,
    protocolContextText: contextText,
    crfContext: {
      formName: crfTitleLine || null,
      fullText: crfContextText || null
    },
    matchedProtocolSection: {
      sectionIndex: best.sectionIndex,
      number: best.number || null,
      title: best.title || null,
      score: best.score,
      titleScore: best.titleScore ?? null,
      contentScore: best.contentScore ?? null,
      matchedBy:
        best.titleScore !== null && best.contentScore !== null
          ? best.titleScore >= best.contentScore
            ? "title"
            : "content"
          : best.titleScore !== null
            ? "title"
            : best.contentScore !== null
              ? "content"
              : null
    },
    matchedCrfForm: bestCrf
      ? {
          formKey: bestCrf.formKey,
          title: bestCrf.title,
          score: bestCrf.score,
          titleScore: bestCrf.titleScore,
          fullTextScore: bestCrf.fullScore,
          matchedBy:
            bestCrf.titleScore !== null && bestCrf.fullScore !== null
              ? bestCrf.titleScore >= bestCrf.fullScore
                ? "title"
                : "full_text"
              : bestCrf.titleScore !== null
                ? "title"
                : bestCrf.fullScore !== null
                  ? "full_text"
                  : null
        }
      : null,
    matchedSpi611Example: spi611Match?.matched
      ? {
          number: spi611Match.matched.number,
          title: spi611Match.matched.title,
          score: spi611Match.matched.score,
          titleEmbeddingModel: spi611Match.matched.titleEmbeddingModel || titleEmbeddingResult?.embeddingModel || null
        }
      : null,
    spi611ExampleText: spi611ExampleText || null
  };
}

module.exports = {
  generateSapSection
};

