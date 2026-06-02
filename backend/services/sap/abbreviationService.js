const OpenAI = require("openai");

// Find the most likely "List of Abbreviations" section among level=1 headers.
async function findAbbreviationSection(sectionedText, client) {
  const levelOne = (sectionedText || [])
    .map((s, idx) => ({ idx, title: s?.title || "", level: s?.level }))
    .filter((s) => s.title && s.level === 1)
    .slice(0, 20);

  if (levelOne.length === 0) return null;

  const normalized = (t) => String(t).toUpperCase().replace(/[^A-Z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
  const direct = levelOne.find((s) => {
    const t = normalized(s.title);
    return t.includes("LIST OF ABBREVIATIONS") || t === "ABBREVIATIONS" || t.includes("ABBREVIATIONS");
  });
  if (direct) return direct.idx;

  // GPT fallback: ask the model to select the best index
  const prompt = [
    "Select the index of the title that most likely means 'List of Abbreviations'.",
    "If none match, return null.",
    "Titles:",
    JSON.stringify(levelOne.map((s) => ({ index: s.idx, title: s.title })))
  ].join("\n\n");

  const completion = await client.chat.completions.create({
    model: process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini",
    temperature: 0,
    messages: [
      { role: "system", content: "You select the best matching title index." },
      { role: "user", content: prompt }
    ]
  });

  const raw = completion?.choices?.[0]?.message?.content || "";
  const jsonText = raw.match(/\{[\s\S]*\}/)?.[0] || "";
  if (jsonText) {
    try {
      const parsed = JSON.parse(jsonText);
      if (Number.isInteger(parsed?.bestIndex)) return parsed.bestIndex;
    } catch (e) {}
  }

  // Fallback: try to parse a plain number from the model output
  const num = raw.match(/\d+/);
  return num ? Number(num[0]) : null;
}

// Extract abbreviations into {term, definition}[] from raw content using GPT.
async function extractAbbreviations(content, client) {
  const prompt = [
    "Extract a list of abbreviations and their explanations from the text.",
    "Return ONLY JSON with this shape:",
    `{ "items": [ { "term": "AE", "definition": "Adverse Event" } ] }`,
    "If a line does not look like a term-definition pair, skip it.",
    "Do not invent terms. Use only the provided text.",
    "Text:",
    content
  ].join("\n\n");

  const completion = await client.chat.completions.create({
    model: process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini",
    temperature: 0,
    messages: [
      { role: "system", content: "You extract structured abbreviation lists from clinical documents." },
      { role: "user", content: prompt }
    ]
  });

  const raw = completion?.choices?.[0]?.message?.content || "";
  const jsonText = raw.match(/\{[\s\S]*\}/)?.[0] || "{}";
  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    throw new Error("Failed to parse abbreviations JSON from OpenAI response");
  }

  return Array.isArray(parsed?.items) ? parsed.items : [];
}

async function generateAbbreviations({ studyNumber, mongoose, openaiApiKey }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const db = mongoose.connection.client.db("llxdocument");
  const studiesCol = db.collection("studies");
  const doc = await studiesCol.findOne(
    { studyNumber, chunkType: "protocol_extraction" },
    { projection: { data: 1, _id: 0 } }
  );

  const sectionedText = doc?.data?.sectionedText || [];
  const client = new OpenAI({ apiKey: openaiApiKey });

  const bestIndex = await findAbbreviationSection(sectionedText, client);
  if (bestIndex === null || bestIndex === undefined) {
    throw new Error("List of abbreviations section not found");
  }

  const content = sectionedText[bestIndex]?.content || "";
  if (!content) {
    throw new Error("Abbreviations section content is empty");
  }

  return await extractAbbreviations(content, client);
}

module.exports = {
  generateAbbreviations
};
