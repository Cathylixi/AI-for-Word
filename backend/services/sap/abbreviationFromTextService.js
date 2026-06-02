const OpenAI = require("openai");

async function extractAbbreviationsFromText({ text, existingTerms = [], openaiApiKey }) {
  if (!text || !String(text).trim()) return [];
  const client = new OpenAI({ apiKey: openaiApiKey });

  const normalized = Array.isArray(existingTerms)
    ? existingTerms.map((t) => String(t || "").trim()).filter(Boolean)
    : [];

  const promptParts = [
    "Identify professional academic terms, medical phrases, or study-specific concepts in the text that are commonly abbreviated in Clinical SAP documents.",
    "For each identified phrase, provide its standard **Abbreviation** as the 'term' and the **Original Full Phrase** as the 'definition'.",
    "Return ONLY JSON with this shape:",
    `{ "items": [ { "term": "AE", "definition": "Adverse Event" }, { "term": "ITT", "definition": "Intent-to-Treat" } ] }`,
    "Rules:",
    "1. The 'term' MUST be the short abbreviation/acronym (e.g., 'PK', 'ORR').",
    "2. The 'definition' MUST be the full original phrase (e.g., 'Pharmacokinetics', 'Objective Response Rate').",
    "3. If the text contains the full phrase but not the abbreviation, generate the standard abbreviation for it.",
    "4. If the text contains only the abbreviation, provide the full phrase if it is standard and clear.",
    "5. Exclude common non-technical English words.",
    "Do not invent non-standard abbreviations. Use only standard clinical/medical abbreviations or those explicitly defined in the text."
  ];

  if (normalized.length > 0) {
    promptParts.push(
      "Do NOT return any term already in this list:",
      JSON.stringify(normalized)
    );
  }

  promptParts.push("Text:", text);

  const completion = await client.chat.completions.create({
    model: process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini",
    temperature: 0,
    messages: [
      { role: "system", content: "You extract structured abbreviation lists from clinical documents." },
      { role: "user", content: promptParts.join("\n\n") }
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

module.exports = {
  extractAbbreviationsFromText
};
