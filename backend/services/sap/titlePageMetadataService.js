const OpenAI = require("openai");

// Extract protocol header metadata from protocol_extraction.sectionedText[0].content
async function generateTitlePageMetadata({ studyNumber, mongoose, openaiApiKey }) {
  if (!studyNumber) throw new Error("Missing studyNumber");
  if (!mongoose?.connection?.client) throw new Error("Database connection is not ready");

  const db = mongoose.connection.client.db("llxdocument");
  const studiesCol = db.collection("studies");
  const doc = await studiesCol.findOne(
    { studyNumber, chunkType: "protocol_extraction" },
    { projection: { data: 1, _id: 0 } }
  );

  const headerText = doc?.data?.sectionedText?.[0]?.content || "";
  if (!headerText) throw new Error("Protocol header content is empty");

  const client = new OpenAI({ apiKey: openaiApiKey });
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";

  // Ask GPT to return strict JSON with the four fields
  const prompt = [
    "Extract the following fields from the protocol header text:",
    "protocolTitle, protocolNumber, protocolVersionDate, documentVersionDate.",
    "Return ONLY JSON with those keys. If a field is missing, return an empty string.",
    "Header text:",
    headerText
  ].join("\n\n");

  const completion = await client.chat.completions.create({
    model,
    temperature: 0,
    messages: [
      { role: "system", content: "You extract structured metadata from clinical protocol headers." },
      { role: "user", content: prompt }
    ]
  });

  const raw = completion?.choices?.[0]?.message?.content || "";
  const jsonText = raw.match(/\{[\s\S]*\}/)?.[0] || "{}";
  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    throw new Error("Failed to parse metadata JSON from OpenAI response");
  }

  return {
    protocolTitle: parsed.protocolTitle || "",
    protocolNumber: parsed.protocolNumber || "",
    protocolVersionDate: parsed.protocolVersionDate || "",
    documentVersionDate: parsed.documentVersionDate || ""
  };
}

module.exports = {
  generateTitlePageMetadata
};
