const fs = require("fs");
const path = require("path");
const mongoose = require("mongoose");
const OpenAI = require("openai");
require("dotenv").config();

const MONGODB_URI = process.env.AZURE_COSMOS_URI || process.env.MONGODB_URI;
const DEFAULT_DB_NAME = process.env.MONGODB_DB_NAME || process.env.DB_NAME || "llxdocument";
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const EMBEDDING_MODEL = process.env.OPENAI_EMBEDDING_MODEL || "text-embedding-3-small";

const SOURCE_FILE = path.join(
  __dirname,
  "..",
  "..",
  "Resource",
  "MockTFL example",
  "tableofcontent_example.txt"
);

function parseMockTflEntries(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const entries = [];
  lines.forEach((line, idx) => {
    const normalized = line.replace(/^\*\s*/, "").trim();
    const match = normalized.match(/^(TABLE|FIGURE|LISTING)\s+([0-9.]+)\s+(.+)$/i);
    if (!match) return;
    entries.push({
      order: idx + 1,
      type: String(match[1] || "").toUpperCase(),
      number: String(match[2] || "").trim(),
      title: String(match[3] || "").trim()
    });
  });
  return entries;
}

async function buildEmbedding(client, entry) {
  const input = `${entry.type} ${entry.number} ${entry.title}`.trim();
  const response = await client.embeddings.create({
    model: EMBEDDING_MODEL,
    input
  });
  return response?.data?.[0]?.embedding || null;
}

async function run() {
  if (!MONGODB_URI) {
    throw new Error("Missing AZURE_COSMOS_URI or MONGODB_URI");
  }
  if (!OPENAI_API_KEY) {
    throw new Error("Missing OPENAI_API_KEY");
  }
  if (!fs.existsSync(SOURCE_FILE)) {
    throw new Error(`Source file not found: ${SOURCE_FILE}`);
  }

  const raw = fs.readFileSync(SOURCE_FILE, "utf8");
  const entries = parseMockTflEntries(raw);
  if (entries.length === 0) {
    throw new Error("No valid MockTFL entries parsed from source file");
  }

  await mongoose.connect(MONGODB_URI, {
    retryWrites: false,
    maxIdleTimeMS: 120000,
    serverSelectionTimeoutMS: 30000,
    socketTimeoutMS: 360000,
    dbName: DEFAULT_DB_NAME
  });

  const client = new OpenAI({ apiKey: OPENAI_API_KEY });
  const col = mongoose.connection.client.db("References").collection("MockTFL_Example");

  let upserted = 0;
  for (const entry of entries) {
    const embedding = await buildEmbedding(client, entry);
    await col.updateOne(
      { type: entry.type, number: entry.number },
      {
        $set: {
          order: entry.order,
          type: entry.type,
          number: entry.number,
          title: entry.title,
          embedding,
          updatedAt: new Date()
        },
        $setOnInsert: {
          createdAt: new Date()
        }
      },
      { upsert: true }
    );
    upserted += 1;
    console.log(`[${upserted}/${entries.length}] upserted ${entry.type} ${entry.number}`);
  }

  console.log(`Done. Upserted ${upserted} MockTFL entries into References.MockTFL_Example.`);
}

run()
  .catch((err) => {
    console.error("Ingest failed:", err);
    process.exitCode = 1;
  })
  .finally(async () => {
    try {
      await mongoose.disconnect();
    } catch (e) {}
  });
