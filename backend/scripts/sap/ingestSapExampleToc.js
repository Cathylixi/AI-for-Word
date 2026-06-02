/**
 * SAP Example TOC Upload Script (Option A: one TOC entry = one document)
 *
 * Target:
 * - Database: References
 * - Collection: SAP_Example
 *
 * What will be stored:
 * - order (1..N)
 * - number (e.g., "1", "1.1", "5.2.7")
 * - title  (e.g., "Introduction")
 *
 * What will NOT be stored:
 * - page number
 * - main text
 *
 * How to run (PowerShell):
 * 1) Make sure you have backend/.env with AZURE_COSMOS_URI (preferred) or MONGODB_URI
 * 2) From AI-for-Word/backend:
 *    node ".\\resource\\sap example\\sap example upload.js"
 */

const path = require("path");
const mongoose = require("mongoose");
const OpenAI = require("openai");

// Load env from AI-for-Word/backend/.env (if present)
try {
  require("dotenv").config({ path: path.join(__dirname, "..", "..", ".env") });
} catch (e) {
  // ignore
}

const REFERENCES_DB_NAME = "References";
const COLLECTION_NAME = "SAP_Example";
const SOURCE = "SAP Example TOC";

const MONGODB_URI = process.env.AZURE_COSMOS_URI || process.env.MONGODB_URI;
const DEFAULT_DB_NAME = process.env.MONGODB_DB_NAME || process.env.DB_NAME || "llxdocument";
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const EMBEDDING_MODEL = process.env.EMBEDDING_MODEL || "text-embedding-3-small";
const GENERATE_EMBEDDINGS = String(process.env.GENERATE_EMBEDDINGS || "true").toLowerCase() === "true";

function die(msg) {
  console.error(`❌ ${msg}`);
  process.exit(1);
}

function parseTocLines(raw) {
  const lines = raw
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const entries = [];
  const warnings = [];

  // Example line formats we accept:
  // 1 Introduction 11
  // 1.2.1 Primary estimand(s) 12
  // 7 (Other) Safety analyses 22
  const re = /^(\d+(?:\.\d+)*)\s+(.+?)\s+(\d+)\s*$/;

  for (const line of lines) {
    // We only store numbered sections. Skip non-numbered lines like:
    // "Table of contents 7", "List of abbreviations 10"
    if (!/^\d/.test(line)) continue;

    const m = line.match(re);
    if (!m) {
      warnings.push(`Unparsed line (skipped): "${line}"`);
      continue;
    }

    const number = m[1];
    const title = m[2].trim();

    entries.push({ number, title });
  }

  // add order
  const docs = entries.map((e, idx) => ({
    source: SOURCE,
    order: idx + 1,
    number: e.number,
    title: e.title,
    createdAt: new Date()
  }));

  return { docs, warnings };
}

async function main() {
  if (!MONGODB_URI) {
    die("Missing AZURE_COSMOS_URI or MONGODB_URI in environment.");
  }
  if (GENERATE_EMBEDDINGS && !OPENAI_API_KEY) {
    die("GENERATE_EMBEDDINGS=true but OPENAI_API_KEY is missing.");
  }

  const RAW_TOC = `
Table of contents  7
List of abbreviations  10
1  Introduction  11
1.1  Study design  11
1.2  Study objectives, endpoints and estimands  11
1.2.1  Primary estimand(s)  12
1.2.2  Secondary estimand(s)  12
2  Analysis sets  12
2.1  Subgroup of interest  13
3  Statistical analyses  13
3.1  General considerations  13
3.1.1  General definitions  14
3.2  Participant demographics and other baseline characteristics  14
3.2.1  Patient disposition  14
3.2.2  Demographics and other baseline characteristics  14
3.3  Treatments  14
3.3.1  Study treatment / compliance  14
3.3.2  Prior, concomitant and post therapies  15
4  Primary endpoint(s) / estimand(s) analysis  15
4.1  Definition of Primary endpoint(s) / estimands  16
4.2  Statistical model, hypothesis, and method of analysis  16
4.3  Handling of intercurrent events of primary estimand (if applicable)  16
4.4  Handling of missing values not related to intercurrent event  17
4.5  Multiplicity adjustment (if applicable)  17
4.6  Sensitivity analyses  17
4.7  Supplementary analyses  18
5  Secondary endpoint(s) / estimand(s) analysis  18
5.1  Efficacy and/or pharmacodynamic endpoint(s)  19
5.1.1  Statistical hypothesis, model, and method of analysis  19
5.1.2  Handling of intercurrent events of secondary estimand (if applicable)  19
5.1.3  Handling of missing values not related to intercurrent event  19
5.1.4  Multiplicity adjustment (if applicable)  19
5.1.5  Sensitivity analyses  19
5.1.6  Supplementary analyses  19
5.2  Safety endpoints (which are Secondary endpoint(s) / estimand(s))  19
5.2.1  Adverse events (which are Secondary)  19
5.2.2  Vital Signs (which are Secondary)  20
5.2.3  12-Lead ECG (which are Secondary)  20
5.2.4  Clinical Laboratory Evaluations (which are Secondary)  20
5.2.5  Other Safety Evaluations (which are Secondary)  20
5.2.6  Immunogenicity (which are Secondary)  20
5.2.7  Resource Utilization  20
5.3  Pharmacokinetics (which are Secondary)  20
5.4  DNA (which are Secondary)  21
5.5  Biomarkers (which are Secondary)  21
5.6  PK/PD relationships (which are Secondary)  21
5.7  Patient reported outcomes (which are Secondary)  21
6  Exploratory endpoint(s) / estimands(s) analysis  22
7  (Other) Safety analyses  22
7.1  Adverse Events  22
7.1.1  AE Adverse events of special interest / grouping of AEs  23
7.2  Deaths  23
7.3  Vital Signs  23
7.4  12-Lead ECG  23
7.5  Clinical Laboratory Evaluations  23
7.6  Other Safety Evaluations  24
7.7  Immunogenicity  24
8  Other analyses  24
8.1  Pharmacokinetics  24
8.2  DNA  24
8.3  Biomarkers  24
8.4  PK/PD relationships  25
8.5  Patient reported outcomes  25
9  Interim analysis  25
10  Sample size determination  26
10.1  Primary endpoint(s)  26
10.2  Secondary endpoint(s)  26
11  Change to protocol specified analyses  26
12  Appendix  26
12.1  Imputation rules  26
12.1.1  Study drug  26
12.1.2  AE date imputation  27
12.1.3  Concomitant medication date imputation  27
12.2  AEs coding/grading  27
12.3  Laboratory parameters derivations  27
12.4  Statistical models  27
12.4.1  Analysis supporting primary objective(s)  28
12.4.2  Analysis supporting secondary objective(s)  28
12.5  Rule of exclusion criteria of analysis sets  28
13  Reference  29
`;

  const { docs, warnings } = parseTocLines(RAW_TOC);
  if (docs.length === 0) die("Parsed 0 TOC entries. Please check the RAW_TOC format.");

  console.log("============================================================");
  console.log("🚀 Uploading SAP Example TOC");
  console.log("============================================================");
  console.log(`Docs to insert: ${docs.length}`);
  console.log(`Embeddings: ${GENERATE_EMBEDDINGS ? `enabled (${EMBEDDING_MODEL})` : "disabled"}`);
  if (warnings.length > 0) {
    console.log(`Warnings: ${warnings.length}`);
    warnings.slice(0, 10).forEach((w) => console.log(`  - ${w}`));
    if (warnings.length > 10) console.log("  - ...");
  }

  // 1) Generate embeddings (batch) before DB writes
  if (GENERATE_EMBEDDINGS) {
    const client = new OpenAI({ apiKey: OPENAI_API_KEY });
    const inputs = docs.map((d) => `${d.number} ${d.title}`);

    console.log("Generating embeddings...");
    const resp = await client.embeddings.create({
      model: EMBEDDING_MODEL,
      input: inputs
    });

    const byIndex = new Map();
    for (const item of resp.data || []) {
      byIndex.set(item.index, item.embedding);
    }

    if (byIndex.size !== docs.length) {
      die(`Embedding count mismatch: expected ${docs.length}, got ${byIndex.size}`);
    }

    for (let i = 0; i < docs.length; i++) {
      docs[i].text = inputs[i];
      docs[i].embeddingModel = EMBEDDING_MODEL;
      docs[i].embedding = byIndex.get(i);
    }
    console.log("Embeddings generated.");
  }

  // Cosmos DB friendly options
  const mongooseOptions = {
    retryWrites: false,
    maxIdleTimeMS: 120000,
    serverSelectionTimeoutMS: 30000,
    socketTimeoutMS: 360000,
    dbName: DEFAULT_DB_NAME
  };

  await mongoose.connect(MONGODB_URI, mongooseOptions);

  try {
    const db = mongoose.connection.client.db(REFERENCES_DB_NAME);
    const collection = db.collection(COLLECTION_NAME);

    // Ensure index for stable querying
    await collection.createIndex({ source: 1, order: 1 }, { unique: true });
    await collection.createIndex({ source: 1, number: 1 }, { unique: true });

    // Replace existing SAP Example TOC docs only (safe)
    const del = await collection.deleteMany({ source: SOURCE });
    console.log(`Deleted existing docs with source="${SOURCE}": ${del.deletedCount}`);

    const result = await collection.insertMany(docs, { ordered: true });
    console.log(`Inserted: ${result.insertedCount}`);

    const count = await collection.countDocuments({ source: SOURCE });
    console.log(`Collection count for source="${SOURCE}": ${count}`);
  } finally {
    await mongoose.disconnect();
  }

  console.log("✅ Done.");
}

main().catch((err) => {
  console.error("❌ Script failed:", err);
  process.exit(1);
});

