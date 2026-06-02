#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { spawnSync } = require("child_process");

const backendNodeModules = path.join(__dirname, "..", "backend", "node_modules");
const dotenv = require(path.join(backendNodeModules, "dotenv"));
const { MongoClient } = require(path.join(backendNodeModules, "mongodb"));
const OpenAI = require(path.join(backendNodeModules, "openai"));

dotenv.config({ path: path.join(__dirname, "..", "backend", ".env") });

const SOURCE_DOCX = path.join(__dirname, "CMAA868D12302_Statistical Analysis Plan_AI.docx");
const SOURCE_RELATIVE_PATH = "example file/CMAA868D12302_Statistical Analysis Plan_AI.docx";
const REFERENCES_DB_NAME = "References";
const COLLECTION_NAME = "SAP_SPI611";
const DOC_KEY = "CMAA868D12302";
const ALIAS = "SPI-611";
const MONGODB_URI = process.env.AZURE_COSMOS_URI || process.env.MONGODB_URI;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const EMBEDDING_MODEL = process.env.OPENAI_EMBEDDING_MODEL || process.env.EMBEDDING_MODEL || "text-embedding-3-small";
const POWERSHELL_EXE = process.platform === "win32" ? "pwsh.exe" : "pwsh";
const DEBUG = String(process.env.DEBUG_SPI611 || "").toLowerCase() === "true";

function fail(message) {
  throw new Error(message);
}

function debugLog(...args) {
  if (DEBUG) {
    console.log("[DEBUG]", ...args);
  }
}

function parentSectionNumber(number) {
  const parts = String(number || "").split(".");
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join(".");
}

function hashPayload(value) {
  return crypto.createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function isValidEmbeddingVector(v) {
  if (!Array.isArray(v) || v.length === 0) return false;
  for (let i = 0; i < Math.min(v.length, 5); i++) {
    if (typeof v[i] !== "number" || !Number.isFinite(v[i])) return false;
  }
  return true;
}

function normalizeTitleForEmbedding(title) {
  return String(title || "")
    .replace(/\s+/g, " ")
    .trim();
}

function runDocxExtraction(docxPath) {
  const tempJsonPath = path.join(os.tmpdir(), `spi611-sap-${Date.now()}.json`);

  const psScript = `
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ParagraphText {
  param(
    [Parameter(Mandatory = $true)] $ParagraphNode,
    [Parameter(Mandatory = $true)] $NamespaceManager
  )

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($t in $ParagraphNode.SelectNodes(".//w:t", $NamespaceManager)) {
    if ($null -ne $t -and $null -ne $t.InnerText) {
      [void]$parts.Add($t.InnerText)
    }
  }

  $text = ($parts.ToArray() -join "")
  if ($null -eq $text) { return "" }
  return $text.Trim()
}

function Get-ParagraphStyleId {
  param(
    [Parameter(Mandatory = $true)] $ParagraphNode,
    [Parameter(Mandatory = $true)] $NamespaceManager
  )

  $styleNode = $ParagraphNode.SelectSingleNode("./w:pPr/w:pStyle", $NamespaceManager)
  if ($null -eq $styleNode) { return $null }

  foreach ($attr in $styleNode.Attributes) {
    if ($attr.LocalName -eq "val") {
      return $attr.Value
    }
  }

  return $null
}

$docxPath = $env:DOCX_PATH
$outPath = $env:OUT_PATH

if (-not (Test-Path -LiteralPath $docxPath)) {
  throw "DOCX file not found: $docxPath"
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("spi611-docx-" + [guid]::NewGuid().ToString("N"))
[System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

try {
  [System.IO.Compression.ZipFile]::ExtractToDirectory($docxPath, $tempDir)

  $documentXmlPath = Join-Path $tempDir "word/document.xml"
  if (-not (Test-Path -LiteralPath $documentXmlPath)) {
    throw "word/document.xml not found in DOCX"
  }

  [xml]$documentXml = Get-Content -LiteralPath $documentXmlPath
  $ns = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

  $body = $documentXml.SelectSingleNode("/w:document/w:body", $ns)
  if ($null -eq $body) {
    throw "DOCX body not found"
  }

  $blocks = New-Object System.Collections.Generic.List[object]

  foreach ($child in $body.ChildNodes) {
    if ($child.LocalName -ne "p") { continue }

    $text = Get-ParagraphText -ParagraphNode $child -NamespaceManager $ns
    if ([string]::IsNullOrWhiteSpace($text)) { continue }

    $styleId = Get-ParagraphStyleId -ParagraphNode $child -NamespaceManager $ns
    $blocks.Add([pscustomobject]@{
      styleId = $styleId
      text = $text
    }) | Out-Null
  }

  $payload = [pscustomobject]@{
    blocks = $blocks
  }

  $payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outPath -Encoding UTF8
}
finally {
  if (Test-Path -LiteralPath $tempDir) {
    Remove-Item -LiteralPath $tempDir -Recurse -Force
  }
}
`;

  const encodedCommand = Buffer.from(psScript, "utf16le").toString("base64");
  const result = spawnSync(
    POWERSHELL_EXE,
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encodedCommand],
    {
      encoding: "utf8",
      env: {
        ...process.env,
        DOCX_PATH: docxPath,
        OUT_PATH: tempJsonPath
      },
      maxBuffer: 20 * 1024 * 1024
    }
  );

  if (result.status !== 0) {
    const stderr = String(result.stderr || "").trim();
    const stdout = String(result.stdout || "").trim();
    fail(`DOCX extraction failed: ${stderr || stdout || `exit code ${result.status}`}`);
  }

  if (!fs.existsSync(tempJsonPath)) {
    fail("DOCX extraction did not produce an output JSON file.");
  }

  try {
    return JSON.parse(fs.readFileSync(tempJsonPath, "utf8"));
  } finally {
    try {
      fs.unlinkSync(tempJsonPath);
    } catch (e) {}
  }
}

function normalizeParagraphText(text) {
  return String(text || "")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function normalizeComparableText(text) {
  return normalizeParagraphText(text)
    .replace(/\s+/g, "")
    .toLowerCase();
}

function isTocStyle(styleId) {
  return /^TOC[1-9]$/i.test(String(styleId || ""));
}

function isLikelyTocLine(text) {
  const s = normalizeParagraphText(text);
  if (!s) return false;
  if (/^table of contents/i.test(s)) return true;
  if (/^list of abbreviations/i.test(s)) return true;
  if (/^list of tables/i.test(s)) return true;
  if (/^list of figures/i.test(s)) return true;
  if (/^references?\b/i.test(s)) return false;
  if (/^\d+(?:\.\d+)*\s+.+\s+\d+$/.test(s)) return true;
  if (/^\d+(?:\.\d+)*\s+[.\s]{2,}\s*\d+$/.test(s)) return true;
  return false;
}

function getStyleLevel(styleId, prefix) {
  const match = String(styleId || "").match(new RegExp(`^${prefix}(\\d+)$`, "i"));
  return match ? Number(match[1]) : null;
}

function detectPlainNumberedHeading(text) {
  const s = normalizeParagraphText(text);
  if (!s) return null;
  if (/^(table|figure)\b/i.test(s)) return null;

  const match = s.match(/^(\d+(?:\.\d+)+)\s+(.+)$/);
  if (!match) return null;

  return {
    number: match[1],
    title: match[2].trim(),
    level: match[1].split(".").length
  };
}

function isOfficialBodyHeading(block) {
  const styleId = String(block?.styleId || "");
  const text = normalizeParagraphText(block?.text);
  const level = getStyleLevel(styleId, "Heading");
  if (!level || level > 5) return false;
  if (!text) return false;
  if (/^(table|figure)\b/i.test(text)) return false;
  return true;
}

function extractTocEntries(blocks) {
  const entries = [];

  for (const block of Array.isArray(blocks) ? blocks : []) {
    const styleId = String(block?.styleId || "");
    const text = normalizeParagraphText(block?.text);
    if (!isTocStyle(styleId)) continue;
    if (!text) continue;
    if (/^table of contents/i.test(text)) continue;
    if (/^list of /i.test(text)) continue;

    entries.push({
      styleId,
      level: getStyleLevel(styleId, "TOC"),
      rawText: text
    });
  }

  return entries;
}

function extractNumberFromTocEntry(tocEntry, bodyHeadingText) {
  if (!tocEntry || !bodyHeadingText) return null;

  const level = tocEntry.level;
  const tocComparable = normalizeComparableText(tocEntry.rawText).replace(/\d+$/, "");
  const headingComparable = normalizeComparableText(bodyHeadingText);
  if (!tocComparable || !headingComparable) return null;
  if (!tocComparable.endsWith(headingComparable)) return null;

  const number = tocComparable.slice(0, tocComparable.length - headingComparable.length);
  if (!number) return null;

  const expectedPattern =
    level <= 1
      ? /^\d+$/
      : new RegExp(`^\\d+(?:\\.\\d+){${level - 1}}$`);
  if (!expectedPattern.test(number)) return null;

  return number;
}

function formatSubheadingText(block) {
  const text = normalizeParagraphText(block?.text);
  if (!text) return "";

  if (isOfficialBodyHeading(block)) {
    const level = getStyleLevel(block.styleId, "Heading");
    if (level && level >= 4) {
      return text;
    }
  }

  return text;
}

function inferNumberFromHeadingLevel(counters, level) {
  const next = Array.isArray(counters) ? counters.slice() : [];
  while (next.length < level) {
    next.push(0);
  }
  next[level - 1] = (next[level - 1] || 0) + 1;
  next.length = level;
  return next;
}

function buildSapStateFromBlocks(blocks) {
  const tocEntries = extractTocEntries(blocks);
  const entries = [];
  const sections = [];
  const frontMatter = [];

  let current = null;
  let counters = [];
  let tocIndex = 0;

  for (const block of Array.isArray(blocks) ? blocks : []) {
    if (isTocStyle(block?.styleId)) continue;

    const text = normalizeParagraphText(block?.text);
    if (!text) continue;

    const plainNumbered = detectPlainNumberedHeading(text);
    const headingLevel = isOfficialBodyHeading(block) ? getStyleLevel(block?.styleId, "Heading") : null;

    if (plainNumbered || headingLevel) {
      let title = text;
      let level = headingLevel;
      let number;

      if (plainNumbered) {
        number = plainNumbered.number;
        title = plainNumbered.title;
        level = plainNumbered.level;
        counters = number.split(".").map((part) => Number(part));
      } else {
        counters = inferNumberFromHeadingLevel(counters, level);
        number = counters.join(".");
      }

      const tocEntry = tocEntries[tocIndex] || null;
      const tocNumber = extractNumberFromTocEntry(tocEntry, title);
      if (tocNumber) {
        tocIndex += 1;
      }

      debugLog("headingCandidate", {
        tocIndex,
        tocText: tocEntry?.rawText || null,
        bodyText: title,
        bodyStyle: block?.styleId || null,
        inferredNumber: number,
        matchedTocNumber: tocNumber || null
      });

      if (current) {
        sections.push({
          sectionNumber: current.sectionNumber,
          title: current.title,
          text: current.paragraphs.join("\n\n").trim()
        });
      }

      const order = entries.length + 1;
      entries.push({
        order,
        number,
        displayNumber: number,
        title,
        level,
        parentNumber: parentSectionNumber(number),
        title_embedded: false,
        tocNumber: tocNumber || null
      });

      current = {
        sectionNumber: number,
        title,
        paragraphs: []
      };
      continue;
    }

    if (current) {
      const contentText = formatSubheadingText(block);
      if (contentText) {
        current.paragraphs.push(contentText);
      }
    } else if (!isLikelyTocLine(text)) {
      frontMatter.push(text);
    }
  }

  if (current) {
    sections.push({
      sectionNumber: current.sectionNumber,
      title: current.title,
      text: current.paragraphs.join("\n\n").trim()
    });
  }

  return {
    schemaVersion: 1,
    saved: true,
    lastSavedAt: new Date().toISOString(),
    entries,
    sections,
    frontMatterText: frontMatter.join("\n\n").trim()
  };
}

async function enrichStateWithTitleEmbeddings(state, existingState) {
  const incomingEntries = Array.isArray(state?.entries) ? state.entries : [];
  if (incomingEntries.length === 0) return state;
  if (!OPENAI_API_KEY) {
    fail("Missing OPENAI_API_KEY in backend/.env. Cannot generate title embeddings for SPI-611 upload.");
  }

  const existingEntries = Array.isArray(existingState?.entries) ? existingState.entries : [];
  const existingByNumber = new Map();
  for (const entry of existingEntries) {
    const number = String(entry?.number || "").trim();
    if (!number) continue;
    existingByNumber.set(number, entry);
  }

  const toEmbed = [];
  const enrichedEntries = incomingEntries.map((entry, idx) => {
    const number = String(entry?.number || "").trim();
    const title = String(entry?.title || "").trim();
    const out = { ...(entry || {}) };
    if (!number || !title) return out;

    const cached = existingByNumber.get(number);
    if (
      cached &&
      String(cached?.title || "").trim() === title &&
      cached?.title_embedded === true &&
      isValidEmbeddingVector(cached?.title_embedding)
    ) {
      out.title_embedded = true;
      out.title_embedding = cached.title_embedding;
      out.title_embedding_model = cached.title_embedding_model || EMBEDDING_MODEL;
      return out;
    }

    out.title_embedded = false;
    delete out.title_embedding;
    delete out.title_embedding_model;
    toEmbed.push({ idx, titleNorm: normalizeTitleForEmbedding(title) });
    return out;
  });

  if (toEmbed.length === 0) {
    return { ...state, entries: enrichedEntries };
  }

  const uniqueTitles = [];
  const titleToIndex = new Map();
  for (const item of toEmbed) {
    if (titleToIndex.has(item.titleNorm)) continue;
    titleToIndex.set(item.titleNorm, uniqueTitles.length);
    uniqueTitles.push(item.titleNorm);
  }

  const client = new OpenAI({ apiKey: OPENAI_API_KEY });
  const resp = await client.embeddings.create({
    model: EMBEDDING_MODEL,
    input: uniqueTitles
  });

  const indexToEmbedding = new Map();
  for (const item of resp?.data || []) {
    indexToEmbedding.set(item.index, item.embedding);
  }
  if (indexToEmbedding.size !== uniqueTitles.length) {
    fail(`Embedding count mismatch: expected ${uniqueTitles.length}, got ${indexToEmbedding.size}`);
  }

  for (const item of toEmbed) {
    const emb = indexToEmbedding.get(titleToIndex.get(item.titleNorm));
    if (!isValidEmbeddingVector(emb)) continue;
    enrichedEntries[item.idx].title_embedded = true;
    enrichedEntries[item.idx].title_embedding = emb;
    enrichedEntries[item.idx].title_embedding_model = EMBEDDING_MODEL;
  }

  return { ...state, entries: enrichedEntries };
}

async function uploadStateToMongo(state) {
  if (!MONGODB_URI) {
    fail("Missing AZURE_COSMOS_URI or MONGODB_URI in backend/.env");
  }

  const contentHash = hashPayload({
    entries: state.entries,
    sections: state.sections,
    frontMatterText: state.frontMatterText
  });

  const client = new MongoClient(MONGODB_URI, {
    retryWrites: false,
    maxIdleTimeMS: 120000,
    serverSelectionTimeoutMS: 30000,
    socketTimeoutMS: 360000
  });

  await client.connect();

  try {
    const db = client.db(REFERENCES_DB_NAME);
    const collection = db.collection(COLLECTION_NAME);

    await collection.createIndex({ docType: 1, docKey: 1 }, { unique: true });

    const existingDoc = await collection.findOne(
      { docType: "state", docKey: DOC_KEY },
      { projection: { "data.entries": 1, _id: 0 } }
    );

    const enrichedState = await enrichStateWithTitleEmbeddings(state, existingDoc?.data || null);

    const sourceStat = fs.statSync(SOURCE_DOCX);
    const now = new Date();
    const embeddedEntryCount = (enrichedState.entries || []).filter((e) => e?.title_embedded === true).length;
    const contentHash = hashPayload({
      entries: enrichedState.entries,
      sections: enrichedState.sections,
      frontMatterText: enrichedState.frontMatterText
    });

    const result = await collection.updateOne(
      { docType: "state", docKey: DOC_KEY },
      {
        $set: {
          alias: ALIAS,
          studyNumber: DOC_KEY,
          docType: "state",
          source: {
            fileName: path.basename(SOURCE_DOCX),
            relativePath: SOURCE_RELATIVE_PATH,
            fileSize: sourceStat.size,
            uploadedAt: now
          },
          data: enrichedState,
          entryCount: enrichedState.entries.length,
          sectionCount: enrichedState.sections.length,
          embeddedEntryCount,
          embeddingModel: EMBEDDING_MODEL,
          contentHash,
          updatedAt: now
        },
        $setOnInsert: {
          createdAt: now
        }
      },
      { upsert: true }
    );

    return {
      result,
      enrichedState,
      embeddedEntryCount
    };
  } finally {
    await client.close();
  }
}

async function main() {
  if (!fs.existsSync(SOURCE_DOCX)) {
    fail(`Source DOCX not found: ${SOURCE_DOCX}`);
  }

  console.log("============================================================");
  console.log("Uploading SPI-611 SAP example to References.SAP_SPI611");
  console.log("============================================================");
  console.log(`Source file: ${SOURCE_DOCX}`);

  const extraction = runDocxExtraction(SOURCE_DOCX);
  const state = buildSapStateFromBlocks(extraction?.blocks || []);
  debugLog("summary", {
    blockCount: Array.isArray(extraction?.blocks) ? extraction.blocks.length : 0,
    entryCount: state.entries.length,
    sectionCount: state.sections.length
  });

  if (!Array.isArray(state.entries) || state.entries.length === 0) {
    fail("No numbered SAP sections were extracted from the DOCX.");
  }

  const upload = await uploadStateToMongo(state);
  const result = upload.result;
  const enrichedState = upload.enrichedState;

  console.log(`Entries extracted : ${enrichedState.entries.length}`);
  console.log(`Sections stored   : ${enrichedState.sections.length}`);
  console.log(`Embedded titles   : ${upload.embeddedEntryCount}`);
  console.log(`Embedding model   : ${EMBEDDING_MODEL}`);
  console.log(`Front matter chars: ${enrichedState.frontMatterText.length}`);
  console.log(`Mongo matchedCount: ${result.matchedCount}`);
  console.log(`Mongo modifiedCount: ${result.modifiedCount}`);
  console.log(`Mongo upsertedId  : ${result.upsertedId || "existing document updated"}`);
  console.log("Done.");
}

main().catch((err) => {
  console.error("Upload failed:", err);
  process.exit(1);
});
