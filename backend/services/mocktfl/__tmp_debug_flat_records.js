const https = require("https");
const path = require("path");
require("dotenv").config({ path: path.join(__dirname, "../../.env") });

const mongoose = require("mongoose");
const OpenAI = require("openai");
const { loadCorrespondingTableInfo, __test } = require("./mockTflFigureAutoService");

function escapeRString(value) {
  return String(value == null ? "" : value)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"');
}

function postRaw(url, payload, timeoutMs = 120000) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const body = JSON.stringify(payload || {});
    const req = https.request(
      parsedUrl,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Content-Length": Buffer.byteLength(body)
        },
        timeout: timeoutMs
      },
      (res) => {
        let raw = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => {
          raw += chunk;
        });
        res.on("end", () => resolve({ statusCode: res.statusCode, raw }));
      }
    );
    req.on("timeout", () => req.destroy(new Error("request timed out")));
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

async function main() {
  const studyNumber = process.env.DEBUG_STUDY_NUMBER || "SPI-611";
  const pureTitle =
    process.env.DEBUG_FIGURE_TITLE ||
    "Cumulative Incidence Plot for Primary Efficacy Endpoint Time to First Composite of VTE Recurrence Event Full Analysis Set";
  const figureType = process.env.DEBUG_FIGURE_TYPE || "cumulative_incidence";
  const correspondingTableType = process.env.DEBUG_TABLE_TYPE || "TABLE";
  const correspondingTableNumber = process.env.DEBUG_TABLE_NUMBER || "14.2.1.1";

  await mongoose.connect(process.env.AZURE_COSMOS_URI, { dbName: "llxdocument" });
  const tableInfo = await loadCorrespondingTableInfo({
    studyNumber,
    tableType: correspondingTableType,
    tableNumber: correspondingTableNumber,
    mongoose
  });

  console.log("tableInfo:", {
    found: !!tableInfo?.found,
    label: tableInfo?.label,
    columnGroups: tableInfo?.columnHeaderConfig?.columnGroups
  });

  const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  const model = process.env.OPENAI_CHAT_MODEL || "gpt-4o-mini";
  const spec = await __test.createSpecByAi({
    pureTitle,
    drug1: "",
    drug2: "",
    tableInfo,
    figureType,
    client,
    model
  });

  const records = Array.isArray(spec?.dataJson?.records) ? spec.dataJson.records : [];
  console.log("ai-json-summary:", {
    variablesFromStep1: spec.variables,
    dataJsonVariables: spec?.dataJson?.variables,
    recordsCount: records.length,
    recordColumns: records[0] ? Object.keys(records[0]) : [],
    firstRecord: records[0] || null
  });

  const prettyJson = JSON.stringify(spec.dataJson, null, 2);
  const singleLineJson = JSON.stringify(spec.dataJson);
  const aiBody = await __test.generateRCodeByAi({
    jsonDataText: prettyJson,
    figureType,
    client,
    model
  });
  const generatedRCode = `library(jsonlite)\njson_data <- "${escapeRString(singleLineJson)}"\n\n${aiBody}`;
  console.log("generated-r-code-start:\n" + generatedRCode.slice(0, 8000));
  console.log("generated-r-code-end");

  const remoteUrl = process.env.MOCKTFL_REMOTE_R_RUNNER_URL;
  const remote = await postRaw(remoteUrl, { rCodeText: generatedRCode });
  console.log("remote-status:", remote.statusCode);
  console.log("remote-body:", remote.raw);
}

main()
  .catch((err) => {
    console.error("debug-error:", err);
    process.exitCode = 1;
  })
  .finally(async () => {
    await mongoose.disconnect().catch(() => {});
  });