/* global Word */

import { insertSapHeader } from "./sapHeader";

function applyHeadingStyle(paragraph, level) {
  try {
    if (Word?.StyleBuiltIn) {
      const map = {
        1: Word.StyleBuiltIn.heading1,
        2: Word.StyleBuiltIn.heading2,
        3: Word.StyleBuiltIn.heading3
      };
      paragraph.styleBuiltIn = map[level] || Word.StyleBuiltIn.heading1;
      return;
    }
  } catch (e) {}
  try {
    paragraph.style = `Heading ${level}`;
  } catch (e) {}
}

function tryRemoveTableBorders(table) {
  try {
    const borders = [
      "Top",
      "Bottom",
      "Left",
      "Right",
      "InsideHorizontal",
      "InsideVertical"
    ];
    borders.forEach((b) => {
      const border = table.getBorder(b);
      border.type = "None";
    });
  } catch (e) {
    // If borders API is not available, leave default borders as-is.
  }
}

const META_FIELDS = [
  { key: "protocolTitle", label: "Protocol Title:" },
  { key: "protocolNumber", label: "Protocol Number:" },
  { key: "protocolVersionDate", label: "Protocol Version, Date:" },
  { key: "documentVersionDate", label: "Document Version, Date:" }
];

function insertTitlePage(body, studyNumber) {
  // Standard SAP header (title + full-width line)
  insertSapHeader(body);

  // Protocol metadata (labels + empty text boxes with spacer rows)
  const metaRows = [];
  META_FIELDS.forEach((f, idx) => {
    metaRows.push([f.label, ""]);
    if (idx < META_FIELDS.length - 1) metaRows.push(["", ""]);
  });

  const metaTable = body.insertTable(metaRows.length, 2, Word.InsertLocation.end, metaRows);
  tryRemoveTableBorders(metaTable);

  let fieldIndex = 0;
  for (let i = 0; i < metaRows.length; i++) {
    const leftCell = metaTable.getCell(i, 0);
    const rightCell = metaTable.getCell(i, 1);
    const isSpacer = metaRows[i][0] === "" && metaRows[i][1] === "";

    if (isSpacer) {
      leftCell.body.getRange().font.set({ name: "Arial", bold: false, size: 10 });
      rightCell.body.getRange().font.set({ name: "Times New Roman", bold: false, size: 10 });
      continue;
    }

    const field = META_FIELDS[fieldIndex++];
    leftCell.body.getRange().font.set({
      name: "Arial",
      bold: true,
      size: 10
    });

    const rightRange = rightCell.body.getRange();
    rightRange.font.set({ name: "Times New Roman", size: 10, bold: false });
    const cc = rightRange.insertContentControl();
    cc.appearance = "BoundingBox";
    cc.placeholderText = " ";
    if (studyNumber && field?.key) {
      cc.tag = `sap-meta:${studyNumber}:${field.key}`;
    }
  }

  // Spacing
  body.insertParagraph("", Word.InsertLocation.end);

  // Signature page title
  const sigTitle = body.insertParagraph("SIGNATURE PAGE", Word.InsertLocation.end);
  applyHeadingStyle(sigTitle, 1);
  sigTitle.font.set({ name: "Arial", bold: true, size: 14, color: "#000000" });

  body.insertParagraph("", Word.InsertLocation.end);

  // Signature block (simple two-column layout)
  const sigRows = [
    ["______________________________", "____________________"],
    ["Name / Title", "Date (DD Mmm YYYY)"]
  ];
  const sigTable = body.insertTable(2, 2, Word.InsertLocation.end, sigRows);
  tryRemoveTableBorders(sigTable);

  // Format signature block
  for (let r = 0; r < 2; r++) {
    for (let c = 0; c < 2; c++) {
      sigTable.getCell(r, c).body.getRange().font.set({
        name: "Times New Roman",
        bold: false,
        size: 12
      });
    }
  }

  // Note: Page breaks and revision history are inserted by the template composer.
}

async function writeTitlePageFields({ studyNumber, values }) {
  if (!studyNumber) return;
  await Word.run(async (context) => {
    for (const field of META_FIELDS) {
      const tag = `sap-meta:${studyNumber}:${field.key}`;
      const ccs = context.document.contentControls.getByTag(tag);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) continue;
      const cc = ccs.items[0];
      cc.insertText(values?.[field.key] || "", "Replace");
      cc.getRange().font.set({ name: "Times New Roman", size: 10, bold: false });
    }
    await context.sync();
  });
}

async function readTitlePageFields({ studyNumber }) {
  if (!studyNumber) return {};
  return await Word.run(async (context) => {
    const result = {};
    for (const field of META_FIELDS) {
      const tag = `sap-meta:${studyNumber}:${field.key}`;
      const ccs = context.document.contentControls.getByTag(tag);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length > 0) {
        const cc = ccs.items[0];
        const range = cc.getRange();
        range.load("text");
        await context.sync();
        result[field.key] = range.text || "";
      } else {
        result[field.key] = "";
      }
    }
    return result;
  });
}

async function clearTitlePageFields({ studyNumber }) {
  if (!studyNumber) return;
  await Word.run(async (context) => {
    for (const field of META_FIELDS) {
      const tag = `sap-meta:${studyNumber}:${field.key}`;
      const ccs = context.document.contentControls.getByTag(tag);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) continue;
      ccs.items[0].insertText("", "Replace");
    }
    await context.sync();
  });
}

export { insertTitlePage, writeTitlePageFields, readTitlePageFields, clearTitlePageFields };
