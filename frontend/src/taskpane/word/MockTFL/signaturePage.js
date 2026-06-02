export * from "../../domains/mocktfl/word/signaturePage";

/* global Word */

// Remove all borders from a table (so it acts as invisible layout grid).
function tryRemoveTableBorders(table) {
  try {
    ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((b) => {
      table.getBorder(b).type = "None";
    });
  } catch (e) {}
}

/**
 * Inserts a signature block:
 *   sectionLabel  (e.g. "Prepared by:" or "Reviewed By:")
 *   [blank spacer]
 *   ──────────────────────────────────   Date
 *                                        ─────────────
 *   Name Line 1 (bold)
 *   Name Line 2 (bold)
 *   ...
 *   [blank spacer]
 */
function insertSignatureBlock(body, { sectionLabel = "", leftDetails = [] } = {}) {
  if (sectionLabel) {
    const lbl = body.insertParagraph(sectionLabel, Word.InsertLocation.end);
    lbl.font.set({ name: "Times New Roman", size: 12, bold: false });
    body.insertParagraph("", Word.InsertLocation.end);
    body.insertParagraph("", Word.InsertLocation.end);
  }

  // Row 0 : signature lines
  // Row 1 : "Date" label (right only)
  // Row 2 : left details (name/title, bold)
  const table = body.insertTable(3, 2, Word.InsertLocation.end, [
    ["", ""],
    ["", "Date"],
    [leftDetails.join("\n"), ""]
  ]);
  tryRemoveTableBorders(table);

  // Row 0: long underline left, short underline right
  table.getCell(0, 0).body.clear();
  const leftLine = table
    .getCell(0, 0)
    .body.insertParagraph(
      "_______________________________________________________________",
      Word.InsertLocation.end
    );
  leftLine.font.set({ name: "Times New Roman", size: 12, bold: false });

  table.getCell(0, 1).body.clear();
  const rightLine = table
    .getCell(0, 1)
    .body.insertParagraph("____________________", Word.InsertLocation.end);
  rightLine.font.set({ name: "Times New Roman", size: 12, bold: false });

  // Row 1: blank left, "Date" right (bold)
  table.getCell(1, 0).body.clear();
  table.getCell(1, 1).body.clear();
  const dateLbl = table.getCell(1, 1).body.insertParagraph("Date", Word.InsertLocation.end);
  dateLbl.font.set({ name: "Times New Roman", size: 12, bold: true });

  // Row 2: left details (bold), right blank
  table.getCell(2, 0).body.clear();
  table.getCell(2, 1).body.clear();
  leftDetails.forEach((line) => {
    const p = table.getCell(2, 0).body.insertParagraph(line, Word.InsertLocation.end);
    p.font.set({ name: "Times New Roman", size: 12, bold: true });
  });

  body.insertParagraph("", Word.InsertLocation.end);
  body.insertParagraph("", Word.InsertLocation.end);
}

function insertMockTflSignaturePage(body) {
  const title = body.insertParagraph("SIGNATURE PAGE", Word.InsertLocation.end);
  title.alignment = "Left";
  title.font.set({ name: "Times New Roman", size: 14, bold: true });

  // Horizontal rule using bottom-border-only 1×1 table
  try {
    const rule = body.insertTable(1, 1, Word.InsertLocation.end, [[""]]);
    ["Top", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((b) => {
      try { rule.getBorder(b).type = "None"; } catch (e) {}
    });
    rule.getBorder("Bottom").type = "Single";
  } catch (e) {}

  body.insertParagraph("", Word.InsertLocation.end);

  insertSignatureBlock(body, {
    sectionLabel: "Prepared by:",
    leftDetails: [
      "Kun Liang, Associated Director/Senior Biostatistician II",
      "LLX Solutions, LLC"
    ]
  });

  insertSignatureBlock(body, {
    sectionLabel: "Reviewed By:",
    leftDetails: [
      "Rong Jiao",
      "Senior Principal Biostatistician",
      "Novartis Pharmaceuticals Corporation"
    ]
  });

  // Third block (no section label, just company)
  insertSignatureBlock(body, {
    leftDetails: ["Novartis Pharmaceuticals Corporation"]
  });
}

export { insertMockTflSignaturePage };
