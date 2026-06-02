export * from "../../domains/mocktfl/word/documentHistory";

/* global Word */

// Inserts a blank Document History page (title + empty table, no study-specific content).
function insertDocumentHistory(body) {
  const title = body.insertParagraph("DOCUMENT HISTORY", Word.InsertLocation.end);
  title.alignment = "Left";
  title.font.set({ name: "Times New Roman", bold: true, size: 14, color: "#000000" });

  body.insertParagraph("", Word.InsertLocation.end);

  const header = ["Date", "Version", "Reason for Update", "Update Summary"];
  const emptyRowCount = 8;
  const rows = [header];
  for (let i = 0; i < emptyRowCount; i++) rows.push(["", "", "", ""]);

  const table = body.insertTable(rows.length, 4, Word.InsertLocation.end, rows);

  // Header row: bold
  for (let c = 0; c < 4; c++) {
    table.getCell(0, c).body.getRange().font.set({
      name: "Times New Roman",
      bold: true,
      size: 11
    });
  }
  // Body rows: normal
  for (let r = 1; r < rows.length; r++) {
    for (let c = 0; c < 4; c++) {
      table.getCell(r, c).body.getRange().font.set({
        name: "Times New Roman",
        bold: false,
        size: 11
      });
    }
  }
}

export { insertDocumentHistory };
