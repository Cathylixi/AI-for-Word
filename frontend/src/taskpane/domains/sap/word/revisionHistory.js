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

// Inserts the Revision History section on a new page.
function insertRevisionHistory(body) {
  insertSapHeader(body);

  const revTitle = body.insertParagraph("REVISION HISTORY", Word.InsertLocation.end);
  applyHeadingStyle(revTitle, 1);
  revTitle.font.set({ name: "Arial", bold: true, size: 14, color: "#000000" });

  const revHeader = ["Version/Date", "Section", "Reason for update", "Changes implemented"];
  const emptyRowCount = 8;
  const revRows = [revHeader];
  for (let i = 0; i < emptyRowCount; i++) {
    revRows.push(["", "", "", ""]);
  }

  const revTable = body.insertTable(revRows.length, 4, Word.InsertLocation.end, revRows);
  // Keep borders visible to match the reference screenshot
  for (let c = 0; c < 4; c++) {
    revTable.getCell(0, c).body.getRange().font.set({
      name: "Arial",
      bold: true,
      size: 10
    });
  }
  // Apply consistent font to the rest of the table
  for (let r = 1; r < revRows.length; r++) {
    for (let c = 0; c < 4; c++) {
      revTable.getCell(r, c).body.getRange().font.set({
        name: "Times New Roman",
        bold: false,
        size: 10
      });
    }
  }
}

export { insertRevisionHistory };
