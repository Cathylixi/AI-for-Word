/* global Word */

// Inserts the standard SAP header (title + full-width line).
function insertSapHeader(body) {
  const titlePara = body.insertParagraph(
    "Statistical Analysis Plan (SAP)",
    Word.InsertLocation.end
  );
  titlePara.alignment = "Centered";
  titlePara.font.set({ name: "Arial", bold: true, size: 16, color: "#000000" });

  // Full-width horizontal line using a 1x1 table with only bottom border.
  const lineTable = body.insertTable(1, 1, Word.InsertLocation.end, [[""]]);
  try {
    // Remove all borders first
    ["Top", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((b) => {
      const border = lineTable.getBorder(b);
      border.type = "None";
    });
    // Keep the bottom border to render a full-width line
    lineTable.getBorder("Bottom").type = "Single";
  } catch (e) {
    // If border APIs are unsupported, keep the default lineTable rendering.
  }

  // Spacing after the line
  body.insertParagraph("", Word.InsertLocation.end);
}

export { insertSapHeader };
