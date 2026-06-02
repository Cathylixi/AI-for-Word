export * from "../../domains/mocktfl/word/divider";

/* global Word */

/**
 * Inserts a full-page divider with large centered text (e.g. "TABLES", "FIGURES", "LISTINGS").
 * The surrounding page break is handled by the caller.
 *
 * @param {object} body - Word body object
 * @param {string} text - The divider label (e.g. "TABLES")
 */
function insertDividerPage(body, text) {
  // Several blank lines to push text toward vertical center visually
  for (let i = 0; i < 12; i++) {
    body.insertParagraph("", Word.InsertLocation.end);
  }

  const p = body.insertParagraph(text, Word.InsertLocation.end);
  p.alignment = "Centered";
  p.font.set({ name: "Times New Roman", bold: false, size: 48, color: "#000000" });

  for (let i = 0; i < 12; i++) {
    body.insertParagraph("", Word.InsertLocation.end);
  }
}

export { insertDividerPage };
