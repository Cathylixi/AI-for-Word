export * from "../../domains/mocktfl/word/toc";

/* global Word */

/**
 * Inserts a Table of Contents page for MockTFL.
 *
 * @param {object} body  - Word body object
 * @param {Array}  tocItems - [{ label: string, page: number }]
 */
function insertMockTflTableOfContents(body, tocItems = []) {
  const title = body.insertParagraph("TABLE OF CONTENTS", Word.InsertLocation.end);
  title.alignment = "Left";
  title.font.set({ name: "Times New Roman", bold: true, size: 14, color: "#000000" });

  body.insertParagraph("", Word.InsertLocation.end);

  if (tocItems.length === 0) {
    const empty = body.insertParagraph("(No sections selected)", Word.InsertLocation.end);
    empty.font.set({ name: "Times New Roman", size: 11, bold: false });
    return;
  }

  // Two-column table: label (left) | page number (right)
  // Each row = one TOC entry
  const rows = tocItems.map((item) => [String(item.label || ""), String(item.page || "")]);
  const table = body.insertTable(rows.length, 2, Word.InsertLocation.end, rows);

  // Remove all borders so it looks like a plain list
  try {
    ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((b) => {
      try { table.getBorder(b).type = "None"; } catch (e) {}
    });
  } catch (e) {}

  // Column widths: label wide, page number narrow
  try {
    table.columns.getFirst().width = 500; // ~6.9 inches
    table.columns.getLast().width = 30;   // ~0.4 inch
  } catch (e) {}

  for (let r = 0; r < rows.length; r++) {
    // Left: entry label
    table.getCell(r, 0).body.getRange().font.set({
      name: "Times New Roman",
      size: 11,
      bold: false
    });
    // Right: page number, right-aligned
    const pageCell = table.getCell(r, 1);
    pageCell.body.getRange().font.set({ name: "Times New Roman", size: 11, bold: false });
    try {
      pageCell.body.paragraphs.getFirst().alignment = "Right";
    } catch (e) {}
  }
}

/**
 * Calculates page numbers for all selected entries given the fixed front matter structure.
 *
 * Front matter pages (in order, each = 1 page):
 *   1. Title Page
 *   2. Confidentiality Statement
 *   3. Signature Page
 *   4. Document History
 *   5. Table of Contents      ← TOC itself, so entries start at page 6
 *
 * Then, for each type group present (TABLE → FIGURE → LISTING):
 *   - 1 divider page  (e.g. "TABLES")
 *   - 1 page per entry
 *
 * @param {Array} entries  - sorted array of { type, number, title }
 * @returns {Array}        - [{ label, page }]
 */
function buildTocItems(entries) {
  const FRONT_MATTER_PAGES = 5; // Title + Conf + Sig + DocHistory + TOC
  let pageCursor = FRONT_MATTER_PAGES;

  const typeOrder = ["TABLE", "FIGURE", "LISTING"];
  const grouped = new Map(typeOrder.map((t) => [t, []]));
  entries.forEach((e) => {
    const t = String(e.type || "").toUpperCase();
    if (grouped.has(t)) grouped.get(t).push(e);
  });

  const tocItems = [];

  typeOrder.forEach((type) => {
    const items = grouped.get(type) || [];
    if (items.length === 0) return;

    // Divider page
    pageCursor += 1;

    items.forEach((item) => {
      pageCursor += 1;
      const label = `${item.type} ${item.number}  ${item.title}`.trim();
      tocItems.push({ label, page: pageCursor });
    });
  });

  return tocItems;
}

export { insertMockTflTableOfContents, buildTocItems };
