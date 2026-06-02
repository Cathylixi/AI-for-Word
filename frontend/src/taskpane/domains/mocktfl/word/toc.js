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

  // Placeholder mode (SAP-like): keep only the title + an empty content control text box.
  // Rationale: we may later generate or insert a native Word TOC / custom TOC in a later step.
  const placeholderPara = body.insertParagraph("", Word.InsertLocation.end);
  const placeholderRange = placeholderPara.getRange();
  placeholderRange.font.set({ name: "Times New Roman", bold: false, size: 12 });
  const cc = placeholderRange.insertContentControl();
  cc.appearance = "BoundingBox";
  cc.tag = "mocktfl-toc-placeholder";
  cc.title = "MockTFL Table of Contents placeholder";
  cc.placeholderText = "TOC placeholder (will be generated/filled later).";
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
