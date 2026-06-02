/* global Word */

/**
 * Configure the MockTFL Page Header.
 * 
 * Layout:
 * ----------------------------------------------------
 * Anthos Therapeutics, Inc., A Novartis Company
 * Protocol no. ANT-007                               Page X of Y
 * ----------------------------------------------------
 * 
 * Logic:
 * - Applied to the document's main section(s).
 * - "differentFirstPageHeaderFooter = true" is set so the Title Page (Page 1) remains blank.
 * - Font: Courier New, 8pt (consistent with TFL body).
 * - Alignment: 
 *   - Line 1: Left
 *   - Line 2: "Protocol..." (Left), "Page..." (Right Aligned via Tab Stop).
 * 
 * @param {Word.RequestContext} context 
 * @param {string} [protocolNumber="ANT-007"] - Optional protocol number.
 */
export async function setupMockTflPageHeader(context, protocolNumber = "ANT-007") {
  // We apply this to all sections in the current selection/document for simplicity,
  // or specifically the section where the MockTFL content resides.
  // Assuming the entire document is generated for MockTFL or we are working on the active section.
  
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  // Iterate over sections (usually just 1 if generated from scratch, or multiple if section breaks exist).
  // We want headers on all TFL pages.
  for (let i = 0; i < sections.items.length; i++) {
    const section = sections.items[i];
    
    // 1. Ensure First Page (Title Page) has NO header/footer.
    // This property makes the First Page header separate from the Primary header.
    // By default, the separate First Page header is blank unless we write to it.
    try {
      section.differentFirstPageHeaderFooter = true;
      await context.sync();
    } catch (e) {
      console.warn("Failed to set differentFirstPageHeaderFooter", e);
    }

    // 2. Access the PRIMARY Header (used for pages 2, 3, ...).
    let header;
    try {
      header = section.getHeader("Primary");
      header.clear();
      await context.sync();
    } catch (e) {
      console.warn("Failed to get/clear header", e);
      continue; // Skip if we can't get header
    }
    
    // 3. Insert Header Table (2 rows x 2 columns) for precise layout
    // Row 1: "Anthos..." (Left) | (Empty)
    // Row 2: "Protocol..." (Left) | "Page X of Y" (Right)
    let headerTable;
    try {
      // Insert a 2x2 table
      headerTable = header.insertTable(2, 2, Word.InsertLocation.start);
      headerTable.autoFitBehavior(Word.AutoFitBehavior.autoFitWindow);
      
      // Clear borders (make it invisible)
      ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach(b => {
        try { headerTable.getBorder(b).type = "None"; } catch(e) {}
      });

      // Set font for the whole table
      headerTable.getRange().font.set({ name: "Courier New", size: 8, bold: false });
      
      await context.sync();

      // --- Row 1, Cell 1: Company Name ---
      const r1c1 = headerTable.getCell(0, 0);
      r1c1.body.clear();
      const p1 = r1c1.body.insertParagraph("Anthos Therapeutics, Inc., A Novartis Company", Word.InsertLocation.start);
      p1.font.set({ name: "Courier New", size: 8, bold: false });
      try { p1.alignment = "Left"; } catch(e) {}

      // --- Row 1, Cell 2: Empty (or merge if needed, but empty works) ---
      const r1c2 = headerTable.getCell(0, 1);
      r1c2.body.clear();

      // --- Row 2, Cell 1: Protocol ---
      const r2c1 = headerTable.getCell(1, 0);
      r2c1.body.clear();
      const p2 = r2c1.body.insertParagraph(`Protocol no. ${protocolNumber}`, Word.InsertLocation.start);
      p2.font.set({ name: "Courier New", size: 8, bold: false });
      try { p2.alignment = "Left"; } catch(e) {}

      // --- Row 2, Cell 2: Page X of Y (Right Aligned) ---
      const r2c2 = headerTable.getCell(1, 1);
      r2c2.body.clear();
      const p3 = r2c2.body.insertParagraph("", Word.InsertLocation.start);
      p3.font.set({ name: "Courier New", size: 8, bold: false });
      try { p3.alignment = "Right"; } catch(e) {}

      // Insert Page Fields into p3
      // "Page "
      const p3Range = p3.getRange(Word.RangeLocation.end);
      const t1 = p3Range.insertText("Page ", Word.InsertLocation.after);
      t1.font.set({ name: "Courier New", size: 8, bold: false });

      // Field: Page Number
      const f1 = t1.insertField(Word.FieldType.page, Word.InsertLocation.after);
      f1.result.font.set({ name: "Courier New", size: 8, bold: false });

      // " of "
      const t2 = f1.result.insertText(" of ", Word.InsertLocation.after);
      t2.font.set({ name: "Courier New", size: 8, bold: false });

      // Field: Total Pages
      const f2 = t2.insertField(Word.FieldType.numPages, Word.InsertLocation.after);
      f2.result.font.set({ name: "Courier New", size: 8, bold: false });

      await context.sync();
    } catch (e) {
      console.warn("Failed to insert header table", e);
    }
  }
}
