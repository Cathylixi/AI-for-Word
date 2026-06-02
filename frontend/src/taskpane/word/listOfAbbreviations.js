/* global Word */

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

// Inserts the LIST OF ABBREVIATIONS title + placeholder.
// The table will be generated later by user action.
function insertListOfAbbreviationsPlaceholder(body) {
  const title = body.insertParagraph("LIST OF ABBREVIATIONS", Word.InsertLocation.end);
  applyHeadingStyle(title, 1);
  title.font.set({ name: "Arial", bold: true, size: 14, color: "#000000" });

  body.insertParagraph("", Word.InsertLocation.end);
  const placeholderPara = body.insertParagraph("", Word.InsertLocation.end);
  const placeholderRange = placeholderPara.getRange();
  placeholderRange.font.set({ name: "Times New Roman", bold: false, size: 12 });
  const cc = placeholderRange.insertContentControl();
  cc.appearance = "BoundingBox";
  cc.tag = "sap-loa-placeholder";
  cc.title = "List of Abbreviations placeholder";
  cc.placeholderText = "Click 'Generate Abbreviations' in the taskpane to create the table.";
}

// Locates the placeholder and replaces it with the abbreviations table.
async function generateAbbreviationsAtPlaceholder(items) {
  await Word.run(async (context) => {
    // 1. Try to find existing table first (in case of re-generation)
    const tableCcs = context.document.contentControls.getByTag("sap-loa-table");
    tableCcs.load("items");
    await context.sync();

    if (tableCcs.items.length > 0) {
      // Clear existing table container
      tableCcs.items[0].delete(false); // delete content but keep... wait, better to delete whole CC and recreate
      await context.sync();
    }

    // 2. Find placeholder
    const placeholderCcs = context.document.contentControls.getByTag("sap-loa-placeholder");
    placeholderCcs.load("items");
    await context.sync();

    let targetRange = null;
    let placeholderCc = null;

    if (placeholderCcs.items.length > 0) {
      placeholderCc = placeholderCcs.items[0];
      targetRange = placeholderCc.getRange();
    } else {
      // Fallback: search for title if placeholder is gone
      const results = context.document.body.search("LIST OF ABBREVIATIONS", { matchCase: true });
      results.load("items");
      await context.sync();
      if (results.items.length > 0) {
        targetRange = results.items[0].paragraph.getNext().getRange();
      } else {
        throw new Error("Abbreviations section not found.");
      }
    }

    // 3. Prepare table data
    const rows = (items || []).map((it) => [it.term || "", it.definition || ""]);
    const tableRows = [["Abbreviation or special term", "Explanation"], ...rows];

    // 4. Insert Table (after placeholder)
    // We insert a NEW content control to wrap the table, ensuring we can find it later for saving/reading.
    const table = targetRange.insertTable(tableRows.length, 2, Word.InsertLocation.after, tableRows);
    
    // 5. Wrap table in a persistent content control
    // Note: insertContentControl must be called on a Range, not a Table object.
    const tableWrapper = table.getRange().insertContentControl();
    tableWrapper.tag = "sap-loa-table";
    tableWrapper.title = "List of Abbreviations Table";
    tableWrapper.appearance = "BoundingBox";

    // 6. Delete placeholder if it exists (now replaced by table)
    if (placeholderCc) {
      placeholderCc.delete(false);
    }

    // 7. Format Table
    // Remove all borders first
    try {
      ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((b) => {
        const border = table.getBorder(b);
        border.type = "None";
      });
    } catch (e) {
      // If border APIs are unsupported, keep default borders.
    }

    // Style header row and add top/bottom borders only on the header cells
    for (let c = 0; c < 2; c++) {
      const headerCell = table.getCell(0, c);
      headerCell.body.getRange().font.set({ name: "Arial", bold: true, size: 10 });
      try {
        headerCell.getBorder("Top").type = "Single";
        headerCell.getBorder("Bottom").type = "Single";
      } catch (e) {
        // If cell borders are unsupported, keep default styling.
      }
    }

    // Body rows (Arial 10)
    for (let r = 1; r < tableRows.length; r++) {
      for (let c = 0; c < 2; c++) {
        table.getCell(r, c).body.getRange().font.set({
          name: "Arial",
          bold: false,
          size: 10
        });
      }
    }

    await context.sync();
  });
}

// Reads the current abbreviations table from the generated table CC.
async function readAbbreviationsFromPlaceholder() {
  return await Word.run(async (context) => {
    // Try finding the generated table first
    const tableCcs = context.document.contentControls.getByTag("sap-loa-table");
    tableCcs.load("items");
    await context.sync();

    let targetCc = null;
    if (tableCcs.items.length > 0) {
      targetCc = tableCcs.items[0];
    } else {
      // Fallback to placeholder (likely empty, but check)
      const ccs = context.document.contentControls.getByTag("sap-loa-placeholder");
      ccs.load("items");
      await context.sync();
      if (ccs.items.length > 0) targetCc = ccs.items[0];
    }

    if (!targetCc) return [];

    const tables = targetCc.tables;
    tables.load("items");
    await context.sync();

    if (tables.items.length === 0) return [];
    const table = tables.items[0];
    const rows = table.rows;
    rows.load("items");
    await context.sync();

    // Skip header row (index 0)
    const items = [];
    for (let i = 1; i < rows.items.length; i++) {
      const cells = rows.items[i].cells;
      cells.load("items");
      await context.sync();
      if (cells.items.length >= 2) {
        const termRange = cells.items[0].body.getRange();
        const defRange = cells.items[1].body.getRange();
        termRange.load("text");
        defRange.load("text");
        await context.sync();
        const term = termRange.text.trim();
        const definition = defRange.text.trim();
        if (term) items.push({ term, definition });
      }
    }
    return items;
  });
}

export { insertListOfAbbreviationsPlaceholder, generateAbbreviationsAtPlaceholder, readAbbreviationsFromPlaceholder };
