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

// Inserts a TOC page title + a plain placeholder text box (content control).
// Rationale: Some Word hosts block programmatic TOC field insertion (InvalidArgument).
// Users can manually insert the real Word TOC afterwards via References -> Table of Contents,
// which will pick up Heading 1-3 styles we apply to section titles.
function insertTableOfContents(body) {
  const title = body.insertParagraph("TABLE OF CONTENTS", Word.InsertLocation.end);
  applyHeadingStyle(title, 1);
  title.font.set({ name: "Arial", bold: true, size: 14, color: "#000000" });

  body.insertParagraph("", Word.InsertLocation.end);
  const placeholderPara = body.insertParagraph("", Word.InsertLocation.end);
  const placeholderRange = placeholderPara.getRange();
  placeholderRange.font.set({ name: "Times New Roman", bold: false, size: 12 });
  const cc = placeholderRange.insertContentControl();
  cc.appearance = "BoundingBox";
  cc.tag = "sap-toc-placeholder";
  cc.title = "Table of Contents placeholder";
  cc.placeholderText =
    "Insert Word Table of Contents here (References -> Table of Contents). Then right-click and Update Field.";
}

async function updateTableOfContents() {
  // No-op for placeholder mode.
  // Real Word TOC insertion/update should be done by user via Word UI.
}

async function generateOrUpdateTocAtPlaceholder() {
  return await Word.run(async (context) => {
    // 1) If a TOC already exists, just update it (safe, no document-structure edits).
    try {
      const tocs = context.document.tablesOfContents;
      tocs.load("items");
      await context.sync();
      if (tocs.items && tocs.items.length > 0) {
        tocs.items.forEach((t) => t.update());
        await context.sync();
        return "updated";
      }
    } catch (e) {
      // Ignore and fall back to selecting placeholder.
    }

    // 2) Otherwise, select the placeholder so the user can insert Word's native TOC
    // via References -> Table of Contents. This avoids risky OOXML insertion that can
    // trigger Word repair dialogs on some hosts.
    const ccs = context.document.contentControls.getByTag("sap-toc-placeholder");
    ccs.load("items");
    await context.sync();
    if (!ccs.items || ccs.items.length === 0) {
      throw new Error("TOC placeholder not found (sap-toc-placeholder).");
    }
    const cc = ccs.items[0];
    const placeholderRange = cc.getRange();
    placeholderRange.select();
    await context.sync();
    return "placeholder_selected";
  });
}

export { insertTableOfContents, updateTableOfContents, generateOrUpdateTocAtPlaceholder };
