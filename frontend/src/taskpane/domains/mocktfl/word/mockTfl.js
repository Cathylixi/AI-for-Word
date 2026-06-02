/* global Word */

async function setupMockTflDocument() {
  await Word.run(async (context) => {
    // 1. Clear all content
    const body = context.document.body;
    body.clear();

    // 2. Set all sections to Landscape (all MockTFL pages use landscape uniformly).
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    sections.items.forEach((section) => {
      section.pageSetup.orientation = Word.PageOrientation.landscape;
      // Optional: Adjust margins if needed for TFLs (e.g. narrow margins)
      // section.pageSetup.bottomMargin = 36; // points (72 pts = 1 inch)
      // section.pageSetup.topMargin = 36;
      // section.pageSetup.leftMargin = 36;
      // section.pageSetup.rightMargin = 36;
    });

    await context.sync();
  });
}

async function setAllSectionsOrientation(orientation) {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    sections.items.forEach((section) => {
      section.pageSetup.orientation = orientation;
    });
    await context.sync();
  });
}

async function setLastSectionOrientation(orientation) {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    if (!sections.items || sections.items.length === 0) return;
    sections.items[sections.items.length - 1].pageSetup.orientation = orientation;
    await context.sync();
  });
}

async function resetPageOrientationToPortrait() {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    sections.items.forEach((section) => {
      section.pageSetup.orientation = Word.PageOrientation.portrait;
    });

    await context.sync();
  });
}

export { setupMockTflDocument, setAllSectionsOrientation, setLastSectionOrientation, resetPageOrientationToPortrait };
