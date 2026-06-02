/* global Word */

import { TAG_PREFIX_BODY } from "./constants";

async function selectSectionBody({ studyNumber, sectionNumber }) {
  await Word.run(async (context) => {
    const tag = `${TAG_PREFIX_BODY}${studyNumber}:${sectionNumber}`;
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return;
    ccs.items[0].select();
    await context.sync();
  });
}

async function readSectionBody({ studyNumber, sectionNumber }) {
  return await Word.run(async (context) => {
    const tag = `${TAG_PREFIX_BODY}${studyNumber}:${sectionNumber}`;
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return "";
    const range = ccs.items[0].getRange();
    range.load("text");
    await context.sync();
    return range.text || "";
  });
}

async function writeSectionBody({ studyNumber, sectionNumber, text }) {
  await Word.run(async (context) => {
    const tag = `${TAG_PREFIX_BODY}${studyNumber}:${sectionNumber}`;
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();
    if (ccs.items.length === 0) return;
    const cc = ccs.items[0];
    cc.insertText(text || "", "Replace");

    // Re-apply formatting to the content control's range after insertion
    // to ensure user-pasted text doesn't override it
    const range = cc.getRange();
    range.font.set({
      name: "Times New Roman",
      size: 12,
      bold: false
    });

    await context.sync();
  });
}

async function writeAllSapBodySections({ studyNumber, sectionsData }) {
  // sectionsData: [{ sectionNumber: "1.1", text: "..." }, ...]
  if (!Array.isArray(sectionsData) || sectionsData.length === 0) return;

  await Word.run(async (context) => {
    const bodies = context.document.body.contentControls;
    bodies.load("items/tag");
    await context.sync();

    for (const item of sectionsData) {
      const tag = `${TAG_PREFIX_BODY}${studyNumber}:${item.sectionNumber}`;
      // Find matching CC in loaded items (avoid calling getByTag in loop if possible, 
      // but getByTag is robust. Let's try matching loaded items first to save roundtrips)
      const target = bodies.items.find(cc => cc.tag === tag);
      
      if (target) {
        target.insertText(item.text || "", "Replace");
        target.getRange().font.set({
          name: "Times New Roman",
          size: 12,
          bold: false
        });
      }
    }
    await context.sync();
  });
}

async function readAllSapBodySections({ studyNumber }) {
  return await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    ccs.items.forEach((cc) => cc.load("tag"));
    await context.sync();

    const prefix = `${TAG_PREFIX_BODY}${studyNumber}:`;
    const targets = ccs.items.filter((cc) => (cc.tag || "").startsWith(prefix));

    // Load text for each target control
    const ranges = targets.map((cc) => cc.getRange());
    ranges.forEach((r) => r.load("text"));
    await context.sync();

    const sections = targets.map((cc, idx) => {
      const tag = cc.tag || "";
      const sectionNumber = tag.slice(prefix.length);
      const text = ranges[idx]?.text || "";
      return { sectionNumber, text };
    });

    return sections;
  });
}

export { selectSectionBody, readSectionBody, writeSectionBody, readAllSapBodySections, writeAllSapBodySections };
