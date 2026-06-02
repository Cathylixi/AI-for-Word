/* global Word */

import { buildMockTflBodyTag } from "../constants";

async function insertMockTflFigure({ studyNumber, type, number, base64Image }) {
  const image = String(base64Image || "").trim();
  if (!image) throw new Error("No generated figure image was provided.");

  await Word.run(async (context) => {
    const tag = buildMockTflBodyTag({ studyNumber, type, number });
    const ccs = context.document.contentControls.getByTag(tag);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) {
      throw new Error(`Target body Content Control for ${type} ${number} was not found.`);
    }

    const placeholderCc = ccs.items[0];
    placeholderCc.clear();
    await context.sync();

    // Figures are rendered by R as a final PNG, so Word only needs a generic image insertion step.
    placeholderCc.insertInlinePictureFromBase64(image, Word.InsertLocation.start);
    await context.sync();
  });
}

export { insertMockTflFigure };
