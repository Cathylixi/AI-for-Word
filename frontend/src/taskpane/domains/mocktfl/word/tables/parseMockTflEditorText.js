const ARM_VALUE_PLACEHOLDER = "xx (xx.x)";

/**
 * Finalize the current block and push it into the output list when valid.
 */
function finalizeCurrentBlock(blocks, currentBlock) {
  const question = String(currentBlock?.question || "").trim();
  const options = Array.isArray(currentBlock?.options)
    ? currentBlock.options.map((opt) => String(opt || "").trim()).filter(Boolean)
    : [];

  if (!question && options.length === 0) return;
  blocks.push({ question, options });
}

/**
 * Normalize a bullet option line by removing the leading "-" marker.
 */
function normalizeOptionLine(line) {
  return String(line || "")
    .replace(/^\s*-\s+/, "")
    .trim();
}

/**
 * Parse the final editable MockTFL Notes text into question blocks.
 *
 * Block rules:
 * - blank line => end current block
 * - first non-bullet line => block question
 * - following non-bullet lines => question continuation (same block)
 * - "- xxx" lines => option rows under the current question
 *
 * IMPORTANT:
 * - We intentionally keep repeated question blocks in their original order.
 * - We do not deduplicate here because the user explicitly wants literal block order preserved.
 */
function parseEditorTextToBlocks(text) {
  const raw = String(text || "");
  if (!raw.trim()) return [];

  const blocks = [];
  let currentBlock = { question: "", options: [] };

  raw.split(/\r\n|\n/g).forEach((line) => {
    const trimmed = String(line || "").trim();
    if (!trimmed) {
      finalizeCurrentBlock(blocks, currentBlock);
      currentBlock = { question: "", options: [] };
      return;
    }

    if (/^\s*-\s+/.test(line)) {
      currentBlock.options.push(normalizeOptionLine(line));
      return;
    }

    if (!currentBlock.question) {
      currentBlock.question = trimmed;
      return;
    }

    // Treat additional non-bullet lines inside the same block as continuations
    // of a long question that wrapped across lines in the editor.
    currentBlock.question = `${currentBlock.question} ${trimmed}`.replace(/\s+/g, " ").trim();
  });

  finalizeCurrentBlock(blocks, currentBlock);
  return blocks;
}

/**
 * Convert parsed question blocks into table rows.
 *
 * Rules:
 * - question rows stay flush-left in column 1
 * - option rows are inserted without the leading "-" and with a visual indent
 * - question blocks are separated by one lightweight blank row
 */
function buildRowsFromBlocks(blocks, columnCount) {
  const rows = [];
  const emptyCells = Array(columnCount - 1).fill("");
  const placeholderCells = Array(columnCount - 1).fill(ARM_VALUE_PLACEHOLDER);

  if (!Array.isArray(blocks) || blocks.length === 0) {
    rows.push(["(No content found in editor)", ...placeholderCells]);
    return rows;
  }

  blocks.forEach((block, idx) => {
    rows.push([
      String(block?.question || "").trim(),
      ...placeholderCells
    ]);

    const options = Array.isArray(block?.options) ? block.options : [];
    options.forEach((optionText) => {
      // Instead of hardcoding spaces, we pass a special invisible prefix \u200B (Zero Width Space)
      // plus a custom indent marker so the OOXML generator can catch it and apply real paragraph indent.
      rows.push([`__INDENT__${String(optionText || "").trim()}`, ...placeholderCells]);
    });

    // Keep a visual separator between question blocks without making the layout
    // as loose as preserving every raw blank line literally.
    if (idx < blocks.length - 1) {
      rows.push([" ", ...emptyCells]);
    }
  });

  return rows;
}

/**
 * Public builder used by the table inserters.
 * We keep the exported function name stable so calling files stay small.
 */
function buildIndentedTableRowsFromEditorText(text, columnCount = 4) {
  return buildRowsFromBlocks(parseEditorTextToBlocks(text), columnCount);
}

export { parseEditorTextToBlocks, buildIndentedTableRowsFromEditorText };
