function uniqueNonEmptyStrings(values) {
  const arr = Array.isArray(values) ? values : [];
  const seen = new Set();
  const out = [];

  arr.forEach((value) => {
    const text = String(value || "").trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    out.push(text);
  });

  return out;
}

function normalizeOptionalString(value) {
  return String(value || "").trim();
}

function normalizeColumnOrientation(value) {
  return String(value || "").trim().toLowerCase() === "horizontal" ? "horizontal" : "vertical";
}

function normalizeHeaderLayers(headerLayers) {
  const arr = Array.isArray(headerLayers) ? headerLayers : [];
  return arr
    .map((item, idx) => ({
      level: Number.isFinite(Number(item?.level)) ? Number(item.level) : idx + 1,
      headers: uniqueNonEmptyStrings(item?.headers)
    }))
    .filter((item) => item.headers.length > 0)
    .sort((a, b) => a.level - b.level);
}

function normalizeColumnHeaderConfigShape(columnHeaderConfig) {
  // Table generators consume one normalized config regardless of how the dialog collected it.
  const orientation = normalizeColumnOrientation(columnHeaderConfig?.orientation);
  const columnGroups = uniqueNonEmptyStrings(columnHeaderConfig?.columnGroups);
  const headerLayers = normalizeHeaderLayers(columnHeaderConfig?.headerLayers);
  const columnGroupSubtitle = normalizeOptionalString(columnHeaderConfig?.columnGroupSubtitle);

  return {
    orientation,
    columnGroups,
    columnGroupSubtitle,
    headerLayers
  };
}

function buildColumnGroupHeaderText({ columnGroup, armSuffixText = "(N=xx)", columnGroupSubtitle = "" }) {
  // Subtitle stays inside the same top cell, so it does not create an extra ruled header row.
  return [String(columnGroup || "").trim(), armSuffixText, normalizeOptionalString(columnGroupSubtitle)]
    .filter(Boolean)
    .join("\u000B");
}

function resolveColumnHeaderConfig({
  columnHeaderConfig,
  selectedColumns,
  defaultLevel1Headers = []
}) {
  const normalizedConfig = normalizeColumnHeaderConfigShape(columnHeaderConfig);
  const selectedLevel1 = uniqueNonEmptyStrings(selectedColumns);
  const fallbackLevel1 = uniqueNonEmptyStrings(defaultLevel1Headers);

  return {
    orientation: normalizedConfig.orientation,
    columnGroups:
      normalizedConfig.columnGroups.length > 0
        ? normalizedConfig.columnGroups
        : selectedLevel1.length > 0
          ? selectedLevel1
          : fallbackLevel1,
    columnGroupSubtitle: normalizedConfig.columnGroupSubtitle,
    headerLayers: normalizedConfig.headerLayers
  };
}

function hasCustomSubLevels(columnHeaderConfig) {
  return normalizeColumnHeaderConfigShape(columnHeaderConfig).headerLayers.length > 0;
}

function buildMultiLevelHeaderRows({
  columnGroups,
  headerLayers,
  armSuffixText = "(N=xx)",
  columnGroupSubtitle = ""
}) {
  const normalizedColumnGroups = uniqueNonEmptyStrings(columnGroups);
  const normalizedHeaderLayers = normalizeHeaderLayers(headerLayers);
  if (normalizedColumnGroups.length === 0 || normalizedHeaderLayers.length === 0) {
    return { rows: [], leafColumnCount: normalizedColumnGroups.length };
  }

  const levels = [normalizedColumnGroups, ...normalizedHeaderLayers.map((item) => item.headers)];
  const leafColumnCount = levels.reduce((acc, levelHeaders) => acc * levelHeaders.length, 1);
  const rows = [];

  for (let levelIdx = 0; levelIdx < levels.length; levelIdx++) {
    const currentLevelHeaders = levels[levelIdx];
    // Each parent header repeats across every leaf under it.
    const blockCount = levels.slice(0, levelIdx).reduce((acc, levelHeaders) => acc * levelHeaders.length, 1);
    const colSpan = levels.slice(levelIdx + 1).reduce((acc, levelHeaders) => acc * levelHeaders.length, 1);
    const row = [""];

    for (let blockIdx = 0; blockIdx < blockCount; blockIdx++) {
      currentLevelHeaders.forEach((headerText) => {
        row.push({
          text:
            levelIdx === 0
              ? buildColumnGroupHeaderText({
                  columnGroup: headerText,
                  armSuffixText,
                  columnGroupSubtitle
                })
              : headerText,
          colSpan
        });
      });
    }

    rows.push(row);
  }

  return { rows, leafColumnCount };
}

export {
  resolveColumnHeaderConfig,
  hasCustomSubLevels,
  buildMultiLevelHeaderRows,
  buildColumnGroupHeaderText
};
