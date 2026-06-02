/**
 * Central R template library for the MockTFL figure workflow (Path A).
 *
 * The AI no longer writes raw R code. Instead it only:
 *   1. picks a `figureFamily` (one of KNOWN_FIGURE_FAMILIES), and
 *   2. maps the existing data columns to that family's required roles.
 *
 * This module turns that mapping into deterministic, pre-tested R code so the
 * remote R runner can never crash on AI-invented statistics or syntax.
 *
 * Contract with the caller:
 * - The caller injects `library(jsonlite)` and the `json_data` string BEFORE the
 *   code returned here. So every template can assume `json_data` exists and
 *   jsonlite is loaded. Templates only need to add their own extra libraries
 *   (survival, ggplot2, ...).
 * - Templates always read observations from `dat$records`.
 * - Templates build the plot into `p` and call `print(p)`.
 */

const KNOWN_FIGURE_FAMILIES = [
  "time_to_event", // survival / cumulative incidence / Kaplan-Meier
  "longitudinal", // mean of a measure over time (optionally by group)
  "categorical", // bar / grouped bar
  "distribution", // boxplot of a numeric measure across categories
  "scatter" // x-y scatter
];

/**
 * Escape a string so it can be embedded inside an R double-quoted string.
 */
function rString(value) {
  return String(value == null ? "" : value)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"');
}

/**
 * Reference a data.frame column by name: df[["col name"]].
 */
function rCol(name) {
  return `df[["${rString(name)}"]]`;
}

function requireMappingColumn(mapping, key, availableColumns, familyLabel) {
  const col = String(mapping?.[key] || "").trim();
  if (!col) {
    throw new Error(`Figure mapping for "${familyLabel}" is missing required column "${key}".`);
  }
  if (Array.isArray(availableColumns) && availableColumns.length && !availableColumns.includes(col)) {
    throw new Error(
      `Figure mapping column "${col}" (role "${key}") does not exist in the data columns: ${availableColumns.join(", ")}.`
    );
  }
  return col;
}

function optionalMappingColumn(mapping, key, availableColumns) {
  const col = String(mapping?.[key] || "").trim();
  if (!col) return "";
  if (Array.isArray(availableColumns) && availableColumns.length && !availableColumns.includes(col)) {
    return "";
  }
  return col;
}

function labelOr(labels, key, fallback) {
  const v = String(labels?.[key] || "").trim();
  return v || fallback;
}

const PREAMBLE = ["dat <- fromJSON(json_data)", "df <- as.data.frame(dat$records)"].join("\n");

/**
 * time_to_event: Kaplan-Meier survival OR cumulative incidence.
 *
 * y-axis:
 *   - survival         -> S(t)        (probability of NOT yet having the event)
 *   - cumulative incid. -> 1 - S(t)   (probability of HAVING had the event)
 *
 * The y-axis kind is inferred from specificType so the same raw data can render
 * either curve correctly.
 */
function buildTimeToEvent({ specificType, mapping, labels, availableColumns }) {
  const timeCol = requireMappingColumn(mapping, "timeColumn", availableColumns, "time_to_event");
  const statusCol = requireMappingColumn(mapping, "statusColumn", availableColumns, "time_to_event");
  const groupCol = optionalMappingColumn(mapping, "groupColumn", availableColumns);

  const typeText = String(specificType || "").toLowerCase();
  const isSurvival = /surviv|kaplan|\bkm\b|overall survival|\bos\b|\bpfs\b/.test(typeText);
  const yExpr = isSurvival ? "plot_df$surv" : "1 - plot_df$surv";
  const defaultYLabel = isSurvival ? "Survival probability" : "Cumulative incidence";

  const xLabel = labelOr(labels, "x", "Time");
  const yLabel = labelOr(labels, "y", defaultYLabel);

  const lines = [
    "library(survival)",
    "library(ggplot2)",
    PREAMBLE,
    `df$.time <- suppressWarnings(as.numeric(${rCol(timeCol)}))`,
    `df$.status <- suppressWarnings(as.numeric(${rCol(statusCol)}))`,
    "df <- df[!is.na(df$.time) & !is.na(df$.status), ]",
    'if (nrow(df) < 2) stop("Not enough valid time-to-event rows to draw a curve.")'
  ];

  if (groupCol) {
    lines.push(
      `df$.grp <- as.character(${rCol(groupCol)})`,
      "fit <- survfit(Surv(.time, .status) ~ .grp, data = df)",
      "s <- summary(fit)",
      'grp <- if (!is.null(s$strata)) sub("^[^=]*=", "", as.character(s$strata)) else rep("All", length(s$time))'
    );
  } else {
    lines.push(
      "fit <- survfit(Surv(.time, .status) ~ 1, data = df)",
      "s <- summary(fit)",
      'grp <- rep("All", length(s$time))'
    );
  }

  lines.push(
    "plot_df <- data.frame(time = s$time, surv = s$surv, n.censor = s$n.censor, group = grp, stringsAsFactors = FALSE)",
    // Anchor every curve at the origin (t = 0, S = 1) so the step starts cleanly.
    "origin <- do.call(rbind, lapply(unique(plot_df$group), function(g) data.frame(time = 0, surv = 1, n.censor = 0, group = g, stringsAsFactors = FALSE)))",
    "plot_df <- rbind(origin, plot_df)",
    "plot_df <- plot_df[order(plot_df$group, plot_df$time), ]",
    `plot_df$yval <- ${yExpr}`,
    "cens <- plot_df[plot_df$n.censor > 0, ]"
  );

  if (groupCol) {
    lines.push(
      "p <- ggplot(plot_df, aes(x = time, y = yval, color = group)) +",
      "  geom_step() +",
      "  geom_point(data = cens, aes(x = time, y = yval, color = group), shape = 1, size = 2) +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}", color = "Group") +`,
      "  theme_minimal()"
    );
  } else {
    lines.push(
      "p <- ggplot(plot_df, aes(x = time, y = yval)) +",
      "  geom_step() +",
      "  geom_point(data = cens, aes(x = time, y = yval), shape = 1, size = 2) +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}") +`,
      "  theme_minimal()"
    );
  }

  lines.push("print(p)");
  return lines.join("\n");
}

/**
 * longitudinal: mean of a numeric measure over an x (often time), optional group.
 */
function buildLongitudinal({ mapping, labels, availableColumns }) {
  const xCol = requireMappingColumn(mapping, "xColumn", availableColumns, "longitudinal");
  const yCol = requireMappingColumn(mapping, "yColumn", availableColumns, "longitudinal");
  const groupCol = optionalMappingColumn(mapping, "groupColumn", availableColumns);

  const xLabel = labelOr(labels, "x", xCol);
  const yLabel = labelOr(labels, "y", `Mean ${yCol}`);

  const lines = [
    "library(ggplot2)",
    PREAMBLE,
    `df$.x <- suppressWarnings(as.numeric(${rCol(xCol)}))`,
    `df$.y <- suppressWarnings(as.numeric(${rCol(yCol)}))`,
    "df <- df[!is.na(df$.x) & !is.na(df$.y), ]",
    'if (nrow(df) < 2) stop("Not enough valid rows for a longitudinal plot.")'
  ];

  if (groupCol) {
    lines.push(
      `df$.grp <- as.character(${rCol(groupCol)})`,
      "agg <- aggregate(.y ~ .x + .grp, data = df, FUN = mean)",
      'names(agg) <- c("x", "grp", "y")',
      "p <- ggplot(agg, aes(x = x, y = y, color = grp)) +",
      "  geom_line() +",
      "  geom_point() +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}", color = "Group") +`,
      "  theme_minimal()"
    );
  } else {
    lines.push(
      "agg <- aggregate(.y ~ .x, data = df, FUN = mean)",
      'names(agg) <- c("x", "y")',
      "p <- ggplot(agg, aes(x = x, y = y)) +",
      "  geom_line() +",
      "  geom_point() +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}") +`,
      "  theme_minimal()"
    );
  }

  lines.push("print(p)");
  return lines.join("\n");
}

/**
 * categorical: bar chart. If valueColumn is provided, bars show the summed value;
 * otherwise bars show the count of rows per category. Optional grouping (dodge).
 */
function buildCategorical({ mapping, labels, availableColumns }) {
  const categoryCol = requireMappingColumn(mapping, "categoryColumn", availableColumns, "categorical");
  const valueCol = optionalMappingColumn(mapping, "valueColumn", availableColumns);
  const groupCol = optionalMappingColumn(mapping, "groupColumn", availableColumns);

  const xLabel = labelOr(labels, "x", categoryCol);
  const yLabel = labelOr(labels, "y", valueCol ? valueCol : "Count");

  const lines = ["library(ggplot2)", PREAMBLE, `df$.cat <- as.character(${rCol(categoryCol)})`];
  if (groupCol) lines.push(`df$.grp <- as.character(${rCol(groupCol)})`);

  if (valueCol) {
    lines.push(`df$.val <- suppressWarnings(as.numeric(${rCol(valueCol)}))`, "df <- df[!is.na(df$.val), ]");
    if (groupCol) {
      lines.push(
        "agg <- aggregate(.val ~ .cat + .grp, data = df, FUN = sum)",
        'names(agg) <- c("cat", "grp", "val")',
        "p <- ggplot(agg, aes(x = cat, y = val, fill = grp)) +",
        '  geom_col(position = "dodge") +',
        `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}", fill = "Group") +`,
        "  theme_minimal()"
      );
    } else {
      lines.push(
        "agg <- aggregate(.val ~ .cat, data = df, FUN = sum)",
        'names(agg) <- c("cat", "val")',
        "p <- ggplot(agg, aes(x = cat, y = val)) +",
        '  geom_col() +',
        `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}") +`,
        "  theme_minimal()"
      );
    }
  } else if (groupCol) {
    lines.push(
      "p <- ggplot(df, aes(x = .cat, fill = .grp)) +",
      '  geom_bar(position = "dodge") +',
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}", fill = "Group") +`,
      "  theme_minimal()"
    );
  } else {
    lines.push(
      "p <- ggplot(df, aes(x = .cat)) +",
      "  geom_bar() +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}") +`,
      "  theme_minimal()"
    );
  }

  lines.push("print(p)");
  return lines.join("\n");
}

/**
 * distribution: boxplot of a numeric value across categories (categoryColumn),
 * optionally filled by group.
 */
function buildDistribution({ mapping, labels, availableColumns }) {
  const valueCol = requireMappingColumn(mapping, "valueColumn", availableColumns, "distribution");
  const categoryCol =
    optionalMappingColumn(mapping, "categoryColumn", availableColumns) ||
    optionalMappingColumn(mapping, "groupColumn", availableColumns);
  const groupCol = optionalMappingColumn(mapping, "groupColumn", availableColumns);

  const xLabel = labelOr(labels, "x", categoryCol || "Group");
  const yLabel = labelOr(labels, "y", valueCol);

  const lines = [
    "library(ggplot2)",
    PREAMBLE,
    `df$.y <- suppressWarnings(as.numeric(${rCol(valueCol)}))`,
    "df <- df[!is.na(df$.y), ]"
  ];

  if (categoryCol) {
    lines.push(`df$.x <- as.character(${rCol(categoryCol)})`);
  } else {
    lines.push('df$.x <- "All"');
  }

  if (groupCol && groupCol !== categoryCol) {
    lines.push(
      `df$.grp <- as.character(${rCol(groupCol)})`,
      "p <- ggplot(df, aes(x = .x, y = .y, fill = .grp)) +",
      "  geom_boxplot() +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}", fill = "Group") +`,
      "  theme_minimal()"
    );
  } else {
    lines.push(
      "p <- ggplot(df, aes(x = .x, y = .y)) +",
      "  geom_boxplot() +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}") +`,
      "  theme_minimal()"
    );
  }

  lines.push("print(p)");
  return lines.join("\n");
}

/**
 * scatter: x-y scatter plot, optional color grouping.
 */
function buildScatter({ mapping, labels, availableColumns }) {
  const xCol = requireMappingColumn(mapping, "xColumn", availableColumns, "scatter");
  const yCol = requireMappingColumn(mapping, "yColumn", availableColumns, "scatter");
  const groupCol = optionalMappingColumn(mapping, "groupColumn", availableColumns);

  const xLabel = labelOr(labels, "x", xCol);
  const yLabel = labelOr(labels, "y", yCol);

  const lines = [
    "library(ggplot2)",
    PREAMBLE,
    `df$.x <- suppressWarnings(as.numeric(${rCol(xCol)}))`,
    `df$.y <- suppressWarnings(as.numeric(${rCol(yCol)}))`,
    "df <- df[!is.na(df$.x) & !is.na(df$.y), ]",
    'if (nrow(df) < 1) stop("Not enough valid rows for a scatter plot.")'
  ];

  if (groupCol) {
    lines.push(
      `df$.grp <- as.character(${rCol(groupCol)})`,
      "p <- ggplot(df, aes(x = .x, y = .y, color = .grp)) +",
      "  geom_point(size = 2) +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}", color = "Group") +`,
      "  theme_minimal()"
    );
  } else {
    lines.push(
      "p <- ggplot(df, aes(x = .x, y = .y)) +",
      "  geom_point(size = 2) +",
      `  labs(x = "${rString(xLabel)}", y = "${rString(yLabel)}") +`,
      "  theme_minimal()"
    );
  }

  lines.push("print(p)");
  return lines.join("\n");
}

const FAMILY_BUILDERS = {
  time_to_event: buildTimeToEvent,
  longitudinal: buildLongitudinal,
  categorical: buildCategorical,
  distribution: buildDistribution,
  scatter: buildScatter
};

/**
 * Build the deterministic R plotting body for a given family + column mapping.
 *
 * @param {Object} args
 * @param {string} args.figureFamily   one of KNOWN_FIGURE_FAMILIES
 * @param {string} [args.specificType] free-text hint (e.g. "cumulative_incidence")
 * @param {Object} args.mapping        column-role mapping chosen by the AI
 * @param {Object} [args.labels]       { x, y } axis labels
 * @param {string[]} [args.availableColumns] column names that actually exist in records
 * @returns {string} R code body (json_data + jsonlite are injected by the caller)
 */
function buildFigureRCode({ figureFamily, specificType, mapping, labels, availableColumns }) {
  const family = String(figureFamily || "").trim();
  const builder = FAMILY_BUILDERS[family];
  if (!builder) {
    throw new Error(
      `Unknown figureFamily "${family}". Must be one of: ${KNOWN_FIGURE_FAMILIES.join(", ")}.`
    );
  }
  if (!mapping || typeof mapping !== "object") {
    throw new Error("Figure mapping object is required to build the R code.");
  }
  return builder({ specificType, mapping: mapping || {}, labels: labels || {}, availableColumns: availableColumns || [] });
}

/**
 * Human-readable schema text injected into the AI prompt so the model knows
 * exactly which family + mapping keys to return. Kept here so the prompt and the
 * builders never drift apart.
 */
function getMappingSchemaHint() {
  return [
    "Available figure families and the mapping keys each one needs:",
    '- "time_to_event" (Kaplan-Meier survival or cumulative incidence): { "timeColumn": <numeric time>, "statusColumn": <1=event, 0=censored>, "groupColumn": <optional treatment arm> }',
    '- "longitudinal" (mean of a measure over time): { "xColumn": <numeric x/time>, "yColumn": <numeric measure>, "groupColumn": <optional> }',
    '- "categorical" (bar / grouped bar): { "categoryColumn": <category>, "valueColumn": <optional numeric; omit to count rows>, "groupColumn": <optional> }',
    '- "distribution" (boxplot): { "valueColumn": <numeric>, "categoryColumn": <optional x category>, "groupColumn": <optional fill> }',
    '- "scatter" (x-y scatter): { "xColumn": <numeric>, "yColumn": <numeric>, "groupColumn": <optional> }'
  ].join("\n");
}

module.exports = {
  KNOWN_FIGURE_FAMILIES,
  buildFigureRCode,
  getMappingSchemaHint
};
