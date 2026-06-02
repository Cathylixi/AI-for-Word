const TAG_PREFIX_MOCKTFL_TITLE = "mocktfl-title:";
const TAG_PREFIX_MOCKTFL_BODY = "mocktfl-body:";

function buildMockTflKey({ type, number }) {
  return `${String(type || "").trim().toUpperCase()}:${String(number || "").trim()}`;
}

function buildMockTflTitleTag({ studyNumber, type, number }) {
  return `${TAG_PREFIX_MOCKTFL_TITLE}${studyNumber}:${buildMockTflKey({ type, number })}`;
}

function buildMockTflBodyTag({ studyNumber, type, number }) {
  return `${TAG_PREFIX_MOCKTFL_BODY}${studyNumber}:${buildMockTflKey({ type, number })}`;
}

export {
  TAG_PREFIX_MOCKTFL_TITLE,
  TAG_PREFIX_MOCKTFL_BODY,
  buildMockTflKey,
  buildMockTflTitleTag,
  buildMockTflBodyTag
};
