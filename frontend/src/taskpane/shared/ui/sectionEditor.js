function showEditor(rowEl, show) {
  if (!rowEl) return;
  rowEl.style.display = show ? "block" : "none";
}

function setEditorText(editor, text) {
  if (editor) editor.value = text || "";
}

function getEditorText(editor) {
  return editor?.value || "";
}

function clearEditor(editor) {
  if (editor) editor.value = "";
}

function hasEditorText(editor) {
  return !!(editor && editor.value && editor.value.trim());
}

function bindEditorInput(editor, handler) {
  if (!editor) return;
  editor.addEventListener("input", handler);
}

export {
  showEditor,
  setEditorText,
  getEditorText,
  clearEditor,
  hasEditorText,
  bindEditorInput
};
