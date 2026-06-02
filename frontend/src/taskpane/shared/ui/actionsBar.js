function setInsertEnabled(btn, enabled) {
  if (btn) btn.disabled = !enabled;
}

function setClearEnabled(btn, enabled) {
  if (btn) btn.disabled = !enabled;
}

function bindInsert(btn, handler) {
  if (!btn) return;
  btn.addEventListener("click", handler);
}

function bindClear(btn, handler) {
  if (!btn) return;
  btn.addEventListener("click", handler);
}

export { setInsertEnabled, setClearEnabled, bindInsert, bindClear };
