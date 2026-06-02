function setAiEnabled(btn, enabled) {
  if (btn) btn.disabled = !enabled;
}

function bindAiClick(btn, handler) {
  if (!btn) return;
  btn.addEventListener("click", handler);
}

export { setAiEnabled, bindAiClick };
