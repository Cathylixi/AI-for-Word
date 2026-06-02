function setStatus(statusEl, text) {
  if (statusEl) statusEl.textContent = text || "";
}

export { setStatus };
