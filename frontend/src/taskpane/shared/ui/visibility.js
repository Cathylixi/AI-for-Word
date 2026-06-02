function showRow(rowEl, show) {
  if (!rowEl) return;
  rowEl.style.display = show ? "block" : "none";
}

function showActions(rowEl, show) {
  if (!rowEl) return;
  rowEl.style.display = show ? "flex" : "none";
}

export { showRow, showActions };
