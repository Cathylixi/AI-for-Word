function setStudyOptions(selectEl, studies) {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Select a study...</option>`;
  studies.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    selectEl.appendChild(opt);
  });
}

function setSectionOptions(selectEl, entries) {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Select a section...</option>`;
  entries.forEach((e) => {
    const opt = document.createElement("option");
    opt.value = e.number;
    const displayNumber = e.displayNumber || e.number;
    opt.textContent = `${displayNumber} ${e.title}`;
    selectEl.appendChild(opt);
  });
}

// Build the "Start working on" dropdown with Title Page + section list.
function setWorkItemOptions(selectEl, entries) {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Select a section...</option>`;

  const titleOpt = document.createElement("option");
  titleOpt.value = "__TITLE_PAGE__";
  titleOpt.textContent = "Title Page";
  selectEl.appendChild(titleOpt);

  const tocOpt = document.createElement("option");
  tocOpt.value = "__TOC__";
  tocOpt.textContent = "Table of Contents";
  selectEl.appendChild(tocOpt);

  entries.forEach((e) => {
    const opt = document.createElement("option");
    opt.value = e.number;
    const displayNumber = e.displayNumber || e.number;
    opt.textContent = `${displayNumber} ${e.title}`;
    selectEl.appendChild(opt);
  });
}

function bindChange(selectEl, handler) {
  if (!selectEl) return;
  selectEl.addEventListener("change", handler);
}

export { setStudyOptions, setSectionOptions, setWorkItemOptions, bindChange };
