# Taskpane Architecture (AI-for-Word)

This folder contains the **Word taskpane frontend**. The codebase is organized by responsibility so that new features can be added predictably.

---

## `taskpane.html`

- **Purpose**: Minimal static markup and mount points (containers) for the taskpane UI.
- **Rule of thumb**: Keep this file simple; most UI behavior should live in `ui/` + `controller/`.

---

## `taskpane.js`

- **Purpose**: Entry point.
- **What it does**: Calls `initApp()` from `controller/appController.js` after `Office.onReady()`.

---

## `controller/`

- **Purpose**: **State machine / orchestration layer**.
- **What belongs here**:
  - Current state (selected `studyNumber`, `task`, current work item like `Title Page` vs a section)
  - Event wiring (dropdown changes, button clicks)
  - High-level flow (e.g., when task becomes `SAP`, generate the template; when user clicks AI, call the right API; when user clicks Insert, write to Word)
- **What does NOT belong here**:
  - Low-level DOM creation (put in `ui/`)
  - Low-level `Word.run` logic (put in `word/`)
  - Raw `fetch` calls (put in `api/`)

---

## `api/`

- **Purpose**: HTTP client wrappers for calling the backend (`https://localhost:4100`).
- **What belongs here**:
  - `client.js`: shared `apiGet` / `apiPost`
  - `studies.js`: study-related endpoints (e.g., existing study numbers)
  - `references.js`: reference-related endpoints (e.g., SAP_Example)
  - `generation.js`: AI-related endpoints (e.g., generate section text, generate title page metadata)
- **Rule of thumb**: Functions here should return plain JSON/values and throw meaningful errors.

---

## `word/`

- **Purpose**: All **Word document manipulation** (Office JS) lives here.
- **What belongs here**:
  - Template creation (inserting title page, inserting section headings + content controls)
  - Reading/writing content controls by tag
  - Clearing template content controls
- **Files**:
  - `templateSap.js`: generate/clear the overall SAP template
  - `titlePage.js`: insert and write/clear title page fields (content controls tagged as `sap-meta:<study>:<field>`)
  - `sections.js`: select/read/write section body content controls (tags like `sap-body:<study>:<sectionNumber>`)
  - `constants.js`: tag prefixes (`sap-title:`, `sap-body:`, `sap-meta:`)
- **Rule of thumb**: No `fetch` and no DOM changes in this folder—only `Word.run`.

---

## `ui/`

- **Purpose**: Taskpane UI helpers (DOM access, show/hide, enable/disable, editors).
- **What belongs here**:
  - DOM lookups (`dom.js`)
  - Status rendering (`status.js`)
  - Show/hide helpers (`visibility.js`)
  - Dropdown population helpers (`selectors.js`)
  - Editor widgets:
    - `sectionEditor.js`: the large textarea for normal SAP sections
    - `titlePageEditor.js`: the structured fields for the Title Page
  - Button helpers (`aiButton.js`, `actionsBar.js`)
- **Rule of thumb**: No `Word.run` and no backend calls here.

---

## `utils/`

- **Purpose**: Shared utilities (pure functions).
- **Example**:
  - `debounce.js`: debounce helper for input events

---

## Where to add new code (quick guide)

- **New backend endpoint?** Add a wrapper in `api/`.
- **New Word document behavior?** Add/extend functions in `word/`.
- **New UI widget or panel?** Add a new module in `ui/`.
- **New workflow / button behavior / state transitions?** Update `controller/appController.js`.

