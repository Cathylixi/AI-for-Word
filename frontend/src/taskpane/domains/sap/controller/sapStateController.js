import {
  readAllSapBodySections,
  writeAllSapBodySections,
  writeSectionBody // Keep for individual updates if needed
} from "../word/sections";
import {
  readTitlePageFields,
  writeTitlePageFields
} from "../word/titlePage";
import {
  readAbbreviationsFromPlaceholder,
  generateAbbreviationsAtPlaceholder
} from "../word/listOfAbbreviations";
import {
  clearSapTemplate,
  generateSapFrontMatter,
  generateSapAfterToc
} from "../word/templateSap";
import {
  loadSapState,
  saveSapState
} from "../api/sapState";
import { setStatus } from "../../../shared/ui/status";
import { getDom } from "../../../shared/ui/dom";
import { showRow } from "../../../shared/ui/visibility";
import { setWorkItemOptions } from "../../../shared/ui/selectors";

async function buildSapSaveState(studyNumber, insertedSapEntries) {
  // 1. Sections content
  const sectionsData = await readAllSapBodySections({ studyNumber });
  // sectionsData is already [{ sectionNumber, text }]

  // 2. Title Page fields
  const titlePageData = await readTitlePageFields({ studyNumber });

  // 3. LOA items
  const loaItems = await readAbbreviationsFromPlaceholder();

  return {
    saved: true,
    lastSavedAt: new Date().toISOString(),
    entries: insertedSapEntries || [], // The structure (numbers, titles, displayNumbers)
    sections: sectionsData, // Store as Array to avoid MongoDB key issues with dots (e.g. "1.1")
    titlePage: titlePageData,
    loa: { items: loaItems },
    schemaVersion: 2 // Bump version for array-based sections
  };
}

async function handleSaveSap({ studyNumber, insertedSapEntries }) {
  const dom = getDom();
  if (!studyNumber) return;
  try {
    setStatus(dom.status, "Saving to database...");
    const state = await buildSapSaveState(studyNumber, insertedSapEntries);
    await saveSapState({ studyNumber, state });
    setStatus(dom.status, "Saved successfully.");
  } catch (e) {
    setStatus(dom.status, `Save failed: ${e.message || e}`);
  }
}

async function restoreSapFromState({ studyNumber, state, onRestored, onProgress }) {
  const dom = getDom();
  if (!state || !state.entries) {
    setStatus(dom.status, "Invalid save state.");
    return;
  }

  const progress = (text) => {
    try {
      if (onProgress) onProgress(text);
    } catch (e) {}
  };

  try {
    setStatus(dom.status, "Restoring SAP document...");
    progress("Clearing existing template...");
    
    // 1. Clear existing
    await clearSapTemplate();

    const entries = state.entries;

    // 2. Re-generate structure (Front Matter)
    progress("Generating title/signature, revision history, and TOC...");
    await generateSapFrontMatter({ studyNumber, entries });

    // 3. Re-generate structure (Body)
    progress("Generating LOA placeholder and section structure...");
    await generateSapAfterToc({ studyNumber, entries });

    // 4. Fill Title Page
    if (state.titlePage) {
      progress("Restoring title page fields...");
      await writeTitlePageFields({ studyNumber, values: state.titlePage });
    }

    // 5. Fill LOA
    if (state.loa?.items && state.loa.items.length > 0) {
      progress("Restoring abbreviations table...");
      await generateAbbreviationsAtPlaceholder(state.loa.items);
    }

    // 6. Fill Sections (Batch write for performance and robustness)
    progress("Restoring section contents...");
    let sectionsToRestore = [];
    if (Array.isArray(state.sections)) {
      // V2 Schema: Array
      sectionsToRestore = state.sections;
    } else if (state.sections && typeof state.sections === "object") {
      // V1 Schema: Object/Map (Legacy support just in case)
      sectionsToRestore = Object.entries(state.sections).map(([k, v]) => ({
        sectionNumber: k,
        text: v
      }));
    }

    if (sectionsToRestore.length > 0) {
      await writeAllSapBodySections({
        studyNumber,
        sectionsData: sectionsToRestore
      });
    }

    // 7. Restore UI
    progress("Finalizing UI...");
    setWorkItemOptions(dom.sectionSelect, entries);
    
    if (onRestored) {
        onRestored(entries);
    }

    setStatus(dom.status, "Restored successfully.");
    progress("Done");
  } catch (e) {
    setStatus(dom.status, `Restore failed: ${e.message || e}`);
    console.error(e);
  }
}

async function checkAndPromptRestore({ studyNumber, onRestoreConfirmed, onCancel }) {
  const dom = getDom();
  try {
    const resp = await loadSapState({ studyNumber });
    if (resp.success && resp.data && resp.data.saved) {
      // Show modal
      const modal = document.getElementById("restore-modal");
      if (modal) {
        modal.style.display = "flex";
        
        const btnYes = document.getElementById("btn-restore-yes");
        const btnNo = document.getElementById("btn-restore-no");
        const dateSpan = document.getElementById("restore-date");
        
        if (dateSpan) {
            const d = new Date(resp.data.lastSavedAt);
            dateSpan.textContent = d.toLocaleString();
        }

        // One-time handlers
        const handleYes = () => {
          modal.style.display = "none";
          cleanup();
          // Show progress modal during restore
          if (dom.restoreProgressModal) dom.restoreProgressModal.style.display = "flex";
          if (dom.restoreProgressText) dom.restoreProgressText.textContent = "Starting...";

          const onProgress = (text) => {
            if (dom.restoreProgressText) dom.restoreProgressText.textContent = text || "";
          };

          restoreSapFromState({ 
            studyNumber, 
            state: resp.data, 
            onRestored: onRestoreConfirmed,
            onProgress
          }).finally(() => {
            if (dom.restoreProgressModal) dom.restoreProgressModal.style.display = "none";
          });
        };

        const handleNo = () => {
          modal.style.display = "none";
          cleanup();
          if (onCancel) onCancel();
        };

        const cleanup = () => {
            btnYes.removeEventListener("click", handleYes);
            btnNo.removeEventListener("click", handleNo);
        };

        btnYes.addEventListener("click", handleYes);
        btnNo.addEventListener("click", handleNo);
      }
    } else {
        if (onCancel) onCancel(); // No save found, proceed as new
    }
  } catch (e) {
    console.warn("Failed to check save state", e);
    if (onCancel) onCancel();
  }
}

export { handleSaveSap, checkAndPromptRestore };
