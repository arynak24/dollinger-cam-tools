import { extractClause }
    from "https://zealous-pond-05d3aaf10.1.azurestaticapps.net/cam/extract.js";

import { normalizeClause } 
    from "https://zealous-pond-05d3aaf10.1.azurestaticapps.net/cam/normalize.js";

import { suggestCategory } 
    from "https://zealous-pond-05d3aaf10.1.azurestaticapps.net/cam/classify.js";

/**
 * Build final ClauseObject used across:
 * - PowerApps ClauseRepository
 * - LeaseLibrary
 * - TenantForms
 */
export async function buildClauseObject() {
    const rawText = await extractClause();
    const normalized = normalizeClause(rawText);
    const autoCategory = suggestCategory(normalized);

    return {
        ClauseID: null,       // Assigned after upload to Flow
        LeaseID: null,        // Optional link to LeaseLibrary
        Source: "TenantForm",

        Text: normalized,
        Notes: document.getElementById("abstractionNotes").value,

        Values: {
            Dollars: document.getElementById("valueDollars").value,
            Percent: document.getElementById("valuePercent").value,
            BaseYear: document.getElementById("valueBaseYear").value,
            Dates: document.getElementById("valueDates").value,
            Other: document.getElementById("valueOther").value,
        },

        Category: document.getElementById("camCategory").value || autoCategory,
        Tags: document.getElementById("camTags").value.split(",").map(t => t.trim()).filter(Boolean),

        PageReference: document.getElementById("pdfPage").value,
        Timestamp: new Date().toISOString(),
        AbstractedBy: "Aryn"
    };
}
