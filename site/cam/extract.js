import { getSelectedText } from "../lib/excel.js";

/**
 * Extract raw clause text.
 * If user pasted text, use that.
 * If user selected cells, extract from Excel.
 */
export async function extractClause() {
    const clauseText = document.getElementById("clauseText").value.trim();

    if (clauseText && clauseText.length > 0) {
        return clauseText;
    }

    // Fallback to Excel-selected text
    const excelText = await getSelectedText();
    return excelText || "";
}
