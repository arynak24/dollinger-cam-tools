import { writeToSelectedCell } from "../lib/excel.js";

/**
 * Inserts ClauseID into the selected row of TenantForm.
 */
export async function insertClauseIDIntoForm(clauseID) {
    if (!clauseID) {
        alert("No ClauseID available.");
        return;
    }
    await writeToSelectedCell(clauseID);
}

/**
 * Opens LeaseLibrary entry for quick review.
 */
export function openLeaseLibraryLink(leaseID) {
    if (!leaseID) {
        alert("No LeaseLibrary ID available.");
        return;
    }
    const url = `https://YOUR_SHAREPOINT_SITE/LeaseLibrary/View.aspx?ID=${leaseID}`;
    window.open(url, "_blank");
}
