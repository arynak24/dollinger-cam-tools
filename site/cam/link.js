import { writeToSelectedCell }
    from "https://zealous-pond-05d3aaf10.1.azurestaticapps.net/lib/excel.js";

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
