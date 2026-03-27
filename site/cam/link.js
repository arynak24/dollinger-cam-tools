// -------------------------------------------------------
// link.js — Excel-only insertion helper
// -------------------------------------------------------

// Write a ClauseID into the currently selected Excel cell
export async function insertClauseIDIntoForm(clauseID) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.getActiveWorksheet();
            const range = context.workbook.getSelectedRange();
            range.values = [[clauseID]];
            await context.sync();
        });

        console.log("✅ ClauseID inserted:", clauseID);
    } catch (err) {
        console.error("❌ Error inserting ClauseID:", err);
        throw err;
    }
}
