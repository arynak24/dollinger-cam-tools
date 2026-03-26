Office.onReady(() => {
    console.log("CAM Add-In Ready.");
});

document.getElementById("btnSaveClause").onclick = saveClause;
document.getElementById("btnMatchClause").onclick = matchClause;
document.getElementById("btnInsertID").onclick = insertClauseID;

/* --------------------------
    COLLECT CLAUSE OBJECT
--------------------------- */
function collectClauseData() {
    return {
        text: document.getElementById("clauseText").value,
        notes: document.getElementById("abstractionNotes").value,

        values: {
            dollars: document.getElementById("valueDollars").value,
            percent: document.getElementById("valuePercent").value,
            baseYear: document.getElementById("valueBaseYear").value,
            dates: document.getElementById("valueDates").value,
            other: document.getElementById("valueOther").value
        },

        category: document.getElementById("camCategory").value,
        tags: document.getElementById("camTags").value.split(",").map(t => t.trim()),
        pdfPage: document.getElementById("pdfPage").value,

        timestamp: new Date().toISOString()
    };
}

/* --------------------------
    SAVE CLAUSE (FLOW/POWERAPPS)
--------------------------- */
async function saveClause() {
    const clause = collectClauseData();
    console.log("Saving clause:", clause);

    // TODO: POST clause to your PowerApps/Dataverse API endpoint
    // fetch("FLOW_URL", { method: "POST", body: JSON.stringify(clause) })

    alert("Stub: Clause saved (send to Flow).");
}

/* --------------------------
    MATCH EXISTING CLAUSES
--------------------------- */
async function matchClause() {
    const text = document.getElementById("clauseText").value;

    // TODO: Query your ClauseRepository API
    console.log("Matching clause text:", text);

    alert("Stub: Would query ClauseRepository and display matches.");
}

/* --------------------------
    INSERT CLAUSE ID IN EXCEL FORM
--------------------------- */
async function insertClauseID() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getSelectedRange();

        // TODO: Replace with real ClauseID returned from ClauseRepository
        range.values = [["CLAUSE-ID-1234"]];

        await context.sync();
    });

    alert("Stub: Inserted ClauseID into selected cell.");
}
