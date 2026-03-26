Office.onReady(() => {
    console.log("CAM Add-In Ready");
});

document.getElementById("btnSaveClause").onclick = saveClause;
document.getElementById("btnMatchClause").onclick = matchClause;
document.getElementById("btnInsertID").onclick = insertClauseID;

async function saveClause() {
    const clauseObj = collectClauseData();
    console.log("Saving clause:", clauseObj);

    // TODO: send to PowerApps / Flow
    alert("Stub: Clause would be saved now.");
}

async function matchClause() {
    const clauseText = document.getElementById("clauseText").value;
    console.log("Matching clause:", clauseText);

    // TODO: query ClauseRepository
    alert("Stub: Would search ClauseRepository for matches.");
}

async function insertClauseID() {
    // TODO: when ClauseRepository returns a ClauseID, insert into Excel
    alert("Stub: Would insert ClauseID into Excel.");
}

function collectClauseData() {
    return {
        text: document.getElementById("clauseText").value,
        notes: document.getElementById("abstractionNotes").value,
        values: {
            dollars: document.getElementById("valueDollars").value,
            percent: document.getElementById("valuePercent").value,
            baseYear: document.getElementById("valueBaseYear").value,
            dates: document.getElementById("valueDates").value,
            other: document.getElementById("valueOther").value,
        },
        category: document.getElementById("camCategory").value,
        tags: document.getElementById("camTags").value.split(",").map(t => t.trim()),
        timestamp: new Date().toISOString()
    };
}
