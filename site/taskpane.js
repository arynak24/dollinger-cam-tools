/* -----------------------------------------------------------
   CAM Add-In: Taskpane Controller
   Works with /cam/* modules & /lib/* helpers.
   This file wires your UI → pipeline → repository → Excel.
------------------------------------------------------------*/

import { buildClauseObject } from "./cam/pipeline.js";
import { findMatchingClauses } from "./cam/match.js";
import { insertClauseIDIntoForm, openLeaseLibraryLink } from "./cam/link.js";
import { postJSON } from "./lib/api.js";

/* PowerApps / Flow endpoints (replace when ready) */
const SAVE_CLAUSE_URL = "YOUR_FLOW_URL/saveClause";
const MATCH_CLAUSE_URL = "YOUR_FLOW_URL/matchClause";
const HISTORY_URL     = "YOUR_FLOW_URL/getClauseHistory";

/* -----------------------------------------------------------
   Initialize Add-In
------------------------------------------------------------*/
Office.onReady(() => {
    console.log("CAM Add-In Ready.");
    wireButtons();
});

/* -----------------------------------------------------------
   Button Event Wiring
------------------------------------------------------------*/
function wireButtons() {
    document.getElementById("btnSaveClause").onclick = saveClause;
    document.getElementById("btnMatchClause").onclick = matchClause;
    document.getElementById("btnInsertID").onclick = insertClauseID;
}

/* -----------------------------------------------------------
   BUILD & SAVE CLAUSE
   - Collect all UI fields
   - Build ClauseObject
   - POST to Flow / PowerApps
   - Update UI history panel
------------------------------------------------------------*/
async function saveClause() {
    try {
        const clauseObj = await buildClauseObject();
        console.log("Saving clause →", clauseObj);

        // Send to Flow / ClauseRepository
        const response = await postJSON(SAVE_CLAUSE_URL, clauseObj);
        console.log("Save response:", response);

        // Display confirmation
        alert("Clause saved successfully!");

        // Insert returned ClauseID into hidden memory for next steps
        clauseObj.ClauseID = response.ClauseID;
        clauseObj.LeaseID  = response.LeaseID;

        // Render updated history
        renderHistory(clauseObj);

    } catch (err) {
        console.error("SaveClause Error:", err);
        alert("Error saving clause. Check console.");
    }
}

/* -----------------------------------------------------------
   MATCH EXISTING CLAUSES
   - Uses text from Clause Verbiage box
   - Shows result in history panel
------------------------------------------------------------*/
async function matchClause() {
    try {
        const text = document.getElementById("clauseText").value.trim();
        if (!text) {
            alert("Paste or select clause text first.");
            return;
        }

        console.log("Matching clause text:", text);

        const matches = await findMatchingClauses(text);

        // Update history panel with match results
        renderMatchResults(matches);

    } catch (err) {
        console.error("MatchClause Error:", err);
        alert("Error matching clause. Check console.");
    }
}

/* -----------------------------------------------------------
   INSERT CLAUSE ID INTO SELECTED TENANT FORM CELL
------------------------------------------------------------*/
async function insertClauseID() {
    const lastHistory = window.__lastClauseHistory;

    if (!lastHistory || !lastHistory.ClauseID) {
        alert("No ClauseID available. Save or match a clause first.");
        return;
    }

    await insertClauseIDIntoForm(lastHistory.ClauseID);
    alert("ClauseID inserted into TenantForm.");
}

/* -----------------------------------------------------------
   HISTORY PANEL RENDERING
------------------------------------------------------------*/
function renderHistory(clauseObj) {
    window.__lastClauseHistory = clauseObj;  // Store for linking

    const div = document.getElementById("historyContent");
    div.innerHTML = `
        <strong>ClauseID:</strong> ${clauseObj.ClauseID ?? "(not assigned)"}<br>
        <strong>LeaseID:</strong> ${clauseObj.LeaseID ?? "(not assigned)"}<br><br>

        <strong>Text:</strong><br>
        <pre>${clauseObj.Text}</pre><br>

        <strong>Notes:</strong><br>
        <pre>${clauseObj.Notes}</pre><br>

        <strong>Values:</strong><br>
        Dollars: ${clauseObj.Values.Dollars}<br>
        Percent: ${clauseObj.Values.Percent}<br>
        BaseYear: ${clauseObj.Values.BaseYear}<br>
        Dates: ${clauseObj.Values.Dates}<br>
        Other: ${clauseObj.Values.Other}<br><br>

        <strong>Category:</strong> ${clauseObj.Category}<br>
        <strong>Tags:</strong> ${clauseObj.Tags.join(", ")}<br><br>

        <strong>PDF Page:</strong> ${clauseObj.PageReference}<br>
        <strong>Timestamp:</strong> ${clauseObj.Timestamp}<br>
        <strong>Abstracted By:</strong> ${clauseObj.AbstractedBy}<br><br>

        ${clauseObj.LeaseID ? `<button id="openLease">Open LeaseLibrary Entry</button>` : ""}
    `;

    // Attach event listener if needed
    const btnOpen = document.getElementById("openLease");
    if (btnOpen) {
        btnOpen.onclick = () => openLeaseLibraryLink(clauseObj.LeaseID);
    }
}

/* -----------------------------------------------------------
   SHOW MATCH RESULTS IN HISTORY PANEL
------------------------------------------------------------*/
function renderMatchResults(matches) {
    const div = document.getElementById("historyContent");

    if (!matches || matches.length === 0) {
        div.innerHTML = "<p>No matching clauses found.</p>";
        return;
    }

    let html = `<strong>Matches Found:</strong><br><br>`;

    matches.forEach((m, i) => {
        html += `
            <div class="match-item">
                <strong>Match #${i + 1}</strong><br>
                ClauseID: ${m.ClauseID}<br>
                Category: ${m.Category}<br>
                <pre>${m.Text}</pre>
                <button class="useClause" data-id="${m.ClauseID}">Use This Clause</button>
                <br><br>
            </div>
        `;
    });

    div.innerHTML = html;

    // Hook up all buttons
    const buttons = document.querySelectorAll(".useClause");
    buttons.forEach(btn => {
        btn.onclick = () => {
            const selectedID = btn.getAttribute("data-id");
            applyMatchedClause(selectedID);
        };
    });
}

/* -----------------------------------------------------------
   USE MATCHED CLAUSE
   - Loads clause details into workspace
   - Preps ClauseID for inserting into TenantForm
------------------------------------------------------------*/
async function applyMatchedClause(clauseID) {
    try {
        const data = await postJSON(HISTORY_URL, { id: clauseID });

        // Update UI fields
        document.getElementById("clauseText").value = data.Text;
        document.getElementById("abstractionNotes").value = data.Notes;

        document.getElementById("valueDollars").value = data.Values.Dollars;
        document.getElementById("valuePercent").value = data.Values.Percent;
        document.getElementById("valueBaseYear").value = data.Values.BaseYear;
        document.getElementById("valueDates").value = data.Values.Dates;
        document.getElementById("valueOther").value = data.Values.Other;

        document.getElementById("camCategory").value = data.Category;
        document.getElementById("camTags").value = data.Tags.join(", ");
        document.getElementById("pdfPage").value = data.PageReference;

        // Render to history panel
        renderHistory(data);

        alert("Clause loaded from repository.");

    } catch (err) {
        console.error("applyMatchedClause Error:", err);
        alert("Error loading clause.");
    }
}
