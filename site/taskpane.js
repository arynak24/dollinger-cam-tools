/* -----------------------------------------------------------
   CAM Add-In: Taskpane Controller
   Wires UI → CAM modules → Flow/PowerApps → Excel.
------------------------------------------------------------*/

import { buildClauseObject } from "./cam/pipeline.js";
import { findMatchingClauses } from "./cam/match.js";
import { insertClauseIDIntoForm, openLeaseLibraryLink } from "./cam/link.js";
import { postJSON } from "./lib/api.js";

/* -----------------------------------------------------------
   ✅ FLOW / POWERAPPS ENDPOINTS  
   (Replace these when Flow URLs are ready)
------------------------------------------------------------*/
const SAVE_CLAUSE_URL  = "YOUR_FLOW_URL/saveClause";
const MATCH_CLAUSE_URL = "YOUR_FLOW_URL/matchClause";
const HISTORY_URL      = "YOUR_FLOW_URL/getClauseHistory";

/* -----------------------------------------------------------
   ✅ Initialize Add-In
------------------------------------------------------------*/
Office.onReady(() => {
    console.log("✅ CAM Add-In Ready.");
    wireButtons();
});

/* -----------------------------------------------------------
   ✅ Attach UI button handlers
------------------------------------------------------------*/
function wireButtons() {
    document.getElementById("btnSaveClause").onclick = saveClause;
    document.getElementById("btnMatchClause").onclick = matchClause;
    document.getElementById("btnInsertID").onclick = insertClauseID;
}

/* -----------------------------------------------------------
   ✅ SAVE CLAUSE → FLOW / CLAUSEREPOSITORY
------------------------------------------------------------*/
async function saveClause() {
    try {
        const clauseObj = await buildClauseObject();
        console.log("Saving Clause:", clauseObj);

        // POST to Flow
        const response = await postJSON(SAVE_CLAUSE_URL, clauseObj);
        console.log("Flow Save Response:", response);

        clauseObj.ClauseID = response.ClauseID || null;
        clauseObj.LeaseID  = response.LeaseID  || null;

        // Store for linking
        window.__lastClauseHistory = clauseObj;

        renderHistory(clauseObj);
        alert("✅ Clause saved to repository!");

    } catch (err) {
        console.error("❌ Save Clause Error:", err);
        alert("Error saving clause. Check console.");
    }
}

/* -----------------------------------------------------------
   ✅ MATCH EXISTING CLAUSES
------------------------------------------------------------*/
async function matchClause() {
    try {
        const text = document.getElementById("clauseText").value.trim();

        if (!text) {
            alert("Paste or extract clause text first.");
            return;
        }

        console.log("Searching for matches:", text);

        const matches = await findMatchingClauses(text);

        renderMatchResults(matches);

    } catch (err) {
        console.error("❌ Match Error:", err);
        alert("Error searching for matches.");
    }
}

/* -----------------------------------------------------------
   ✅ INSERT CLAUSEID INTO TENANTFORM (selected cell)
------------------------------------------------------------*/
async function insertClauseID() {
    const last = window.__lastClauseHistory;

    if (!last || !last.ClauseID) {
        alert("No ClauseID found. Save or match a clause first.");
        return;
    }

    await insertClauseIDIntoForm(last.ClauseID);
    alert("✅ ClauseID inserted into tenant form.");
}

/* -----------------------------------------------------------
   ✅ HISTORY PANEL — Show current clause object
------------------------------------------------------------*/
function renderHistory(clauseObj) {
    window.__lastClauseHistory = clauseObj;

    const div = document.getElementById("historyContent");

    div.innerHTML = `
        <strong>ClauseID:</strong> ${clauseObj.ClauseID ?? "(pending)"}<br>
        <strong>LeaseID:</strong> ${clauseObj.LeaseID ?? "(none yet)"}<br><br>

        <strong>Clause Text:</strong><br>
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

        ${clauseObj.LeaseID ? 
            `<button id="openLease" class="primary">Open LeaseLibrary Entry</button>` 
            : ""}
    `;

    // Add handler if button exists
    const btn = document.getElementById("openLease");
    if (btn) {
        btn.onclick = () => openLeaseLibraryLink(clauseObj.LeaseID);
    }
}

/* -----------------------------------------------------------
   ✅ HISTORY PANEL — Show match results
------------------------------------------------------------*/
function renderMatchResults(matches) {
    const div = document.getElementById("historyContent");

    if (!matches || matches.length === 0) {
        div.innerHTML = "<p>No matching clauses found.</p>";
        return;
    }

    let html = `<strong>Matching Clauses:</strong><br><br>`;

    matches.forEach((m, i) => {
        html += `
            <div class="match-item">
                <strong>Match #${i + 1}</strong><br>
                ClauseID: ${m.ClauseID}<br>
