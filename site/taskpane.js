// ------------------------------------------------------------
// CAM Add-In: Taskpane Controller (SharePoint REST Architecture)
// ------------------------------------------------------------

// Imports
import { buildClauseObject } from "./cam/pipeline.js";
import { insertClauseIDIntoForm } from "./cam/link.js";
import { spCreate, spGet, spGetByID } from "./lib/sharepoint.js";

// -----------------------
// SharePoint Settings
// -----------------------
const SP_SITE = "https://dpmproperties.sharepoint.com/sites/DollingerCAMProject";
const SP_LIST = "ClauseRepository";

// ------------------------------------------------------------
// Initialize Add-In
// ------------------------------------------------------------
Office.onReady(() => {
    console.log("✅ CAM Add-In Ready");
    wireButtons();
});

// Wire UI buttons to functions
function wireButtons() {
    document.getElementById("btnSaveClause").onclick = saveClause;
    document.getElementById("btnMatchClause").onclick = matchClause;
    document.getElementById("btnInsertID").onclick = insertClauseID;
}

// ------------------------------------------------------------
// SAVE CLAUSE (Create item in SharePoint)
// ------------------------------------------------------------
async function saveClause() {
    try {
        // Build clause object from UI fields
        const clause = await buildClauseObject();

        // Map to SharePoint column names
        const body = {
            ClauseText: clause.Text,
            Notes: clause.Notes,
            PageNumber: clause.PageReference,
            DocType: clause.Category,
            Tags: clause.Tags.join(", "),
            Dollars: clause.Values.Dollars,
            Percent: clause.Values.Percent,
            BaseYear: clause.Values.BaseYear,
            Dates: clause.Values.Dates,
            OtherValues: clause.Values.Other
        };

        // Create item in SharePoint
        const item = await spCreate(SP_SITE, SP_LIST, body);

        // Store new ClauseID
        clause.ClauseID = item.ID;
        window.__lastClauseHistory = clause;

        renderHistory(clause);

        alert(`✅ Clause saved! New ClauseID = ${item.ID}`);

    } catch (err) {
        console.error("❌ SaveClause Error:", err);
        alert("Error saving clause to SharePoint. See console.");
    }
}

// ------------------------------------------------------------
// MATCH EXISTING CLAUSES (SharePoint GET + client filter)
// ------------------------------------------------------------
async function matchClause() {
    try {
        const text = document.getElementById("clauseText").value.trim();
        if (!text) {
            alert("Paste clause text to match.");
            return;
        }

        // Pull minimal fields for faster matching
        const results = await spGet(SP_SITE, SP_LIST, "?$select=ID,ClauseText,DocType");

        // Client-side filtering because ClauseText is multiline
        const matches = results.value
            .filter(item =>
                item.ClauseText?.toLowerCase().includes(text.toLowerCase())
            )
            .map(item => ({
                ClauseID: item.ID,
                Text: item.ClauseText,
                Category: item.DocType
            }));

        renderMatchResults(matches);

    } catch (err) {
        console.error("❌ MatchClause Error:", err);
        alert("Could not query SharePoint. See console.");
    }
}

// ------------------------------------------------------------
// GET CLAUSE DETAILS (SharePoint GET by ID)
// ------------------------------------------------------------
async function applyMatchedClause(clauseID) {
    try {
        const item = await spGetByID(SP_SITE, SP_LIST, clauseID);

        const clause = {
            ClauseID: item.ID,
            Text: item.ClauseText,
            Notes: item.Notes,
            Category: item.DocType,
            Tags: item.Tags,
            PageReference: item.PageNumber,
            Values: {
                Dollars: item.Dollars,
                Percent: item.Percent,
                BaseYear: item.BaseYear,
                Dates: item.Dates,
                Other: item.OtherValues
            }
        };

        // Save for Excel insertion
        window.__lastClauseHistory = clause;

        // Populate UI fields
        document.getElementById("clauseText").value = clause.Text;
        document.getElementById("abstractionNotes").value = clause.Notes;

        document.getElementById("valueDollars").value = clause.Values.Dollars;
        document.getElementById("valuePercent").value = clause.Values.Percent;
        document.getElementById("valueBaseYear").value = clause.Values.BaseYear;
        document.getElementById("valueDates").value = clause.Values.Dates;
        document.getElementById("valueOther").value = clause.Values.Other;

        document.getElementById("camCategory").value = clause.Category;
        document.getElementById("camTags").value = clause.Tags;
        document.getElementById("pdfPage").value = clause.PageReference;

        renderHistory(clause);

        alert(`✅ Loaded ClauseID ${clauseID}`);

    } catch (err) {
        console.error("❌ GetClauseHistory Error:", err);
        alert("Could not load clause details. See console.");
    }
}

// ------------------------------------------------------------
// INSERT CLAUSEID INTO EXCEL
// ------------------------------------------------------------
async function insertClauseID() {
    const last = window.__lastClauseHistory;

    if (!last || !last.ClauseID) {
        alert("No ClauseID found. Save or load a clause first.");
        return;
    }

    await insertClauseIDIntoForm(last.ClauseID);
    alert("✅ ClauseID written into the tenant form.");
}

// ------------------------------------------------------------
// HISTORY PANEL
// ------------------------------------------------------------
function renderHistory(clause) {
    const div = document.getElementById("historyContent");

    div.innerHTML = `
        <strong>ClauseID:</strong> ${clause.ClauseID ?? "(pending)"}<br><br>

        <strong>Clause Text:</strong><br>
        <pre>${clause.Text}</pre><br>

        <strong>Notes:</strong><br>
        <pre>${clause.Notes}</pre><br>

        <strong>Values:</strong><br>
        Dollars: ${clause.Values.Dollars}<br>
        Percent: ${clause.Values.Percent}<br>
        Base Year: ${clause.Values.BaseYear}<br>
        Dates: ${clause.Values.Dates}<br>
        Other: ${clause.Values.Other}<br><br>

        <strong>Category:</strong> ${clause.Category}<br>
        <strong>Tags:</strong> ${clause.Tags}<br>
        <strong>PDF Page:</strong> ${clause.PageReference}<br>
    `;
}

// ------------------------------------------------------------
// MATCH RESULTS PANEL
// ------------------------------------------------------------
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
                Category: ${m.Category}<br>
                <pre>${m.Text}</pre>
                <button class="useClauseBtn" data-id="${m.ClauseID}">
                    ✅ Use This Clause
                </button>
                <hr>
            </div>
        `;
    });

    div.innerHTML = html;

    document.querySelectorAll(".useClauseBtn").forEach(btn => {
        btn.onclick = () => applyMatchedClause(btn.dataset.id);
    });
}
