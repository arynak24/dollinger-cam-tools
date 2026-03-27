// -------------------------------------------------------
// match.js — simple text matcher for Excel-only add-in
// -------------------------------------------------------

/**
 * Finds matching clauses based on text similarity.
 * This is a lightweight helper used by matchClause()
 * inside taskpane.js.
 */
export function findMatches(searchText, items) {
    if (!searchText || !items) return [];

    const needle = searchText.toLowerCase();

    return items
        .filter(item => item.ClauseText?.toLowerCase().includes(needle))
        .map(item => ({
            ClauseID: item.ID,
            Text: item.ClauseText,
            Category: item.DocType
        }));
}
