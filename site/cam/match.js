import { postJSON } from "../lib/api.js";

/**
 * Query ClauseRepository for possible matches.
 */
export async function findMatchingClauses(text) {
    if (!text) return [];

    try {
        const response = await postJSON("YOUR_FLOW_URL/matchClause", { text });
        return response.matches || [];
    } catch (e) {
        console.error("Match error:", e);
        return [];
    }
}
