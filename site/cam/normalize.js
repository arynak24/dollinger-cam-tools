/**
 * Minimal normalization — preserves user intentions.
 * Removes invisible characters & trims excess whitespace.
 */
export function normalizeClause(text) {
    if (!text) return "";

    return text
        .replace(/\r\n/g, "\n")
        .replace(/\t/g, "    ")
        .trim();
}
