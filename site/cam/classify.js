/**
 * Suggest a CAM category based on keywords.
 * Always overridable by the user.
 */
export function suggestCategory(text) {
    const lower = text.toLowerCase();

    if (lower.includes("operating expense") || lower.includes("opex"))
        return "Operating Expenses";
    if (lower.includes("tax"))
        return "Taxes";
    if (lower.includes("insurance"))
        return "Insurance";
    if (lower.includes("utility") || lower.includes("electric"))
        return "Utilities";
    if (lower.includes("capital"))
        return "Capital Expenditures";
    if (lower.includes("admin") || lower.includes("management fee"))
        return "Admin/Management Fees";
    if (lower.includes("exclude") || lower.includes("notwithstanding"))
        return "Exclusions";

    return "";
}
