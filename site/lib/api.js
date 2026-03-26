/**
 * POST JSON to an API endpoint (PowerApps / Flow).
 */
export async function postJSON(url, data) {
    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
    });

    if (!response.ok) {
        throw new Error(`API POST failed: ${response.status}`);
    }

    return await response.json();
}
