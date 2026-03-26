// -----------------------------------------------------------
// SharePoint REST Helper
// Works in M365 Default Environment (No Premium Licensing)
// -----------------------------------------------------------

/**
 * GET items from a SharePoint list with optional OData query.
 */
export async function spGet(siteUrl, listName, filter = "") {
    const url = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/items${filter}`;

    const response = await fetch(url, {
        method: "GET",
        headers: {
            "Accept": "application/json;odata=nometadata"
        },
        credentials: "include"
    });

    if (!response.ok) {
        throw new Error(`SharePoint GET failed: ${response.status}`);
    }

    return await response.json();
}

/**
 * POST a new item into a SharePoint list.
 */
export async function spCreate(siteUrl, listName, bodyObj) {
    const url = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/items`;

    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata"
        },
        credentials: "include",
        body: JSON.stringify(bodyObj)
    });

    if (!response.ok) {
        throw new Error(`SharePoint POST failed: ${response.status}`);
    }

    return await response.json();
}

/**
 * GET a single item by ID.
 */
export async function spGetByID(siteUrl, listName, id) {
    const url = `${siteUrl}/_api/web/lists/getByTitle('${listName}')/items(${id})`;

    const response = await fetch(url, {
        method: "GET",
        headers: {
            "Accept": "application/json;odata=nometadata"
        },
        credentials: "include"
    });

    if (!response.ok) {
        throw new Error(`SharePoint GET by ID failed: ${response.status}`);
    }

    return await response.json();
}
