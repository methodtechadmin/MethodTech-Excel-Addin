/* global fetch, console */

import { catalogToExcelFunctions } from "./registerFunctions";

/**
 * Push catalog metadata to the local dev server so /functions.json
 * includes every MTECH.* name for Excel IntelliSense.
 * Excel only reloads metadata on add-in restart.
 */
export async function publishCatalogMetadata(catalog) {
  const functions = catalogToExcelFunctions(catalog).filter((item) => item && item.id && item.name);

  try {
    const response = await fetch("/__methodtech/catalog-metadata", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify({ functions }),
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `Metadata publish failed (${response.status})`);
    }

    const result = await response.json();
    console.log(`Published ${result.count || functions.length} catalog function(s) to functions.json metadata.`);
    return { updated: true, count: result.count || functions.length };
  } catch (error) {
    console.warn("Could not publish catalog metadata for IntelliSense:", error);
    return { updated: false, count: 0, error: error.message };
  }
}
