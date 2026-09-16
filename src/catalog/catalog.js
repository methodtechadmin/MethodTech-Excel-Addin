/* global OfficeRuntime, console */

import { djangoFetch } from "../auth/http";

const CATALOG_STORAGE_KEY = "methodtech.catalog";

/**
 * @returns {Promise<object[]>}
 */
export async function fetchCatalog() {
  const data = await djangoFetch("/api/microsoft/excel/catalog/", {
    method: "GET",
  });

  const list = Array.isArray(data)
    ? data
    : Array.isArray(data.functions)
      ? data.functions
      : Array.isArray(data.results)
        ? data.results
        : null;

  if (!list) {
    throw new Error("Catalog response was not a function list.");
  }

  await saveCatalog(list);
  console.log(`MethodTech catalog loaded: ${list.length} function(s)`);
  return list;
}

export async function saveCatalog(list) {
  globalThis.__methodTechCatalog = list;

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(CATALOG_STORAGE_KEY, JSON.stringify(list));
    }
  } catch (error) {
    console.warn("Could not persist catalog.", error);
  }
}

export async function getCatalog() {
  if (Array.isArray(globalThis.__methodTechCatalog)) {
    return globalThis.__methodTechCatalog;
  }

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const raw = await OfficeRuntime.storage.getItem(CATALOG_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        globalThis.__methodTechCatalog = parsed;
        return parsed;
      }
    }
  } catch (error) {
    console.warn("Could not read saved catalog.", error);
  }

  return [];
}

export async function getCatalogEntry(idOrName) {
  const catalog = await getCatalog();
  const key = String(idOrName || "").toUpperCase();
  return (
    catalog.find(
      (item) =>
        String(item.id || "").toUpperCase() === key ||
        String(item.name || "").toUpperCase() === key ||
        String(item.method_name || "").toUpperCase() === key
    ) || null
  );
}
