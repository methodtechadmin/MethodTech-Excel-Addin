/* global OfficeRuntime, console, fetch */

import { getDjangoBaseUrl } from "../auth/config";
import { djangoFetch } from "../auth/http";

const CATALOG_STORAGE_KEY = "methodtech.catalog";
const CATALOG_PATH = "/api/microsoft/excel/catalog/";

function catalogListFromResponse(data) {
  if (Array.isArray(data)) {
    return data;
  }
  if (data && Array.isArray(data.functions)) {
    return data.functions;
  }
  if (data && Array.isArray(data.results)) {
    return data.results;
  }
  return null;
}

async function storeCatalog(list) {
  await saveCatalog(list);
  console.log(`MethodTech catalog loaded: ${list.length} function(s)`);
  return list;
}

/**
 * Load the function list without a MethodTech session.
 * Excel must associate these names before a cell calculates, including when
 * Edge cannot complete a silent sign-in.
 * @returns {Promise<object[]>}
 */
export async function fetchPublicCatalog() {
  const base = getDjangoBaseUrl().replace(/\/$/, "");
  const response = await fetch(`${base}${CATALOG_PATH}`, {
    method: "GET",
    headers: { Accept: "application/json" },
  });
  const text = await response.text();
  let data = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = null;
    }
  }
  if (!response.ok) {
    const message =
      (data && (data.detail || data.error || data.message)) ||
      `Catalog request failed (${response.status})`;
    throw new Error(typeof message === "string" ? message : JSON.stringify(message));
  }

  const list = catalogListFromResponse(data);
  if (!list) {
    throw new Error("Catalog response was not a function list.");
  }
  return storeCatalog(list);
}

/**
 * @returns {Promise<object[]>}
 */
export async function fetchCatalog() {
  const data = await djangoFetch(CATALOG_PATH, {
    method: "GET",
  });

  const list = catalogListFromResponse(data);
  if (!list) {
    throw new Error("Catalog response was not a function list.");
  }
  return storeCatalog(list);
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
