/* global CustomFunctions, console */

import { getDjangoBaseUrl } from "../auth/config";
import { djangoFetch } from "../auth/http";
import { getAuthHeadersContext } from "../auth/session";
import { getCatalog, getCatalogEntry } from "./catalog";

/**
 * Flatten Excel range values into a plain JS array when needed.
 */
function normalizeParamValue(value, paramType) {
  if (value === null || value === undefined || value === "") {
    return undefined;
  }

  if (paramType === "array" && Array.isArray(value)) {
    if (value.length > 0 && Array.isArray(value[0])) {
      return value.map((row) => (row.length === 1 ? row[0] : row));
    }
    return value;
  }

  if (paramType === "boolean") {
    if (typeof value === "boolean") {
      return value;
    }
    if (typeof value === "string") {
      const lowered = value.trim().toLowerCase();
      if (lowered === "true" || lowered === "1" || lowered === "yes") {
        return true;
      }
      if (lowered === "false" || lowered === "0" || lowered === "no") {
        return false;
      }
    }
    return Boolean(value);
  }

  return value;
}

/**
 * Turn API result into something Excel can show in cells (value or 2D spill range).
 */
function formatResultForExcel(result) {
  if (result === null || result === undefined) {
    return "";
  }

  // Common API wrappers: { data: ... } / { result: ... } / { results: ... }
  if (result && typeof result === "object" && !Array.isArray(result)) {
    if (Object.prototype.hasOwnProperty.call(result, "data")) {
      return formatResultForExcel(result.data);
    }
    if (Object.prototype.hasOwnProperty.call(result, "result")) {
      return formatResultForExcel(result.result);
    }
    if (Object.prototype.hasOwnProperty.call(result, "results")) {
      return formatResultForExcel(result.results);
    }
  }

  if (typeof result !== "object") {
    return result;
  }

  if (Array.isArray(result)) {
    if (result.length === 0) {
      return "";
    }

    // Already a matrix
    if (Array.isArray(result[0])) {
      return result;
    }

    // Array of objects → header row + value rows for Excel spill
    if (result[0] !== null && typeof result[0] === "object") {
      const keys = Object.keys(result[0]);
      return [keys, ...result.map((row) => keys.map((key) => {
        const value = row[key];
        if (value === null || value === undefined) {
          return "";
        }
        if (typeof value === "object") {
          return JSON.stringify(value);
        }
        return value;
      }))];
    }

    // Simple 1D list → column
    return result.map((value) => [value]);
  }

  // Plain object → two-column key/value
  return Object.keys(result).map((key) => {
    const value = result[key];
    return [key, value !== null && typeof value === "object" ? JSON.stringify(value) : value];
  });
}

/**
 * Every METHODTECH.* call goes through /api/microsoft/excel/execute/
 * with the catalog function JSON + Excel argument data.
 */
export async function invokeCatalogFunction(entry, args) {
  if (!entry || !(entry.id || entry.name)) {
    throw new Error("Catalog entry is missing id/name.");
  }

  const { username } = await getAuthHeadersContext();
  const params = Array.isArray(entry.parameters) ? entry.parameters : [];
  const data = {};

  params.forEach((param, index) => {
    const raw = args[index];
    const normalized = normalizeParamValue(raw, param.type);
    if (normalized === undefined) {
      if (param.required) {
        throw new Error(`Missing required parameter: ${param.name}`);
      }
      return;
    }
    data[param.name] = normalized;
  });

  const payload = {
    id: entry.id,
    name: entry.name,
    description: entry.description,
    service: entry.service,
    method_name: entry.method_name,
    url_name: entry.url_name,
    http_method: entry.http_method,
    parameters: entry.parameters,
    result: entry.result,
    username,
    data,
  };

  const response = await djangoFetch("/api/microsoft/excel/execute/", {
    method: "POST",
    body: JSON.stringify(payload),
  });

  return formatResultForExcel(response);
}

/**
 * Invoke by catalog id/name, e.g. "MARKETCAP".
 */
export async function invokeCatalogById(idOrName, args) {
  let entry = await getCatalogEntry(idOrName);
  if (!entry) {
    throw new Error(`Function ${idOrName} is not in the MethodTech catalog. Open the task pane to refresh.`);
  }
  return invokeCatalogFunction(entry, args);
}

/**
 * Map catalog parameter types to Excel custom function metadata.
 */
export function catalogParamToExcel(param) {
  const excelParam = {
    name: param.name,
    description: param.description || param.name,
    optional: param.required === false,
  };

  switch (param.type) {
    case "number":
      excelParam.type = "number";
      break;
    case "boolean":
      excelParam.type = "boolean";
      break;
    case "string":
      excelParam.type = "string";
      break;
    case "array":
      excelParam.type = "any";
      excelParam.dimensionality = "matrix";
      break;
    default:
      excelParam.type = "any";
      break;
  }

  return excelParam;
}

/**
 * Convert one catalog item into Excel functions.json metadata.
 */
export function catalogItemToExcelMetadata(item) {
  const id = String(item.id || item.name || "").toUpperCase();
  const name = String(item.name || item.id || "").toUpperCase();
  const resultType = item.result && item.result.type === "array" ? "any" : item.result && item.result.type === "number" ? "number" : item.result && item.result.type === "string" ? "string" : "any";

  const meta = {
    id,
    name,
    description: item.description || name,
    parameters: (item.parameters || []).map(catalogParamToExcel),
    result: {
      type: resultType,
    },
  };

  if (item.result && item.result.type === "array") {
    meta.result.dimensionality = "matrix";
  }

  return meta;
}

/**
 * Convert full catalog to Excel functions.json "functions" array.
 */
export function catalogToExcelFunctions(catalog) {
  return (catalog || []).map(catalogItemToExcelMetadata);
}

/**
 * Bind each catalog function to CustomFunctions so =METHODTECH.MARKETCAP() works.
 */
export async function registerCatalogFunctions() {
  const catalog = await getCatalog();
  if (!catalog.length) {
    console.warn("No catalog functions to register.");
    return [];
  }

  const registered = [];

  catalog.forEach((item) => {
    const id = String(item.id || item.name || "").toUpperCase();
    if (!id) {
      return;
    }

    const handler = async function catalogHandler() {
      const args = Array.prototype.slice.call(arguments);
      return invokeCatalogFunction(item, args);
    };

    try {
      if (typeof CustomFunctions !== "undefined" && CustomFunctions.associate) {
        CustomFunctions.associate(id, handler);
        registered.push(id);
      }
    } catch (error) {
      console.warn(`Could not associate ${id}:`, error);
    }
  });

  console.log(
    "Registered MethodTech functions:",
    registered.map((id) => `METHODTECH.${id}`).join(", ")
  );

  // Helpful for debugging / UI
  globalThis.__methodTechRegisteredFunctions = registered;
  globalThis.__methodTechBaseUrl = getDjangoBaseUrl();

  return registered;
}
