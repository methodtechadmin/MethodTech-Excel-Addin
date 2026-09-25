/* global CustomFunctions, console */

import { getDjangoBaseUrl, FUNCTIONS_NAMESPACE } from "../auth/config";
import { djangoFetch } from "../auth/http";
import { getAuthHeadersContext } from "../auth/session";
import { getCatalog, getCatalogEntry } from "./catalog";

/**
 * Convert an Excel range into a backend-friendly array payload.
 *
 * - 1D list / single column → ["v1", "v2", ...]  (no headers)
 * - Multi-column with header row → [{ "colA": v, "colB": v }, ...]
 * - Multi-column with only one row (no header+data) → leave as matrix
 *
 * @param {unknown} value Excel argument
 * @param {{ forceObjectRows?: boolean }} [options]
 *        forceObjectRows: always treat row 0 as headers (for params like `data`)
 */
function excelRangeToBackendArray(value, options = {}) {
  if (!Array.isArray(value) || value.length === 0) {
    return value;
  }

  // Already a list of objects (unlikely from Excel, but keep as-is)
  if (value[0] !== null && typeof value[0] === "object" && !Array.isArray(value[0])) {
    return value;
  }

  // Ensure matrix shape: each row is an array
  const matrix = value.map((row) => (Array.isArray(row) ? row : [row]));
  const colCount = matrix.reduce((max, row) => Math.max(max, row.length), 0);

  // Single column → flat list of values (1D / no header objects)
  if (colCount <= 1) {
    return matrix.map((row) => row[0]);
  }

  // Multi-column with only one row → keep as [[v1, v2, ...]]
  if (matrix.length < 2) {
    return matrix;
  }

  const headers = matrix[0].map((header) => String(header ?? "").trim());
  const nonEmptyHeaders = headers.filter((header) => header.length > 0);
  const fieldNameHeaders = nonEmptyHeaders.filter((header) =>
    /^[A-Za-z_][A-Za-z0-9_]*$/.test(header)
  );

  // Majority of non-empty first-row cells look like field names (amfi_mf_code, fund_name).
  // Or caller forces object rows (e.g. CALCULATESCHEMEBENCHMARK param `data`).
  const looksLikeFieldNames =
    fieldNameHeaders.length >= 2 &&
    fieldNameHeaders.length >= Math.ceil(nonEmptyHeaders.length * 0.7);

  if (!options.forceObjectRows && !looksLikeFieldNames) {
    // e.g. [["ICICI ...", "Large & Mid Cap"], ["HDFC ...", "Flexi Cap"]]
    return matrix;
  }

  if (nonEmptyHeaders.length === 0) {
    return matrix;
  }

  const rowsAsObjects = matrix.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      if (!header) {
        return;
      }
      const cell = row[index];
      obj[header] = cell === null || cell === undefined ? "" : cell;
    });
    return obj;
  });

  console.log(
    `[MethodTech] Converted ${rowsAsObjects.length} Excel row(s) → [{column: value}, ...] headers:`,
    nonEmptyHeaders
  );

  return rowsAsObjects;
}

/**
 * Flatten Excel range values into a plain JS array when needed.
 * @param {unknown} value
 * @param {string} paramType
 * @param {{ paramName?: string }} [meta]
 */
function normalizeParamValue(value, paramType, meta = {}) {
  if (value === null || value === undefined || value === "") {
    return undefined;
  }

  if (paramType === "array" || paramType === "any") {
    if (Array.isArray(value)) {
      // Param named `data` is a table body for APIs like calculate_scheme_benchmark(data=...)
      const forceObjectRows = String(meta.paramName || "").toLowerCase() === "data";
      return excelRangeToBackendArray(value, { forceObjectRows });
    }
    // Single cell for array-like payloads
    if (paramType === "array") {
      return [value];
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
 * Excel custom functions with result.dimensionality=matrix must return a 2D array.
 * Returning a plain string/number often shows as #CALC!.
 */
function toExcelMatrix(value) {
  if (value === null || value === undefined || value === "") {
    return [[`${FUNCTIONS_NAMESPACE}: empty result`]];
  }
  if (typeof value !== "object") {
    return [[String(value)]];
  }
  if (Array.isArray(value)) {
    if (value.length === 0) {
      return [[`${FUNCTIONS_NAMESPACE}: empty result`]];
    }
    if (Array.isArray(value[0])) {
      return value;
    }
    // 1D → column
    return value.map((cell) => [cell]);
  }
  return [[String(value)]];
}

/**
 * Turn API result into something Excel can show in cells (value or 2D spill range).
 */
function formatResultForExcel(result) {
  if (result === null || result === undefined) {
    return toExcelMatrix(`${FUNCTIONS_NAMESPACE}: empty result`);
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
    return toExcelMatrix(result);
  }

  if (Array.isArray(result)) {
    if (result.length === 0) {
      return toExcelMatrix(`${FUNCTIONS_NAMESPACE}: empty result`);
    }

    // Already a matrix
    if (Array.isArray(result[0])) {
      return result;
    }

    // Array of objects → header row + value rows for Excel spill
    if (result[0] !== null && typeof result[0] === "object") {
      const keys = Object.keys(result[0]);
      if (!keys.length) {
        return toExcelMatrix(`${FUNCTIONS_NAMESPACE}: empty object rows`);
      }
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

  // Plain object with no keys
  const objectKeys = Object.keys(result);
  if (!objectKeys.length) {
    return toExcelMatrix(`${FUNCTIONS_NAMESPACE}: empty object`);
  }

  // Plain object → two-column key/value
  return objectKeys.map((key) => {
    const value = result[key];
    return [key, value !== null && typeof value === "object" ? JSON.stringify(value) : value];
  });
}

/**
 * Every MTECH.* call goes through /api/microsoft/excel/execute/
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
    const normalized = normalizeParamValue(raw, param.type, { paramName: param.name });
    if (normalized === undefined) {
      if (param.required) {
        throw new Error(`Missing required parameter: ${param.name}`);
      }
      return;
    }
    data[param.name] = normalized;
  });

  // Exact JSON body the Data API method should receive, e.g.
  // calculate_scheme_benchmark(data=...) → { "data": [ {...}, {...} ] }
  // Django MUST post this object (not the inner list alone).
  const request_body = { ...data };

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
    request_body,
  };

  const response = await djangoFetch("/api/microsoft/excel/execute/", {
    method: "POST",
    body: JSON.stringify(payload),
  });

  console.log(`${FUNCTIONS_NAMESPACE}.${entry.id || entry.name} Data API body must be:`, request_body);
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
    case "any":
      // Required so Excel can pass A1:J14 into the function.
      // Without dimensionality, Excel shows #CALC! "Unliftable Array".
      excelParam.type = "any";
      excelParam.dimensionality = "matrix";
      break;
    default:
      excelParam.type = "any";
      excelParam.dimensionality = "matrix";
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

  // Required by AppSource: Excel "Help on this function" uses helpUrl.
  const helpUrl = item.helpUrl || item.help_url || "https://www.methodtech.in/";

  const meta = {
    id,
    name,
    description: item.description || name,
    helpUrl,
    parameters: (item.parameters || []).map(catalogParamToExcel),
    result: {
      type: resultType,
    },
  };

  // Allow spilling tables for array/any results (avoids #CALC! on 2D returns)
  if (item.result && (item.result.type === "array" || item.result.type === "any" || resultType === "any")) {
    meta.result.type = "any";
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
 * Bind each catalog function to CustomFunctions so =MTECH.MARKETCAP() works.
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
      try {
        return await invokeCatalogFunction(item, args);
      } catch (error) {
        const message =
          (error && error.message) ||
          (typeof error === "string" ? error : "MethodTech function failed");
        console.error(`${FUNCTIONS_NAMESPACE}.${id} failed:`, error);
        // Must return a 2D array when result is declared as matrix (avoids #CALC!).
        return [[`${FUNCTIONS_NAMESPACE} error`], [message]];
      }
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
    registered.map((id) => `${FUNCTIONS_NAMESPACE}.${id}`).join(", ")
  );

  // Helpful for debugging / UI
  globalThis.__methodTechRegisteredFunctions = registered;
  globalThis.__methodTechBaseUrl = getDjangoBaseUrl();

  return registered;
}
