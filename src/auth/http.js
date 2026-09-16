/* global fetch */

import { getDjangoBaseUrl } from "./config";
import { getAuthHeadersContext } from "./session";

/**
 * Headers for every authenticated Django request:
 * - Authorization: Bearer <access>
 * - Username: <email from sign-in>
 */
export async function buildAuthHeaders(extra = {}) {
  const { accessToken, username } = await getAuthHeadersContext();

  return {
    Authorization: `Bearer ${accessToken}`,
    Username: username,
    Accept: "application/json",
    ...extra,
  };
}

/**
 * Authenticated fetch against DJANGO_BASE_URL + path.
 * @param {string} path e.g. "/api/microsoft/excel/catalog/"
 * @param {RequestInit} [options]
 */
export async function djangoFetch(path, options = {}) {
  const base = getDjangoBaseUrl();
  const normalizedPath = path.startsWith("/") ? path : `/${path}`;
  const url = `${base}${normalizedPath}`;

  const headers = await buildAuthHeaders(options.headers || {});
  if (options.body && !headers["Content-Type"]) {
    headers["Content-Type"] = "application/json";
  }

  const response = await fetch(url, {
    ...options,
    headers,
  });

  const text = await response.text();
  let data = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = text;
    }
  }

  if (!response.ok) {
    const message =
      (data && (data.detail || data.error || data.message)) ||
      `Request failed (${response.status}) ${normalizedPath}`;
    throw new Error(typeof message === "string" ? message : JSON.stringify(message));
  }

  return data;
}
