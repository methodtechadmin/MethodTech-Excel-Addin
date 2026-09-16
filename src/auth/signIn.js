/* global fetch, console */

import { getDjangoBaseUrl } from "./config";
import { getMicrosoftAccessToken } from "./microsoftToken";
import { getSession, saveSession } from "./session";
import { fetchCatalog } from "../catalog/catalog";
import { registerCatalogFunctions } from "../catalog/registerFunctions";
import { publishCatalogMetadata } from "../catalog/publishMetadata";

let signInPromise = null;

/**
 * Call Django Excel sign-in with the Microsoft identity token, then load catalog.
 * Django accepts:
 *   Authorization: Bearer <Microsoft access token>
 *   X-Microsoft-Authorization: Bearer <Microsoft access token>
 *   body: { "access_token": "<Microsoft access token>" }
 * @returns {Promise<object>}
 */
export async function signInToDjango(options = {}) {
  const interactive = options.interactive !== false;
  const microsoftToken = await getMicrosoftAccessToken({ interactive });
  if (!microsoftToken) {
    throw new Error("Microsoft identity token was empty.");
  }

  const url = `${getDjangoBaseUrl()}/api/microsoft/excel/sign-in/`;

  let response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: `Bearer ${microsoftToken}`,
        "X-Microsoft-Authorization": `Bearer ${microsoftToken}`,
      },
      body: JSON.stringify({
        access_token: microsoftToken,
      }),
    });
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    throw new Error(
      `Cannot reach API at ${url} (${detail}). Check CORS allows ${typeof location !== "undefined" ? location.origin : "the add-in origin"} and that staging API is up.`
    );
  }

  const text = await response.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`Sign-in returned non-JSON (${response.status}): ${text.slice(0, 200)}`);
  }

  if (!response.ok) {
    throw new Error(data.detail || data.error || `Sign-in failed (${response.status})`);
  }

  if (!data.access || !data.email) {
    throw new Error("Sign-in response missing access token or email.");
  }

  const session = {
    access: data.access,
    refresh: data.refresh || "",
    access_token_duration: data.access_token_duration,
    email: data.email,
    sidebarConfigTag: data.sidebarConfigTag,
  };

  await saveSession(session);
  console.log("MethodTech signed in as", session.email);

  const catalog = await fetchCatalog();
  // Publish metadata FIRST so Excel's waiting functions.json request
  // can complete with the full suggestion list immediately.
  const published = await publishCatalogMetadata(catalog);
  const registered = await registerCatalogFunctions();
  session.catalogCount = catalog.length;
  session.registeredFunctions = registered;
  session.metadataPublished = Boolean(published && published.updated);
  session.metadataCount = published && published.count ? published.count : 0;

  return session;
}

/**
 * Sign in once on load (or reuse saved session). Safe to call from task pane and functions.
 * @param {{ force?: boolean }} [options]
 * @returns {Promise<object>}
 */
export async function ensureSignedIn(options = {}) {
  if (!options.force) {
    const existing = await getSession();
    if (existing && existing.access && existing.email) {
      try {
        const catalog = await fetchCatalog();
        await publishCatalogMetadata(catalog);
        await registerCatalogFunctions();
      } catch (error) {
        console.warn("Catalog refresh failed; using existing session.", error);
      }
      return existing;
    }
  }

  if (!signInPromise) {
    signInPromise = signInToDjango({
      interactive: options.interactive !== false,
    }).finally(() => {
      signInPromise = null;
    });
  }

  return signInPromise;
}
