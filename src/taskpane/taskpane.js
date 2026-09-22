/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

import { ensureSignedIn } from "../auth/signIn";
import { FUNCTIONS_NAMESPACE } from "../auth/config";
import { fetchCatalog, getCatalog } from "../catalog/catalog";
import { registerCatalogFunctions } from "../catalog/registerFunctions";
import { publishCatalogMetadata } from "../catalog/publishMetadata";

function setAuthStatus(message, isError = false) {
  const el = document.getElementById("auth-status");
  if (!el) {
    return;
  }
  el.textContent = message;
  el.style.color = isError ? "#a4262c" : "#107c10";
}

function setCatalogRefreshStatus(message, isError = false) {
  const el = document.getElementById("catalog-refresh-status");
  if (!el) {
    return;
  }
  el.style.display = message ? "block" : "none";
  el.textContent = message || "";
  el.style.color = isError ? "#a4262c" : "#107c10";
}

function renderCatalogFunctions(catalog) {
  const list = document.getElementById("catalog-functions");
  if (!list) {
    return;
  }

  list.innerHTML = "";

  if (!catalog.length) {
    const empty = document.createElement("li");
    empty.className = "ms-ListItem";
    empty.textContent = "No catalog functions returned.";
    list.appendChild(empty);
    return;
  }

  catalog.forEach((item) => {
    const name = String(item.name || item.id || "").toUpperCase();
    const li = document.createElement("li");
    li.className = "ms-ListItem";
    li.innerHTML = `<span class="ms-font-m"><b>=${FUNCTIONS_NAMESPACE}.${name}()</b> — ${item.description || item.method_name || ""}</span>`;
    list.appendChild(li);
  });
}

async function refreshCatalogFromApi() {
  const button = document.getElementById("refresh-catalog");
  if (button) {
    button.disabled = true;
    button.textContent = "Refreshing…";
  }
  setCatalogRefreshStatus("Calling catalog API…");

  try {
    await ensureSignedIn({ force: false, interactive: true });
    const catalog = await fetchCatalog();
    await publishCatalogMetadata(catalog);
    const registered = await registerCatalogFunctions();
    renderCatalogFunctions(catalog);
    setCatalogRefreshStatus(
      `Catalog refreshed: ${catalog.length} function(s), ${registered.length} registered.`
    );
  } catch (error) {
    console.error("Catalog refresh failed:", error);
    setCatalogRefreshStatus(
      `Refresh failed: ${(error && error.message) || String(error)}`,
      true
    );
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "Refresh catalog";
    }
  }
}

Office.onReady(async () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const msButton = document.getElementById("microsoft-signin");
  if (msButton) {
    msButton.style.display = "none";
  }

  const refreshButton = document.getElementById("refresh-catalog");
  if (refreshButton) {
    refreshButton.addEventListener("click", () => {
      refreshCatalogFromApi();
    });
  }

  setAuthStatus("Using current Excel account…");

  try {
    const session = await ensureSignedIn({ force: true, interactive: true });
    const catalog = await getCatalog();
    renderCatalogFunctions(catalog);
    setAuthStatus(`Signed in as ${session.email}`);
  } catch (error) {
    console.error("Excel identity / Django sign-in failed:", error);
    const detail =
      (error && error.message) ||
      (error && error.code != null && `Office auth error ${error.code}`) ||
      (error && String(error)) ||
      "unknown error";
    setAuthStatus(`Sign-in failed: ${detail}`, true);
  }
});
