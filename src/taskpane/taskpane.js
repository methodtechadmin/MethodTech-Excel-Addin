/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

import { ensureSignedIn } from "../auth/signIn";
import { getCatalog } from "../catalog/catalog";

function setAuthStatus(message, isError = false) {
  const el = document.getElementById("auth-status");
  if (!el) {
    return;
  }
  el.textContent = message;
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
    li.innerHTML = `<span class="ms-font-m"><b>=METHODTECH.${name}()</b> — ${item.description || item.method_name || ""}</span>`;
    list.appendChild(li);
  });
}

Office.onReady(async () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const msButton = document.getElementById("microsoft-signin");
  if (msButton) {
    msButton.style.display = "none";
  }

  setAuthStatus("Using current Excel account…");

  try {
    const session = await ensureSignedIn({ force: true, interactive: true });
    const catalog = await getCatalog();
    renderCatalogFunctions(catalog);
    setAuthStatus(`Signed in as ${session.email}`);
  } catch (error) {
    console.error("Excel identity / Django sign-in failed:", error);
    setAuthStatus(`Sign-in failed: ${error.message}`, true);
  }
});
