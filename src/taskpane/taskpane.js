/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

import { ensureSignedIn } from "../auth/signIn";
import { clearSession, getSession } from "../auth/session";
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

function setAuthButtons({ signedIn }) {
  const signInBtn = document.getElementById("microsoft-signin");
  const signOutBtn = document.getElementById("microsoft-signout");

  if (signInBtn) {
    signInBtn.style.display = signedIn ? "none" : "inline-block";
    signInBtn.disabled = false;
    signInBtn.textContent = "Sign in";
  }
  if (signOutBtn) {
    signOutBtn.style.display = signedIn ? "inline-block" : "none";
    signOutBtn.disabled = false;
  }
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

async function showSignedInState() {
  const catalog = await getCatalog();
  renderCatalogFunctions(catalog);
  setAuthStatus("Signed in");
  setAuthButtons({ signedIn: true });
}

async function signInWithExcelAccount() {
  const signInBtn = document.getElementById("microsoft-signin");
  if (signInBtn) {
    signInBtn.disabled = true;
    signInBtn.textContent = "Signing in…";
  }

  setAuthStatus("Using current Excel account…");

  try {
    await ensureSignedIn({ force: true, interactive: true });
    await showSignedInState();
  } catch (error) {
    console.error("Excel identity / Django sign-in failed:", error);
    const detail =
      (error && error.message) ||
      (error && error.code != null && `Office auth error ${error.code}`) ||
      (error && String(error)) ||
      "unknown error";
    setAuthStatus(`Sign-in failed: ${detail}`, true);
    setAuthButtons({ signedIn: false });
  }
}

async function signOut() {
  const signOutBtn = document.getElementById("microsoft-signout");
  if (signOutBtn) {
    signOutBtn.disabled = true;
  }

  try {
    await clearSession();
    renderCatalogFunctions([]);
    setCatalogRefreshStatus("");
    setAuthStatus("Signed out. Sign in to use MethodTech functions.");
    setAuthButtons({ signedIn: false });
  } catch (error) {
    console.error("Sign-out failed:", error);
    setAuthStatus(`Sign-out failed: ${(error && error.message) || String(error)}`, true);
    setAuthButtons({ signedIn: true });
  }
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
    setAuthStatus("Signed in");
    setAuthButtons({ signedIn: true });
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

function showAppBody() {
  const sideload = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideload) {
    sideload.style.display = "none";
  }
  if (appBody) {
    appBody.style.display = "flex";
  }
}

async function initTaskPane() {
  showAppBody();

  const signInBtn = document.getElementById("microsoft-signin");
  if (signInBtn) {
    signInBtn.addEventListener("click", () => {
      signInWithExcelAccount();
    });
  }

  const signOutBtn = document.getElementById("microsoft-signout");
  if (signOutBtn) {
    signOutBtn.addEventListener("click", () => {
      signOut();
    });
  }

  const refreshButton = document.getElementById("refresh-catalog");
  if (refreshButton) {
    refreshButton.addEventListener("click", () => {
      refreshCatalogFromApi();
    });
  }

  // Restore session if present; otherwise wait for Sign in button.
  setAuthButtons({ signedIn: false });
  setAuthStatus("Sign in to use MethodTech functions.");

  try {
    const existing = await getSession();
    if (existing && existing.access && existing.email) {
      await ensureSignedIn({ force: false, interactive: false });
      await showSignedInState();
    }
  } catch (error) {
    console.warn("Could not restore previous session:", error);
    setAuthButtons({ signedIn: false });
    setAuthStatus("Sign in to use MethodTech functions.");
  }
}

// Always show the real UI (never leave reviewers on the old "sideload" placeholder).
showAppBody();

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    initTaskPane().catch((error) => {
      console.error("Task pane init failed:", error);
      showAppBody();
      setAuthStatus(`Startup error: ${(error && error.message) || String(error)}`, true);
    });
  });
} else {
  initTaskPane().catch((error) => {
    console.error("Task pane init failed (no Office.js):", error);
  });
}
