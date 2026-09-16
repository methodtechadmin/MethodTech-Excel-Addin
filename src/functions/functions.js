/* global console, Office */

import { ensureSignedIn } from "../auth/signIn";

/**
 * MethodTech custom functions come only from the Django catalog.
 * They are registered at runtime after sign-in + catalog load.
 */

// Shared runtime: sign in and register catalog functions as soon as Office is ready
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    ensureSignedIn({ interactive: false }).catch((error) => {
      console.warn("Background MethodTech sign-in/catalog failed:", error);
    });
  });
}
