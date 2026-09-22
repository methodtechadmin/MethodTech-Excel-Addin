/* global process */

/** Excel custom function prefix: =MTECH.MARKETCAP() — must match manifest Functions.Namespace */
export const FUNCTIONS_NAMESPACE = "MTECH";

export function getDjangoBaseUrl() {
  // In npm start (dev), this is "" so requests go to https://localhost:3001
  // and webpack proxies /api + /v1 to DJANGO_BASE_URL from .env.local
  return (process.env.DJANGO_BASE_URL || "").replace(/\/$/, "");
}

export function getMicrosoftAppId() {
  return (process.env.MICROSOFT_APP_ID || "").trim();
}

/** Matches Django / Azure Application ID URI audience. */
export function getMicrosoftApiScope() {
  return (
    process.env.MICROSOFT_API_SCOPE ||
    "api://612da946-27be-427c-b997-2279839c4e56/access_as_user"
  ).trim();
}

/**
 * common = work + personal Microsoft accounts (matches "All Microsoft account users").
 * Or set MICROSOFT_TENANT_ID to a specific tenant GUID.
 */
export function getMicrosoftAuthority() {
  const tenant = (process.env.MICROSOFT_TENANT_ID || "common").trim();
  return `https://login.microsoftonline.com/${tenant}`;
}

export function getMicrosoftRedirectUri() {
  if (typeof window !== "undefined" && window.location && window.location.origin) {
    return `${window.location.origin}/taskpane.html`;
  }
  return (process.env.MICROSOFT_REDIRECT_URI || "https://localhost:3001/taskpane.html").trim();
}
