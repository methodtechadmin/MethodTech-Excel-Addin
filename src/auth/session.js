/* global OfficeRuntime */

const STORAGE_KEY = "methodtech.auth.session";

/**
 * @typedef {Object} AuthSession
 * @property {string} access
 * @property {string} refresh
 * @property {number} [access_token_duration]
 * @property {string} email
 * @property {string} [sidebarConfigTag]
 */

/**
 * @returns {AuthSession | null}
 */
export function getSessionSync() {
  return globalThis.__methodTechSession || null;
}

/**
 * @returns {Promise<AuthSession | null>}
 */
export async function getSession() {
  if (globalThis.__methodTechSession) {
    return globalThis.__methodTechSession;
  }

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      const raw = await OfficeRuntime.storage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        globalThis.__methodTechSession = parsed;
        return parsed;
      }
    }
  } catch (error) {
    console.warn("Could not read saved auth session.", error);
  }

  return null;
}

/**
 * @param {AuthSession} session
 */
export async function saveSession(session) {
  globalThis.__methodTechSession = session;

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(STORAGE_KEY, JSON.stringify(session));
    }
  } catch (error) {
    console.warn("Could not persist auth session.", error);
  }
}

export async function clearSession() {
  globalThis.__methodTechSession = null;

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      await OfficeRuntime.storage.removeItem(STORAGE_KEY);
    }
  } catch (error) {
    console.warn("Could not clear auth session.", error);
  }
}

/**
 * Values later API calls should use.
 * @returns {Promise<{ accessToken: string, email: string, username: string }>}
 */
export async function getAuthHeadersContext() {
  const session = await getSession();
  if (!session || !session.access || !session.email) {
    throw new Error("Not signed in. Open the MethodTech task pane and wait for sign-in.");
  }

  return {
    accessToken: session.access,
    email: session.email,
    username: session.email,
  };
}
