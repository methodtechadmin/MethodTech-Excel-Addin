/* global Office, OfficeRuntime, console */

/**
 * Use the identity already signed into Excel (Office SSO).
 * No MSAL popup inside the add-in.
 */
export async function getMicrosoftAccessToken() {
  // Always create a NEW options object per call. Office.js mutates the object
  // and attaches a callback; reusing it throws:
  // "Callback cannot be specified both in argument list and in optional object."
  const createOptions = () => ({
    allowSignInPrompt: true,
    allowConsentPrompt: true,
    forMSGraphAccess: false,
  });

  const formatOfficeError = (error) => {
    if (!error) {
      return "unknown Office auth error";
    }
    if (error.message) {
      return error.code != null ? `${error.message} (code ${error.code})` : error.message;
    }
    if (error.code != null) {
      return `Office auth error code ${error.code}`;
    }
    try {
      return JSON.stringify(error);
    } catch {
      return String(error);
    }
  };

  let runtimeError = null;

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken) {
      return await OfficeRuntime.auth.getAccessToken(createOptions());
    }
  } catch (error) {
    runtimeError = error;
    console.warn("OfficeRuntime.auth.getAccessToken failed, trying Office.auth.", error);
  }

  try {
    if (typeof Office !== "undefined" && Office.auth && Office.auth.getAccessToken) {
      return await Office.auth.getAccessToken(createOptions());
    }
  } catch (error) {
    throw new Error(
      `Excel SSO failed: ${formatOfficeError(error)}. For local debug, Entra Application ID URI must include api://localhost:3001/1218ba26-b58d-4829-91e4-5be4bbb88461`
    );
  }

  if (runtimeError) {
    throw new Error(
      `Excel SSO failed: ${formatOfficeError(runtimeError)}. For local debug, Entra Application ID URI must include api://localhost:3001/1218ba26-b58d-4829-91e4-5be4bbb88461`
    );
  }

  throw new Error(
    "Could not read the current Excel/Microsoft account. Make sure you are signed into Excel and the add-in WebApplicationInfo Resource matches Azure Application ID URI (api://localhost:3001/<app-id> for local)."
  );
}
