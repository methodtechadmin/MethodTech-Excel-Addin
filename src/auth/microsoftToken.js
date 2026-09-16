/* global Office, OfficeRuntime, console */

/**
 * Use the identity already signed into Excel (Office SSO).
 * No MSAL popup inside the add-in.
 */
export async function getMicrosoftAccessToken() {
  const options = {
    allowSignInPrompt: true,
    forMSGraphAccess: false,
  };

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken) {
      return await OfficeRuntime.auth.getAccessToken(options);
    }
  } catch (error) {
    console.warn("OfficeRuntime.auth.getAccessToken failed, trying Office.auth.", error);
  }

  if (typeof Office !== "undefined" && Office.auth && Office.auth.getAccessToken) {
    return await Office.auth.getAccessToken(options);
  }

  throw new Error(
    "Could not read the current Excel/Microsoft account. Make sure you are signed into Excel and the add-in WebApplicationInfo Resource matches Azure Application ID URI (api://<add-in-host>/<app-id>)."
  );
}
