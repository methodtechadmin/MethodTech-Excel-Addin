/* global console, Office */

import { fetchPublicCatalog } from "../catalog/catalog";
import { registerCatalogFunctions } from "../catalog/registerFunctions";

/**
 * Connect catalog names as soon as Excel starts. Do not wait for sign-in.
 * A silent sign-in failure in Edge leaves the fx list populated and the cell as #NAME?.
 */
const catalogReady = fetchPublicCatalog().catch((error) => {
  console.warn("MethodTech catalog load failed before functions could connect:", error);
  throw error;
});

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    catalogReady
      .then(() => registerCatalogFunctions())
      .catch((error) => {
        console.warn("MethodTech functions were not connected:", error);
      });
  });
}
