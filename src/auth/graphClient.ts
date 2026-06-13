import { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import { graphScopes } from "./msalConfig";

export function createGraphClient(instance: IPublicClientApplication, account: AccountInfo) {
  return Client.init({
    authProvider: async (done) => {
      try {
        const response = await instance.acquireTokenSilent({ scopes: graphScopes, account });
        done(null, response.accessToken);
      } catch (silentError) {
        try {
          const response = await instance.acquireTokenPopup({ scopes: graphScopes, account });
          done(null, response.accessToken);
        } catch (popupError) {
          done(toError(popupError, silentError), null);
        }
      }
    }
  });
}

function toError(popupError: unknown, silentError: unknown): Error {
  if (popupError instanceof Error) return popupError;
  if (silentError instanceof Error) return silentError;
  return new Error("Microsoft sign-in is required to continue.");
}
