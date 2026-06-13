import { Configuration } from "@azure/msal-browser";

export const graphScopes = [
  "User.Read",
  "Chat.Read",
  "Calendars.Read",
  "Mail.Read",
  "Files.Read.All"
];

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_MS_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_MS_TENANT_ID || "common"}`,
    redirectUri: import.meta.env.VITE_MS_REDIRECT_URI || "http://localhost:5173"
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};
