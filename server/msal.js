import { ConfidentialClientApplication } from "@azure/msal-node";
import { config, getRedirectUri } from "./config.js";

export function createMsalClient() {
  if (!config.entra.clientId || !config.entra.clientSecret) {
    return null;
  }

  return new ConfidentialClientApplication({
    auth: {
      clientId: config.entra.clientId,
      clientSecret: config.entra.clientSecret,
      authority: `https://login.microsoftonline.com/${config.entra.tenantId}`,
    },
  });
}

export function getAuthCodeUrlRequest() {
  return {
    scopes: config.entra.scopes,
    redirectUri: getRedirectUri(),
  };
}

export function getTokenRequest(code) {
  return {
    code,
    scopes: config.entra.scopes,
    redirectUri: getRedirectUri(),
  };
}
