import { ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from "@microsoft/microsoft-graph-client";

export function createADGraphClient() {
  const envKeys = Object.keys(process.env);
  if (!envKeys.includes("ADTenantId") || !envKeys.includes("ADClientId") || !envKeys.includes("ADSecret")) {
    return "Tenant, client, or secret is not configured";
  }
  return createGraphClient(process.env["ADTenantId"], process.env["ADClientId"], process.env["ADSecret"]);
}

export function createB2cGraphClient() {
  const envKeys = Object.keys(process.env);
  if (!envKeys.includes("B2cTenantId") || !envKeys.includes("B2cClientId") || !envKeys.includes("B2cSecret")) {
    return "Tenant, client, or secret is not configured";
  }
  return createGraphClient(process.env["B2cTenantId"], process.env["B2cClientId"], process.env["B2cSecret"]);
}

function createGraphClient(tenantId: string, clientId: string, secret: string): Client | string {
  const envKeys = Object.keys(process.env);
  if (!envKeys.includes("B2cTenantId") || !envKeys.includes("B2cClientId") || !envKeys.includes("B2cSecret")) {
    return "Tenant, client, or secret is not configured";
  }

  const tokenCredential = new ClientSecretCredential(tenantId, clientId, secret);
  const authProvider = new TokenCredentialAuthenticationProvider(tokenCredential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });
  const client = Client.initWithMiddleware({
    debugLogging: true,
    authProvider,
  });
  return client;
}
