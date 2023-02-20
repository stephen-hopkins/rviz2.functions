import { ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from "@microsoft/microsoft-graph-client";

export function createGraphClient(): Client | string {
  const envKeys = Object.keys(process.env);
  if (
    !envKeys.includes("TenantId") ||
    !envKeys.includes("ClientId") ||
    !envKeys.includes("Secret")
  ) {
    return "Tenant, client, or secret is not configured";
  }

  const tokenCredential = new ClientSecretCredential(
    process.env["TenantId"],
    process.env["ClientId"],
    process.env["Secret"]
  );
  const authProvider = new TokenCredentialAuthenticationProvider(
    tokenCredential,
    { scopes: ["https://graph.microsoft.com/.default"] }
  );
  const client = Client.initWithMiddleware({
    debugLogging: true,
    authProvider,
  });
  return client;
}
