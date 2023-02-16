import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { Client } from "@microsoft/microsoft-graph-client";

const tenantId = "5b20d9ae-2307-4ebf-94e0-b9493842a5da";
const clientId = "e62af7e9-a1d3-49a9-94db-cfb00a229ef1";
const secret = "zcy8Q~GF93yHp3L5HrCHviIDyxZ1A24xw.8tucnQ";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  var envKeys = Object.keys(process.env);
  if (
    !envKeys.includes("TenantId") ||
    !envKeys.includes("ClientId") ||
    !envKeys.includes("Secret")
  ) {
    context.log.error("Tenant, client, or secret is not configured");
    context.res = { status: 404 };
    return;
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

  try {
    const graphRes = await client.api("/users/").select(userFields).get();
    if (Array.isArray(graphRes.value)) {
      const results = graphRes.value.map((u) => {
        return {
          id: u.id,
          displayName: u.displayName,
          givenName: u.givenName,
          surname: u.surname,
          jobTitle: u.jobTitle,
        } as UserDetail;
      });
      context.res = { body: results };
    }
  } catch (error: any) {
    context.log.error("Error calling graph API", error);
    context.res = { status: 400 };
  }
};

type UserDetail = {
  id: string;
  displayName: string;
  givenName: string;
  surname: string;
  email: string;
  jobTitle: string;
  internalExternal: string;
  level: string;
  subscriptionStatus: string;
};

const userFields =
  "id, displayName, identities, givenName, jobTitle, surname, extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External, extension_bb8c6b6c043d4874965ca693e13a3e1e_Level, extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status";

export default httpTrigger;
