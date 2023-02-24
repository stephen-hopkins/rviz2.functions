import { AzureFunction, Context } from "@azure/functions";
import { createB2cGraphClient } from "../Helpers/graphclient";
import { createRvizUser } from "../Models/RvizUser";

const httpTrigger: AzureFunction = async function (context: Context): Promise<void> {
  const graphClient = createB2cGraphClient();
  if (typeof graphClient === "string") {
    context.log.error(graphClient);
    context.res = { status: 404 };
    return;
  }

  try {
    const graphRes = await graphClient.api("/users/").select(userFields).get();

    if (Array.isArray(graphRes.value)) {
      const results = graphRes.value.map(createRvizUser);
      context.res = { body: results };
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    context.log.error("Error calling graph API", error);
    context.res = { status: 400 };
  }
};

const userFields =
  "id, displayName, identities, givenName, jobTitle, surname, mail, userPrincipalName, extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External, extension_bb8c6b6c043d4874965ca693e13a3e1e_Level, extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status";

export default httpTrigger;
