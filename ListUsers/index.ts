import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { createGraphClient } from "../Helpers/graphclient";
import { RvizUser } from "../Models/RvizUser";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  const graphClient = createGraphClient();
  if (typeof graphClient === "string") {
    context.log.error(graphClient);
    context.res = { status: 404 };
    return;
  }

  try {
    const graphRes = await graphClient.api("/users/").select(userFields).get();
    if (Array.isArray(graphRes.value)) {
      const results = graphRes.value.map((u) => {
        return {
          id: u.id,
          displayName: u.displayName,
          givenName: u.givenName,
          surname: u.surname,
          jobTitle: u.jobTitle,
        } as RvizUser;
      });
      context.res = { body: results };
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    context.log.error("Error calling graph API", error);
    context.res = { status: 400 };
  }
};

const userFields =
  "id, displayName, identities, givenName, jobTitle, surname, extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External, extension_bb8c6b6c043d4874965ca693e13a3e1e_Level, extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status";

export default httpTrigger;
