import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { createADGraphClient } from "../Helpers/graphclient";
import { createRvizUser } from "../Models/RvizUser";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {

    const graphClient = createADGraphClient();
    if (typeof graphClient === "string") {
      context.log.error(graphClient);
      context.res = { status: 404 };
      return;
    }

    try {
        const emailSearch = context.req.query.email;
        if (emailSearch && emailSearch !== "") {
          const filter = `startsWith(mail, '${emailSearch}') or startswith(userPrincipalName, '${emailSearch}')`;
          const userFields = "id, displayName, identities, givenName, jobTitle, surname, mail, userPrincipalName";
          const graphRes = await graphClient.api("/users").filter(filter).select(userFields).get();

          if (Array.isArray(graphRes.value)) {
            const results = graphRes.value.map(createRvizUser);
            context.res = { body: results };
          }
        } 
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (error: any) {
        context.log.error("Error calling graph API", error);
        context.res = { status: 400 };
      }
};

export default httpTrigger;