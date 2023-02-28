import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { createB2cGraphClient } from "../Helpers/graphclient";
import { validate } from "../Helpers/validate";
import { createExistingB2cUser } from "../Models/B2cUser";
import { RvizUserFields } from "../Models/RvizUser";
import enforceAdmin from "../Helpers/enforceAdmin";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  if (!enforceAdmin(context)) {
    return;
  }

  const editUserId = context.bindingData.id;
  const editUser = req.body as RvizUserFields;

  console.log(editUser);
  if (!validate(editUser, ["displayName", "givenName", "surname", "internalExternal", "level", "subscriptionStatus"])) {
    context.res = {
      status: 400,
      body: "Invalid object provided",
    };
    return;
  }

  const graphClient = createB2cGraphClient();
  if (typeof graphClient === "string") {
    context.log.error(graphClient);
    context.res = { status: 404 };
    return;
  }

  const b2cUser = createExistingB2cUser(editUser);

  try {
    await graphClient.api(`/users/${editUserId}`).patch(b2cUser);
    context.res = {
      status: 200,
    };
    return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    context.log.error("Error calling graph API", error);
    context.res = {
      status: 400,
      body: error.message ?? error,
    };
  }
};

export default httpTrigger;
