import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { createGraphClient } from "../Helpers/graphclient";
import { validate } from "../Helpers/validate";
import { createNewB2cUser } from "../Models/B2cUser";
import { NewRvizUser } from "../Models/RvizUser";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  const newUser = req.body as NewRvizUser;
  if (
    !validate(newUser, [
      "displayName",
      "givenName",
      "surname",
      "email",
      "internalExternal",
      "level",
      "subscriptionStatus",
      "password",
    ])
  ) {
    context.res = {
      status: 400,
      body: "Invalid object provided",
    };
    return;
  }

  const graphClient = createGraphClient();
  if (typeof graphClient === "string") {
    context.log.error(graphClient);
    context.res = { status: 404 };
    return;
  }

  const newB2cUser = createNewB2cUser(newUser);

  try {
    await graphClient.api("/users").post(newB2cUser);
    // the returned object from the post call doesn't have much in it so let's return nothing for now
    context.res = {
      status: 201,
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
