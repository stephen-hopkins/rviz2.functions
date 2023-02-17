import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { createGraphClient } from "../Helpers/graphclient";
import { createB2cUser } from "../Models/B2cUser";
import { createRvizUser, NewRvizUser } from "../Models/RvizUser";

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

  const newB2cUser = createB2cUser(newUser);

  try {
    const createdB2cUser = await graphClient.api("/user").post(newB2cUser);
    const createdRvizUser = createRvizUser(createdB2cUser);
    context.res = {
      body: createdRvizUser,
    };
    return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    context.log.error("Error calling graph API", error);
    context.res = { status: 400 };
  }
};

export default httpTrigger;
