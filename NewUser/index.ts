import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { createADGraphClient, createB2cGraphClient } from "../Helpers/graphclient";
import { validate } from "../Helpers/validate";
import { createExistingB2cUser, createNewB2cExternalUser, createNewB2cInternalUser } from "../Models/B2cUser";
import { NewRvizUser } from "../Models/RvizUser";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  const graphClient = createB2cGraphClient();
  if (typeof graphClient === "string") {
    context.log.error(graphClient);
    context.res = { status: 404 };
    return;
  }
  const newUser = req.body as NewRvizUser;

  if (newUser.internalExternal === "Internal") {
    const result = await addADUser(newUser, graphClient, context);
    context.res = result;
  } else {
    const result = await addB2cUser(newUser, graphClient, context);
    context.res = result;
  }
};

async function addADUser(newUser: NewRvizUser, graphClient: Client, context: Context) {
  if (
    !validate(newUser, [
      "displayName",
      "givenName",
      "surname",
      "email",
      "internalExternal",
      "level",
      "subscriptionStatus",
    ])
  ) {
    return {
      status: 400,
      body: "Invalid object provided",
    };
  }

  const redirectUrl = process.env.AppURL;
  if (typeof redirectUrl === "undefined" || redirectUrl === "") {
    return {
      status: 404,
      body: "Configuration error in Azure functions",
    };
  }

  try {
    // to setup an external ad login (rviz internal user) we need tenant id and object id of user
    const adGraphClient = createADGraphClient() as Client;
    if (typeof adGraphClient === "string") {
      context.log.error(adGraphClient);
      context.res = { status: 404 };
      return;
    }

    const searchResults = await adGraphClient
      .api("/users")
      .filter(`mail eq '${newUser.email}' or userPrincipalName eq '${newUser.email}'`)
      .select("id")
      .get();

    if (!Array.isArray(searchResults.value) || searchResults.value.length === 0) {
      return {
        status: 400,
        body: "No users found for that email address",
      };
    }

    if (searchResults.value.length > 1) {
      return {
        status: 400,
        body: "Multiple users found for that email address",
      };
    }

    const adTenantId = process.env["ADTenantId"];
    const clientId = searchResults.value[0].id as string;

    const b2cUser = createNewB2cInternalUser(newUser, adTenantId, clientId);
    await graphClient.api("/users").post(b2cUser);

    return {
      status: 201,
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    context.log.error("Error inviting user to b2c", error);
    return {
      status: 502,
    };
  }
}

async function addB2cUser(newUser: NewRvizUser, graphClient: Client, context: Context) {
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
    return {
      status: 400,
      body: "Invalid object provided",
    };
  }

  const newB2cUser = createNewB2cExternalUser(newUser);

  try {
    await graphClient.api("/users").post(newB2cUser);
    // the returned object from the post call doesn't have much in it so let's return nothing for now
    return {
      status: 201,
    };

    return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } catch (error: any) {
    context.log.error("Error calling graph API", error);
    return {
      status: 400,
      body: "Error saving b2c user",
    };
  }
}

export default httpTrigger;
