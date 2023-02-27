import { B2cUser } from "./B2cUser";

// A copy of this file is also in the web project.  This is bad practice but unfortunately at the moment
// there is no way for CRA to import from outside it's src folder without ejecting
// That doesn't seem worth it at the moment
// Best solution would be to publish this to npm, either public or private (paid)

export type RvizUserFields = {
  displayName: string;
  givenName: string;
  surname: string;
  email: string;
  jobTitle: string;
  internalExternal: InternalExternal;
  level: UserLevel;
  subscriptionStatus: SubscriptionStatus;
};

export type NewRvizUser = RvizUserFields & {
  password: string;
};

export type RvizUser = RvizUserFields & {
  id: string;
};

export const internalExternalValues = ["Internal", "External"] as const;
export type InternalExternal = (typeof internalExternalValues)[number];

export const userLevelValues = ["Admin", "Super User", "User"] as const;
export type UserLevel = (typeof userLevelValues)[number];

export const subscriptionStatusValues = ["Free", "Trial", "Pro"];
export type SubscriptionStatus = (typeof subscriptionStatusValues)[number];

export function createRvizUser(b2cUser: B2cUser) {
  let mail = b2cUser.mail;
  if (!mail || mail === "") {
    const rvizOn = b2cUser.identities.find((i) => i.issuer === "rvizcouk.onmicrosoft.com");
    if (rvizOn) {
      mail = rvizOn.issuerAssignedId;
    }
  }

  return {
    id: b2cUser.id,
    displayName: b2cUser.displayName,
    givenName: b2cUser.givenName,
    surname: b2cUser.surname,
    email: mail,
    jobTitle: b2cUser.jobTitle,
    internalExternal: b2cUser.extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External ?? "",
    level: b2cUser.extension_bb8c6b6c043d4874965ca693e13a3e1e_Level ?? "",
    subscriptionStatus: b2cUser.extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status ?? "",
  } as RvizUser;
}
