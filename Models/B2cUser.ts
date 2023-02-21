import { InternalExternal, NewRvizUser, RvizUserFields, SubscriptionStatus, UserLevel } from "./RvizUser";

export type B2cUser = {
  id: string;
  displayName: string;
  givenName: string;
  surname: string;
  jobTitle: string;
  extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External: InternalExternal;
  extension_bb8c6b6c043d4874965ca693e13a3e1e_Level: UserLevel;
  extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status: SubscriptionStatus;
  identities: ObjectIdentity[];
  passwordProfile: PasswordProfile;
  passwordPolicies: PasswordPolicy;
};

type ObjectIdentity = {
  signInType: SignInType;
  issuer: string;
  issuerAssignedId: string;
};

type SignInType = "emailAddress";

type PasswordProfile = {
  password: string;
  forceChangePasswordNextSignIn: boolean;
};

type PasswordPolicy = "DisableStrongPassword" | "DisablePasswordExpiration";

export function createNewB2cUser(rvizUser: NewRvizUser) {
  return {
    displayName: rvizUser.displayName,
    givenName: rvizUser.givenName,
    surname: rvizUser.surname,
    jobTitle: rvizUser.jobTitle,
    extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External: rvizUser.internalExternal,
    extension_bb8c6b6c043d4874965ca693e13a3e1e_Level: rvizUser.level,
    extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status: rvizUser.subscriptionStatus,
    identities: [
      {
        signInType: "emailAddress",
        issuer: "rviz2.onmicrosoft.com",
        issuerAssignedId: rvizUser.email,
      },
    ],
    passwordProfile: {
      password: rvizUser.password,
      forceChangePasswordNextSignIn: true,
    },
    passwordPolicies: "DisablePasswordExpiration",
  } as B2cUser;
}

export function createExistingB2cUser(rvizUser: RvizUserFields) {
  return {
    displayName: rvizUser.displayName,
    givenName: rvizUser.givenName,
    surname: rvizUser.surname,
    jobTitle: rvizUser.jobTitle,
    extension_bb8c6b6c043d4874965ca693e13a3e1e_Internal_External: rvizUser.internalExternal,
    extension_bb8c6b6c043d4874965ca693e13a3e1e_Level: rvizUser.level,
    extension_bb8c6b6c043d4874965ca693e13a3e1e_Subscription_Status: rvizUser.subscriptionStatus,
  } as B2cUser;
}
