import { HttpRequestHeaders, Context } from "@azure/functions";
import jwtDecode from "jwt-decode";

export default function enforceAdmin(context: Context) {
  if (!isAdmin(context.req.headers)) {
    context.log.error("Attempt to access admin endpoint by non admin user");
    context.res = { status: 401 };
    return false;
  }
  return true;
}

function isAdmin(headers: HttpRequestHeaders) {
  const authHeader = headers["authentication"];
  if (authHeader && authHeader !== "") {
    const token = jwtDecode(headers["authentication"]) as object;
    // eslint-disable-next-line no-prototype-builtins
    if (token.hasOwnProperty("extension_Level")) {
      const level = token["extension_Level"] as string;
      if (level === "Super User" || level === "Admin") {
        return true;
      }
    }
  }
  return false;
}
