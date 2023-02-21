import { HttpRequestHeaders } from "@azure/functions";
import jwtDecode from "jwt-decode";

export default function isAdmin(headers: HttpRequestHeaders) {
  const token = jwtDecode(headers["authentication"]) as object;
  // eslint-disable-next-line no-prototype-builtins
  if (token.hasOwnProperty("extension_Level")) {
    const level = token["extension_Level"] as string;
    if (level === "Super User" || level === "Admin") {
      return true;
    }
  }
  return false;
}
