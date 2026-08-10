import { getRemoteUrl, sendError } from "./_remote.js";

export default async function handler(request, response) {
  try {
    const remote = await fetch(getRemoteUrl(request), { redirect: "follow" });
    response.status(remote.status);
    response.setHeader(
      "content-type",
      remote.headers.get("content-type") || "application/pdf",
    );
    if (!remote.ok) {
      response.end();
      return;
    }
    response.end(Buffer.from(await remote.arrayBuffer()));
  } catch (error) {
    sendError(response, error);
  }
}
