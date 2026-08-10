import { getRemoteUrl, sendError } from "./_remote.js";

export default async function handler(request, response) {
  try {
    const remote = await fetch(getRemoteUrl(request), {
      method: "HEAD",
      redirect: "follow",
    });
    response.status(remote.status);
    response.setHeader(
      "content-type",
      remote.headers.get("content-type") || "application/pdf",
    );
    response.end();
  } catch (error) {
    sendError(response, error);
  }
}
