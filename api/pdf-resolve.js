import { getRemoteUrl, sendError } from "./_remote.js";

export default async function handler(request, response) {
  try {
    const remote = await fetch(getRemoteUrl(request), { redirect: "follow" });
    if (!remote.ok) {
      response.status(remote.status).json({ url: null });
      return;
    }
    const html = await remote.text();
    const matches = [
      ...html.matchAll(/(?:href|src|data)=["']([^"']+\.pdf(?:\?[^"']*)?)["']/gi),
    ];
    const resolved = matches.map((match) =>
      new URL(match[1], remote.url).toString(),
    );
    response.status(200).json({ url: resolved[0] || null });
  } catch (error) {
    sendError(response, error);
  }
}
