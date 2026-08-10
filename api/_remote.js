export function getRemoteUrl(req) {
  const value = new URL(req.url, "https://vercel.local").searchParams.get("url");
  if (!value || !/^https?:\/\//i.test(value))
    throw new Error("Only HTTP/HTTPS URLs are supported.");
  return value;
}

export function sendError(response, error) {
  response.status(502).json({ error: error.message || "Remote request failed." });
}
