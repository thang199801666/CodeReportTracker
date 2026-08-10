import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

async function getRemoteUrl(request) {
  const value = new URL(request.url || "", "http://localhost").searchParams.get(
    "url",
  );
  if (!value || !/^https?:\/\//i.test(value))
    throw new Error("Only HTTP/HTTPS URLs are supported.");
  return value;
}

async function resolveEmbeddedPdf(url) {
  const response = await fetch(url, { redirect: "follow" });
  const html = await response.text();
  const matches = [
    ...html.matchAll(/(?:href|src|data)=["']([^"']+\.pdf(?:\?[^"']*)?)["']/gi),
  ];
  const resolved = matches.map((match) =>
    new URL(match[1], response.url).toString(),
  );
  return resolved[0] || null;
}

function pdfProxy() {
  const handler = (prefix) => async (request, response, next) => {
    if (!request.url?.startsWith(prefix)) return next();
    try {
      const remoteUrl = await getRemoteUrl(request);
      const remote = await fetch(remoteUrl, {
        method: prefix.endsWith("head") ? "HEAD" : "GET",
        redirect: "follow",
      });
      if (!remote.ok) {
        response.statusCode = remote.status;
        response.end();
        return;
      }
      response.statusCode = 200;
      response.setHeader(
        "content-type",
        remote.headers.get("content-type") || "application/pdf",
      );
      if (prefix.endsWith("head")) response.end();
      else response.end(Buffer.from(await remote.arrayBuffer()));
    } catch (error) {
      response.statusCode = 502;
      response.setHeader("content-type", "application/json");
      response.end(JSON.stringify({ error: error.message }));
    }
  };
  const resolveHandler = async (request, response, next) => {
    if (!request.url?.startsWith("/api/pdf-resolve")) return next();
    try {
      const remoteUrl = await getRemoteUrl(request);
      const resolvedUrl = await resolveEmbeddedPdf(remoteUrl);
      response.statusCode = resolvedUrl ? 200 : 404;
      response.setHeader("content-type", "application/json");
      response.end(JSON.stringify({ url: resolvedUrl }));
    } catch (error) {
      response.statusCode = 502;
      response.setHeader("content-type", "application/json");
      response.end(JSON.stringify({ error: error.message }));
    }
  };
  return {
    name: "pdf-proxy",
    configureServer(server) {
      server.middlewares.use(handler("/api/pdf-head"));
      server.middlewares.use(handler("/api/pdf-download"));
      server.middlewares.use(resolveHandler);
    },
    configurePreviewServer(server) {
      server.middlewares.use(handler("/api/pdf-head"));
      server.middlewares.use(handler("/api/pdf-download"));
      server.middlewares.use(resolveHandler);
    },
  };
}

export default defineConfig({ plugins: [react(), pdfProxy()] });
