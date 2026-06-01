// Minimal HTTPS static server for the prebuilt task pane.
//
// The published package ships `dist/` (already webpack-bundled) instead
// of a webpack toolchain, so `npx markwright` can't reuse the dev
// server. This serves those static files over HTTPS on :3000 using the
// same office-addin-dev-certs CA that Word trusts — office-addin-debugging
// launches it via `--dev-server` and waits for the port before sideloading.
import { createServer } from "node:https";
import { createReadStream, existsSync, statSync } from "node:fs";
import { dirname, extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import devCerts from "office-addin-dev-certs";

const here = dirname(fileURLToPath(import.meta.url));
const distDir = join(here, "..", "dist");
const port = Number(process.env.PORT) || 3000;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".json": "application/json; charset=utf-8",
};

// Word will only load the add-in over HTTPS from a CA it trusts, so make
// sure the local dev cert is installed before we bind the socket.
await devCerts.ensureCertificatesAreInstalled();
const httpsOptions = await devCerts.getHttpsServerOptions();

createServer(httpsOptions, (req, res) => {
  // Mirror the dev server's open CORS policy — office.js is loaded from
  // Microsoft's CDN, cross-origin to localhost.
  res.setHeader("Access-Control-Allow-Origin", "*");

  const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  let filePath = join(distDir, normalize(urlPath));
  if (filePath.endsWith("/") || (existsSync(filePath) && statSync(filePath).isDirectory())) {
    filePath = join(filePath, "taskpane.html");
  }

  // Refuse anything that escapes dist/ (path traversal) or doesn't exist.
  if (!filePath.startsWith(distDir) || !existsSync(filePath)) {
    res.statusCode = 404;
    res.end("Not found");
    return;
  }

  res.setHeader("Content-Type", MIME[extname(filePath)] || "application/octet-stream");
  createReadStream(filePath).pipe(res);
}).listen(port, () => {
  console.log(`Markwright serving dist/ on https://localhost:${port}`);
});
