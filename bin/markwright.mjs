#!/usr/bin/env node
// `npx markwright` — serve the prebuilt task pane locally and sideload it
// into Word desktop in one step. End users never touch a shared-folder
// catalog or the wef directory; office-addin-debugging (fetched on demand
// via npx, so it isn't a dependency) handles registration and launch,
// while bin/serve.mjs hosts dist/ over HTTPS.
import { spawn } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
const manifest = join(root, "dist", "manifest.xml");
const serve = join(here, "serve.mjs");
const command = process.argv[2] || "start";

if (!existsSync(manifest)) {
  console.error(
    "Markwright: dist/manifest.xml is missing — the package looks unbuilt. " +
      "From a source checkout, run `npm run build` first."
  );
  process.exit(1);
}

const npx = process.platform === "win32" ? "npx.cmd" : "npx";

function run(args) {
  const child = spawn(npx, args, { stdio: "inherit" });
  child.on("error", (err) => {
    console.error(`Markwright: could not launch office-addin-debugging via npx — ${err.message}`);
    process.exit(1);
  });
  child.on("exit", (code) => process.exit(code ?? 0));
}

if (command === "stop") {
  run(["--yes", "office-addin-debugging@latest", "stop", manifest]);
} else if (command === "start") {
  // Quote both paths: process.execPath and the script path can contain
  // spaces, and office-addin-debugging runs this string through a shell.
  const devServer = `"${process.execPath}" "${serve}"`;
  run([
    "--yes",
    "office-addin-debugging@latest",
    "start",
    manifest,
    "desktop",
    "--app",
    "word",
    "--no-debug",
    "--dev-server",
    devServer,
    "--dev-server-port",
    "3000",
  ]);
} else {
  console.error(`Markwright: unknown command "${command}". Usage: markwright [start|stop]`);
  process.exit(1);
}
