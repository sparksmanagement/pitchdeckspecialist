#!/usr/bin/env node
// Build a prospect's decks and, when a Drive credential is configured, upload them to
// Pitch Decks/<Client>/. By default this builds BOTH the full 14-slide deck and the
// 4-slide short deck (the house standard since Oct 2026).
//   npm run deck -- prospects/<slug>.yml [more...]          full + short
//   npm run deck -- prospects/<slug>.yml --full              full only
//   npm run deck -- prospects/<slug>.yml --short             short only
const { spawnSync } = require("child_process");
const path = require("path");
const argv = process.argv.slice(2);
const files = argv.filter((a) => !a.startsWith("--"));
const only = argv.includes("--short") ? ["--short"] : argv.includes("--full") ? [] : null;
const variants = only ? [only] : [[], ["--short"]];
if (!files.length) { console.error("usage: npm run deck -- prospects/<slug>.yml [--full | --short]"); process.exit(1); }
const run = (script, args) => spawnSync(process.execPath, [path.join(__dirname, script), ...args], { stdio: "inherit" }).status || 0;
for (const f of files) for (const v of variants) {
  const code = run("build_deck.js", [f, ...v]);
  if (code !== 0) process.exit(code);
}
const hasDrive = ["DRIVE_UPLOAD_URL", "GOOGLE_SERVICE_ACCOUNT_JSON", "GOOGLE_APPLICATION_CREDENTIALS", "GOOGLE_OAUTH_TOKEN", "DRIVE_AUTH_VIA_PROXY"].some((k) => process.env[k]);
if (!hasDrive) { console.log("Built. Drive upload skipped: no credential configured (see drive/README.md)."); process.exit(0); }
let status = 0;
for (const v of variants) status = run("publish_to_drive.js", [...files, ...v]) || status;
process.exit(status);
