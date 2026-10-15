#!/usr/bin/env node
// Build a deck and, when a Drive credential is configured, upload it to Pitch Decks/<Client>/.
//   npm run deck -- prospects/<slug>.yml [more...] [--short]
const { spawnSync } = require("child_process");
const path = require("path");
const flags = process.argv.slice(2).filter((a) => a.startsWith("--"));
const files = process.argv.slice(2).filter((a) => !a.startsWith("--"));
if (!files.length) { console.error("usage: npm run deck -- prospects/<slug>.yml [--short]"); process.exit(1); }
for (const f of files) {
  const r = spawnSync(process.execPath, [path.join(__dirname, "build_deck.js"), f, ...flags], { stdio: "inherit" });
  if (r.status !== 0) process.exit(r.status || 1);
}
const hasDrive = ["DRIVE_UPLOAD_URL", "GOOGLE_SERVICE_ACCOUNT_JSON", "GOOGLE_APPLICATION_CREDENTIALS", "GOOGLE_OAUTH_TOKEN", "DRIVE_AUTH_VIA_PROXY"].some((k) => process.env[k]);
if (!hasDrive) { console.log("Built. Drive upload skipped: no credential configured (see drive/README.md)."); process.exit(0); }
const p = spawnSync(process.execPath, [path.join(__dirname, "publish_to_drive.js"), ...files, ...flags], { stdio: "inherit" });
process.exit(p.status || 0);
