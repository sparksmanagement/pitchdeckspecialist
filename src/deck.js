#!/usr/bin/env node
// Build a deck and, when a Drive credential is configured, upload it to Pitch Decks/<Client>/.
//   npm run deck -- prospects/<slug>.yml [more...]
const { spawnSync } = require("child_process");
const path = require("path");
const files = process.argv.slice(2);
if (!files.length) { console.error("usage: npm run deck -- prospects/<slug>.yml"); process.exit(1); }
for (const f of files) {
  const r = spawnSync(process.execPath, [path.join(__dirname, "build_deck.js"), f], { stdio: "inherit" });
  if (r.status !== 0) process.exit(r.status || 1);
}
const hasDrive = ["DRIVE_UPLOAD_URL", "GOOGLE_SERVICE_ACCOUNT_JSON", "GOOGLE_APPLICATION_CREDENTIALS", "GOOGLE_OAUTH_TOKEN", "DRIVE_AUTH_VIA_PROXY"].some((k) => process.env[k]);
if (!hasDrive) { console.log("Built. Drive upload skipped: no credential configured (see drive/README.md)."); process.exit(0); }
const p = spawnSync(process.execPath, [path.join(__dirname, "publish_to_drive.js"), ...files], { stdio: "inherit" });
process.exit(p.status || 0);
