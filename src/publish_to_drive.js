#!/usr/bin/env node
/**
 * Upload a built deck to the client's folder in Google Drive.
 *
 *   node src/publish_to_drive.js prospects/<slug>.yml [more prospect files...]
 *
 * Auth: a Google service account. Set GOOGLE_SERVICE_ACCOUNT_JSON (the key file's JSON
 * as a string) or GOOGLE_APPLICATION_CREDENTIALS (path to the key file), and share the
 * "Pitch Decks" Drive folder (brand/config.yml → drive.pitch_decks_folder) with the
 * service account's email as Editor.
 *
 * For each prospect: find-or-create a folder named after prospect.name under the Pitch
 * Decks folder, then upload decks/Sparks-x-<slug>.pptx (replacing a file of the same
 * name so Drive keeps version history).
 */
const fs = require("fs");
const path = require("path");
const yaml = require("js-yaml");
const { google } = require("googleapis");

const ROOT = path.resolve(__dirname, "..");
const brand = yaml.load(fs.readFileSync(path.join(ROOT, "brand/config.yml"), "utf8"));
const ROOT_FOLDER = brand.drive && brand.drive.pitch_decks_folder;
const files = process.argv.slice(2);
if (!files.length || !ROOT_FOLDER) {
  console.error("usage: node src/publish_to_drive.js prospects/<slug>.yml  (needs drive.pitch_decks_folder in brand/config.yml)");
  process.exit(1);
}

function auth() {
  const scopes = ["https://www.googleapis.com/auth/drive"];
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    return new google.auth.GoogleAuth({ credentials: creds, scopes });
  }
  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) return new google.auth.GoogleAuth({ scopes });
  console.error("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS (service account with Drive access).");
  process.exit(2);
}

const esc = (s) => String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'");

async function findOrCreateFolder(drive, name, parent) {
  const q = `name = '${esc(name)}' and '${parent}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
  const res = await drive.files.list({ q, fields: "files(id, name)", supportsAllDrives: true, includeItemsFromAllDrives: true });
  if (res.data.files && res.data.files.length) return res.data.files[0];
  const created = await drive.files.create({
    requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: [parent] },
    fields: "id, name", supportsAllDrives: true,
  });
  console.log(`Created folder "${name}"`);
  return created.data;
}

async function upload(drive, folderId, filePath) {
  const name = path.basename(filePath);
  const mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const q = `name = '${esc(name)}' and '${folderId}' in parents and trashed = false`;
  const existing = await drive.files.list({ q, fields: "files(id, name)", supportsAllDrives: true, includeItemsFromAllDrives: true });
  const media = { mimeType, body: fs.createReadStream(filePath) };
  if (existing.data.files && existing.data.files.length) {
    const id = existing.data.files[0].id;
    await drive.files.update({ fileId: id, media, supportsAllDrives: true });
    return { id, updated: true };
  }
  const res = await drive.files.create({ requestBody: { name, parents: [folderId] }, media, fields: "id", supportsAllDrives: true });
  return { id: res.data.id, updated: false };
}

(async () => {
  const drive = google.drive({ version: "v3", auth: auth() });
  for (const f of files) {
    const P = yaml.load(fs.readFileSync(path.resolve(f), "utf8"));
    const client = (P.prospect && P.prospect.name) || path.basename(f, ".yml");
    const slug = path.basename(f, ".yml").replace(/^_/, "");
    const deck = path.join(ROOT, "decks", `Sparks-x-${slug}.pptx`);
    if (!fs.existsSync(deck)) { console.error(`No deck at ${deck} — run npm run build -- ${f} first`); process.exitCode = 1; continue; }
    const folder = await findOrCreateFolder(drive, client, ROOT_FOLDER);
    const { id, updated } = await upload(drive, folder.id, deck);
    console.log(`${updated ? "Updated" : "Uploaded"} ${path.basename(deck)} → ${client}/  https://drive.google.com/file/d/${id}/view`);
  }
})().catch((e) => { console.error(e.message || e); process.exit(1); });
