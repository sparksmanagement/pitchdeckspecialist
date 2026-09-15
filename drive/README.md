# Drive upload — one-time setup (about 3 minutes)

Decks upload themselves to **Pitch Decks / <Client> /** in Google Drive once this is in
place. It uses a tiny Google Apps Script web app that runs as your Google account, so
there is nothing to renew and no Google Cloud project to manage.

1. Open https://script.google.com and click **New project**. Delete the sample code and
   paste the contents of `drive/AppsScript.gs`. Change `SHARED_KEY` to any long random
   string (a password generator is fine). Name the project "Sparks Pitch Decks Uploader"
   and save.
2. Click **Deploy → New deployment**. Under *Select type* choose **Web app**. Set
   *Execute as*: **Me**, *Who has access*: **Anyone**. Click **Deploy**, approve the
   permissions prompt (Drive access for your own account), then copy the **Web app URL**.
3. At claude.ai/code, open the **Sparks** environment dialog and add two environment
   variables:
   ```
   DRIVE_UPLOAD_URL=<the Web app URL from step 2>
   DRIVE_UPLOAD_KEY=<the SHARED_KEY you set in step 1>
   ```
   Save. New sessions pick this up automatically.
4. If the environment uses a **Custom** network allow-list, it must include
   `script.google.com` and `script.googleusercontent.com` (Google answers the upload
   through a redirect to the second host). Without the second one the upload still
   completes, but the script can't read back the file link.

From then on `npm run deck -- prospects/<slug>.yml` builds the deck and uploads it, and
`npm run publish -- prospects/<slug>.yml` uploads an already-built deck. Re-uploading
replaces the previous copy (the old one goes to Drive's trash).

Anyone with both the URL and the key can put files in your Pitch Decks folder, so treat
the key like a password. To rotate it, change `SHARED_KEY` in the script, redeploy
(**Deploy → Manage deployments → edit → New version**), and update the variable.

Alternatives the same script supports, if you ever prefer them: a Google service account
(`GOOGLE_SERVICE_ACCOUNT_JSON` or `GOOGLE_APPLICATION_CREDENTIALS`, folder shared with
the service account as Editor) or a short-lived user token (`GOOGLE_OAUTH_TOKEN`).
