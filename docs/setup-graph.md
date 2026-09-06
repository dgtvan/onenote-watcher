# Microsoft Entra App Registration — Step by Step

This is the one-time setup that lets the watcher read your notebooks' server-side state through
Microsoft Graph. It is free, takes ~5 minutes, and needs no admin rights on your PC.

## Answer to "Delegated or Application permissions?"

**Delegated permissions.** And **use Microsoft Graph, not the legacy "OneNote" API** shown in your
screenshot.

Why:

- **Delegated** = "act as the signed-in user." The watcher signs in as *you* and reads *your*
  notebooks. That is exactly this app. **Application permissions** = a headless daemon with no
  user, using its own app secret; for personal Microsoft accounts it is **not supported at all**
  (app-only access needs an organizational tenant and an admin to grant tenant-wide consent). So
  Application permissions is both wrong for the design and impossible for your consumer OneDrive
  notebooks.
- **Graph, not legacy OneNote API:** your own screenshot says it — *"OneNote APIs are available via
  the Microsoft Graph API. You may want to consider using Microsoft Graph instead."* The legacy
  `https://onenote.com` API is deprecated; Graph is the supported path and the endpoints in
  `reference-data-sources.md` are Graph endpoints. Remove the legacy "OneNote" permission if you
  already added it.

So on that screen: click **← All APIs**, choose **Microsoft Graph**, then **Delegated
permissions**, then tick **Notes.Read** (and **offline_access**). Details below.

## Exact steps

Do this at <https://entra.microsoft.com> (or <https://portal.azure.com> → *App registrations*),
signed in with your **personal Microsoft account** (dgtvan.ai@gmail.com).

### 1. Create the registration (if not already done)

- *App registrations* → **New registration**.
- **Name:** `OneNote Sync Watcher` (any name).
- **Supported account types:** select **"Personal Microsoft accounts only"**.
  (If you want it to also work with a work account later, pick "Accounts in any org directory and
  personal Microsoft accounts". For your consumer notebooks, "Personal … only" is correct.)
- **Redirect URI:** leave blank for now — set it in step 3.
- **Register.**

### 2. API permissions

- *API permissions* → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
- Search and tick: **`Notes.Read`** (read your OneNote notebooks) and **`offline_access`**
  (lets the watcher refresh its token silently so you sign in only once).
- **Add permissions.**
- You do **not** need "Grant admin consent" — for a personal account these are
  user-consentable; you'll approve them once at first sign-in.

> If you ever want the watcher to *fix* things (e.g. trigger a sync) rather than only read,
> that would be a different scope — not needed for this app. Keep it to `Notes.Read`.

### 3. Allow the device-code / public-client sign-in

- *Authentication* → **Add a platform** → **Mobile and desktop applications**.
- Tick the suggested redirect URI
  **`https://login.microsoftonline.com/common/oauth2/nativeclient`**.
- Scroll to **Advanced settings** → **Allow public client flows** → set to **Yes**.
  (This enables the device-code flow the watcher uses: it shows you a short code, you paste it at
  <https://microsoft.com/devicelogin>, and it never handles your password.)
- **Save.**

### 4. Copy two values into `config.ini`

- *Overview* page → copy **Application (client) ID** → paste into `config.ini` under
  `[graph] client_id = …`.
- Leave `[graph] tenant = consumers` (this is correct for a personal Microsoft account).

That's everything. No client secret is needed (public-client device-code flow uses none).

## What you'll see at first run

The watcher prints/points you to: *"To sign in, open <https://microsoft.com/devicelogin> and enter
code `XXXX-XXXX`."* You approve once; the token is cached (encrypted with Windows DPAPI, per-user)
so subsequent launches are silent until the refresh token eventually expires (~90 days of
inactivity), at which point the tray icon turns yellow and asks you to sign in again.

## If sign-in fails

| Symptom | Fix |
|---|---|
| "The app doesn't exist / wrong account type" | account-types must include *Personal Microsoft accounts*; re-check step 1 |
| "AADSTS7000218 / public client not enabled" | *Allow public client flows* = Yes (step 3) |
| "Need admin approval" | you picked Application permissions or an admin-only scope — switch to **Delegated** `Notes.Read` |
| consent screen lists scary permissions | it should only ask for *Read your OneNote notebooks* + *Maintain access to data you've given it access to* |
