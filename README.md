# BESI GitHub Pages entry page

This folder contains the public entry page for the BetterEdge Solutions Inc. workforce portal. GitHub Pages serves static HTML. The portal's login, records, and spreadsheet operations run separately as a Google Apps Script web app.

## Set up

1. Deploy the workforce portal from the existing `Code.gs` and `Index.html` in Google Apps Script as a web app. Set access to the intended users, then copy the deployed URL ending in `/exec`.
2. Open this folder's `index.html`. Replace `PASTE_APPS_SCRIPT_WEB_APP_URL_HERE` with that exact `/exec` URL. Save the file.
3. Copy `index.html` to the root of your `BESI-APP` repository (alongside the existing repository files, if any). You may also copy this README. In GitHub Desktop, commit and push to `main`.
4. In the repository on GitHub, open **Settings → Pages**. Under **Build and deployment**, choose **Deploy from a branch**, select **main** and **/(root)**, then save.
5. Open the Pages URL shown in Settings → Pages. Click **Open Management Portal** and verify the deployed Apps Script portal opens and its sign-in works.

If you later redeploy the Apps Script web app with a different URL, update `PORTAL_URL` in `index.html` and push again.

The GitHub Pages page is public. Keep employee data, passwords, spreadsheet IDs, and server-side application code out of the public repository. The CSS brand mark is a light-blue and pink approximation because the uploaded logo file was not available to embed.
