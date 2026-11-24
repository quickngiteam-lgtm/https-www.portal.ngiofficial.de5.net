# NGI Leader's Portal — GitHub Pages + Google Apps Script

This repository contains the static website files for GitHub Pages. The Google Apps Script backend is kept separately in `backend/Code.gs`; GitHub Pages does not execute `.gs` files.

## Files

- `index.html` — portal frontend; original layout and styling retained.
- `manifest.json` — installable PWA metadata.
- `pwabuilder-sw.js` — service worker for the app shell. It does not cache API responses.
- `icons/` — PWA icons.
- `backend/Code.gs` — Apps Script backend source; deploy this separately as a Web App.

## Publish on GitHub Pages

1. Create a GitHub repository, for example `ngi-leaders-portal`.
2. Upload the **contents** of this folder to the repository root (not the enclosing folder).
3. Open **Settings → Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select branch `main` and folder `/ (root)`, then Save.
6. Wait for GitHub Pages to publish the site. The URL will usually be `https://YOUR-USERNAME.github.io/ngi-leaders-portal/`.

## Deploy the backend

1. Open `script.google.com` and create/open the Apps Script project connected to the portal's existing Google Spreadsheet.
2. Copy the contents of `backend/Code.gs` into the Apps Script editor, preserving any existing project configuration and spreadsheet IDs.
3. Select **Deploy → New deployment → Web app**.
4. Choose the required execution identity and access setting for your organization's security policy, then deploy.
5. Copy the Web App `/exec` URL and compare it with the `API_URL` constant near the beginning of the frontend script in `index.html`. Update it only if the deployed endpoint is different.
6. If you change the backend later, use **Deploy → Manage deployments → Edit → New version → Deploy**.

## Important checks before going live

- Do not replace the existing Google Spreadsheet or its data.
- Confirm the Apps Script Web App URL, permissions, and authorized account.
- Test login, master OTP delivery/verification, member profile/photo loading, attendance submission, admin attendance view, and all role-specific actions.
- GitHub Pages is static hosting; it cannot run `Code.gs` or protect secrets embedded in browser JavaScript.
- The current frontend contains an Apps Script endpoint URL. Treat it as a public endpoint and enforce authorization in Apps Script; never put passwords, private keys, or service-account secrets in `index.html`.

Syntax/archive checks do not prove that live email delivery, spreadsheet permissions, or deployed Apps Script behavior are working. Verify those in the deployed environment.
