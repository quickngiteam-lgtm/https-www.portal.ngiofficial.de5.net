# NGI Portal (frontend)

Static site: upload these files to a GitHub repo and enable GitHub Pages (or use any static host).
Do NOT upload `Code.gs` to a public repo; paste it into the Apps Script project instead.

## Backend setup (Google Apps Script, signed in as quickngiteam@gmail.com)
1. Replace the project's Code.gs with the supplied Code.gs.
2. Authorize the Apps Script project to access Sheets, Mail, Drive, and external requests. The first API request initializes the required sheets.
3. Deploy > Manage deployments > edit the existing Web app deployment > New version (keeps the same /exec URL).
4. Script properties (Project Settings), optional, for push notifications:
   - `ONESIGNAL_APP_ID`, `ONESIGNAL_API_KEY` (from your OneSignal app linked to Median)
   - `PORTAL_URL` (your GitHub Pages URL)
5. Open the site on your phone: the Authorized Admin Login button appears only until a master device exists.
   Send the code to the NGI email, verify, and that device becomes the Master Device.
6. Add more admin devices: Admin Devices tab > Authorize New Device, then on the new device tap the login logo 7 times and enter the code.

Master-device recovery requires access to the deployed Apps Script project.
