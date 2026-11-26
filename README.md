# NGI Portal

This repository hosts the static portal frontend on GitHub Pages. Keep the Apps Script backend and its configuration in the private Apps Script project.

## Release order

1. Apply the matching backend update in the private Apps Script project. Deploy a new version of the existing web app so the frontend keeps the same endpoint.
2. Check the backend with a test member and test administrator. Confirm that records and attendance enter Pending Approvals and that the admin can approve or delete a test entry.
3. Merge the frontend pull request and confirm GitHub Pages serves the new version.
4. On an installed test device, verify offline entry, Sync All, and the green online indicator. Test a push announcement with a subscribed test device in a test OneSignal app before broadcasting to the production app.

## Announcement push

For the Median app, configure its OneSignal plugin and set the corresponding `ONESIGNAL_APP_ID` and `ONESIGNAL_API_KEY` in the private Apps Script Script Properties. Keep the API key out of GitHub. Browser-installed PWAs use a separate Web Push configuration.
