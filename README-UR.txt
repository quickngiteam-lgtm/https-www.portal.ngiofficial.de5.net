NGI Leader's Portal — GitHub Pages package

GitHub Pages ke liye is folder ke tamam contents repository ke root mein upload karein.
GitHub: Settings > Pages > Deploy from a branch > main > /(root) > Save.

Website frontend: index.html
PWA config: manifest.json
Service worker: pwabuilder-sw.js
Icons: icons/
Google Apps Script backend: backend/Code.gs

Zaroori: Code.gs GitHub Pages par execute nahi hoti. Isay script.google.com par Web App ke taur par alag deploy karein. Frontend ke index.html mein API_URL ko deployed Apps Script /exec URL se match rakhein. Existing spreadsheet ID/data ko replace na karein.

Live OTP, Google Spreadsheet permissions, login aur attendance ko deployment ke baad test karna zaroori hai.
