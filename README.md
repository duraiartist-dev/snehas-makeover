# Sneha's Makeover - Bridal Booking Website

Responsive website for a Madurai bridal makeup artist with a real booking system:
customers see date availability and send a request, the owner confirms or cancels
from a password-protected dashboard. Data lives in Google Sheets (Apps Script backend).

Live: https://snehas-makeover.netlify.app  |  Dashboard: /admin.html

## Features
- Date availability check (blocked dates + already confirmed dates)
- Booking request saved to Google Sheet, then WhatsApp opens with the details and a booking ID
- Owner dashboard: login, Confirm / Cancel, WhatsApp the customer, block leave dates
- Spam protection (honeypot, per-phone limit), login lockout, input sanitising

## Tech
HTML, CSS, JavaScript, Google Apps Script, Google Sheets, Netlify

## Folder
index.html | admin.html | images/ | apps-script/Code.gs

## Setup (one time)
1. Create a Google Sheet "Sneha Bookings". Extensions > Apps Script. Paste apps-script/Code.gs.
2. Project Settings: set time zone to Asia/Kolkata.
3. Run the `setup` function once (allow permissions). The two sheets are created.
4. Project Settings > Script properties > add `ADMIN_PASSWORD` = a strong password.
5. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the /exec URL.
6. Paste that URL into `API_URL` in index.html and admin.html.
7. Push to GitHub, Netlify auto-deploys.

After changing Code.gs: Deploy > Manage deployments > edit > New version (the URL stays the same).
