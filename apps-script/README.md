# Centrum Přítomnost — Google Apps Script setup

This backend is dedicated to the Centrum Přítomnost booking flow.

## Ownership
Create and deploy the Apps Script while signed in as **centrumpritomnost@gmail.com**. The script uses that account's primary Google Calendar and sends notifications from that account.

## 1. Create the Apps Script project
1. Open https://script.google.com/ while signed in as centrumpritomnost@gmail.com.
2. Create a new standalone project named **Centrum Přítomnost – Rezervace**.
3. Replace the starter code in Code.gs with the contents of this folder's Code.gs.
4. In Project Settings, set the time zone to **Europe/Prague**.
5. Save the project.

## 2. Initialize the spreadsheet and calendar access
1. Run the function **setupCentrum** from the Apps Script editor.
2. Approve the requested Google Sheets, Calendar and email permissions.
3. Confirm that the spreadsheet contains **Requests** and **Settings** tabs.
4. Confirm the function reports that the primary calendar is accessible.

The Requests tab stores one row per booking request. The Settings tab documents the configured owner email, calendar and time zone.

## 3. Deploy
Deploy as a **Web app**:
- Execute as: **Me (centrumpritomnost@gmail.com)**
- Who has access: **Anyone**

Copy the Web app URL. Do not publish any credentials or authorization tokens in the website source.

## Booking behavior
- Requests are always created with status **PENDING**.
- A pending calendar event is created to temporarily block that requested time and reduce double-booking risk.
- The administrator receives an email with the request details.
- The customer receives an acknowledgement explaining that the request is not yet confirmed.
- The administrator must manually contact the customer to confirm or suggest another time.

## Important implementation note
The website integration must use a browser-safe transport for Apps Script Web Apps. Do not assume that a normal cross-origin JSON fetch will work: Apps Script redirects and CORS behavior need to be handled deliberately. The site-side booking UI should be connected only after the Web App URL is available and the complete request/availability flow has been tested.

## Current configuration
- Spreadsheet ID: 1h1uEVdEO_8-4c6YwMMuHA3t9SvqY3Nqofc8nSPiPY8k
- Administrator: centrumpritomnost@gmail.com
- Calendar: primary calendar of the deploying account
- Time zone: Europe/Prague
