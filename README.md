# CSDL Student Assistant Attendance Tracker

Web-based time tracking system for Student Assistants under the Center for 
Student Development and Leadership (CSDL), PHINMA University of Iloilo.

## Features
- Time In / Time Out per student, per AM/PM session
- Course and Year Level based student selection
- Auto time-out safety net (lunch cutoff, end-of-day cutoff)
- Approval workflow (Pending / Approved / Rejected) via Google Sheets menu
- Attendance history view per student
- Daily auto-archiving of records to a "History" sheet

## Tech Stack
- Google Apps Script (backend, `Code.gs`)
- HTML/CSS/JavaScript (frontend, `Index.html`)
- Google Sheets (database)

## Setup
1. Create a Google Sheet
2. Open Extensions → Apps Script
3. Paste `Code.gs` and `Index.html` into the project
4. Deploy as a Web App
5. Run `setupSheetDesign()` once to format the sheets
6. Set up triggers: `createDailyArchiveTrigger()` and `createAutoTimeOutTrigger()`

## Note
This project is for internal use by CSDL, PHINMA University of Iloilo.
