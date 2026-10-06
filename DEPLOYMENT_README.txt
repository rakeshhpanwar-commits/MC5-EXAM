ASTI EXAM PORTAL — COMPLETE TWO-BACKEND PACKAGE
================================================

ARCHITECTURE
------------
A. PAPER GOOGLE SPREADSHEET / PAPER BACKEND
   File: PAPER_GOOGLE_SHEET_BACKEND.gs
   Purpose: Stores question papers only.
   Paper Spreadsheet ID configured in the script:
   Stored automatically by setupPaperBackend(); no hard-coded Spreadsheet ID is used.

   It maintains exactly these seven paper-code sheets:
   BASIC
   SCREENER(INT)
   SCREENER(REC)
   INSTRUCTOR
   BASIC REFRESHER
   INDUCTION
   OJT

   Each sheet can contain multiple papers. Column A is PaperName.

B. STUDENT GOOGLE SPREADSHEET / STUDENT BACKEND
   File: STUDENT_GOOGLE_SHEET_BACKEND.gs
   Purpose: Students, batches, assignments, schedules, modules, results,
   permanent history, live tracker, notices, materials, etc.
   It does not receive question uploads from the portal.

FRONTEND
--------
File: INDEX.html

Configured Web Apps:
Student backend:
https://script.google.com/macros/s/AKfycby98naN8XiOe2af9fpz_YFYaQpRUeWkF_OCQCk-VVUv41eZQeol2KCZ4zYzM2qEYufK/exec

Paper backend:
https://script.google.com/macros/s/AKfycbxagb8k0giXYlJ5VQdCuCxJ-KHTBiS7IiIZKJGgouaR8HCCa549Ru_VaVjhRbhmvWmc/exec

QUESTION PAPER WORKFLOW
-----------------------
Question Papers -> enter Paper Name -> select Paper Code -> select XLSX -> Upload.
The XLSX is sent to the Paper backend and is written to the selected Paper Code
sheet in the Paper Google Spreadsheet. Uploading the same Paper Code + Paper Name
replaces only that paper.

MANAGE CONTENT
--------------
Manage Content -> Paper Code -> Paper Name -> existing assignment settings.
Paper Code values are available locally in the frontend. Paper Name values are
loaded from the Paper Google Spreadsheet through the Paper backend.

DEPLOYMENT
----------
1. Open the dedicated PAPER Google Spreadsheet (the spreadsheet where the seven paper-code sheets must be stored).
2. Open Extensions -> Apps Script.
3. Put ONLY PAPER_GOOGLE_SHEET_BACKEND.gs in that project.
4. Save the project. From the function dropdown select setupPaperBackend and click Run. Approve Google authorization if requested. This stores the ACTUAL Paper Spreadsheet ID in Script Properties and creates the seven paper-code sheets.
5. Deploy -> New deployment -> Web app.
5. Execute as: Me.
6. Who has access: Anyone.
7. Copy the deployment URL if it changes.
8. Open the STUDENT Google Spreadsheet.
9. Open Extensions -> Apps Script.
10. Put ONLY STUDENT_GOOGLE_SHEET_BACKEND.gs in that project.
11. Deploy -> New deployment -> Web app.
12. Execute as: Me.
13. Who has access: Anyone.
14. If either URL changes, update only the corresponding constant in INDEX.html:
       RECORDS_API_URL = Student URL
       QUESTIONS_API_URL = Paper URL
15. Deploy INDEX.html to Netlify.

IMPORTANT
---------
Do not put PAPER_GOOGLE_SHEET_BACKEND.gs into the Student Spreadsheet project.
Do not put STUDENT_GOOGLE_SHEET_BACKEND.gs into the Paper Spreadsheet project.
Do not add another backend .gs file to the package.

ADMIN PASSWORD
--------------
The frontend synchronizes an admin password change to both backends.
Initial/default password in the supplied backend code is: Rakesh
After changing it through the portal, the new password is stored in both
Apps Script projects.

LOCAL VALIDATION PERFORMED
--------------------------
- INDEX.html: all JavaScript <script> blocks passed Node syntax checking.
- PAPER_GOOGLE_SHEET_BACKEND.gs: passed JavaScript syntax checking after treating
  the .gs source as JavaScript.
- STUDENT_GOOGLE_SHEET_BACKEND.gs: passed JavaScript syntax checking after treating
  the .gs source as JavaScript.
- Package contains exactly one Student backend and one Paper backend.

LIVE GOOGLE VALIDATION
----------------------
A live Apps Script request could not be executed from this build environment,
so deployment/account permissions must be verified in your Google account.
After deployment, first test these actions:
1. Open Paper backend with action=testPaperSpreadsheet.
2. Confirm the response shows the actual Paper Spreadsheet ID/name and the seven Paper Code sheets.
3. If it reports that the Paper Spreadsheet is not configured, run setupPaperBackend() once in the Paper Spreadsheet Apps Script project.
3. Upload one sample XLSX and verify its PaperName rows in the selected code sheet.
4. Open Manage Content and confirm Paper Name loads after Paper Code selection.
5. Login with a student and verify the exam loads the selected Paper Code + Paper Name.


FAST UPLOAD: The admin Excel uploader now parses XLSX in the browser and posts only question data to the PAPER backend action savePaperQuestions. No 30-second Paper Name polling is used.


REGISTERED STUDENT ONE-SHEET MODEL
----------------------------------
The Admin -> Registered Students first section now accepts:
1. Batch Name
2. Paper Code / Programme Course Code dropdown
3. XLSX student file
4. Upload Excel Batch

The student registration data is stored in ONE Google Sheet tab named:
Registered Students

Columns:
Serial Number | CISF NO | Rank | Name | Paper Code | Batch Name

Re-uploading the same Batch Name + Paper Code replaces only that registration.
Deleting a registered batch removes its rows from the same Registered Students tab.

No new Students_<COURSE> sheets are created by the new registration uploader.
Existing legacy Students_* sheets are left untouched for backward compatibility.

LOGIN COMPATIBILITY
--------------------
New registrations use CISF NO as the initial student password because the requested
six-column Registered Students sheet intentionally contains no Password column.
The existing first-login password-change flow is preserved; changed passwords are
stored in Apps Script Script Properties, not in the six-column sheet.

DEPLOYMENT UPDATE REQUIRED
--------------------------
Replace the deployed STUDENT_GOOGLE_SHEET_BACKEND.gs with the version in this package
and deploy a NEW Web App version. The existing Paper backend does not need to change.


FINAL CONFIGURATION UPDATE (2026-09-22):
Student/Records Web App URL: https://script.google.com/macros/s/AKfycbzzKk98j4u1e_p7fWXHbu6LiepBio_QB_CQTfH4xf_jPMmrq07ss8PiGMouNHr1HKA_gw/exec
Question Paper Web App URL: https://script.google.com/macros/s/AKfycbxPiPd8tv7bdW-3Z1OM81ZfUPuNmZTJG3MqvBtV-Ma3UV3JS7-MURTxMVY4wbu2zMfozQ/exec
Admin password default: Rakesh


STUDENT BACKEND - NEW GOOGLE SHEET SETUP (IMPORTANT)
1. Create/open the NEW Student Google Sheet.
2. Extensions > Apps Script.
3. Replace the Student backend with STUDENT_GOOGLE_SHEET_BACKEND.gs from this package.
4. Save.
5. In Apps Script, select setupStudentBackend and click Run once. Authorize if prompted.
6. Confirm the Google Sheet now contains Registered Students with these headers:
   Serial Number | CISF NO | Rank | Name | Paper Code | Batch Name
7. Deploy > Manage deployments > Edit the Web App > create a new version and Deploy.
8. Execute as Me; Who has access: Anyone.
9. The frontend Student URL must point to this deployment.
IMPORTANT: This version has NO old hard-coded Student spreadsheet ID. It uses the Google Sheet to which the Apps Script is bound.
