# EZ-Ensemble Plumbing Project

EZ-Ensemble is our capstone project for the BCA music department. This demo shows one Flute audition score from the teacher’s screen to Google Sheets and back to the student’s screen, which displays the total score.
## What we used

- **React** displays the role picker, teacher scoring table, and student total.
- **Microsoft OAuth** handles sign-in. The app uses Microsoft Graph to read the signed-in user’s name and email.
- **Google Apps Script** hosts the page, handles the sign-in callback, checks the selected role, and reads and writes scores.
- **Google Sheets** stores the instrument and its three scores.

We edit the project in VS Code, track changes with Git, and use `clasp` to upload the code to Apps Script. The React page is in `Index.html`, the backend is in `Code.js`, and the Apps Script settings are in `appsscript.json`.




























## Run the demo (if you need it)

[Open EZ-Ensemble](https://script.google.com/macros/s/AKfycbxQqzQhhDTYwFN8HAExT7jS7oQfUZTrB_O7DIUeIcsmz8ZZQMUVeI3DB87MWFUqfzl6Ig/exec). You can also open `Index.html` from the project folder and click **Open EZ-Ensemble**. The demo needs an internet connection. Use the project owner’s Google account to open the private web app, then sign in with an approved Microsoft demo account. You can choose either role to try both views.

1. Choose **Teacher**, click **Sign in with Microsoft**, and follow **Continue to Microsoft**.
2. Enter **8** for Tone, **9** for Pitch, and **7** for Rhythm. Click **Update Spreadsheet**. In the **EZ-Ensemble Plumbing Demo** spreadsheet, the **Audition Scores** tab now has one Flute row. Saving again updates that same row.
3. Change Tone to **10** directly in the sheet. Back in the app, click **Retrieve Spreadsheet Information**. The page should show the saved scores of **10, 9, and 7**.
4. Click **Log out** and sign in as **Student**. You should see **Flute total: 26**. Students can refresh the total. Only the teacher role can see the individual scores and edit them.