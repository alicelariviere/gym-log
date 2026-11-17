# Gym log

A phone app for logging gym weights. It reads the workouts from a Google Sheet and writes each change back into the same cells, so anyone who prefers the sheet still sees it as before.

- **Workouts:** Legs, Pull and Push tabs. Any sheet tab with an "A weight" column shows up.
- **Logging:** big + and − buttons, or tap a number to type it. Saves about a second after you stop tapping.
- **History:** each save also adds a row (date, workout, exercise, person, weight) to the sheet's History tab, and the app shows a trend per exercise.
- **No signal:** the app still opens, keeps your changes on the phone and saves them when you're back online.

## How it fits together

- `index.html`: the app (installable from Chrome: menu > Install app).
- `apps-script/Code.gs`: the Google Apps Script that reads and writes the sheet. The live copy runs in your Google account; this copy has the key removed.
- Setup: the app is pointed at the script once per phone with a private setup link. Nothing private is stored in this repo.
