# Workflows

No deployment workflow is enabled. Legacy workflows require audit before retirement or reuse.

## capture-live-source.yml (read-only)

Pulls the live Apps Script project with `clasp pull` (never `push`/`deploy`), fingerprints it, compares it with `main`, and commits the result to the `live-capture` branch. It runs daily and on demand (Actions → "Capture live Apps Script source" → Run workflow). Diff `main...live-capture` to see drift.

Setup (one time):
1. Enable the Apps Script API for the owning account: https://script.google.com/home/usersettings
2. Locally: `npm i -g @google/clasp@2.4.2 && clasp login`, then copy the contents of `~/.clasprc.json`.
3. Repo Settings → Secrets and variables → Actions:
   - Secret `CLASPRC_JSON` = that file's contents.
   - Variable `APPS_SCRIPT_ID` = Script ID from the Apps Script editor → Project Settings.

The fingerprint here is SHA-256 over the sorted per-file SHA-256 list; it is not necessarily the same method as the baseline fingerprint in the root README.
