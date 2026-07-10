# AI Video Prompt Builder v1.0 — Apps Script Layer

Status: controlled prototype. This does not constitute production, technical, legal or launch approval.

## Purpose

This bound Google Apps Script adds a safe `Copy Prompt + Open Target Tool` workflow to the live Google Sheet. It does not inject text into third-party websites.

## Included

- `Code.gs` — custom menu, prompt retrieval, tool routing, archive action and experiment-draft action.
- `Sidebar.html` — copy/open/archive interface.
- `appsscript.json` — V8 runtime manifest.

## Install into the live Google Sheet

1. Open the AI Video workbook.
2. Choose **Extensions → Apps Script**.
3. Replace the default script with the contents of `Code.gs`.
4. Add an HTML file named `Sidebar` and paste in `Sidebar.html`.
5. In **Project Settings**, enable the manifest if required and replace it with `appsscript.json`.
6. Save and reload the spreadsheet.
7. Authorise the script when Google requests access.
8. Use **AI Video Tools → Open Prompt Panel**.

## Required sheet contract

The workbook must contain:

- `00_Live_Builder`
- `04_Experiments`
- `07_Final_Prompts`

Required named ranges:

- `PROMPT_FINAL`
- `CANONICAL_BRIEF`
- `MIDJOURNEY_PROMPT`
- `GROK_PROMPT`
- `VIDEO_PROMPT`
- `HEDRA_HEYGEN_PROMPT`
- `IMMERSITY_PROMPT`

## Safety boundary

The script copies prompts to the clipboard and opens the official target site in a new tab. It does not attempt cross-site DOM access or automatic prompt-box injection.

## Test checklist

- Change `00_Live_Builder!B5` and confirm Final Prompt changes.
- Open the sidebar and switch between all prompt outputs.
- Copy each output and paste it into a plain-text editor.
- Open the selected target tool.
- Archive a prompt and confirm a new row appears in `07_Final_Prompts`.
- Create an experiment draft and confirm a new row appears in `04_Experiments`.
- Confirm no existing rows are overwritten.
