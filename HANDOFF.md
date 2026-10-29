# Handoff - martinson-matching
Last updated: 2026-10-29 · Machine: gov laptop · Agent: Claude Code · Branch: main

Public repo: keep this file free of personal data, internal links and sign-in details.

## Objective
Run the honors-program coaching match end to end with minimal work for the program lead: survey ->
automatic matching -> a Teams group chat per coach/coachee pair with an intro message.

## Done
- Matcher (web page on GitHub Pages, CLI `run_match.py`, Office Script `MatchCadets.ts`), form spec, tests.
- Repo pushed to GitHub; GitHub Pages site live (`docs/`).
- Microsoft Form built and matches `form-questions.md` (feedback edits applied 2026-10-13).
- Forms-synced results workbook created; Office Scripts confirmed enabled in the tenant; script installed.
- Older flow (coach gets a Teams deep link + copy-paste intro from a Matches.xlsx) saved and tested.
- New flow (run script -> create Teams group chat per pair -> post intro) saved; a test run reached the
  script successfully in the cloud (failed only because there was a single response).
- 2026-10-29: agent context files added (AGENTS.md, HANDOFF.md, CLAUDE.md, .gitignore secret rules).

## In progress
- Zero-touch flow: still on a manual trigger and not yet tested with two real participants.

## Next steps
1. Sign in (Microsoft + GitHub) in the browser the user clicks through; confirm the Form, synced workbook
   and both flows still open and their connections are valid.
2. Finish the zero-touch flow: set class roles (coach/coachee classes), switch the manual trigger to a
   Recurrence trigger at the survey deadline, and run a group-chat test with two real volunteers.
3. Share the Form with the participating classes; after the deadline, run the flow (or the website /
   `run_match.py` fallback) and spot-check pairs.
4. After any matcher change: run tests, `python build_site.py`, commit, push so GitHub Pages updates.
5. Fix or delete the stale `test/run_test.mjs` (fake workbook lacks `getTables()`).

## Open questions / blockers
- Survey deadline date (needed for the Recurrence trigger).
- Which classes coach vs. are coached this semester.
- The flow owner is auto-added to every group chat; is that acceptable or should a service account own it?
- Should the older "send coach intro links" flow be kept as a fallback or retired?

## Commands that matter
- `node test/test_matcher.cjs; node test/test_forms_headers.cjs; node test/run_officescript.mjs`
- `python build_site.py` (rebuild `docs/index.html` after editing `site/`)
- `python run_match.py <responses.xlsx> --coach 2029 --coachee 2030`

## Machine / hardware notes
- Microsoft 365 parts are reachable only from the government laptop's school account.
- GitHub CLI on the gov laptop is installed via winget but not on PATH.
- Sign-in popups for Power Automate connections need the user's own click in the browser.

## Session pointers
- Background and decisions: AGENTS.md "Decisions".
