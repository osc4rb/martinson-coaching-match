# Handoff - martinson-matching
Last updated: 2026-11-05 · Machine: C29-2MQ517152H (gov laptop) · Agent: Claude Code · Branch: main

Public repo: keep this file free of personal data, internal links, flow IDs and sign-in details.

## Objective
Zero-touch coaching match: the program lead only posts the form link in their Teams channel. On the
match date (Tue 2026-11-17, flow fires daily ~1:30 PM Mountain) the flow matches 2027->2029 and
2028->2030 (several coachees per coach) and creates one Teams group chat per coach with an intro.

## Done (all verified live 2026-11-05)
- `MatchCadets.ts`: pairings, balanced multi-coachee matching, Settings sheet, one group per coach,
  intro without the "meet this week" line. Before the date each run refreshes the "Matches" sheet
  with the would-be pairs (plain table, no status text; nothing sent); on the date chats are created
  once and stamped.
- Form questions shortened (live form verified via the forms API); script reads old + new headers.
- Flow (scheduled, On, real Forms-synced workbook, runNow blank): Run script -> Apply to each ->
  Create a chat (members = memberEmails minus the flow owner, via expression) -> Post message.
- Live tests: 1 coach + 1 coachee chat; 1 coach + 2 coachees with the flow owner as a coachee
  (needed the owner-exclusion fix - Teams rejected duplicate members). Both created + intro posted.
- Tried auto-removing the flow owner from chats they aren't in: impossible with the standard Teams
  connector (see AGENTS.md Decisions). Owner stays in every chat.
- Real-workbook run before the date: returned nothing, wrote the preview. It currently has 1 response
  (class 2027) - likely an old test submission; the lead should confirm/delete it.
- Old Martinson flows are Off. Tests: 22 Office Script checks + matcher tests pass.

## Next steps
1. Lead posts the form link; checks the workbook's Matches tab any time for the preview.
2. Lead removes any test responses from the real workbook/form before 11/17.
3. On/after 2026-11-17 13:30 MT check the run history; Settings!H7 gets stamped.
4. Optional cleanup: test workbooks "MHP Match TEST*.xlsx" in OneDrive; test chats in Teams.

## Open questions / blockers
- Office Script `runNow` parameter arrives undefined at runtime. For a forced test, temporarily
  install a copy of the script with `force = true`, run, then reinstall the real one.
- Edge case: if the flow owner is a coach with exactly one coachee, the chat would be 1:1 (untested).

## Commands that matter
- `node test/test_matcher.cjs; node test/test_forms_headers.cjs; node test/run_officescript.mjs`
- `python build_site.py` (rebuild `docs/index.html` after editing `site/`)
- Reinstall script after editing: `.osts` body = `MatchCadets.ts` verbatim (keep parameterInfo/apiInfo).

## Machine / hardware notes
- Microsoft 365 parts are only reachable from the gov laptop's school account (built-in browser pane).
- Commit as the owner's GitHub noreply address (`git config user.email` is already set in this repo).

## Session pointers
- Claude Code desktop session d437bb9e on the gov laptop.
