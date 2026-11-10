# Handoff - martinson-matching
Last updated: 2026-11-10 · Machine: C29-2MQ517152H (gov laptop) · Agent: Claude Code · Branch: main

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
- Form renamed "MHP Coaching Matcher", no description, short question titles; Minor lists languages
  (Arabic..Spanish); grad question is "Education goal" (Master's/PhD/Med/Law/Not sure/Not planning).
  Script reads old + new headers. 2030s now have Teams accounts (as of 11/09).
- Flow (scheduled, On, real Forms-synced workbook, runNow blank): Run script -> Apply to each ->
  Create a chat (members = memberEmails minus the flow owner, via expression) -> Post message.
- Live tests: 1 coach + 1 coachee chat; 1 coach + 2 coachees with the flow owner as a coachee
  (needed the owner-exclusion fix - Teams rejected duplicate members). Both created + intro posted.
- Tried auto-removing the flow owner from chats they aren't in: impossible with the standard Teams
  connector (see AGENTS.md Decisions). Owner stays in every chat.
- Real-workbook run before the date: returned nothing, wrote the preview. Its first response (class 2027, from Oct) is a REAL coach - never delete it.
- Old Martinson flows are Off. Tests: 22 Office Script checks + matcher tests pass.

## Next steps
0. STATUS 2026-11-10: form link sent to all 4 classes (forms.cloud.microsoft/Pages/ResponsePage.aspx?id=<same form id>);
   responses are coming in. Do not delete any responses.
1. TODO NOW (owner asked 2026-11-10): review window before chats go out on 11/17.
   a. MatchCadets.ts: on the match date, `due` only if Mountain time >= 13:25 (before that = preview only).
      Today: due = force || (matchOn !== null && mountainToday() >= matchOn), so ANY run on 11/17
      (incl. a manual Test) sends. Add mountainMinutes() (UTC-7 in Nov), keep days after the date due.
      Add tests (match day 13:05 -> no groups, Matches written; 13:30 -> groups). Reinstall .osts
      (body = MatchCadets.ts verbatim; keep parameterInfo/apiInfo), commit + push (noreply author).
   b. Flow "Martinson: auto match + group chats" Recurrence: run at 13:05 AND 13:30 Mountain daily
      (classic designer ?v3=false; "At these minutes" 5,30). 13:05 refreshes Matches, 13:30 sends.
   c. Forms: set end date 11/17/2026 1:00 PM (Settings -> End date) so 13:05 preview == final pairs.
   d. Verify with a flow Test run before 11/17 (sends nothing) and read the Matches tab.
   Browser pane needs the owner to sign in to Microsoft again first.
2. On 11/17 at ~13:05 the owner reviews pairs (Matches tab, or ask the agent to read the synced
   workbook in OneDrive root read-only); 13:30 run sends.
3. Optional: scale test with a few hundred fake people (speed + even splits); manual pair override feature.

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
