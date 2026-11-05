# Handoff - martinson-matching
Last updated: 2026-11-05 · Machine: C29-2MQ517152H (gov laptop) · Agent: Claude Code · Branch: main

Public repo: keep this file free of personal data, internal links, flow IDs and sign-in details.

## Objective
Zero-touch coaching match: the program lead only posts the form link in their Teams channel. On the
match date (Tue 2026-11-17, flow fires daily ~1:30 PM Mountain) the flow matches 2027->2029 and
2028->2030 (several coachees per coach) and creates one Teams group chat per coach with an intro.

## Done (all verified 2026-11-05)
- `MatchCadets.ts`: class pairings, balanced multi-coachee Hungarian matching, Settings sheet
  (pairings A5:B16, weights D5:E16, Match on H5 = 11/17/2026, max per coach H6, "Chats created at" H7),
  names cleaned from school email, one group per coach. Installed in OneDrive as `MatchCadets.osts`.
- Two cloud-only bugs found by the live test and fixed: (1) no workbook API calls inside array
  callbacks (runtime error `filterAsync is not a function`); (2) Teams rejects `:` in chat titles,
  so titles are "MHP Coaching - <coach>".
- Live end-to-end test on a TEST workbook (2 volunteers): group chat created with coach + coachee
  (+ flow owner), intro posted by Workflows. Post step needed ~5 min of automatic retries while the
  new chat came online; it succeeded.
- Production flow (scheduled, On) now points at the real Forms-synced workbook; runNow blank. A test
  run on the real workbook succeeded and correctly did nothing (before match date).
- The two older Martinson flows are turned Off.
- `test/run_officescript.mjs`: 16 checks, all pass; other tests pass.
- GitHub: only the owner shows as contributor (see AGENTS.md Conventions).

## In progress
- Nothing mid-flight.

## Next steps
1. Lead posts the form link in the channel; responses flow into the synced workbook automatically.
2. On/after 2026-11-17 13:30 MT check the flow's run history once; the Settings sheet's
   "Chats created at" cell gets stamped when chats are created.
3. To rerun after a change: clear Settings!H7 in the real workbook.
4. Optional: owner deletes the `*-old-delete-me` private repos on GitHub.

## Open questions / blockers
- Office Script `runNow` parameter arrives undefined at runtime (hand-written parameterInfo in
  .osts). Not needed in production: the Settings date controls everything. For a forced test,
  temporarily install a copy of the script with `force = true`, run, then reinstall the real one.

## Commands that matter
- `node test/test_matcher.cjs; node test/test_forms_headers.cjs; node test/run_officescript.mjs`
- `python build_site.py` (rebuild `docs/index.html` after editing `site/`)
- `python run_match.py <responses.xlsx> --coach 2029 --coachee 2030` (offline fallback)
- Reinstall script after editing: `.osts` body = `MatchCadets.ts` verbatim (keep parameterInfo/apiInfo).

## Machine / hardware notes
- Microsoft 365 parts (Form, workbook, script, flows, Teams) are only reachable from the gov laptop's
  school account. Built-in browser pane in the Claude desktop app is signed in there.
- Commit as the owner's GitHub noreply address (`git config user.email` is already set in this repo).

## Session pointers
- Claude Code desktop session d437bb9e on the gov laptop.
