# Martinson Coaching Match

Pairs Martinson Honors coaches with coachees for the semester from a survey. Any class can sign up, and
the program lead picks which classes coach and which get coached.

## Each semester

1. Share the form (Forms → Collect responses) with the classes taking part.
2. When it closes: Forms → Responses → Open results in Excel → File → Save a copy / Download.
3. Open the matcher site, drop the file in, set each class to Coach / Coachee / Skip, review.
4. Download Matches.xlsx and save it over `OneDrive - USAFA/Martinson Coaching Match/Matches.xlsx`.
5. Power Automate → My flows → *Martinson: send coach intro links* → Run.

## Matching

Each person's checkbox answers are compared with Jaccard overlap per question, weighted
(coaching focus, major, research = 3; grad school, career, fun = 2; future path, opportunities,
issues, research keywords = 1.5; minor = 1). The free-text research answer is compared by shared
keywords. The Hungarian algorithm then finds the pairing with the highest total score for the whole group.
Every coach gets at least one coachee before anyone gets two. Loads differ by at most
`ceil(coachees / coaches)`. Weights are adjustable on the page.

## Files

| Path | What |
|---|---|
| `site/` | Matcher page source (`index.html`, `matcher.js`, fake `sample-data.js`) |
| `docs/index.html` | Bundled single-file page served by GitHub Pages (`python build_site.py`) |
| `form-questions.md` | The form's questions and options (the live form matches this) |
| `MatchCadets.ts` | Same matching as an Excel Office Script (optional alternative) |
| `test/` | `node test/test_matcher.cjs`, `node test/test_forms_headers.cjs` |

The sample data is fake (`@example.com`). No real responses are stored in this repo.
