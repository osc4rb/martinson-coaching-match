# Martinson Honors Coaching Match: setup guide

**How it works**

```
Microsoft Form ──(Flow 1: on each submit)──► Excel "Responses" table (OneDrive)
                                                   │
                       (Flow 2: you click Run) ──► Office Script "MatchCadets"
                                                   │  scores every 29er↔30er pair,
                                                   │  finds best overall pairing
                                                   ▼
                                             "Matches" sheet ◄── you review / hand-swap
                                                   │
                       (Flow 3: you click Run) ──► intro email to each pair
```

Power Automate is good at moving data around but bad at "find the best pairing for
everyone", so that part runs as an **Office Script** (TypeScript inside Excel for the web)
that Power Automate calls.

| File | What it is |
|---|---|
| `form-questions.md` | Every question + exact answer options to put in the Form |
| `Martinson_Matching_TEMPLATE.xlsx` | Empty workbook with the `Responses` table (real one) |
| `Martinson_Matching_SAMPLE.xlsx` | Same thing with 10 fake 29ers + 14 fake 30ers (for testing) |
| `MatchCadets.ts` | The matching script you paste into Excel |
| `build_files.py`, `test/` | Regenerates the files above / tests the script locally |

---

## Step 1: Build the form (~20 min)

1. Go to **forms.office.com** → **New Form** → title it *Martinson Honors Coaching Match*.
2. Add the questions in `form-questions.md` **in that order, with option text typed exactly**.
   Matching compares option text, so `AI / machine learning` and `AI/ML` count as different answers.
   - Checkbox questions: turn on **Multiple answers**.
   - Where it says "pick up to 3", put that in the question text.
3. **... → Settings**: *Only people in my organization can respond*, *Record name*, *One response per person*.
4. Don't send it yet.

## Step 2: Put the workbook in OneDrive

1. Upload `Martinson_Matching_TEMPLATE.xlsx` to OneDrive (USAFA), e.g. folder
   `Martinson Coaching Match`, and rename it `Martinson_Matching.xlsx`.
2. Also upload `Martinson_Matching_SAMPLE.xlsx` for testing.

> Should this outlive you running it? Put it in a SharePoint/Teams site the program owns instead
> of your personal OneDrive, and pick that site as the Location in every flow step below.

## Step 3: Add the matching script

1. Open `Martinson_Matching_SAMPLE.xlsx` in **Excel for the web** (browser, not desktop).
2. **Automate** tab → **New Script**.
3. Delete the starter code, paste all of `MatchCadets.ts`, rename the script **MatchCadets**, **Save**.
4. Click **Run**. You should get a new **Matches** sheet: 14 coachees, every coach has 1–2, plus a
   "Shared Interests" column and runner-up coaches for each coachee.

The script is saved to *your* account, so it'll show up for any workbook, including the real one.

> **No Automate tab?** USAFA IT may have Office Scripts turned off. Ask the Help Desk to enable
> "Office Scripts" for you, or use the fallback at the bottom.

## Step 4: Flow 1 to save each response (automated)

**make.powerautomate.com → Create → Automated cloud flow**, name it *Martinson: save responses*,
trigger **When a new response is submitted** (Microsoft Forms).

1. **When a new response is submitted** → Form Id: your form.
2. **+ New step → Get response details** (Microsoft Forms) → Form Id: same form; Response Id:
   dynamic content *Response Id*.
3. **+ New step → Add a row into a table** (Excel Online (Business))
   - Location: OneDrive for Business · Document Library: OneDrive
   - File: `/Martinson Coaching Match/Martinson_Matching.xlsx` · Table: `Responses`
   - Each column now appears as a box. Fill each one with the matching question from the dynamic
     content of *Get response details* (Email → *Responders' Email* works too).
4. **Save**, submit one test response yourself, and check that a row appears in the workbook.
   Delete that test row afterward.

Checkbox answers arrive as `["A","B"]`. Leave them like that; the script handles it.

## Step 5: Flow 2 to generate matches (manual)

**Create → Instant cloud flow** → *Martinson: generate matches* → trigger **Manually trigger a flow**.

1. **+ New step → Run script** (Excel Online (Business))
   - Location / Document Library / File: the real `Martinson_Matching.xlsx`
   - Script: **MatchCadets**
2. *(Optional)* **Send an email (V2)** to yourself: "Matches ready, go review."

Run it after the form closes. **Running it again rebuilds the Matches sheet from scratch**, so any
hand edits are lost. Do your edits after the final run.

## Step 6: Review the matches

Open the workbook → **Matches** sheet.
- **Score** (0–100) is relative. Real data scores higher than the random sample, and what matters
  is how pairs compare to each other.
- To hand-swap a pair, edit the Coach / Coach Email cells. "Runner-up Coaches" suggests options.
- Rows saying `(unmatched - alternate)` are extra coaches (only if there are more 29ers than 30ers).
  Delete those rows or pair them with someone by hand.

## Step 7: Flow 3 to send intro emails (manual)

**Create → Instant cloud flow** → *Martinson: send match emails* → **Manually trigger a flow**.

1. **List rows present in a table** (Excel Online (Business)) → same file, Table: `MatchTable`.
2. **Apply to each** → output: *value*. Inside it:
   1. **Condition**: *Coachee Email* **is not equal to** (leave blank) → continue only in **If yes**.
   2. **If yes → Send an email (V2)** (Office 365 Outlook):
      - **To:** *Coach Email*`;`*Coachee Email*
      - **CC:** program director's email
      - **Subject:** `Martinson Honors Coaching Match: ` *Coach* ` & ` *Coachee*
      - **Body:**
        ```
        Hi [Coach] and [Coachee],

        You've been paired for this semester's Martinson Honors coaching program!
          Coach (C/O 2029): [Coach]
          Coachee (C/O 2030): [Coachee]

        From your survey answers, you have in common:
        [Shared Interests]

        [Coach], please reach out within the next week to set up your first meeting.
        ```
        (Put the dynamic content where each `[...]` is.)

**Test this before the real run.** Point it at the SAMPLE workbook and temporarily set **To:**
to your own email. The sample addresses are fake `@example.com` addresses. Then switch the file
back to the real workbook and put **To:** back.

---

## Tuning the matching

At the top of `MatchCadets.ts`, the `FIELDS` list sets how much each question counts (`weight`).
Right now coaching focus, major, and research count the most (3). Fun, grad school, and career
count 2, and everything else counts less. Set a weight to `0` to ignore a question. Edit the
script in Excel (Automate → MatchCadets), save, and rerun Flow 2.

If you add, rename, or remove a question, the Excel column header must match the `header` in
`FIELDS`. The script lists any missing columns by name if they don't match.

## Fallback if Office Scripts is blocked

Flow 1 and Flow 3 still work. For the matching step, download the workbook and send it to Claude:
the same script runs locally (`node test/run_test.mjs`, pointed at your data) and can write the
Matches sheet for you. Then upload it back and run Flow 3.
