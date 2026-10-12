/**
 * Martinson Honors Program: Coach/Coachee Matcher (Office Script)
 *
 * Reads the "Responses" table, scores how similar every coach is to every
 * coachee, then finds the best overall pairing for the whole group (Hungarian
 * algorithm), not just greedy first-come-first-served.
 *
 * Writes a "Matches" sheet (table "MatchTable") and returns the matches so
 * Power Automate can use them.
 *
 * Unequal class sizes:
 *   - More coachees than coaches: every coach gets at least 1, and nobody
 *     gets more than ceil(coachees / coaches).
 *   - More coaches than coachees: extra coaches are listed as unmatched
 *     alternates at the bottom of the sheet.
 */

// ===================== CONFIG: edit to taste =====================
const RESPONSES_TABLE = "Responses";
const MATCHES_SHEET = "Matches";
const MATCHES_TABLE = "MatchTable";

const COACH_CLASS = "2029";   // anything in "Class Year" containing this is a coach
const COACHEE_CLASS = "2030"; // anything containing this is a coachee

const NAME_COL = "Name";
const EMAIL_COL = "Email";
const CLASS_COL = "Class Year";

// kind "list" = checkbox/choice answers (exact option overlap)
// kind "text" = free text (shared keyword overlap, weight it lower)
// Higher weight = matters more. Set weight to 0 to ignore a question.
const FIELDS: FieldConfig[] = [
  { header: "Coaching Focus", label: "Coaching focus", weight: 3, kind: "list" },
  { header: "Major", label: "Major", weight: 3, kind: "list" },
  { header: "Research Areas", label: "Research", weight: 3, kind: "list" },
  { header: "Grad School Plans", label: "Grad school", weight: 2, kind: "list" },
  { header: "AF Career Interests", label: "AF/SF career", weight: 2, kind: "list" },
  { header: "Fun", label: "For fun", weight: 2, kind: "list" },
  { header: "Future Path", label: "Future path", weight: 1.5, kind: "list" },
  { header: "Opportunities", label: "Opportunities", weight: 1.5, kind: "list" },
  { header: "Research Interests", label: "Research keywords", weight: 1.5, kind: "text" },
  { header: "Minor", label: "Minor", weight: 1, kind: "list" },
];
// =================================================================

interface FieldConfig {
  header: string;
  label: string;
  weight: number;
  kind: string;
}

interface Person {
  name: string;
  email: string;
  // per field: lowercase key -> display text
  answers: Map<string, string>[];
}

interface Match {
  coachName: string;
  coachEmail: string;
  coacheeName: string;
  coacheeEmail: string;
  score: number;
  shared: string;
}

function main(workbook: ExcelScript.Workbook): Match[] {
  const table = workbook.getTable(RESPONSES_TABLE);
  if (!table) {
    throw new Error(`No table named "${RESPONSES_TABLE}" in this workbook.`);
  }

  const headers = table.getHeaderRowRange().getValues()[0].map((h) => String(h).trim().toLowerCase());
  const col = (name: string): number => headers.indexOf(name.trim().toLowerCase());

  const missing: string[] = [NAME_COL, EMAIL_COL, CLASS_COL]
    .concat(FIELDS.filter((f) => f.weight > 0).map((f) => f.header))
    .filter((h) => col(h) < 0);
  if (missing.length > 0) {
    throw new Error(`Missing column(s) in "${RESPONSES_TABLE}": ${missing.join(", ")}`);
  }

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Dedupe by email: if someone submitted twice, keep their latest row.
  const coachesByEmail = new Map<string, Person>();
  const coacheesByEmail = new Map<string, Person>();
  for (const row of rows) {
    const name = String(row[col(NAME_COL)]).trim();
    const email = String(row[col(EMAIL_COL)]).trim();
    const classYear = String(row[col(CLASS_COL)]);
    if (!name && !email) continue;

    const person: Person = {
      name: name,
      email: email,
      answers: FIELDS.map((f) => {
        const raw = String(row[col(f.header)]);
        return f.kind === "text" ? keywords(raw) : parseList(raw);
      }),
    };
    const key = (email || name).toLowerCase();
    if (classYear.includes(COACH_CLASS)) coachesByEmail.set(key, person);
    else if (classYear.includes(COACHEE_CLASS)) coacheesByEmail.set(key, person);
  }
  const coaches = Array.from(coachesByEmail.values());
  const coachees = Array.from(coacheesByEmail.values());
  if (coaches.length === 0 || coachees.length === 0) {
    throw new Error(`Need at least one ${COACH_CLASS} and one ${COACHEE_CLASS} response (found ${coaches.length} and ${coachees.length}).`);
  }

  // score[i][j] = similarity (0-100) of coachee i and coach j
  const score: number[][] = coachees.map((ce) => coaches.map((co) => similarity(ce, co)));

  // Each coach gets `cap` slots. Filling a coach's first slot earns a big
  // bonus, so every coach gets someone before anyone gets a second coachee.
  const cap = Math.max(1, Math.ceil(coachees.length / coaches.length));
  const FIRST_SLOT_BONUS = 1000;
  const slotCoach: number[] = [];
  const slotIsFirst: boolean[] = [];
  for (let k = 0; k < cap; k++) {
    for (let j = 0; j < coaches.length; j++) {
      slotCoach.push(j);
      slotIsFirst.push(k === 0);
    }
  }
  const cost: number[][] = coachees.map((_, i) =>
    slotCoach.map((j, s) => -(score[i][j] + (slotIsFirst[s] ? FIRST_SLOT_BONUS : 0)))
  );
  const assignment = hungarian(cost);

  const matches: Match[] = [];
  const coachLoad: number[] = coaches.map(() => 0);
  const values: (string | number)[][] = [];
  coachees.forEach((ce, i) => {
    const j = slotCoach[assignment[i]];
    const co = coaches[j];
    coachLoad[j]++;
    const s = Math.round(score[i][j]);
    const shared = sharedText(ce, co);
    matches.push({
      coachName: co.name,
      coachEmail: co.email,
      coacheeName: ce.name,
      coacheeEmail: ce.email,
      score: s,
      shared: shared,
    });

    // Runner-up coaches, so the director can hand-swap if they want.
    const alts = coaches
      .map((c, jj) => ({ name: c.name, s: score[i][jj], jj: jj }))
      .filter((a) => a.jj !== j)
      .sort((a, b) => b.s - a.s)
      .slice(0, 2)
      .map((a) => `${a.name} (${Math.round(a.s)})`)
      .join("; ");

    values.push([co.name, co.email, ce.name, ce.email, s, shared, alts]);
  });

  values.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  coaches.forEach((co, j) => {
    if (coachLoad[j] === 0) {
      values.push([co.name, co.email, "(unmatched - alternate)", "", 0, "", ""]);
    }
  });

  writeMatchesSheet(workbook, values);
  return matches;
}

// ---------- scoring ----------

function similarity(a: Person, b: Person): number {
  let total = 0;
  let weightSum = 0;
  FIELDS.forEach((f, k) => {
    if (f.weight <= 0) return;
    weightSum += f.weight;
    total += f.weight * jaccard(a.answers[k], b.answers[k]);
  });
  return weightSum === 0 ? 0 : (100 * total) / weightSum;
}

function jaccard(a: Map<string, string>, b: Map<string, string>): number {
  if (a.size === 0 || b.size === 0) return 0;
  let inter = 0;
  a.forEach((_, key) => {
    if (b.has(key)) inter++;
  });
  return inter / (a.size + b.size - inter);
}

function sharedText(a: Person, b: Person): string {
  const parts: string[] = [];
  FIELDS.forEach((f, k) => {
    if (f.weight <= 0) return;
    const common: string[] = [];
    a.answers[k].forEach((display, key) => {
      if (b.answers[k].has(key) && !isNoneAnswer(key)) common.push(display);
    });
    if (common.length > 0) parts.push(`${f.label}: ${common.slice(0, 4).join(", ")}`);
  });
  return parts.join(" | ");
}

function isNoneAnswer(key: string): boolean {
  return /^(none|n\/a|not sure|undecided|no|other)\b/.test(key);
}

// ---------- parsing ----------

// Handles Forms/Power Automate checkbox output: ["A","B"]  and Excel-sync
// output: A;B;  (and plain single answers).
function parseList(raw: string): Map<string, string> {
  const out = new Map<string, string>();
  let s = (raw || "").trim();
  if (!s || s === "null" || s === "undefined") return out;
  let items: string[] = [];
  if (s.startsWith("[")) {
    try {
      items = (JSON.parse(s) as string[]).map((x) => String(x));
    } catch (e) {
      items = s.replace(/^\[|\]$/g, "").split(/","|;/);
    }
  } else {
    items = s.split(/[;\n]/);
  }
  for (const item of items) {
    const display = item.replace(/^"|"$/g, "").trim();
    if (display) out.set(display.toLowerCase(), display);
  }
  return out;
}

const STOPWORDS = new Set<string>(
  ("the and for with that this are was were have has had not but you your our their they them " +
    "from into about how what which who whom why when where can could would should will also " +
    "very really just like interested interest interests research want study studying learn " +
    "learning things thing stuff more most some any all its it's i'm im lot lots").split(" ")
);

function keywords(raw: string): Map<string, string> {
  const out = new Map<string, string>();
  const words = (raw || "").toLowerCase().split(/[^a-z0-9+#]+/);
  for (let w of words) {
    if (w.length < 3 || STOPWORDS.has(w)) continue;
    if (w.length > 4 && w.endsWith("s") && !w.endsWith("ss")) w = w.slice(0, -1);
    out.set(w, w);
  }
  return out;
}

// ---------- assignment (Hungarian algorithm, min-cost, rows <= cols) ----------

function hungarian(cost: number[][]): number[] {
  const n = cost.length;
  const m = cost[0].length;
  const INF = 1e18;
  const u: number[] = new Array(n + 1).fill(0);
  const v: number[] = new Array(m + 1).fill(0);
  const p: number[] = new Array(m + 1).fill(0);
  const way: number[] = new Array(m + 1).fill(0);
  for (let i = 1; i <= n; i++) {
    p[0] = i;
    let j0 = 0;
    const minv: number[] = new Array(m + 1).fill(INF);
    const used: boolean[] = new Array(m + 1).fill(false);
    do {
      used[j0] = true;
      const i0 = p[j0];
      let delta = INF;
      let j1 = 0;
      for (let j = 1; j <= m; j++) {
        if (used[j]) continue;
        const cur = cost[i0 - 1][j - 1] - u[i0] - v[j];
        if (cur < minv[j]) {
          minv[j] = cur;
          way[j] = j0;
        }
        if (minv[j] < delta) {
          delta = minv[j];
          j1 = j;
        }
      }
      for (let j = 0; j <= m; j++) {
        if (used[j]) {
          u[p[j]] += delta;
          v[j] -= delta;
        } else {
          minv[j] -= delta;
        }
      }
      j0 = j1;
    } while (p[j0] !== 0);
    do {
      const j1 = way[j0];
      p[j0] = p[j1];
      j0 = j1;
    } while (j0 !== 0);
  }
  const ans: number[] = new Array(n).fill(-1);
  for (let j = 1; j <= m; j++) {
    if (p[j] !== 0) ans[p[j] - 1] = j - 1;
  }
  return ans;
}

// ---------- output ----------

function writeMatchesSheet(workbook: ExcelScript.Workbook, values: (string | number)[][]) {
  const old = workbook.getWorksheet(MATCHES_SHEET);
  if (old) old.delete();
  const ws = workbook.addWorksheet(MATCHES_SHEET);

  const header: (string | number)[] = [
    "Coach", "Coach Email", "Coachee", "Coachee Email", "Score", "Shared Interests", "Runner-up Coaches",
  ];
  const all = [header].concat(values);
  const range = ws.getRangeByIndexes(0, 0, all.length, header.length);
  range.setValues(all);

  const t = ws.addTable(range, true);
  t.setName(MATCHES_TABLE);
  range.getFormat().autofitColumns();
  ws.getRange("F:F").getFormat().setColumnWidth(420);
  ws.getRange("F:F").getFormat().setWrapText(true);
  ws.activate();
}
