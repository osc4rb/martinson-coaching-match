/**
 * Martinson Honors Program: coach/coachee matcher (Office Script).
 *
 * Runs on the Excel workbook that Microsoft Forms keeps in sync with the
 * "Martinson Honors Coaching Match" form (first table on the first sheet).
 * Called by the Power Automate flow "Martinson: match + group chats".
 *
 * coachClasses / coacheeClasses: comma lists like "2028,2029" and "2030".
 * Leave both empty to use the default: youngest class is coached, every
 * older class coaches.
 *
 * Returns one row per pair; also writes them to a "Matches" sheet.
 */

interface Match {
  coachName: string;
  coachEmail: string;
  coachClass: string;
  coacheeName: string;
  coacheeEmail: string;
  coacheeClass: string;
  score: number;
  shared: string;
  intro: string;
}

interface Field {
  header: string;
  question: string;
  label: string;
  weight: number;
  isText: boolean;
}

interface Person {
  name: string;
  email: string;
  cls: string;
  answers: Map<string, string>[];
}

// header = short column name; question = start of the Forms question text.
const FIELDS: Field[] = [
  { header: "Coaching Focus", question: "What do you want help with, or want your coach to focus on?", label: "Coaching focus", weight: 3, isText: false },
  { header: "Major", question: "Major (if undeclared: majors you're considering)", label: "Major", weight: 3, isText: false },
  { header: "Research Areas", question: "Research areas you're interested in", label: "Research", weight: 3, isText: false },
  { header: "Grad School Plans", question: "Grad school / fellowships you're interested in", label: "Grad school", weight: 2, isText: false },
  { header: "AF Career Interests", question: "Air Force / Space Force career fields you're interested in", label: "AF/SF career", weight: 2, isText: false },
  { header: "Fun", question: "What do you do for fun?", label: "For fun", weight: 2, isText: false },
  { header: "Future Path", question: "Long-term (after your commitment), where do you see yourself?", label: "Future path", weight: 1.5, isText: false },
  { header: "Opportunities", question: "Opportunities you want to pursue at USAFA", label: "Opportunities", weight: 1.5, isText: false },
  { header: "Research Interests", question: "In a sentence or two, what topics are you curious about?", label: "Research keywords", weight: 1.5, isText: true },
  { header: "Minor", question: "Minor(s) you have or are considering", label: "Minor", weight: 1, isText: false },
];

function main(workbook: ExcelScript.Workbook, coachClasses?: string, coacheeClasses?: string): Match[] {
  const tables = workbook.getTables();
  if (tables.length === 0) throw new Error("No table found. Open the form's results in Excel first.");
  const table = tables[0];
  const headers = table.getHeaderRowRange().getValues()[0].map((h) => String(h));
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();

  const nameCol = findColumn(headers, ["Name", "Full name"]);
  const emailCol = findColumn(headers, ["Email", "Responders' Email", "USAFA email"]);
  const classCol = findColumn(headers, ["Class year", "Class Year"]);
  if (emailCol < 0 || classCol < 0) throw new Error("Couldn't find the Email or Class year column.");
  const fieldCols = FIELDS.map((f) => findColumn(headers, [f.header, f.question]));

  // Everyone, deduped by email (a later response replaces an earlier one).
  const byEmail = new Map<string, Person>();
  for (const row of rows) {
    const email = String(row[emailCol]).trim();
    if (!email) continue;
    let name = nameCol >= 0 ? String(row[nameCol]).trim() : "";
    if (!name || name.toLowerCase() === "anonymous") name = nameFromEmail(email);
    const person: Person = {
      name: name,
      email: email,
      cls: classOf(String(row[classCol])),
      answers: FIELDS.map((f, k) => {
        const raw = fieldCols[k] >= 0 ? String(row[fieldCols[k]]) : "";
        return f.isText ? keywords(raw) : parseList(raw);
      }),
    };
    byEmail.set(email.toLowerCase(), person);
  }
  const people = Array.from(byEmail.values());

  let coachSet = splitList(coachClasses);
  let coacheeSet = splitList(coacheeClasses);
  if (coachSet.length === 0 && coacheeSet.length === 0) {
    const years = Array.from(new Set(people.map((p) => p.cls))).filter((c) => /^20\d\d$/.test(c)).sort();
    coacheeSet = years.slice(-1);
    coachSet = years.slice(0, -1);
  }
  const coaches = people.filter((p) => coachSet.indexOf(p.cls) >= 0);
  const coachees = people.filter((p) => coacheeSet.indexOf(p.cls) >= 0);
  if (coaches.length === 0 || coachees.length === 0) {
    throw new Error(`Need at least one coach (${coachSet.join(",")}) and one coachee (${coacheeSet.join(",")}). Found ${coaches.length} and ${coachees.length}.`);
  }

  // score[i][j]: similarity (0-100) of coachee i and coach j
  const score: number[][] = coachees.map((ce) => coaches.map((co) => similarity(ce, co)));

  // Each coach gets `cap` slots; filling a coach's first slot earns a big bonus
  // so every coach gets someone before anyone gets a second coachee.
  const cap = Math.max(1, Math.ceil(coachees.length / coaches.length));
  const slotCoach: number[] = [];
  const slotFirst: boolean[] = [];
  for (let k = 0; k < cap; k++) {
    for (let j = 0; j < coaches.length; j++) {
      slotCoach.push(j);
      slotFirst.push(k === 0);
    }
  }
  const cost: number[][] = coachees.map((_, i) => slotCoach.map((j, s) => -(score[i][j] + (slotFirst[s] ? 1000 : 0))));
  const assignment = hungarian(cost);

  const matches: Match[] = coachees.map((ce, i) => {
    const j = slotCoach[assignment[i]];
    const co = coaches[j];
    const common = sharedItems(ce, co);
    const sharedShort = common.length > 0 ? common.slice(0, 4).join(", ") : "a lot to talk about";
    return {
      coachName: co.name,
      coachEmail: co.email,
      coachClass: co.cls,
      coacheeName: ce.name,
      coacheeEmail: ce.email,
      coacheeClass: ce.cls,
      score: Math.round(score[i][j]),
      shared: sharedShort,
      intro: `Welcome to Martinson Honors coaching! ${co.name} (Class of ${co.cls}) is coaching ${ce.name} (Class of ${ce.cls}) this semester. You two have in common: ${sharedShort}. ${firstName(co.name)}, kick things off by setting up a time to meet this week.`,
    };
  });
  matches.sort((a, b) => a.coachName.localeCompare(b.coachName));

  writeMatches(workbook, matches);
  return matches;
}

// ---------- matching helpers ----------

function similarity(a: Person, b: Person): number {
  let total = 0;
  let wsum = 0;
  FIELDS.forEach((f, k) => {
    wsum += f.weight;
    total += f.weight * jaccard(a.answers[k], b.answers[k]);
  });
  return wsum === 0 ? 0 : (100 * total) / wsum;
}

function jaccard(a: Map<string, string>, b: Map<string, string>): number {
  if (a.size === 0 || b.size === 0) return 0;
  let inter = 0;
  a.forEach((_, key) => {
    if (b.has(key)) inter++;
  });
  return inter / (a.size + b.size - inter);
}

function sharedItems(a: Person, b: Person): string[] {
  const out: string[] = [];
  FIELDS.forEach((f, k) => {
    if (f.isText) return;
    a.answers[k].forEach((display, key) => {
      if (b.answers[k].has(key) && !isNone(key) && out.indexOf(display) < 0) out.push(display);
    });
  });
  return out;
}

function isNone(key: string): boolean {
  return /^(none|n\/a|not sure|not interested|not planning|undecided|no|other)\b/.test(key);
}

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

// ---------- parsing helpers ----------

function norm(s: string): string {
  return s.trim().toLowerCase().replace(/[’']/g, "'").replace(/\s+/g, " ");
}

function findColumn(headers: string[], candidates: string[]): number {
  const h = headers.map(norm);
  for (const c of candidates) {
    const i = h.indexOf(norm(c));
    if (i >= 0) return i;
  }
  for (const c of candidates) {
    const n = norm(c);
    if (n.length < 12) continue;
    const prefix = n.slice(0, 30);
    for (let i = 0; i < h.length; i++) {
      if (h[i].indexOf(prefix) === 0) return i;
    }
  }
  return -1;
}

// Forms writes checkbox answers as "A;B;" (or ["A","B"] from Power Automate).
function parseList(raw: string): Map<string, string> {
  const out = new Map<string, string>();
  const s = (raw || "").trim();
  if (!s || s === "null" || s === "undefined") return out;
  const items = s.replace(/^\[|\]$/g, "").split(/[;\n]|","/);
  for (const item of items) {
    const display = item.replace(/^"|"$/g, "").trim();
    if (display) out.set(display.toLowerCase(), display);
  }
  return out;
}

const STOPWORDS: string[] = ("the and for with that this are was were have has had not but you your our their they them " +
  "from into about how what which who whom why when where can could would should will also very really just like " +
  "interested interest interests research want study studying learn learning things thing stuff more most some any " +
  "all its it's i'm im lot lots").split(" ");

function keywords(raw: string): Map<string, string> {
  const out = new Map<string, string>();
  const words = (raw || "").toLowerCase().split(/[^a-z0-9+#]+/);
  for (let w of words) {
    if (w.length < 3 || STOPWORDS.indexOf(w) >= 0) continue;
    if (w.length > 4 && w.endsWith("s") && !w.endsWith("ss")) w = w.slice(0, -1);
    out.set(w, w);
  }
  return out;
}

function classOf(raw: string): string {
  const m = raw.match(/20\d\d/);
  return m ? m[0] : raw.trim();
}

function splitList(s: string): string[] {
  return (s || "").split(",").map((x) => x.trim()).filter((x) => x.length > 0);
}

// "C29Oscar.Basaula@afacademy.af.edu" -> "Oscar Basaula"
function nameFromEmail(email: string): string {
  const local = email.split("@")[0].replace(/^C\d{2}/i, "");
  return local.split(/[._]/).filter((x) => x.length > 0).map((x) => x.charAt(0).toUpperCase() + x.slice(1)).join(" ");
}

function firstName(name: string): string {
  return name.split(/\s+/)[0];
}

function writeMatches(workbook: ExcelScript.Workbook, matches: Match[]) {
  const old = workbook.getWorksheet("Matches");
  if (old) old.delete();
  const ws = workbook.addWorksheet("Matches");
  const values: (string | number)[][] = [["Coach", "Coach Class", "Coach Email", "Coachee", "Coachee Class", "Coachee Email", "Score", "In common"]];
  for (const m of matches) {
    values.push([m.coachName, m.coachClass, m.coachEmail, m.coacheeName, m.coacheeClass, m.coacheeEmail, m.score, m.shared]);
  }
  const range = ws.getRangeByIndexes(0, 0, values.length, values[0].length);
  range.setValues(values);
  range.getFormat().autofitColumns();
}
