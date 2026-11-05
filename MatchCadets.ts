/**
 * Martinson Honors Program: coaching matcher (Office Script).
 *
 * Runs on the Excel workbook that Microsoft Forms keeps in sync with the coaching survey.
 * Called once a day by the Power Automate flow "Martinson: auto match + group chats".
 *
 * Everything the program lead changes lives on the "Settings" sheet (yellow cells):
 *   - Class pairings: which class coaches which (default 2027 -> 2029, 2028 -> 2030)
 *   - Question weights: how much each question counts (0 = ignore it)
 *   - Match on: the date the chats go out; Max coachees per coach (blank = even split)
 *   - Chats created at: filled in after a run; clear it to run again
 * The sheet is created with defaults the first time the script runs.
 *
 * runNow = "yes" ignores the date and the "already created" stamp (for a test run).
 * Returns one entry per coach: who goes in the group chat, its title and the intro message.
 * Also writes every pairing to a "Matches" sheet.
 */

interface Group {
  coachName: string;
  coachEmail: string;
  coachClass: string;
  coacheeCount: number;
  memberEmails: string;
  chatTitle: string;
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

const SETTINGS_SHEET = "Settings";
const MATCHES_SHEET = "Matches";

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

const DEFAULT_PAIRINGS: string[][] = [["2027", "2029"], ["2028", "2030"]];
const DEFAULT_MATCH_ON = "11/17/2026"; // pre-filled on the Settings sheet; the lead can change it there

function main(workbook: ExcelScript.Workbook, runNow?: string): Group[] {
  const settings = readSettings(workbook);
  const force = String(runNow === undefined || runNow === null ? "" : runNow).trim().toLowerCase() === "yes";
  console.log(`runNow=${JSON.stringify(runNow)} force=${force} matchOn=${settings.matchOn} today=${mountainToday()} sentAt=${settings.sentAt} pairings=${JSON.stringify(settings.pairings)}`);

  if (!force) {
    if (settings.matchOn === null) return [];
    if (mountainToday() < settings.matchOn) return [];
    if (settings.sentAt) return [];
  }

  const responses = findResponsesTable(workbook);
  if (!responses) throw new Error("Couldn't find the form responses table in this workbook.");
  const people = readPeople(responses);
  console.log(`responses=${people.length} classes=${people.map((p) => p.cls).join(",")}`);

  const groups: Group[] = [];
  const rows: (string | number)[][] = [];
  const claimed = new Set<string>();
  for (const pairing of settings.pairings) {
    const coaches = people.filter((p) => pairing.coachClasses.indexOf(p.cls) >= 0);
    const coachees = people.filter((p) => pairing.coacheeClasses.indexOf(p.cls) >= 0 && !claimed.has(p.email.toLowerCase()));
    if (coaches.length === 0 || coachees.length === 0) continue;
    coachees.forEach((c) => claimed.add(c.email.toLowerCase()));

    const assignment = assign(coaches, coachees, settings.weights, settings.maxPerCoach);
    coaches.forEach((co, j) => {
      const mine = coachees
        .map((ce, i) => ({ ce: ce, i: i }))
        .filter((x) => assignment.coachOf[x.i] === j);
      if (mine.length === 0) return;
      const lines: string[] = [];
      for (const x of mine) {
        const common = sharedItems(x.ce, co, settings.weights);
        const commonText = common.length > 0 ? common.slice(0, 4).join(", ") : "a lot to talk about";
        lines.push(`&bull; <b>${x.ce.name}</b> (Class of ${x.ce.cls}): in common with your coach: ${commonText}`);
        rows.push([co.name, co.cls, co.email, x.ce.name, x.ce.cls, x.ce.email, Math.round(assignment.score[x.i][j]), commonText]);
      }
      groups.push({
        coachName: co.name,
        coachEmail: co.email,
        coachClass: co.cls,
        coacheeCount: mine.length,
        memberEmails: [co.email].concat(mine.map((x) => x.ce.email)).join(";"),
        chatTitle: `MHP Coaching - ${co.name}`.replace(/:/g, ""), // Teams rejects ':' in chat titles
        intro:
          `<p>Welcome to your Martinson Honors coaching group!</p>` +
          `<p><b>Coach:</b> ${co.name} (Class of ${co.cls})</p>` +
          `<p><b>Coachee${mine.length === 1 ? "" : "s"}:</b><br>${lines.join("<br>")}</p>` +
          `<p>${firstName(co.name)}, kick things off by setting up a time for everyone to meet this week.</p>`,
      });
    });
  }

  writeMatches(workbook, rows);
  if (groups.length > 0) stampSent(workbook);
  return groups;
}

// ---------- settings sheet ----------

interface Pairing {
  coachClasses: string[];
  coacheeClasses: string[];
}

interface Settings {
  pairings: Pairing[];
  weights: Map<string, number>;
  matchOn: number | null;
  maxPerCoach: number;
  sentAt: string;
}

function readSettings(workbook: ExcelScript.Workbook): Settings {
  let ws = workbook.getWorksheet(SETTINGS_SHEET);
  if (!ws) ws = createSettings(workbook);

  // Pairings: rows with the same coach class are merged.
  const byCoach = new Map<string, string[]>();
  for (const r of ws.getRange("A5:B16").getValues()) {
    const coach = classOf(String(r[0]));
    const coachee = classOf(String(r[1]));
    if (!coach || !coachee) continue;
    const list = byCoach.get(coach) || [];
    if (list.indexOf(coachee) < 0) list.push(coachee);
    byCoach.set(coach, list);
  }
  const pairings: Pairing[] = [];
  byCoach.forEach((coachees, coach) => pairings.push({ coachClasses: [coach], coacheeClasses: coachees }));

  const weights = new Map<string, number>();
  for (const r of ws.getRange("D5:E16").getValues()) {
    const label = String(r[0]).trim();
    const w = Number(r[1]);
    if (label) weights.set(label, isFinite(w) && w >= 0 ? w : 0);
  }

  const v = ws.getRange("H5:H7").getValues();
  const max = Number(v[1][0]);
  return {
    pairings: pairings,
    weights: weights,
    matchOn: toDayNumber(v[0][0]),
    maxPerCoach: isFinite(max) && max >= 1 ? Math.floor(max) : 0,
    sentAt: String(v[2][0]).trim(),
  };
}

function createSettings(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  const ws = workbook.addWorksheet(SETTINGS_SHEET);
  ws.getRange("A1").setValue("Martinson coaching settings: change the yellow cells. The daily flow reads this sheet.");
  ws.getRange("A1").getFormat().getFont().setBold(true);

  ws.getRange("A4:B4").setValues([["Coach class", "Coaches this class"]]);
  const pairs: string[][] = DEFAULT_PAIRINGS.map((p) => [p[0], p[1]]);
  while (pairs.length < 12) pairs.push(["", ""]);
  ws.getRange("A5:B16").setValues(pairs);

  ws.getRange("D4:E4").setValues([["Question", "Weight (0 = ignore)"]]);
  const w: (string | number)[][] = FIELDS.map((f) => [f.label, f.weight]);
  while (w.length < 12) w.push(["", ""]);
  ws.getRange("D5:E16").setValues(w);

  ws.getRange("G4:H4").setValues([["Setting", "Value"]]);
  ws.getRange("G5:G7").setValues([["Match on (date)"], ["Max coachees per coach (blank = even split)"], ["Chats created at (clear to run again)"]]);

  for (const a of ["A4:B4", "D4:E4", "G4:H4"]) ws.getRange(a).getFormat().getFont().setBold(true);
  for (const a of ["A5:B16", "E5:E16", "H5:H6"]) ws.getRange(a).getFormat().getFill().setColor("#FFF2CC");
  ws.getRange("H5").setNumberFormatLocal("mm/dd/yyyy");
  ws.getRange("H5").setValue(DEFAULT_MATCH_ON);
  ws.getRange("A:H").getFormat().autofitColumns();
  return ws;
}

function stampSent(workbook: ExcelScript.Workbook) {
  const ws = workbook.getWorksheet(SETTINGS_SHEET);
  if (ws) ws.getRange("H7").setValue(new Date().toISOString().replace("T", " ").slice(0, 16) + " UTC");
}

// Excel day number for a date cell (number) or a typed date (text); null if empty/invalid.
function toDayNumber(v: string | number | boolean): number | null {
  if (typeof v === "number" && v > 0) return Math.floor(v);
  const s = String(v).trim();
  if (!s) return null;
  const t = Date.parse(s);
  if (isNaN(t)) return null;
  return Math.floor(t / 86400000) + 25569;
}

// Today's date in Mountain time as an Excel day number (the service runs in UTC).
function mountainToday(): number {
  return Math.floor((Date.now() - 7 * 3600000) / 86400000) + 25569;
}

// ---------- responses ----------

function findResponsesTable(workbook: ExcelScript.Workbook): ExcelScript.Table | undefined {
  // Plain loop: Office Scripts can't call workbook APIs inside array-method callbacks.
  const tables = workbook.getTables();
  for (let i = 0; i < tables.length; i++) {
    const name = tables[i].getWorksheet().getName();
    if (name !== SETTINGS_SHEET && name !== MATCHES_SHEET) return tables[i];
  }
  return undefined;
}

function readPeople(table: ExcelScript.Table): Person[] {
  const headers = table.getHeaderRowRange().getValues()[0].map((h) => String(h));
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const nameCol = findColumn(headers, ["Name", "Full name"]);
  const emailCol = findColumn(headers, ["Email", "Responders' Email", "USAFA email"]);
  const classCol = findColumn(headers, ["Class year", "Class Year"]);
  if (emailCol < 0 || classCol < 0) throw new Error("Couldn't find the Email or Class year column.");
  const fieldCols = FIELDS.map((f) => findColumn(headers, [f.header, f.question]));

  // Deduped by email: a later response replaces an earlier one.
  const byEmail = new Map<string, Person>();
  for (const row of rows) {
    const email = String(row[emailCol]).trim();
    if (!email || email.indexOf("@") < 0) continue;
    // Forms records the Teams display name ("Last, First M C3C USAF ..."); a name built from the
    // school email ("C29First.Last@...") reads better, so prefer that when the email has that shape.
    let name = nameCol >= 0 ? String(row[nameCol]).trim() : "";
    const fromEmail = nameFromEmail(email);
    if (/\./.test(email.split("@")[0]) || !name || name.toLowerCase() === "anonymous") name = fromEmail;
    byEmail.set(email.toLowerCase(), {
      name: name,
      email: email,
      cls: classOf(String(row[classCol])),
      answers: FIELDS.map((f, k) => {
        const raw = fieldCols[k] >= 0 ? String(row[fieldCols[k]]) : "";
        return f.isText ? keywords(raw) : parseList(raw);
      }),
    });
  }
  return Array.from(byEmail.values());
}

// ---------- matching ----------

interface Assignment {
  coachOf: number[];
  score: number[][];
}

// Every coach gets a coachee before anyone gets a second; nobody gets more than
// maxPerCoach (or an even split when blank). Coachees beyond total capacity stay unmatched.
function assign(coaches: Person[], coachees: Person[], weights: Map<string, number>, maxPerCoach: number): Assignment {
  const score: number[][] = coachees.map((ce) => coaches.map((co) => similarity(ce, co, weights)));
  const even = Math.ceil(coachees.length / coaches.length);
  const cap = Math.max(1, maxPerCoach > 0 ? maxPerCoach : even);
  const slotCoach: number[] = [];
  const slotRank: number[] = [];
  for (let k = 0; k < cap; k++) {
    for (let j = 0; j < coaches.length; j++) {
      slotCoach.push(j);
      slotRank.push(k);
    }
  }
  // Filling earlier slots earns a bonus, which keeps loads even.
  const bonus = (k: number) => 1000 * (cap - k);
  const nSlots = slotCoach.length;
  // Hungarian needs rows <= cols; pad with "unmatched" columns when coachees outnumber slots.
  const extra = Math.max(0, coachees.length - nSlots);
  const cost: number[][] = coachees.map((_, i) => {
    const row = slotCoach.map((j, s) => -(score[i][j] + bonus(slotRank[s])));
    for (let e = 0; e < extra; e++) row.push(0);
    return row;
  });
  const a = hungarian(cost);
  return { coachOf: a.map((s) => (s >= 0 && s < nSlots ? slotCoach[s] : -1)), score: score };
}

function similarity(a: Person, b: Person, weights: Map<string, number>): number {
  let total = 0;
  let wsum = 0;
  FIELDS.forEach((f, k) => {
    const w = weights.has(f.label) ? (weights.get(f.label) as number) : f.weight;
    if (w <= 0) return;
    wsum += w;
    total += w * jaccard(a.answers[k], b.answers[k]);
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

function sharedItems(a: Person, b: Person, weights: Map<string, number>): string[] {
  const out: string[] = [];
  FIELDS.forEach((f, k) => {
    if (f.isText) return;
    if (weights.has(f.label) && (weights.get(f.label) as number) <= 0) return;
    a.answers[k].forEach((display, key) => {
      if (b.answers[k].has(key) && !isNone(key) && out.indexOf(display) < 0) out.push(display);
    });
  });
  return out;
}

function isNone(key: string): boolean {
  return /^(none|n\/a|not sure|not interested|not planning|undecided|no|other)\b/.test(key);
}

// Min-cost assignment (rows <= cols). Returns the column chosen for each row.
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

// "C29First.Last@afacademy.af.edu" -> "First Last"
function nameFromEmail(email: string): string {
  const local = email.split("@")[0].replace(/^C\d{2}/i, "");
  return local.split(/[._]/).filter((x) => x.length > 0).map((x) => x.charAt(0).toUpperCase() + x.slice(1)).join(" ");
}

function firstName(name: string): string {
  return name.split(/\s+/)[0];
}

function writeMatches(workbook: ExcelScript.Workbook, rows: (string | number)[][]) {
  const old = workbook.getWorksheet(MATCHES_SHEET);
  if (old) old.delete();
  const ws = workbook.addWorksheet(MATCHES_SHEET);
  const values: (string | number)[][] = [["Coach", "Coach Class", "Coach Email", "Coachee", "Coachee Class", "Coachee Email", "Score", "In common"]];
  for (const r of rows) values.push(r);
  const range = ws.getRangeByIndexes(0, 0, values.length, values[0].length);
  range.setValues(values);
  range.getFormat().autofitColumns();
}
