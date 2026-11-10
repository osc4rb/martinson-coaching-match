// Runs MatchCadets.ts locally against a fake ExcelScript workbook shaped like the Forms
// "Open results in Excel" workbook (real question headers, "A;B;" checkbox values).
// Usage: node test/run_officescript.mjs
import { readFileSync, writeFileSync, unlinkSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const script = readFileSync(join(here, "..", "MatchCadets.ts"), "utf8");
const harness = `
const s = JSON.parse(require("node:fs").readFileSync(${JSON.stringify(join(here, "sample.json"))}, "utf8"));
const map: Record<string, string> = {
  "Coaching Focus": "What do you want help with, or want your coach to focus on? (pick up to 5)",
  "Major": "Major (if undeclared: majors you're considering) (pick up to 3)",
  "Minor": "Minor(s) you have or are considering",
  "Grad School Plans": "Grad school / fellowships you're interested in",
  "AF Career Interests": "Air Force / Space Force career fields you're interested in (pick up to 4)",
  "Future Path": "Long-term (after your commitment), where do you see yourself?",
  "Opportunities": "Opportunities you want to pursue at USAFA",
  "Research Areas": "Research areas you're interested in (pick up to 4)",
  "Research Interests": "In a sentence or two, what topics are you curious about?",
  "Fun": "What do you do for fun?",
  "Class Year": "Class year",
};
const headers = ["Id", "Start time", "Completion time", "Email", "Name", "Last modified time"];
const keep = s.headers.map((h: string, i: number) => [h, i]).filter(([h]: any) => map[h]);
for (const [h] of keep as any) headers.push(map[h]);
const toForms = (v: string) => v.startsWith("[") ? JSON.parse(v).join(";") + ";" : v;
// 28 fake people: 3 x 2027, 3 x 2028, 11 x 2029, 11 x 2030 (few coaches, many coachees).
const cls = (k: number) => k < 3 ? "2027" : k < 6 ? "2028" : k < 17 ? "2029" : "2030";
const classIdx = keep.findIndex(([h]: any) => h === "Class Year");
const rows = s.rows.map((r: any[], k: number) => {
  const vals = (keep as any).map(([h, i]: any) => toForms(String(r[i])));
  vals[classIdx] = cls(k);
  return [String(k + 1), "", "", k === 0 ? "C27Jane.Doe@example.com" : "c" + k + "@example.com", k === 0 ? "Doe, Jane Q C1C USAF USAFA CW/CS01" : (k % 4 === 0 ? "" : "Person " + k), "", ...vals];
});

// --- tiny cell store so the script can read/write sheet ranges ---
const colN = (c: string) => c.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function cells(addr: string): [number, number][][] {
  const m = addr.match(/^([A-Z]+)(\\d+)?(?::([A-Z]+)(\\d+)?)?$/)!;
  const c1 = colN(m[1]), r1 = Number(m[2] || 1), c2 = colN(m[3] || m[1]), r2 = Number(m[4] || m[2] || 1);
  const out: [number, number][][] = [];
  for (let r = r1; r <= r2; r++) { const row: [number, number][] = []; for (let c = c1; c <= c2; c++) row.push([r, c]); out.push(row); }
  return out;
}
function makeSheet(name: string) {
  const store = new Map<string, any>();
  const fmt = () => ({ autofitColumns() {}, getFont: () => ({ setBold() {} }), getFill: () => ({ setColor() {} }) });
  const range = (addr: string) => ({
    getValues: () => cells(addr).map((row) => row.map(([r, c]) => store.has(r + "," + c) ? store.get(r + "," + c) : "")),
    setValues: (v: any[][]) => cells(addr).forEach((row, i) => row.forEach(([r, c], j) => store.set(r + "," + c, v[i][j]))),
    setValue: (v: any) => { const [[[r, c]]] = cells(addr) as any; store.set(r + "," + c, v); },
    setNumberFormatLocal() {}, getFormat: fmt,
  });
  return { name, store, getName: () => name, getRange: range, delete() { sheets.delete(name); },
    getRangeByIndexes: (r0: number, c0: number, nr: number, nc: number) => ({
      setValues: (v: any[][]) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) store.set((r0 + i + 1) + "," + (c0 + j + 1), v[i][j]); },
      getFormat: fmt }) };
}
const sheets = new Map<string, any>();
const formSheet = makeSheet("Form1"); sheets.set("Form1", formSheet);
const wb: any = {
  getTables: () => [{ getWorksheet: () => formSheet, getHeaderRowRange: () => ({ getValues: () => [headers] }), getRangeBetweenHeaderAndTotal: () => ({ getValues: () => rows }) }],
  getWorksheet: (n: string) => sheets.get(n),
  addWorksheet: (n: string) => { const w = makeSheet(n); sheets.set(n, w); return w; },
};
const settings = () => sheets.get("Settings");
const set = (addr: string, v: any) => settings().getRange(addr).setValue(v);
let fails = 0;
const check = (ok: boolean, msg: string) => { console.log((ok ? "PASS " : "FAIL ") + msg); if (!ok) fails++; };

let g = main(wb, "");
check(!!settings() && String(settings().getRange("H5").getValues()[0][0]) === "11/17/2026", "first run creates Settings with match date 11/17/2026 pre-filled");
check(String(settings().getRange("A1").getValues()[0][0]) === "", "Settings sheet has no extra instruction text");
set("H5", "");
check(main(wb, "").length === 0, "blank match date: nothing happens");
check(settings().getRange("A5:B6").getValues().map((r: any) => r.join(">")).join(",") === "2027>2029,2028>2030", "default pairings 2027>2029, 2028>2030");

set("H5", "2999-01-01");
check(main(wb, "").length === 0, "future match date: no chats returned");
const mws = () => sheets.get("Matches");
check(!!mws() && String(mws().getRange("A1").getValues()[0][0]) === "Coach" && String(mws().getRange("A2").getValues()[0][0]).length > 0, "before the date: Matches sheet lists the would-be pairs (no extra text)");
check(String(settings().getRange("H7").getValues()[0][0]) === "", "preview does not stamp 'Chats created at'");

set("H5", "2020-01-01");
check(main(wb, "").length === 0, "past date but not approved (H8 blank): nothing sent");
set("H8", "yes");
g = main(wb, "");
const loads = g.map((x: any) => x.coachClass + ":" + x.coacheeCount);
check(g.length === 6, "past match date: one group chat per coach (6 coaches) -> " + loads.join(" "));
check(g.filter((x: any) => x.coachClass === "2027").reduce((n: number, x: any) => n + x.coacheeCount, 0) === 11, "all 11 2029s go to the 3 2027 coaches");
check(g.filter((x: any) => x.coachClass === "2028").reduce((n: number, x: any) => n + x.coacheeCount, 0) === 11, "all 11 2030s go to the 3 2028 coaches");
check(Math.max(...g.map((x: any) => x.coacheeCount)) - Math.min(...g.map((x: any) => x.coacheeCount)) <= 1, "loads are even (3-4 each)");
check(g.every((x: any) => x.memberEmails.split(";").length === x.coacheeCount + 1), "chat members = coach + their coachees");
check(g.every((x: any) => x.chatTitle.indexOf(":") < 0), "chat titles have no ':' (Teams rejects it)");
check(String(settings().getRange("H7").getValues()[0][0]).length > 0, "run stamps 'Chats created at'");
check(main(wb, "").length === 0, "second daily run does nothing (no duplicate chats)");
check(String(mws().getRange("A1").getValues()[0][0]) === "Coach", "after chats exist, the Matches record is left alone");
check(main(wb, "yes").length === 6, "runNow=yes ignores the date/stamp (test runs)");

set("H6", 2);
g = main(wb, "yes");
check(Math.max(...g.map((x: any) => x.coacheeCount)) === 2, "max coachees per coach = 2 is respected");
set("H6", "");

settings().getRange("E5:E16").setValues([[0],[0],[0],[0],[0],[0],[0],[0],[0],[5],[""],[""]]);
g = main(wb, "yes");
check(g.length === 6 && g.every((x: any) => !/Coaching focus|Major:/.test(x.intro)), "weights from Settings are used (0 = ignored)");
check(g.some((x: any) => x.coachName === "Jane Doe"), "Teams-style display name becomes 'Jane Doe' from the school email");
// Match day: nothing is sent until the lead approves (Settings!H8 = YES); then the reviewed
// (possibly edited) Matches sheet is what gets sent.
const realNow = Date.now;
set("H5", "11/17/2026"); set("H7", ""); set("H8", "");
Date.now = () => Date.UTC(2026, 10, 17, 20, 5);   // 13:05 Mountain on the match day
check(main(wb, "").length === 0 && String(mws().getRange("A2").getValues()[0][0]).length > 0 && String(settings().getRange("H7").getValues()[0][0]) === "", "match day, not approved: pairs shown on Matches, nothing sent");
check(String(settings().getRange("G8").getValues()[0][0]).indexOf("YES") > 0, "Settings has the 'Send chats? (type YES)' cell");
// Lead edit in the sheet: move the coachee in row 2 to the coach of the last row; delete row 3.
const lastRow = (() => { let r = 2; while (String(mws().getRange("C" + (r + 1)).getValues()[0][0])) r++; return r; })();
const movedCoachee = String(mws().getRange("F2").getValues()[0][0]);
const newCoach = String(mws().getRange("C" + lastRow).getValues()[0][0]);
const droppedCoachee = String(mws().getRange("F3").getValues()[0][0]);
mws().getRange("C2").setValue(newCoach);
mws().getRange("C3").setValue(""); mws().getRange("F3").setValue("");
set("H8", "yes");
Date.now = () => Date.UTC(2026, 10, 17, 20, 30);
let sent = main(wb, "");
const total = (gs: any[]) => gs.reduce((n: number, x: any) => n + x.coacheeCount, 0);
const holder = sent.find((x: any) => x.memberEmails.split(";").indexOf(movedCoachee) > 0);
check(sent.length > 0 && !!holder && holder.coachEmail === newCoach, "approved: send uses the edited Matches sheet (coachee moved to the new coach)");
check(sent.every((x: any) => x.memberEmails.split(";").indexOf(droppedCoachee) < 0) && total(sent) === 21, "approved: a deleted Matches row is left out (21 of 22 coachees sent)");
check(main(wb, "").length === 0, "after sending: no duplicates");
// Day-of edits via OVERRIDES (set in the installed copy): drop one coachee, move another.
set("H7", "");
const someCoachee = String(mws().getRange("F4").getValues()[0][0]);
OVERRIDES.push([movedCoachee, ""], [someCoachee, newCoach]);
sent = main(wb, "");
check(sent.every((x: any) => x.memberEmails.split(";").indexOf(movedCoachee) < 0), "OVERRIDES: [coachee, blank] leaves that coachee out");
check(sent.find((x: any) => x.coachEmail === newCoach)!.memberEmails.split(";").indexOf(someCoachee) > 0, "OVERRIDES: [coachee, coach] moves them to that coach");
OVERRIDES.length = 0;
Date.now = realNow;
set("H8", "");

// Standing requests: pin a pair and cap one coach; everyone else stays matched and even.
{
  set("H6", "");
  const base = main(wb, "yes").filter((x: any) => x.coachClass === "2028");
  const A = base[0].coachEmail, B = base[1];
  const ceB = B.memberEmails.split(";")[1];
  PINS.push([ceB, A]); COACH_CAPS.push([A.split("@")[0], 2]);
  const g2 = main(wb, "yes").filter((x: any) => x.coachClass === "2028");
  const gA = g2.find((x: any) => x.coachEmail === A);
  const others = g2.filter((x: any) => x.coachEmail !== A).map((x: any) => x.coacheeCount);
  check(gA.memberEmails.split(";").indexOf(ceB) > 0, "PINS: pinned coachee is with their coach");
  check(gA.coacheeCount === 2, "COACH_CAPS: capped coach gets exactly the cap (2)");
  check(g2.reduce((n: number, x: any) => n + x.coacheeCount, 0) === 11 && Math.max(...others) - Math.min(...others) <= 1, "caps: every 2030 still matched; other coaches stay even -> " + others.join(","));
  PINS.push(["new.cadet@example.com", A]);
  const g3 = main(wb, "yes").filter((x: any) => x.coachClass === "2028");
  const gA3 = g3.find((x: any) => x.coachEmail === A);
  check(gA3.memberEmails.indexOf("new.cadet@example.com") > 0 && gA3.coacheeCount === 2, "PINS: a pinned coachee who never filled the form is added (cap still holds)");
  PINS.length = 0; COACH_CAPS.length = 0;
  // Neither filled out the form: both given by school email -> their own group with the cap.
  PINS.push(["C30Pat.Kim@example.com", "C28Lee.Ray@example.com"]); COACH_CAPS.push(["lee.ray", 4]);
  const g4 = main(wb, "yes").filter((x: any) => x.coachClass === "2028");
  const gL = g4.find((x: any) => x.coachEmail === "C28Lee.Ray@example.com");
  check(!!gL && gL.memberEmails.indexOf("C30Pat.Kim@example.com") > 0 && gL.coacheeCount <= 4 && gL.coachName === "Lee Ray", "PINS: coach and coachee who never filled the form still get their group (cap 4)");
  check(g4.reduce((n: number, x: any) => n + x.coacheeCount, 0) === 12, "PINS: all 11 real 2030s + the pinned one are matched");
  PINS.length = 0; COACH_CAPS.length = 0;
}

// The live form was reworded to short titles; matching must not change.
const shortTitle: Record<string, string> = {
  "What do you want help with, or want your coach to focus on? (pick up to 5)": "What do you want help with? (up to 5)",
  "Major (if undeclared: majors you're considering) (pick up to 3)": "Major (up to 3 if undecided)",
  "Minor(s) you have or are considering": "Minor",
  "Grad school / fellowships you're interested in": "Education goal",
  "Air Force / Space Force career fields you're interested in (pick up to 4)": "Career fields you want (up to 4)",
  "Long-term (after your commitment), where do you see yourself?": "Long-term plans",
  "Opportunities you want to pursue at USAFA": "What do you want to do at USAFA?",
  "Research areas you're interested in (pick up to 4)": "Research areas (up to 4)",
  "In a sentence or two, what topics are you curious about?": "What are you curious about?",
};
settings().getRange("E5:E16").setValues([[3],[3],[3],[2],[2],[2],[1.5],[1.5],[1.5],[1],[""],[""]]);
const before = JSON.stringify(main(wb, "yes"));
let swapped = 0;
for (let i = 0; i < headers.length; i++) if (shortTitle[headers[i]]) { headers[i] = shortTitle[headers[i]]; swapped++; }
const after = JSON.stringify(main(wb, "yes"));
check(swapped === 9 && before === after, "short question titles give the same matches as the old wording (" + swapped + " renamed)");
console.log("\\nexample intro:\\n" + g[0].intro.replace(/<br>|<\\/p>/g, "\\n").replace(/<[^>]+>/g, "").replace(/&bull;/g, "-"));
console.log(fails === 0 ? "\\nALL PASS" : "\\n" + fails + " FAILED");
`;
const tmp = join(here, "_os.cts");
writeFileSync(tmp, script + "\n" + harness);
try { execFileSync(process.execPath, [tmp], { stdio: "inherit" }); } finally { unlinkSync(tmp); }
