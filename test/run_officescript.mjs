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
set("H5", "");
check(main(wb, "").length === 0, "blank match date: nothing happens");
check(settings().getRange("A5:B6").getValues().map((r: any) => r.join(">")).join(",") === "2027>2029,2028>2030", "default pairings 2027>2029, 2028>2030");

set("H5", "2999-01-01");
check(main(wb, "").length === 0, "future match date: no chats returned");
const mws = () => sheets.get("Matches");
check(!!mws() && /^PREVIEW ONLY/.test(String(mws().getRange("A1").getValues()[0][0])) && String(mws().getRange("A1").getValues()[0][0]).indexOf("01/01/2999") >= 0, "before the date: Matches sheet says PREVIEW ONLY + the date");
check(String(mws().getRange("A3").getValues()[0][0]) === "Coach" && String(mws().getRange("A4").getValues()[0][0]).length > 0, "preview lists the would-be pairs (header row 3, pairs from row 4)");
check(String(settings().getRange("H7").getValues()[0][0]) === "", "preview does not stamp 'Chats created at'");

set("H5", "2020-01-01");
g = main(wb, "");
const loads = g.map((x: any) => x.coachClass + ":" + x.coacheeCount);
check(g.length === 6, "past match date: one group chat per coach (6 coaches) -> " + loads.join(" "));
check(g.filter((x: any) => x.coachClass === "2027").reduce((n: number, x: any) => n + x.coacheeCount, 0) === 11, "all 11 2029s go to the 3 2027 coaches");
check(g.filter((x: any) => x.coachClass === "2028").reduce((n: number, x: any) => n + x.coacheeCount, 0) === 11, "all 11 2030s go to the 3 2028 coaches");
check(Math.max(...g.map((x: any) => x.coacheeCount)) - Math.min(...g.map((x: any) => x.coacheeCount)) <= 1, "loads are even (3-4 each)");
check(g.every((x: any) => x.memberEmails.split(";").length === x.coacheeCount + 1), "chat members = coach + their coachees");
check(g.every((x: any) => x.chatTitle.indexOf(":") < 0), "chat titles have no ':' (Teams rejects it)");
check(String(settings().getRange("H7").getValues()[0][0]).length > 0, "run stamps 'Chats created at'");
check(/^Group chats created/.test(String(mws().getRange("A1").getValues()[0][0])), "on the date: Matches sheet says 'Group chats created'");
check(main(wb, "").length === 0, "second daily run does nothing (no duplicate chats)");
check(/^Group chats created/.test(String(mws().getRange("A1").getValues()[0][0])), "after chats exist, the Matches record is left alone");
check(main(wb, "yes").length === 6, "runNow=yes ignores the date/stamp (test runs)");

set("H6", 2);
g = main(wb, "yes");
check(Math.max(...g.map((x: any) => x.coacheeCount)) === 2, "max coachees per coach = 2 is respected");
set("H6", "");

settings().getRange("E5:E16").setValues([[0],[0],[0],[0],[0],[0],[0],[0],[0],[5],[""],[""]]);
g = main(wb, "yes");
check(g.length === 6 && g.every((x: any) => !/Coaching focus|Major:/.test(x.intro)), "weights from Settings are used (0 = ignored)");
check(g.some((x: any) => x.coachName === "Jane Doe"), "Teams-style display name becomes 'Jane Doe' from the school email");
console.log("\\nexample intro:\\n" + g[0].intro.replace(/<br>|<\\/p>/g, "\\n").replace(/<[^>]+>/g, "").replace(/&bull;/g, "-"));
console.log(fails === 0 ? "\\nALL PASS" : "\\n" + fails + " FAILED");
`;
const tmp = join(here, "_os.cts");
writeFileSync(tmp, script + "\n" + harness);
try { execFileSync(process.execPath, [tmp], { stdio: "inherit" }); } finally { unlinkSync(tmp); }
