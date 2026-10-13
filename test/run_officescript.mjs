// Runs MatchCadets.ts locally with a fake ExcelScript workbook shaped like the
// Forms "Open in Excel" sheet (real question headers, "A;B;" checkbox values).
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
  "Anything Else": "Anything else your coach/coachee should know? (optional, not used for matching)",
  "Class Year": "Class year",
};
const headers = ["Id", "Start time", "Completion time", "Email", "Name", "Last modified time"];
const keep = s.headers.map((h: string, i: number) => [h, i]).filter(([h]: any) => map[h]);
for (const [h] of keep as any) headers.push(map[h]);
const toForms = (v: string) => v.startsWith("[") ? JSON.parse(v).join(";") + ";" : v;
const rows = s.rows.map((r: any[], k: number) => [String(k + 1), "", "", r[1], k % 3 === 0 ? "" : r[0], "", ...(keep as any).map(([h, i]: any) => toForms(String(r[i])))]);
let written: any = null;
const rng = (v?: any) => ({ getValues: () => v, setValues: (x: any) => { written = x; }, getFormat: () => ({ autofitColumns() {} }) });
const wb: any = {
  getTables: () => [{ getHeaderRowRange: () => rng([headers]), getRangeBetweenHeaderAndTotal: () => rng(rows) }],
  getWorksheet: () => undefined,
  addWorksheet: () => ({ getRangeByIndexes: () => rng() }),
};
const def = main(wb, "", "");
console.log("default roles ->", def.length, "pairs; first:", def[0].coachName, "(" + def[0].coachClass + ") ->", def[0].coacheeName, "(" + def[0].coacheeClass + ")");
console.log("intro:", def[0].intro);
const custom = main(wb, "2028", "2029,2030");
const load: Record<string, number> = {}; custom.forEach((m: any) => load[m.coachName] = (load[m.coachName] || 0) + 1);
console.log("2028 coach 2029+2030 ->", custom.length, "pairs; loads", Object.values(load).join(","));
console.log("blank-name fallback example:", def.map((m: any) => m.coachName).concat(def.map((m: any) => m.coacheeName)).find((n: string) => /^[A-Z][a-z]+ [A-Z]/.test(n)));
console.log("Matches sheet rows:", written.length - 1);
`;
const tmp = join(here, "_os.cts");
writeFileSync(tmp, script + "\n" + harness);
try { execFileSync(process.execPath, [tmp], { stdio: "inherit" }); } finally { unlinkSync(tmp); }
