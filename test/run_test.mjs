// Runs MatchCadets.ts locally against test/sample.json with a fake Excel
// workbook, so the matching logic can be checked without Excel.
// Usage: node test/run_test.mjs
import { readFileSync, writeFileSync, unlinkSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const script = readFileSync(join(here, "..", "MatchCadets.ts"), "utf8");
const harness = `
const data = JSON.parse(require("node:fs").readFileSync(${JSON.stringify(join(here, "sample.json"))}, "utf8"));
let written: any = null;
const fakeRange = (v?: any) => ({
  getValues: () => v,
  setValues: (x: any) => { written = x; },
  getFormat: () => ({ autofitColumns() {}, setColumnWidth() {}, setWrapText() {} }),
});
const workbook: any = {
  getTable: (n: string) => n === "Responses" ? {
    getHeaderRowRange: () => fakeRange([data.headers]),
    getRangeBetweenHeaderAndTotal: () => fakeRange(data.rows),
  } : undefined,
  getWorksheet: () => undefined,
  addWorksheet: () => ({
    getRangeByIndexes: () => fakeRange(),
    addTable: () => ({ setName() {} }),
    getRange: () => fakeRange(),
    activate() {},
  }),
};
const result = main(workbook);
console.table(written.slice(1).map((r: any[]) => ({ coach: r[0], coachee: r[2], score: r[4], shared: String(r[5]).slice(0, 70) })));
const load: Record<string, number> = {};
for (const m of result) load[m.coachName] = (load[m.coachName] || 0) + 1;
console.log("coachees per coach:", load);
console.log("matched coachees:", result.length, "of", data.rows.filter((r: any[]) => r[2] === "2030").length);
`;
const tmp = join(here, "_combined.cts");
writeFileSync(tmp, script + "\n" + harness);
try { execFileSync(process.execPath, [tmp], { stdio: "inherit" }); } finally { unlinkSync(tmp); }
