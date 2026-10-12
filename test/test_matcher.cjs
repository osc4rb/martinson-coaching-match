// Checks site/matcher.js: sample data run + Hungarian vs brute force.
const M = require("../site/matcher.js");
const data = require("./sample.json");
const weights = Object.fromEntries(M.FIELDS.map((f) => [f.key, f.weight]));
const { coaches, coachees, missing } = M.readPeople(data.headers, data.rows, { coachClass: "2029", coacheeClass: "2030" });
console.log("coaches", coaches.length, "coachees", coachees.length, "missing", missing);
const r = M.match(coaches, coachees, weights);
const load = {}; r.assignment.forEach((j) => (load[j] = (load[j] || 0) + 1));
console.log("loads", Object.values(load).sort().join(","), "every coach used:", Object.keys(load).length === coaches.length);
// brute force on random 6x6 matrices
const perms = (a) => a.length < 2 ? [a] : a.flatMap((x, i) => perms([...a.slice(0, i), ...a.slice(i + 1)]).map((p) => [x, ...p]));
let ok = true;
for (let t = 0; t < 200; t++) {
  const n = 6, c = Array.from({ length: n }, () => Array.from({ length: n }, () => Math.round(Math.random() * 100)));
  const h = M.hungarian(c).reduce((s, j, i) => s + c[i][j], 0);
  const b = Math.min(...perms([...Array(n).keys()]).map((p) => p.reduce((s, j, i) => s + c[i][j], 0)));
  if (h !== b) ok = false;
}
console.log("hungarian optimal on 200 random cases:", ok);
