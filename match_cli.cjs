// Runs the same matcher as the website, from the command line.
// Usage: node match_cli.cjs <in.json> <out.json> [coachClasses] [coacheeClasses]
//   in.json  = { headers: [...], rows: [[...], ...] }  (Forms export, first sheet)
//   classes  = comma lists like "2028,2029" and "2030". Default: youngest class is coached, all older classes coach.
const fs = require("fs");
const M = require("./site/matcher.js");

const [inPath, outPath, coachArg, coacheeArg] = process.argv.slice(2);
const data = JSON.parse(fs.readFileSync(inPath, "utf8"));
const { people, missing } = M.readPeople(data.headers, data.rows);

const years = [...new Set(people.map((p) => p.cls))].filter((c) => /^20\d\d$/.test(c)).sort();
const coachClasses = coachArg ? coachArg.split(",") : years.slice(0, -1);
const coacheeClasses = coacheeArg ? coacheeArg.split(",") : years.slice(-1);
const coaches = people.filter((p) => coachClasses.includes(p.cls));
const coachees = people.filter((p) => coacheeClasses.includes(p.cls));

const weights = Object.fromEntries(M.FIELDS.map((f) => [f.key, f.weight]));
const TEMPLATE = "Hi {coacheeFirst}! I'm {coach}, your Martinson Honors coach this semester. Looks like we have some things in common: {shared}. Want to meet up this week to talk classes, research, and whatever else is on your mind?";
const first = (n) => String(n).split(/\s+/)[0];

const m = M.match(coaches, coachees, weights);
const pairs = coachees.map((ce, i) => {
  const co = coaches[m.assignment[i]];
  const parts = M.shared(ce, co, weights);
  const flat = parts.flatMap((p) => p.items).filter((x, k, a) => a.indexOf(x) === k).slice(0, 4);
  const message = M.fillTemplate(TEMPLATE, {
    coach: co.name, coachee: ce.name, coachFirst: first(co.name), coacheeFirst: first(ce.name),
    shared: flat.length ? flat.join(", ") : "a lot to talk about",
  });
  return {
    coach: co.name, coachClass: co.cls, coachEmail: co.email,
    coachee: ce.name, coacheeClass: ce.cls, coacheeEmail: ce.email,
    score: Math.round(m.score[i][m.assignment[i]]),
    shared: parts.map((p) => `${p.label}: ${p.items.join(", ")}`).join(" | "),
    message, link: M.teamsLink(ce.email, message),
  };
}).sort((a, b) => a.coach.localeCompare(b.coach));

const load = {};
pairs.forEach((p) => (load[p.coach] = (load[p.coach] || 0) + 1));
const idle = coaches.filter((c) => !load[c.name]).map((c) => c.name);
fs.writeFileSync(outPath, JSON.stringify({ coachClasses, coacheeClasses, missing, idle, pairs }, null, 1));
console.log(`coaches ${coaches.length} (${coachClasses}) | coachees ${coachees.length} (${coacheeClasses}) | pairs ${pairs.length} | missing columns: ${missing.join(", ") || "none"}`);
