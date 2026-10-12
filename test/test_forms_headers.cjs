// The real Forms export headers (question titles as built in the live form).
const M = require("../site/matcher.js");
const headers = ["Id", "Start time", "Completion time", "Email", "Name", "Last modified time",
  "Class year",
  "What do you want help with, or want your coach to focus on? (pick up to 3)",
  "Major (if undeclared: majors you're considering) (pick up to 3)",
  "Minor(s) you have or are considering",
  "Grad school / fellowships you're interested in",
  "Air Force / Space Force career fields you're interested in (pick up to 4)",
  "Long-term (after your commitment), where do you see yourself?",
  "Opportunities you want to pursue at USAFA",
  "Research areas you're interested in (pick up to 4)",
  "In a sentence or two, what topics are you curious about?",
  "What do you do for fun?",
  "Anything else your coach/coachee should know? (optional, not used for matching)"];
const row = (email, name, cls, major, fun) => ["1", "", "", email, name, "", cls, "Grad school planning;", major + ";", "None / undecided;", "Rhodes;", "Cyber;", "Industry / tech;", "Internships;", "AI / machine learning;", "drone swarms and AI", fun + ";", ""];
const { people, missing } = M.readPeople(headers, [
  row("a@x", "Coach A", "2029", "Computer Science", "Running"),
  row("b@x", "Coach B", "2028", "History", "Chess / board games"),
  row("c@x", "Cadet C", "2030", "Computer Science", "Running"),
  row("d@x", "Cadet D", "2030", "History", "Chess / board games"),
]);
console.log("missing columns:", missing);
const coaches = people.filter((p) => p.cls !== "2030"), coachees = people.filter((p) => p.cls === "2030");
const m = M.match(coaches, coachees, Object.fromEntries(M.FIELDS.map((f) => [f.key, f.weight])));
coachees.forEach((c, i) => console.log(c.name, "->", coaches[m.assignment[i]].name, Math.round(m.score[i][m.assignment[i]])));
