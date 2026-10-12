// Martinson Coaching Match: matching engine (runs entirely in the browser).
// Same algorithm as MatchCadets.ts (the Excel Office Script).
(function (root) {
  "use strict";

  // header: short column name (Power Automate workbook)
  // question: the Microsoft Forms question text (Forms "Open in Excel" export)
  const FIELDS = [
    { key: "focus", header: "Coaching Focus", question: "If you're coaching: what can you help with?", label: "Coaching focus", weight: 3, kind: "list" },
    { key: "major", header: "Major", question: "Major (if undeclared: majors you're considering)", label: "Major", weight: 3, kind: "list" },
    { key: "research", header: "Research Areas", question: "Research areas you're interested in", label: "Research", weight: 3, kind: "list" },
    { key: "grad", header: "Grad School Plans", question: "Grad school / fellowships you're interested in", label: "Grad school", weight: 2, kind: "list" },
    { key: "career", header: "AF Career Interests", question: "Air Force / Space Force career fields you're interested in", label: "AF/SF career", weight: 2, kind: "list" },
    { key: "fun", header: "Fun", question: "What do you do for fun?", label: "For fun", weight: 2, kind: "list" },
    { key: "future", header: "Future Path", question: "Long-term (after your commitment), where do you see yourself?", label: "Future path", weight: 1.5, kind: "list" },
    { key: "opps", header: "Opportunities", question: "Opportunities you want to pursue at USAFA", label: "Opportunities", weight: 1.5, kind: "list" },
    { key: "issues", header: "Issues", question: "Issues you care about", label: "Issues", weight: 1.5, kind: "list" },
    { key: "text", header: "Research Interests", question: "In a sentence or two, what topics are you curious about?", label: "Research keywords", weight: 1.5, kind: "text" },
    { key: "minor", header: "Minor", question: "Minor(s) you have or are considering", label: "Minor", weight: 1, kind: "list" },
  ];
  const ID_COLS = {
    name: ["Name", "Full name"],
    email: ["Email", "USAFA email", "Responders' Email"],
    classYear: ["Class Year", "Class year"],
  };

  const norm = (s) => String(s == null ? "" : s).trim().toLowerCase().replace(/[’']/g, "'").replace(/\s+/g, " ");

  function findColumn(headers, candidates) {
    const h = headers.map(norm);
    for (const c of candidates) {
      const i = h.indexOf(norm(c));
      if (i >= 0) return i;
    }
    for (const c of candidates) {
      const n = norm(c);
      if (n.length < 12) continue;
      const i = h.findIndex((x) => x.startsWith(n.slice(0, 30)));
      if (i >= 0) return i;
    }
    return -1;
  }

  function parseList(raw) {
    const out = new Map();
    const s = String(raw == null ? "" : raw).trim();
    if (!s || s === "null" || s === "undefined") return out;
    let items;
    if (s.startsWith("[")) {
      try { items = JSON.parse(s).map(String); } catch (e) { items = s.replace(/^\[|\]$/g, "").split(/","|;/); }
    } else {
      items = s.split(/[;\n]/);
    }
    for (const item of items) {
      const display = item.replace(/^"|"$/g, "").trim();
      if (display) out.set(display.toLowerCase(), display);
    }
    return out;
  }

  const STOPWORDS = new Set(("the and for with that this are was were have has had not but you your our their they them " +
    "from into about how what which who whom why when where can could would should will also very really just like " +
    "interested interest interests research want study studying learn learning things thing stuff more most some any " +
    "all its it's i'm im lot lots").split(" "));

  function keywords(raw) {
    const out = new Map();
    for (let w of String(raw == null ? "" : raw).toLowerCase().split(/[^a-z0-9+#]+/)) {
      if (w.length < 3 || STOPWORDS.has(w)) continue;
      if (w.length > 4 && w.endsWith("s") && !w.endsWith("ss")) w = w.slice(0, -1);
      out.set(w, w);
    }
    return out;
  }

  function jaccard(a, b) {
    if (!a.size || !b.size) return 0;
    let inter = 0;
    a.forEach((_, k) => { if (b.has(k)) inter++; });
    return inter / (a.size + b.size - inter);
  }

  const isNone = (k) => /^(none|n\/a|not sure|undecided|no|other)\b/.test(k);

  function similarity(a, b, weights) {
    let total = 0, wsum = 0;
    FIELDS.forEach((f, k) => {
      const w = weights[f.key];
      if (!(w > 0)) return;
      wsum += w;
      total += w * jaccard(a.answers[k], b.answers[k]);
    });
    return wsum ? (100 * total) / wsum : 0;
  }

  function shared(a, b, weights) {
    const parts = [];
    FIELDS.forEach((f, k) => {
      if (!(weights[f.key] > 0)) return;
      const common = [];
      a.answers[k].forEach((display, key) => { if (b.answers[k].has(key) && !isNone(key)) common.push(display); });
      if (common.length) parts.push({ label: f.label, items: common.slice(0, 4) });
    });
    return parts;
  }

  // Min-cost assignment, rows <= cols. Returns col index per row.
  function hungarian(cost) {
    const n = cost.length, m = cost[0].length, INF = 1e18;
    const u = new Array(n + 1).fill(0), v = new Array(m + 1).fill(0);
    const p = new Array(m + 1).fill(0), way = new Array(m + 1).fill(0);
    for (let i = 1; i <= n; i++) {
      p[0] = i;
      let j0 = 0;
      const minv = new Array(m + 1).fill(INF), used = new Array(m + 1).fill(false);
      do {
        used[j0] = true;
        const i0 = p[j0];
        let delta = INF, j1 = 0;
        for (let j = 1; j <= m; j++) {
          if (used[j]) continue;
          const cur = cost[i0 - 1][j - 1] - u[i0] - v[j];
          if (cur < minv[j]) { minv[j] = cur; way[j] = j0; }
          if (minv[j] < delta) { delta = minv[j]; j1 = j; }
        }
        for (let j = 0; j <= m; j++) {
          if (used[j]) { u[p[j]] += delta; v[j] -= delta; } else minv[j] -= delta;
        }
        j0 = j1;
      } while (p[j0] !== 0);
      do { const j1 = way[j0]; p[j0] = p[j1]; j0 = j1; } while (j0 !== 0);
    }
    const ans = new Array(n).fill(-1);
    for (let j = 1; j <= m; j++) if (p[j] !== 0) ans[p[j] - 1] = j - 1;
    return ans;
  }

  // Class year as a short label: "2029", or the raw answer if no year in it.
  function classOf(raw) {
    const m = String(raw == null ? "" : raw).match(/20\d\d/);
    return m ? m[0] : String(raw == null ? "" : raw).trim();
  }

  // rows: array of arrays, headers: array of strings. Returns everyone, deduped by email (last wins).
  function readPeople(headers, rows) {
    const cols = {
      name: findColumn(headers, ID_COLS.name),
      email: findColumn(headers, ID_COLS.email),
      classYear: findColumn(headers, ID_COLS.classYear),
    };
    const fieldCols = FIELDS.map((f) => findColumn(headers, [f.header, f.question]));
    const missing = [];
    if (cols.name < 0) missing.push("Name");
    if (cols.email < 0) missing.push("Email");
    if (cols.classYear < 0) missing.push("Class year");
    FIELDS.forEach((f, k) => { if (fieldCols[k] < 0) missing.push(f.header); });

    const byKey = new Map();
    for (const row of rows) {
      const get = (i) => (i < 0 ? "" : row[i] == null ? "" : String(row[i]));
      const name = get(cols.name).trim(), email = get(cols.email).trim();
      if (!name && !email) continue;
      byKey.set((email || name).toLowerCase(), {
        name, email, cls: classOf(get(cols.classYear)) || "Unknown",
        answers: FIELDS.map((f, k) => (f.kind === "text" ? keywords(get(fieldCols[k])) : parseList(get(fieldCols[k])))),
      });
    }
    return { people: [...byKey.values()], missing };
  }

  // Teams deep link: opens a 1:1 chat with `email`, message pre-typed (sender just hits Send).
  function teamsLink(email, message) {
    return "https://teams.microsoft.com/l/chat/0/0?users=" + encodeURIComponent(email) + "&message=" + encodeURIComponent(message);
  }

  function fillTemplate(tpl, v) {
    return tpl.replace(/\{(\w+)\}/g, (all, k) => (k in v ? v[k] : all));
  }

  function match(coaches, coachees, weights) {
    if (!coaches.length || !coachees.length) throw new Error("Need at least one coach and one coachee.");
    const score = coachees.map((ce) => coaches.map((co) => similarity(ce, co, weights)));
    const cap = Math.max(1, Math.ceil(coachees.length / coaches.length));
    const BONUS = 1000;
    const slotCoach = [], slotFirst = [];
    for (let k = 0; k < cap; k++) for (let j = 0; j < coaches.length; j++) { slotCoach.push(j); slotFirst.push(k === 0); }
    const cost = coachees.map((_, i) => slotCoach.map((j, s) => -(score[i][j] + (slotFirst[s] ? BONUS : 0))));
    const a = hungarian(cost);
    return { score, cap, assignment: coachees.map((_, i) => slotCoach[a[i]]) };
  }

  const api = { FIELDS, parseList, keywords, jaccard, similarity, shared, hungarian, readPeople, match, findColumn, classOf, teamsLink, fillTemplate };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Matcher = api;
})(this);
