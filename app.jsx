const { useEffect, useMemo, useState } = React;
/* XLSX is a global from the xlsx CDN script */

/*
  Jira Migration Validator — V1 (browser-only).
  Upload source (iTrack) + target (Jira Cloud) CSV exports; the comparison runs
  entirely in the browser (no server, no MCP, nothing uploaded anywhere). Produces
  the dashboard, per-bucket detail, an overall score, a "reach 100%" checklist,
  and an Excel report. This is the exact V1 flow; V2/V3 later swap the file parse
  for live MCP calls behind the same engine + UI.

  Field map / matching mirror the Python engine (config.yaml). Edit FIELDS / KEYS
  to match your real export column names.
*/

const KEYS = { key: "Issue key", project: "Project key", mwid: "Migrated Work Item ID", created: "Created" };
const FIELDS = {
  status:               { col: "Status",                             cmp: "exact",  w: 3, sev: "high" },
  issue_type:           { col: "Issue Type",                         cmp: "exact",  w: 2, sev: "high" },
  parent_key:           { col: "Parent",                             cmp: "exact",  w: 2, sev: "high" },
  assignee:             { col: "Assignee",                           cmp: "exact",  w: 2, sev: "med"  },
  reporter:             { col: "Reporter",                           cmp: "exact",  w: 1, sev: "low"  },
  description:          { col: "Description",                        cmp: "text",   w: 2, sev: "med"  },
  acceptance_criteria:  { col: "Custom field (Acceptance Criteria)", cmp: "text",   w: 1, sev: "med"  },
  priority:             { col: "Priority",                           cmp: "exact",  w: 1, sev: "med"  },
  resolution:           { col: "Resolution",                         cmp: "exact",  w: 1, sev: "med"  },
  story_points:         { col: "Custom field (Story Points)",        cmp: "number", w: 1, sev: "med"  },
  due_date:             { col: "Due Date",                           cmp: "date",   w: 1, sev: "low"  },
  components:           { col: "Component/s",                        cmp: "set",    w: 1, sev: "low"  },
  fix_versions:         { col: "Fix Version/s",                      cmp: "set",    w: 1, sev: "low"  },
  labels:               { col: "Labels",                             cmp: "set",    w: 1, sev: "low"  },
  links:                { col: "Outward issue link",                 cmp: "set",    w: 1, sev: "low"  },
  sprint_name:          { col: "Sprint",                             cmp: "text",   w: 1, sev: "low"  },
  jira_align_team_name: { col: "Jira Align Team Name",               cmp: "text",   w: 1, sev: "low"  },
  comments:             { col: "Comment",                            cmp: "count",  w: 1, sev: "low"  },
  comment_count:        { col: "Comment",                            cmp: "count",  w: 1, sev: "low"  },
  attachment_count:     { col: "Attachment",                         cmp: "count",  w: 1, sev: "low"  },
};
const SCORING = { coverage: 0.4, field: 0.6, pass: 95, warn: 85 };
const FIELD_NAMES = Object.keys(FIELDS);
const TOTAL_W = FIELD_NAMES.reduce((s, f) => s + FIELDS[f].w, 0);

// Landing-page sources. Each has its own value-mapping files (status, priority…).
const SOURCES = [
  { id: "itrack", name: "iTrack", desc: "Jira Data Center" },
  { id: "attjira", name: "ATT Jira", desc: "Jira Server" },
  { id: "ado", name: "Azure DevOps", desc: "Azure DevOps" },
];
// Fields whose source values are normalised to Cloud values via a mapping file.
const MAPPABLE = ["status", "priority", "issue_type"];
// Fields compared by user identity (via the user list), not raw string match.
const USER_FIELDS = ["assignee", "reporter"];

// Preloaded per-source value mappings (rules keyed by normalised source value).
// Replace these with your real mapping files (or fetch them from /public at load).
const DEFAULT_MAPPINGS = {
  itrack: {
    status: { name: "iTrack_to_Jira_Clous_Status_Mappings.xlsx", rules: { "new": "New", "open": "New", "backlog": "New", "ready": "New", "in development": "In Progress", "sprint to do": "In Progress", "reopened": "In Progress", "on hold": "In Progress", "need more info": "In Progress", "in progress": "In Progress", "ready for release": "In Progress", "in review": "In Progress", "waiting for qa": "In Progress", "in qa": "In Progress", "dev complete": "In Progress", "to do": "In Progress", "demo-present": "In Progress", "test complete": "In Progress", "resolved": "In Progress", "released": "In Progress", "cancelled": "Cancelled", "closed": "Closed", "complete": "Closed", "accepted": "Closed" } },
    priority: { name: "Itrack_to_Jira_Cloud-priority-field-mapping.xlsx", rules: { major: "High", trivial: "Medium", minor: "Low", blocker: "", critical: "Critical" } },
    issue_type: { name: "Itrack_to_Jira_Cloud-priority-field-mapping.xlsx", rules: { epic: "Feature", feature: "Feature", story: "Story", task: "Task", "sub-task": "Sub-task", bug: "Defect" } },
  },
  attjira: {
    status: { name: "ATT Jira status mapping", rules: { "ready for acceptance": ["Ready for Review"], "in progress": ["In Progress"], "in review": ["In Progress"], "blocked": ["In Progress"], "delivered": ["Dev Complete"], "in pst": ["Dev Complete"], "in qa": ["Dev Complete"], "deleted": ["Cancelled"], "accepted": ["Closed"], "new": ["New", "Ready"], "done": ["Closed"] } },
    // priority: not required — ATT Jira and Cloud share the same priorities
    // issue_type: not mapped — same type names as Cloud
  },
  ado: {
    status: { name: "ADO_to_JC_State_Mapping.xlsx", rules: { "feature|product - defining requirements": "New", "feature|architect solutioning": "Analyzing", "feature|ready for pi": "Ready", "feature|implementing": "In Progress", "feature|test complete": "Ready for Review", "feature|accepted": "Accepted", "feature|deployed": "Accepted", "feature|released": "Accepted", "feature|canceled": "Canceled", "user story|pending approval": "New", "user story|ready to start": "Ready", "user story|implementing": "In Development", "user story|dev complete": "Dev Complete", "user story|test complete": "Ready for Review", "user story|accepted": "Closed", "user story|canceled": "Canceled", "task|not started": "New", "task|in progress": "In Progress", "task|done": "Closed", "task|canceled": "Deleted", "bug|<not used>": "New", "bug|open": "Analysis", "bug|accepted": "Closed", "bug|canceled": "Rejected/Returned" } },
    issue_type: { name: "ADO_to_JC_Field_Mapping.xlsx", rules: { feature: "Feature", "user story": "User Story", task: "Sub-Task", bug: "Defect" } },
    priority: { name: "ADO_to_JC_priority", rules: { "": "Low", "0": "Low", "1": "Low", "2": "Medium", "3": "High", "4": "Critical", "5": "Critical", "6": "Critical", "7": "Critical", "8": "Critical", "9": "Critical", "10": "Critical" } },
  },
};
// ADO uses different export column names — per-source SOURCE-side column overrides.
const SOURCE_COLUMNS = {
  ado: { summary: "Title", issue_type: "Work Item Type", status: "State", priority: "Priority",
    assignee: "Assigned To", reporter: "Created By", story_points: "Points", labels: "Tags",
    description: "Description", acceptance_criteria: "Acceptance Criteria", parent_key: "Parent",
    due_date: "Target Date" },
};
// Target (Jira Cloud) columns that differ for a given source's migrated items.
const TARGET_COLUMNS = {
  ado: { priority: "Business Priority" },
};
const SOURCE_KEYS = {
  ado: { key: "ID", project: "Area Path" },
};
// Fields whose mapping is keyed by (type|value) instead of value alone.
const SOURCE_COMPOSITE = { ado: ["status"] };
// Shared user list (common to every source). A small sample is inlined so the app
// works standalone; the FULL org list (28k+ rows) should be bundled as a static
// asset and loaded via REFERENCE_USERS_URL rather than inlined.
const REFERENCE_USERS_URL = "./reference/users.json";   // compact alias map (bundled static asset)
const DEFAULT_USER_ROWS = [
  ["Name", "ATTUID", "Email"],
  ["LEGG, JEREMY", "jl870j", "jl870j@att.com"],
  ["BAICH, RICH", "jz6055", "jz6055@att.com"],
  ["CLARK, ALAINA", "az9067", "az9067@att.com"],
  ["DEATS, JON", "jd365n", "jd365n@att.com"],
  ["ARCOT, RAMESH", "ra038a", "ra038a@att.com"],
  ["BHATT, PREMA", "pb6115", "pb6115@att.com"],
  ["BOPPANA, SRINIVAS", "sq6918", "sq6918@att.com"],
  ["CHITHAMBARAM, RAMASAMY", "cr597a", "cr597a@att.com"],
  ["DONATH, CARL", "cd6456", "cd6456@att.com"],
  ["ERICKSON, MATTHEW", "me3237", "me3237@att.com"],
  ["GUPTA, ANURADHA", "ag661p", "ag661p@att.com"],
];

/* ---------------- engine (JS port of engine.py essentials) ---------------- */
function parseCsv(text) {
  const rows = [];
  let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; }
      else if (c === '"') q = false;
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === ",") { row.push(cell); cell = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += c;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((x) => x.trim() !== ""));
}

function readJira(text) {
  const rows = parseCsv(text);
  if (!rows.length) return [];
  const headers = rows[0].map((h) => h.trim());
  return rows.slice(1).map((r) => {
    const acc = {};
    headers.forEach((h, i) => {
      const v = (r[i] || "").trim();
      if (!(h in acc)) acc[h] = v;
      else if (v) acc[h] = acc[h] ? acc[h] + ";" + v : v;   // repeated Jira columns -> joined
    });
    return acc;
  });
}

function parseMapping(text) {
  const rows = parseCsv(text), out = {};
  rows.forEach((r, i) => {
    const a = (r[0] || "").trim(), b = (r[1] || "").trim();
    if (i === 0 && /source|from|itrack|value/i.test(a) && /target|to|jira|cloud/i.test(b)) return; // header row
    if (a) out[norm(a)] = b;    // normalised source value -> expected Cloud value
  });
  return out;
}

function buildUserMap(rowsArrays) {
  // Each row = one person; every non-empty cell (email, ID, name, name+ID…) is an
  // alias for that person. Two values are the same user if they share a row.
  const map = {};
  rowsArrays.forEach((r, gi) => {
    if (gi === 0) return;                       // skip header row
    r.forEach((cell) => { const v = norm(cell); if (v) map[v] = gi; });
  });
  return map;
}

function resolveUser(v, map) {
  const k = norm(v);
  return (map && map[k] !== undefined) ? "user#" + map[k] : k;   // fall back to raw value if unknown
}

function toIssues(raw, side, source) {
  const colOverride = side === "source" ? (SOURCE_COLUMNS[source] || {}) : (TARGET_COLUMNS[source] || {});
  const kcol = { ...KEYS, ...(side === "source" ? (SOURCE_KEYS[source] || {}) : {}) };
  const sumCol = colOverride.summary || "Summary";
  return raw.map((row) => {
    const fields = {};
    FIELD_NAMES.forEach((f) => { fields[f] = row[colOverride[f] || FIELDS[f].col] || ""; });
    return { key: (row[kcol.key] || "").trim(), project: (row[kcol.project] || "").trim(), summary: (row[sumCol] || "").trim(), fields, raw: row };
  });
}

const norm = (s) => String(s ?? "").trim().replace(/\s+/g, " ").toLowerCase();
function cmp(field, a, b) {
  const kind = FIELDS[field].cmp;
  if (kind === "number") { const x = parseFloat(a || 0), y = parseFloat(b || 0); return (isNaN(x) || isNaN(y)) ? norm(a) === norm(b) : x === y; }
  if (kind === "set") {
    const S = (v) => new Set(String(v || "").split(";").map((x) => x.trim().toLowerCase()).filter(Boolean));
    const A = S(a), B = S(b); return A.size === B.size && [...A].every((x) => B.has(x));
  }
  if (kind === "date") {
    const d = (v) => { const x = new Date(String(v).trim()); return isNaN(x) ? norm(v) : x.toISOString().slice(0, 10); };
    return d(a) === d(b);
  }
  if (kind === "count") {
    const n = (v) => String(v || "").split(";").map((x) => x.trim()).filter(Boolean).length;
    return n(a) === n(b);
  }
  return norm(a) === norm(b);   // text + exact both normalise
}

function fieldApplies(f, issue) {
  const t = norm(issue.fields.issue_type);
  if (f === "story_points") return t === "story" || t === "user story";        // points only on stories
  if (f === "jira_align_team_name") return !(t === "feature" || t === "epic"); // excluded for features/epics
  return true;
}

function validate(source, target, mappings, userMap, composite) {
  // match: MWID on target, else summary
  const byMwid = {};
  target.forEach((t) => { const m = (t.raw[KEYS.mwid] || "").trim(); if (m) (byMwid[m] = byMwid[m] || []).push(t); });
  const bySummary = {};
  target.forEach((t) => { const k = norm(t.summary); (bySummary[k] = bySummary[k] || []).push(t); });

  const used = new Set();
  const pairs = [], missing = [];
  source.forEach((s) => {
    let t = (byMwid[s.key] || []).find((x) => !used.has(x));
    if (!t) t = (bySummary[norm(s.summary)] || []).find((x) => !used.has(x));
    if (t) { used.add(t); pairs.push([s, t]); } else missing.push(s);
  });

  // duplicates: >1 target sharing a source key (MWID)
  const dupMap = {};
  Object.entries(byMwid).forEach(([k, list]) => { if (list.length > 1) dupMap[k] = list.map((x) => x.key); });
  const dupKeys = new Set(Object.keys(dupMap));
  const extra = target.filter((t) => !used.has(t) && !dupKeys.has((t.raw[KEYS.mwid] || "").trim()));

  // per-issue field comparison
  const issueRows = [], discrepancies = [];
  let hierarchyMissing = 0;
  pairs.forEach(([s, t]) => {
    if (norm(s.fields.parent_key) && !norm(t.fields.parent_key)) hierarchyMissing++;   // source has a parent, target doesn't
    let mw = 0, applW = 0; const per = {}; const mism = [];
    FIELD_NAMES.forEach((f) => {
      if (!fieldApplies(f, s)) { per[f] = null; return; }   // not relevant for this work item type
      applW += FIELDS[f].w;
      let sv = s.fields[f];
      let expected, mapped;
      const map = mappings && mappings[f];
      if (map) {
        const mk = (composite && composite.includes(f)) ? norm(s.fields.issue_type) + "|" + norm(sv) : norm(sv);
        mapped = map[mk];
      }
      let ok;
      if (Array.isArray(mapped)) {                     // valid-target-set: match if target is any allowed value
        const set = new Set(mapped.map((x) => norm(x)));
        ok = set.has(norm(t.fields[f]));
        expected = mapped.join(" / ");
      } else if (USER_FIELDS.includes(f)) {            // identity comparison + data-quality flags
        const un = (v) => { const n = norm(v); return n === "" || n === "unassigned"; };
        const unm = (v) => userMap && !un(v) && norm(v) !== "" && userMap[norm(v)] === undefined;
        if (un(s.fields[f]) || un(t.fields[f])) { ok = false; expected = "Flag: Unassigned"; }
        else if (unm(s.fields[f]) || unm(t.fields[f])) { ok = false; expected = "Flag: Unmatched account"; }
        else { ok = resolveUser(s.fields[f], userMap) === resolveUser(t.fields[f], userMap); expected = s.fields[f]; }
      } else {
        if (mapped !== undefined) sv = mapped;         // single-target translation
        ok = cmp(f, sv, t.fields[f]); expected = sv;
      }
      per[f] = ok;
      if (ok) mw += FIELDS[f].w;
      else { mism.push(f); discrepancies.push({ itrackKey: s.key, jiraKey: t.key, field: f, sourceValue: String(s.fields[f]), expected: String(expected), targetValue: String(t.fields[f]), severity: FIELDS[f].sev }); }
    });
    issueRows.push({ source_key: s.key, target_key: t.key, project: s.project, fieldScore: applW ? Math.round(1000 * mw / applW) / 10 : 100, per, mism, summary: s.summary });
  });

  const matchedClean = issueRows.filter((r) => !r.mism.length).length;
  const differences = issueRows.length - matchedClean;
  const duplicates = Object.values(dupMap).reduce((s, v) => s + v.length, 0);

  // global score
  const coverage = source.length ? 100 * pairs.length / source.length : 0;
  const fieldAcc = issueRows.length ? issueRows.reduce((s, r) => s + r.fieldScore, 0) / issueRows.length : 0;
  const fidelity = SCORING.coverage * coverage + SCORING.field * fieldAcc;

  const summary = {
    itrack_count: source.length, jira_count: target.length, migrated: pairs.length,
    pending_migration: missing.length, jira_only: extra.length, differences, matched: matchedClean,
    hierarchy_missing: hierarchyMissing,
    duplicates, all_fields: FIELD_NAMES.length,
    coverage: Math.round(coverage * 10) / 10, field_accuracy: Math.round(fieldAcc * 10) / 10,
    score: Math.round(fidelity * 10) / 10,
    field_pct: Object.fromEntries(FIELD_NAMES.map((f) => {
      const appl = issueRows.filter((r) => r.per[f] === true || r.per[f] === false);
      const m = appl.filter((r) => r.per[f] === true).length;
      return [f, appl.length ? Math.round(1000 * m / appl.length) / 10 : null];
    })),
  };

  const details = {
    pending_migration: missing.map((i) => ({ "iTrack key": i.key, Summary: i.summary, Type: i.fields.issue_type, Status: i.fields.status, Assignee: i.fields.assignee, Priority: i.fields.priority, "Expected result": "Create this issue in Jira Cloud" })),
    jira_only: extra.map((i) => ({ "Jira key": i.key, Summary: i.summary, Type: i.fields.issue_type, Status: i.fields.status, "Expected result": "Verify — no matching source item" })),
    matched: issueRows.filter((r) => !r.mism.length).map((r) => ({ "iTrack key": r.source_key, "Jira key": r.target_key, Summary: r.summary, "Field score": r.fieldScore + "%", "Expected result": "Matches source — no action" })),
    differences: discrepancies.map((d) => ({ "iTrack key": d.itrackKey, "Jira key": d.jiraKey, Field: d.field, "iTrack value": d.sourceValue, "Expected in Cloud": d.expected, "Jira value (actual)": d.targetValue, Severity: d.severity, "Expected result": d.expected.startsWith("Flag:") ? d.expected.replace("Flag:", "⚑") : `Set ${d.field} to "${d.expected}"` })),
    duplicates: Object.entries(dupMap).flatMap(([, list]) => list.map((k) => ({ "Jira key": k, "Matched with": list.filter((x) => x !== k).join(";"), "Detection reason": "source-key-conflict", "Expected result": "Keep one; delete the duplicate(s)" }))),
    all_fields: FIELD_NAMES.map((f) => {
      const appl = issueRows.filter((r) => r.per[f] === true || r.per[f] === false);
      const m = appl.filter((r) => r.per[f]).length;
      return { Field: f, Matched: m, Mismatched: appl.length - m, "Match score": appl.length ? Math.round(1000 * m / appl.length) / 10 + "%" : "n/a", "Target": "100%" };
    }),
  };
  return { summary, details };
}

/* ---------------- UI ---------------- */
const PALETTES = {
  dark: { bg: "#4c4f55", panel: "#565a61", border: "#6b6f78", card: "#565a61", text: "#ffffff", muted: "#dfe3e8", faint: "#b8bdc4", field: "#45484d", accent: "#5b9dff", head: "#565a61", rowAlt: "#515560" },
  light: { bg: "#d7dbe4", panel: "#ffffff", border: "#bcc3d0", card: "#ffffff", text: "#272d3b", muted: "#515a6b", faint: "#7f8797", field: "#eef1f6", accent: "#2f4d94", head: "#e2e6ee", rowAlt: "#eceff5" },
  cb: { bg: "#2b2b2b", panel: "#333333", border: "#474747", card: "#333333", text: "#ededed", muted: "#a8a8a8", faint: "#757575", field: "#262626", accent: "#e69f00", head: "#333333", rowAlt: "#2f2f2f" },
};
const RAMPS = {
  dark: { blue: "#4d9fff", purple: "#c08cff", green: "#2ee06e", teal: "#25e0c8", amber: "#ffc23d", red: "#ff6b7d", indigo: "#8aa0ff" },
  light: { blue: "#3d59a1", purple: "#8250df", green: "#2e8b6f", teal: "#3f8f88", amber: "#b7791f", red: "#c0485a", indigo: "#5a6fb0" },
  cb: { blue: "#56b4e9", purple: "#cc79a7", green: "#3fae90", teal: "#3fae90", amber: "#e69f00", red: "#e07a3f", indigo: "#4a90c9" },
};
let C = PALETTES.dark;
let RAMP = RAMPS.dark;
const TILES = [
  ["pending_migration", "Pending migration", "teal"], ["differences", "Differences", "amber"],
  ["duplicates", "Duplicates", "purple"], ["jira_only", "Jira only", "indigo"],
  ["matched", "Matched", "green"], ["all_fields", "All fields", "blue"],
];

function MultiFileDrop({ label, files, onAdd, onRemove }) {
  const read = (fileList) => {
    Promise.all([...fileList].map((f) => new Promise((res) => {
      const r = new FileReader(); r.onload = () => res({ name: f.name, text: r.result }); r.readAsText(f);
    }))).then(onAdd);
  };
  return (
    <div style={{ flex: 1, minWidth: 300 }}>
      <div style={{ color: C.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>{label}</div>
      <label style={{ display: "block", cursor: "pointer" }}>
        <div style={{ border: `1.5px dashed ${files.length ? C.accent : C.border}`, borderRadius: 12, padding: "22px 18px",
          background: C.field, textAlign: "center", color: files.length ? C.text : C.muted, fontSize: 13.5 }}>
          {files.length ? `${files.length} file${files.length > 1 ? "s" : ""} selected — click to add more…` : "Click to choose Jira CSV export(s)…"}
        </div>
        <input type="file" accept=".csv" multiple style={{ display: "none" }}
          onChange={(e) => { read(e.target.files); e.target.value = ""; }} />
      </label>
      {files.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
          {files.map((f, i) => (
            <span key={i} style={{ display: "inline-flex", alignItems: "center", gap: 7, padding: "5px 10px",
              background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12.5, color: C.text }}>
              {f.name}
              <button onClick={() => onRemove(i)} aria-label="Remove"
                style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 15, lineHeight: 1, padding: 0 }}>×</button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

function Card({ title, value, desc, color }) {
  return (
    <div style={{ flex: 1, minWidth: 200, background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "18px 20px" }}>
      <div style={{ color, fontSize: 13.5, fontWeight: 600 }}>{title}</div>
      <div style={{ fontSize: 34, fontWeight: 700, margin: "8px 0 6px" }}>{value?.toLocaleString?.() ?? value}</div>
      <div style={{ color: C.faint, fontSize: 12 }}>{desc}</div>
    </div>
  );
}

function csvOf(rows) {
  if (!rows || !rows.length) return "";
  const cols = Object.keys(rows[0]);
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  return [cols.map(esc).join(","), ...rows.map((r) => cols.map((c) => esc(r[c])).join(","))].join("\n");
}

// ---- run log: every comparison records a compact entry for the Migration Report ----
// Local (this machine) by default. Set RUNS_API to a backend to aggregate ALL users.
const RUNS_API = "";   // e.g. "/api/runs"  (GET = all runs, POST = append one)
let memRuns = null;
function loadRuns() {
  if (memRuns) return memRuns;
  try { memRuns = JSON.parse(localStorage.getItem("cc.runs") || "[]"); } catch { memRuns = []; }
  return memRuns;
}
function saveRun(rec) {
  const runs = loadRuns(); runs.push(rec); if (runs.length > 5000) runs.shift(); memRuns = runs;
  try { localStorage.setItem("cc.runs", JSON.stringify(runs)); } catch { /* memory only */ }
  if (RUNS_API) { try { fetch(RUNS_API, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(rec) }); } catch { /* best-effort */ } }
}
function fetchRuns() {
  if (RUNS_API) return fetch(RUNS_API).then((r) => r.json()).catch(() => loadRuns());
  return Promise.resolve(loadRuns());
}
function getUser() { try { return localStorage.getItem("cc.user") || "unknown"; } catch { return "unknown"; } }
function setUserName(n) { try { localStorage.setItem("cc.user", n); } catch { /* ignore */ } }

// ---- validation tracker ----
const STAGES = ["Backlog", "In Validation", "Validation Complete", "Review with DRO", "Launched", "Done", "Cancelled"];
const statusForStage = (st) => st;   // status IS the validation stage
function inferStage(status) {
  const s = String(status || "").trim().toLowerCase();
  const map = { backlog: "Backlog", "in validation": "In Validation", "validation complete": "Validation Complete", "review with dro": "Review with DRO", launched: "Launched", done: "Done", cancelled: "Cancelled", canceled: "Cancelled", new: "Backlog", "": "Backlog" };
  return map[s] || "Backlog";
}
const DEFAULT_TRACKER = [{"id":"JCJAMT-7","issueKey":"JCJAMT-7","summary":"CDO Data Inventory","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"ps7299","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-9","issueKey":"JCJAMT-9","summary":"EDF-BDE","tool":"LeanKit - net new","droPortfolio":"Cohen","programArtPoc":"Sreeni Kaparthi","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-10","issueKey":"JCJAMT-10","summary":"CloudRunner-Historic Data","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Valarie Littles","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-12","issueKey":"JCJAMT-12","summary":"DDoS Defense Orchestrator","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Patrick Burke, Savitha Iyer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-13","issueKey":"JCJAMT-13","summary":"CSO EFORC","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Patrick Burke, Jen Delisle","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-14","issueKey":"JCJAMT-14","summary":"CSO Agentless","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Patrick Burke, Brittany Vassallo","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-15","issueKey":"JCJAMT-15","summary":"[Fresh] CSO Edge Security","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Patrick Burke","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-16","issueKey":"JCJAMT-16","summary":"FirstNet & NG911","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Heather Lydon hl5773","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-17","issueKey":"JCJAMT-17","summary":"IGLOO Enhancements","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Nilsa Samol","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-18","issueKey":"JCJAMT-18","summary":"Fiber Mapping Phase 3","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Nilsa Samol","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-19","issueKey":"JCJAMT-19","summary":"ASEoD Speed Floors","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Nilsa Samol, Rachelle Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-20","issueKey":"JCJAMT-20","summary":"Pre-Sales,Contact, Credit,Payment,Pricing","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-21","issueKey":"JCJAMT-21","summary":"Platform Capabilities","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-22","issueKey":"JCJAMT-22","summary":"OWS & Wholesale","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Adnan Rizvi","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-23","issueKey":"JCJAMT-23","summary":"OWS & Whole Sale Billing Vendor Management","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Adnan Rizvi (sr728x), Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-24","issueKey":"JCJAMT-24","summary":"Order Capability, Data","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-25","issueKey":"JCJAMT-25","summary":"Order Capability","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-26","issueKey":"JCJAMT-26","summary":"NGGN-I","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-27","issueKey":"JCJAMT-27","summary":"Next Gen Pricing","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Hana Anderson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-28","issueKey":"JCJAMT-28","summary":"IYOE","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-29","issueKey":"JCJAMT-29","summary":"IPTF, MPCFNNOW","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-30","issueKey":"JCJAMT-30","summary":"INDIRECT","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-31","issueKey":"JCJAMT-31","summary":"Hyperloop, ACVP","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Sridhar Reddy, sx4748","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-32","issueKey":"JCJAMT-32","summary":"Frontend BCSS","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-33","issueKey":"JCJAMT-33","summary":"FOBPM","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-34","issueKey":"JCJAMT-34","summary":"FirstNet, CREOLA, MPCFNNOW, FNMIS","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Prathibha Murthy","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-35","issueKey":"JCJAMT-35","summary":"EPC","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-36","issueKey":"JCJAMT-36","summary":"BTPD - Data Fabric (EDF)","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Anshul Mishra/Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-38","issueKey":"JCJAMT-38","summary":"Enhanced Internet (SDW)","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-39","issueKey":"JCJAMT-39","summary":"Enhanced Internet","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-40","issueKey":"JCJAMT-40","summary":"E2E, PVT","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Devaraju Muniraju","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-41","issueKey":"JCJAMT-41","summary":"Dreamworks ART","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-42","issueKey":"JCJAMT-42","summary":"DOMS, T2R","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-43","issueKey":"JCJAMT-43","summary":"DOMS","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-44","issueKey":"JCJAMT-44","summary":"Data Capability","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-45","issueKey":"JCJAMT-45","summary":"Customer Pricing","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ratna Bhargavi Batchu","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-46","issueKey":"JCJAMT-46","summary":"CSPS - SE","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-47","issueKey":"JCJAMT-47","summary":"CREOLA, IPTF","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-48","issueKey":"JCJAMT-48","summary":"BTPD - Service Delivery Mobility","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Todd Meyer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-49","issueKey":"JCJAMT-49","summary":"Business Contact Center Transformation","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Michael Burke","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-50","issueKey":"JCJAMT-50","summary":"COMS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Praveen Pola","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-51","issueKey":"JCJAMT-51","summary":"COMMON CAPABILITIES","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Raja Shekar Bollam","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-52","issueKey":"JCJAMT-52","summary":"CLEAR","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Namratha Ambekar","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-53","issueKey":"JCJAMT-53","summary":"CDIS","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-54","issueKey":"JCJAMT-54","summary":"CDEX","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Punit Shah","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-55","issueKey":"JCJAMT-55","summary":"CDCOG","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-56","issueKey":"JCJAMT-56","summary":"Business Center","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Swathi Nagamahanti sn767m","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-57","issueKey":"JCJAMT-57","summary":"BRASS (DSL)","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-58","issueKey":"JCJAMT-58","summary":"BizOps Service Delivery","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Mike Smith (ms8737)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-59","issueKey":"JCJAMT-59","summary":"BizOps CASA - Customer Advocacy & Service AssuranceBizOps SD - Service Delivery","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Mike Smith (ms8737)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-704","issueKey":"JCJAMT-704","summary":"[Configuration] CTX - 147 ARTs","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-705","issueKey":"JCJAMT-705","summary":"[Configuration] CTX - 1092 Teams","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-136","issueKey":"JCJAMT-136","summary":"Test Data Automation Platform","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Shabina Abdulkareem","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-168","issueKey":"JCJAMT-168","summary":"DICA","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Clint Lewis","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-165","issueKey":"JCJAMT-165","summary":"DBA Monitoring","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Olga Pogostkina","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-167","issueKey":"JCJAMT-167","summary":"Application Recovery NFR","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Rafael Benitez","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-61","issueKey":"JCJAMT-61","summary":"Bill Trigger","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-62","issueKey":"JCJAMT-62","summary":"Beyond Now","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-63","issueKey":"JCJAMT-63","summary":"BBW","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-64","issueKey":"JCJAMT-64","summary":"AOSTE","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-65","issueKey":"JCJAMT-65","summary":"ACVP","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-66","issueKey":"JCJAMT-66","summary":"A&BR and CUSTOM PRICING","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ratna Bhargavi Batchu, rb257q","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-67","issueKey":"JCJAMT-67","summary":"A&BR","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ratna Bhargavi Batchu","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-68","issueKey":"JCJAMT-68","summary":"24 Hour Internet","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-69","issueKey":"JCJAMT-69","summary":"(FED) - OMX","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-70","issueKey":"JCJAMT-70","summary":"(FED) - OCX","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-71","issueKey":"JCJAMT-71","summary":"GenAI Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"Abdullah Riaz (ar280d)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-72","issueKey":"JCJAMT-72","summary":"Ntelagent","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Kurt Brubaker","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-73","issueKey":"JCJAMT-73","summary":"WP-WCM","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Kurt Brubaker","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-74","issueKey":"JCJAMT-74","summary":"ASKME","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Shilpa Shetty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-75","issueKey":"JCJAMT-75","summary":"youKnow","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Rujuta Patel","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-76","issueKey":"JCJAMT-76","summary":"Modern Delivery Platform","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Kumar Singisethi","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-114","issueKey":"JCJAMT-114","summary":"ISPNOW","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"KISA, MATUS","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-177","issueKey":"JCJAMT-177","summary":"Gateway as a Service","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Jennifer McLain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-869","issueKey":"JCJAMT-869","summary":"Storage Control Plane","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Jennifer Mclain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-868","issueKey":"JCJAMT-868","summary":"OS Engineering","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Jennifer Mclain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-176","issueKey":"JCJAMT-176","summary":"Application Authorization Framework","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Jennifer McLain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-360","issueKey":"JCJAMT-360","summary":"EASE Platform Architecture","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Nancy Greenwell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1097","issueKey":"JCJAMT-1097","summary":"WKRUAN","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdhelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-362","issueKey":"JCJAMT-362","summary":"DICA","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Clint Lewis","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-359","issueKey":"JCJAMT-359","summary":"EASE SRE","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Natalya Fridman","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1016","issueKey":"JCJAMT-1016","summary":"OS Engineering","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Jennifer McClain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-957","issueKey":"JCJAMT-957","summary":"Hyper Automation Platform","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Jennifer Mclain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-866","issueKey":"JCJAMT-866","summary":"FlexCloud","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Jennifer Mclain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-867","issueKey":"JCJAMT-867","summary":"Kubernetes as a Service","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Jennifer Mclain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-999","issueKey":"JCJAMT-999","summary":"AO-Calc","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Manish Kukreja","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-363","issueKey":"JCJAMT-363","summary":"CSV - DATE","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Shilpa Shetty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-470","issueKey":"JCJAMT-470","summary":"WMS-NT Transformation","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"ROWAN, JOHN","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-471","issueKey":"JCJAMT-471","summary":"Mobility Transformation","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"ROWAN, JOHN (jr1368)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-930","issueKey":"JCJAMT-930","summary":"Finance Internal (FINT)","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga (vs3452)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-948","issueKey":"JCJAMT-948","summary":"External & Legal Affairs (ELA)","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga (vs3452)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-965","issueKey":"JCJAMT-965","summary":"Glanceable Ticket","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Bhavini Patel","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-977","issueKey":"JCJAMT-977","summary":"CDO IDW","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"st639b","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-792","issueKey":"JCJAMT-792","summary":"CSO Third Party Risk Management","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Scott Bickhaus/Aaron Bostick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-781","issueKey":"JCJAMT-781","summary":"CSO Scorecard","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Steve Tenhoor","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-78","issueKey":"JCJAMT-78","summary":"SPT-ATT Mobile App","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ranit Chrust","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-79","issueKey":"JCJAMT-79","summary":"ISAAC","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"MAHMOOD, BUSHRA","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-80","issueKey":"JCJAMT-80","summary":"DOATLAS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Delphine Knaff","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-81","issueKey":"JCJAMT-81","summary":"DTATLAS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Dennis Collins","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-82","issueKey":"JCJAMT-82","summary":"CNIOATLAS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Dennis Collins","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-83","issueKey":"JCJAMT-83","summary":"ServiceNow","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-84","issueKey":"JCJAMT-84","summary":"Mobile Testing Apps (Mtaas)","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-86","issueKey":"JCJAMT-86","summary":"Corporate Systems Payments","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-87","issueKey":"JCJAMT-87","summary":"Network Device Testing","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-88","issueKey":"JCJAMT-88","summary":"Wired Networks","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-89","issueKey":"JCJAMT-89","summary":"Network Analytics","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-90","issueKey":"JCJAMT-90","summary":"Network Operations","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-91","issueKey":"JCJAMT-91","summary":"Application Development","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-92","issueKey":"JCJAMT-92","summary":"Service Delivery","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-93","issueKey":"JCJAMT-93","summary":"Network Data Services","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-94","issueKey":"JCJAMT-94","summary":"Service Provisioning","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-95","issueKey":"JCJAMT-95","summary":"Network - Data Services","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-96","issueKey":"JCJAMT-96","summary":"Mobility Services","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-97","issueKey":"JCJAMT-97","summary":"Network Core Services","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-98","issueKey":"JCJAMT-98","summary":"Network API Monetization","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-99","issueKey":"JCJAMT-99","summary":"Platform Operations - dup space name, get alternate?","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-100","issueKey":"JCJAMT-100","summary":"Integration Services","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-101","issueKey":"JCJAMT-101","summary":"EASE Infrastructure Platform Release Train","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-102","issueKey":"JCJAMT-102","summary":"Field Enablement (ISAAC - WFE)","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-104","issueKey":"JCJAMT-104","summary":"TechDev Europe","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-105","issueKey":"JCJAMT-105","summary":"PRP","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"KAPA FORGACOVA, MIROSLAVA","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-106","issueKey":"JCJAMT-106","summary":"PRC","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"KAPA FORGACOVA, MIROSLAVA (mk1521), Jozef Hrabovecky JH093V","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-107","issueKey":"JCJAMT-107","summary":"Product Inventory","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"ADAMCZYK, ANETA IZABELA (aa759e)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-108","issueKey":"JCJAMT-108","summary":"TDE DSO guild","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-109","issueKey":"JCJAMT-109","summary":"C360","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV (rb393f)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-110","issueKey":"JCJAMT-110","summary":"DRS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"RAKUSANOVA, ALEXANDRA (ar902s)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-111","issueKey":"JCJAMT-111","summary":"SD3 - ASEoD","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"PIVOVARNK, KAMIL","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-112","issueKey":"JCJAMT-112","summary":"NEXTGENDC","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"PIVOVARNK, KAMIL (kp0801)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-113","issueKey":"JCJAMT-113","summary":"TDE QA guild","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-115","issueKey":"JCJAMT-115","summary":"MCU","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"VITIKOVA, ANNA","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-116","issueKey":"JCJAMT-116","summary":"SDWOFS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"ADAMCZYK, ANETA IZABELA (aa759e)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-117","issueKey":"JCJAMT-117","summary":"CIMS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"RAKUSANOVA, ALEXANDRA (ar902s)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-118","issueKey":"JCJAMT-118","summary":"GESS Brno","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"MICHALCAK, JAN","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-119","issueKey":"JCJAMT-119","summary":"ASEOD","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"NEMCOVA, EVA","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-121","issueKey":"JCJAMT-121","summary":"Production Support","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Rakesh Patnaik","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-122","issueKey":"JCJAMT-122","summary":"ATLASNATIV","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Matt Heacock","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-123","issueKey":"JCJAMT-123","summary":"FITTREX","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Amir Alisic","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-124","issueKey":"JCJAMT-124","summary":"Fleet (FLT)","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Angie Fisher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-125","issueKey":"JCJAMT-125","summary":"RAPIDS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Nick Picclocca","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-126","issueKey":"JCJAMT-126","summary":"NetOnePortal","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Nazima Syed","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-127","issueKey":"JCJAMT-127","summary":"CSV - EXCOMM","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bhavneet Tiwana","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-128","issueKey":"JCJAMT-128","summary":"SHIFT","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vanessa Blinder","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-129","issueKey":"JCJAMT-129","summary":"RAPIDUI","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ginny Beyer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-130","issueKey":"JCJAMT-130","summary":"RAPIDSE","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ginny Beyer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-131","issueKey":"JCJAMT-131","summary":"NX - MWL","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vanessa Blinder","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-132","issueKey":"JCJAMT-132","summary":"OPSTOOLS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Nazima Syed","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-133","issueKey":"JCJAMT-133","summary":"FASTWFBPM","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Nazima Syed","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-134","issueKey":"JCJAMT-134","summary":"ERCWDT","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Nazima Syed","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-135","issueKey":"JCJAMT-135","summary":"Continuous Testing & Virtualization","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Shabina Abdulkareem","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-137","issueKey":"JCJAMT-137","summary":"Legacy Workload Scheduler Transformation","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Yisacc Domez; Anna Deen Thomas/David Lambert","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-138","issueKey":"JCJAMT-138","summary":"Messaging Platform Solutions","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Didi Xie; Michael Fafore; Steve Orr; Shweta Dalvi","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-139","issueKey":"JCJAMT-139","summary":"CAR Team","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Corey Gunnell; Joseph Shadwick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-171","issueKey":"JCJAMT-171","summary":"Oracle Utility Server (OUS)","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Joel Thatcher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-142","issueKey":"JCJAMT-142","summary":"NextGEN Portal / ask&GET","tool":"iTrack","droPortfolio":"Green","programArtPoc":"John Drury","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-143","issueKey":"JCJAMT-143","summary":"One Desk Central (ODC)","tool":"iTrack","droPortfolio":"Green","programArtPoc":"John Drury","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-144","issueKey":"JCJAMT-144","summary":"Legacy DB Transformation","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Yisacc Demoz","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-145","issueKey":"JCJAMT-145","summary":"Observability","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Charles Prigmore","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-146","issueKey":"JCJAMT-146","summary":"ContactHub","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV (rb393f)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-147","issueKey":"JCJAMT-147","summary":"Product Relationship Manager","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV (rb393f)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-148","issueKey":"JCJAMT-148","summary":"DBOR Migration & Retirement","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Edward Hernandez","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-149","issueKey":"JCJAMT-149","summary":"RAPID TELEMATICS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bhavneet Tiwana","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-150","issueKey":"JCJAMT-150","summary":"CSV - WMTENT","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bhavneet Tiwana","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-151","issueKey":"JCJAMT-151","summary":"FLEETATO","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bhavneet Tiwana","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-152","issueKey":"JCJAMT-152","summary":"CSV - L1WEBAPPS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bhavneet Tiwana","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-153","issueKey":"JCJAMT-153","summary":"National Scheduling - NEST","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ginny Beyer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-154","issueKey":"JCJAMT-154","summary":"Moose","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"O'NEILL, JENNY","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-155","issueKey":"JCJAMT-155","summary":"CSV - A&R","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"O'NEILL, JENNY","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-156","issueKey":"JCJAMT-156","summary":"CSV - NETOPSWEB","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"O'NEILL, JENNY","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-157","issueKey":"JCJAMT-157","summary":"CAPTOOLS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-158","issueKey":"JCJAMT-158","summary":"PIMMOD","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-159","issueKey":"JCJAMT-159","summary":"CSV - TEOMOD","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-160","issueKey":"JCJAMT-160","summary":"CSV - ABWSD","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-161","issueKey":"JCJAMT-161","summary":"DBOR Migration & Retirementown","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-162","issueKey":"JCJAMT-162","summary":"CSV -MYSOLSCRT","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-163","issueKey":"JCJAMT-163","summary":"CSV - MYSOLNS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-164","issueKey":"JCJAMT-164","summary":"CSV - BSET","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Bart Marmorstone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-166","issueKey":"JCJAMT-166","summary":"Application Recovery","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Rafael Benitez","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-141","issueKey":"JCJAMT-141","summary":"CTX Infra Support (EASE Team 1)","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Kelson Melekottu Kurian","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-169","issueKey":"JCJAMT-169","summary":"DBA Setup","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Joel Thatcher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-170","issueKey":"JCJAMT-170","summary":"AutoSRM- DOD","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Joel Thatcher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-972","issueKey":"JCJAMT-972","summary":"CTX Infra Support (EASE Team 2)","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Kelson Melekottu Kurian","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-172","issueKey":"JCJAMT-172","summary":"Database of Databases","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Joel Thatcher, Olga P","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-173","issueKey":"JCJAMT-173","summary":"SDBA Internal Tools","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Joel Thatcher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-175","issueKey":"JCJAMT-175","summary":"ATTCOMPASS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Tobi Coleman","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-178","issueKey":"JCJAMT-178","summary":"[Fresh Start] CloudRunner-Foundational","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Valarie Littles (vl1055@att.com)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-179","issueKey":"JCJAMT-179","summary":"[Fresh Start] CloudRunner-Enhancing","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Valarie Littles (vl1055@att.com)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-180","issueKey":"JCJAMT-180","summary":"[Fresh Start] CSO Cloud & Data Security – SaaS Security","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Jennifer Gao","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-181","issueKey":"JCJAMT-181","summary":"CSO Security Policy Governance and Engagement Transformation Solutions","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Sarah A Williams","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-182","issueKey":"JCJAMT-182","summary":"CSO Network Hardening","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Patrick Burke","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-183","issueKey":"JCJAMT-183","summary":"CSO Cyber Threat Analytics: Jan – Dec 2026","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Chris Heist","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-184","issueKey":"JCJAMT-184","summary":"[Fresh] CSO - Cloud & Data Security – Cloud Security","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Craig Tope","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-185","issueKey":"JCJAMT-185","summary":"CSO BISO Program: Jan 2026 - Dec 2026","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Sarah A Williams","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-186","issueKey":"JCJAMT-186","summary":"Astra – AI Security Posture Management: Jan-Dec 2026","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Paul Farkas","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-187","issueKey":"JCJAMT-187","summary":"CSO Cyber Threat Operations","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Christina Monteleone","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-335","issueKey":"JCJAMT-335","summary":"CSO Workforce","tool":"ADO Boards","droPortfolio":"Baich","programArtPoc":"Mona Shamma, Kevin Kim","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-336","issueKey":"JCJAMT-336","summary":"TRACE","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Ben Hanley","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-337","issueKey":"JCJAMT-337","summary":"STEM","tool":"ADO Boards","droPortfolio":"Markus","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-338","issueKey":"JCJAMT-338","summary":"SPT-ATT Mobile App","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Ranit Chrust","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-339","issueKey":"JCJAMT-339","summary":"CSV - SmartChat","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Rujuta Patel","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-341","issueKey":"JCJAMT-341","summary":"CSV - Communications and Notifications","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Meena Kshirsagar","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-355","issueKey":"JCJAMT-355","summary":"Enterprise Architecture Office (EAO)- Domain Architecture","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Donald Pericful","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-344","issueKey":"JCJAMT-344","summary":"Hybrid Cloud Enablement","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Azim Kasimov","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1000","issueKey":"JCJAMT-1000","summary":"AI Enablement and Automation (AIEA)","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Prakash Thiruvenkatam","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-346","issueKey":"JCJAMT-346","summary":"GWS Transformation","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Jeannine Sullivan","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-340","issueKey":"JCJAMT-340","summary":"Oracle Service Cloud","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-342","issueKey":"JCJAMT-342","summary":"LevelUP","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"DIANA BINNY, Ohad Penso, Richard Cook, Tal Lahav","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-343","issueKey":"JCJAMT-343","summary":"Legacy HR","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-347","issueKey":"JCJAMT-347","summary":"FirstNet & NG911","tool":"ADO Boards","droPortfolio":"Cohen","programArtPoc":"Heather Lydon (hl5773)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-348","issueKey":"JCJAMT-348","summary":"Field Enablement (DLE - DMP)","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-349","issueKey":"JCJAMT-349","summary":"Field Enablement (DICE)","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-386","issueKey":"JCJAMT-386","summary":"26501-SEI","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Lilach Yaniv (ly359m)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-976","issueKey":"JCJAMT-976","summary":"CDO Core Mainframe","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"st639b","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-978","issueKey":"JCJAMT-978","summary":"CDO EDT Archival","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"st639b","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-969","issueKey":"JCJAMT-969","summary":"[Fresh Start] Billing Tracker","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Aneta Adamczyk (aa759e)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-876","issueKey":"JCJAMT-876","summary":"[Fresh Start] CDO GEN AI Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ar280d & ec898v","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-716","issueKey":"JCJAMT-716","summary":"GWS Legacy","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-646","issueKey":"JCJAMT-646","summary":"5GINA Program - Train A","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-708","issueKey":"JCJAMT-708","summary":"5GINA Program Level Train B","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-351","issueKey":"JCJAMT-351","summary":"Enterprise Architecture Office (EAO)- Transformation Business Assessment","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-352","issueKey":"JCJAMT-352","summary":"Enterprise Architecture Office (EAO)- Taxonomy","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-350","issueKey":"JCJAMT-350","summary":"Enterprise Architecture Office (EAO)- UFD Engagements","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-353","issueKey":"JCJAMT-353","summary":"Enterprise Architecture Office (EAO)- Operational Excellence","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-354","issueKey":"JCJAMT-354","summary":"Enterprise Architecture Office (EAO)- EA Gov Council","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-356","issueKey":"JCJAMT-356","summary":"Enterprise Architecture Office (EAO)- Business Architecture","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Bill Blumberg","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-358","issueKey":"JCJAMT-358","summary":"Enterprise Architecture Office (EAO)- App Tiering","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Maher Abdelhaq","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-361","issueKey":"JCJAMT-361","summary":"Dispatch","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Aharon, Arie","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-364","issueKey":"JCJAMT-364","summary":"Database Architecture & Engineering (DAE)","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Clint Lewis","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-357","issueKey":"JCJAMT-357","summary":"Enterprise Architecture Office (EAO)- AskArchitect","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Lori Mcdaniel","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-365","issueKey":"JCJAMT-365","summary":"Data Analytics & Solutions Hub","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-366","issueKey":"JCJAMT-366","summary":"Customer Workforce","tool":"ADO Boards","droPortfolio":"Baich","programArtPoc":"Mona Shamma, Kevin Kim","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-367","issueKey":"JCJAMT-367","summary":"CSO IAM External","tool":"ADO Boards","droPortfolio":"Baich","programArtPoc":"Mona Shamma, Kevin Kim","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-368","issueKey":"JCJAMT-368","summary":"CTX Total Secure","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Ranit Chrust","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-369","issueKey":"JCJAMT-369","summary":"CSHR & F/WIT ART","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-370","issueKey":"JCJAMT-370","summary":"CSHR & F/Unified ART","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-371","issueKey":"JCJAMT-371","summary":"CSHR & F/Shared Services","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-372","issueKey":"JCJAMT-372","summary":"CSHR & F/L&D ART","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-373","issueKey":"JCJAMT-373","summary":"CSHR & F","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Susan O'Day (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-374","issueKey":"JCJAMT-374","summary":"Cricket","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-375","issueKey":"JCJAMT-375","summary":"COU","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Katie Baxter","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-376","issueKey":"JCJAMT-376","summary":"Corporate HR and Payroll","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-377","issueKey":"JCJAMT-377","summary":"Cloud Foundation","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Xavier Allen, Nimish Buch, Jennifer McLain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-378","issueKey":"JCJAMT-378","summary":"CSO Breach and Attack Simulation (BAS) Project","tool":"ADO Boards","droPortfolio":"Baich","programArtPoc":"Rich Caiati, Daniel Carson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-379","issueKey":"JCJAMT-379","summary":"Azure DevOps Enhancements","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Yehoshuva Arasavelli","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-380","issueKey":"JCJAMT-380","summary":"AWS Cloud Platform Mainframe Migration","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Nancy Greenwell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-381","issueKey":"JCJAMT-381","summary":"Atlas","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Delphine Knaff","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-382","issueKey":"JCJAMT-382","summary":"ACTP","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Ben Hanely bh4695","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-383","issueKey":"JCJAMT-383","summary":"37372-ADAS","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Aleksandra Korkus","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-385","issueKey":"JCJAMT-385","summary":"32629-SSM","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"VULLUM, SOUMYA","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-387","issueKey":"JCJAMT-387","summary":"AFSCP","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-388","issueKey":"JCJAMT-388","summary":"STRMSYND (Streamline Syndicate)","tool":"ATT Jira","droPortfolio":"Cohen","programArtPoc":"Preston Landell (pl603t)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-466","issueKey":"JCJAMT-466","summary":"ESAP","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"SRI TELLAPUR","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-389","issueKey":"JCJAMT-389","summary":"SDCICLOUD (BTPD - NetBond Advanced)","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Lindsey Ghorban, ARIC WALKER, JIGISH SHAH","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-390","issueKey":"JCJAMT-390","summary":"Hyperloop,HYLINC, ACVP","tool":"ATT Jira","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-392","issueKey":"JCJAMT-392","summary":"Hyperloop","tool":"ATT Jira","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-393","issueKey":"JCJAMT-393","summary":"Towers & Poles Transformation","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Liz Pham","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-707","issueKey":"JCJAMT-707","summary":"Digital Leasing","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Liz Pham","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-394","issueKey":"JCJAMT-394","summary":"[Fresh Start] OSS_4_BSSE_CopperTrain","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-395","issueKey":"JCJAMT-395","summary":"DevOps-Orange Train","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Mark Thomson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-415","issueKey":"JCJAMT-415","summary":"WOODLAND","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-399","issueKey":"JCJAMT-399","summary":"Royal Train Circuit Emulation","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Harley Stack & Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-400","issueKey":"JCJAMT-400","summary":"Network Service Assurance AI/ML","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-402","issueKey":"JCJAMT-402","summary":"LAYER3-DEPRECIATION","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-413","issueKey":"JCJAMT-413","summary":"ROYALS TRAIN OCODON","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick, Frank (Joe) Watson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-419","issueKey":"JCJAMT-419","summary":"PHOENIX","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-691","issueKey":"JCJAMT-691","summary":"RAN Automation","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Carsten Lund","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-645","issueKey":"JCJAMT-645","summary":"[Fresh Start] Signaling Transport","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-689","issueKey":"JCJAMT-689","summary":"ASR9K Violet Train","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-660","issueKey":"JCJAMT-660","summary":"[Fresh Start] DCU_Lavender_Train","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-422","issueKey":"JCJAMT-422","summary":"SOLAR TRAIN","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-423","issueKey":"JCJAMT-423","summary":"STELLAR TRAIN","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-418","issueKey":"JCJAMT-418","summary":"DevOps Blue Train","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick, Ben Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-424","issueKey":"JCJAMT-424","summary":"[Fresh Start] AGN Core","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-425","issueKey":"JCJAMT-425","summary":"P2B/P2O","tool":"ATT Jira","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-426","issueKey":"JCJAMT-426","summary":"Consolidated Compliance","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Steve Tenhoor","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-427","issueKey":"JCJAMT-427","summary":"GRP-AXIS","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Daniel Hirata/Rama Vakiti/Satya Grandhi/Madhu Navandar/Srinivasu Appalaneni/Radhikesh Dhakal/Tiep Bui/Tihomir Guentchev/ Russ O'Neal/Neha Dhir/Mitchell Marx/Judy Draper","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-429","issueKey":"JCJAMT-429","summary":"ACCESS-OPS","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Daniel Hirata/Madhu Navandar/Srinivasu Appalaneni","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-428","issueKey":"JCJAMT-428","summary":"GP NEON Operations","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Daniel HirataSrinivasu Appalaneni","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-430","issueKey":"JCJAMT-430","summary":"Hybrid Cloud Enablement","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Aparna YeddulaSampath Ramakrishnappa","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-431","issueKey":"JCJAMT-431","summary":"ISO- FinOps","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Brendan CarlinCamilla Bishop","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-432","issueKey":"JCJAMT-432","summary":"AI & ML Solutions","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Brendan CarlinPranoy Behera","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-433","issueKey":"JCJAMT-433","summary":"Neuron Express","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Lynn Fuhrmann","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-434","issueKey":"JCJAMT-434","summary":"CPUC Pole Management","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Ashish Sethi (as7139)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-435","issueKey":"JCJAMT-435","summary":"Network Operations","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-436","issueKey":"JCJAMT-436","summary":"Project Vector 3049","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-437","issueKey":"JCJAMT-437","summary":"Network Operations (BUSINESS)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-438","issueKey":"JCJAMT-438","summary":"Formation Tool","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-439","issueKey":"JCJAMT-439","summary":"IAMS","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Lohit Ananda","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-440","issueKey":"JCJAMT-440","summary":"Managed Services","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-441","issueKey":"JCJAMT-441","summary":"SDN Suite","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-442","issueKey":"JCJAMT-442","summary":"GenAI","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-443","issueKey":"JCJAMT-443","summary":"SDN Controller Platform","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-444","issueKey":"JCJAMT-444","summary":"Field Enablement (Canopi)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-445","issueKey":"JCJAMT-445","summary":"Service Delivery","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-446","issueKey":"JCJAMT-446","summary":"Finshed Goods Suite of Applications","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-447","issueKey":"JCJAMT-447","summary":"Network Service Assurance","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-448","issueKey":"JCJAMT-448","summary":"Network Systems - Resliancy Engineering Services","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-449","issueKey":"JCJAMT-449","summary":"In house Applications - Mitch Gunnels","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-450","issueKey":"JCJAMT-450","summary":"Cloud Network Operations","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-451","issueKey":"JCJAMT-451","summary":"OVALS","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-452","issueKey":"JCJAMT-452","summary":"Automation Platforms Development","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Naresh Mukkara","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-454","issueKey":"JCJAMT-454","summary":"Ticket2Ride","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"RAKUSANOVA, ALEXANDRA (ar902s)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-455","issueKey":"JCJAMT-455","summary":"RKG","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"KESEG, RADIM","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-456","issueKey":"JCJAMT-456","summary":"WESPHR, STSPHR, SSSPHR, SPHRLANCER, MCMRN, FSSPHR","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"TABORSKY, VOJTECH (vt2210)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-457","issueKey":"JCJAMT-457","summary":"UMNAST, SSSPHR, SPHRLANCER.SHSPHR, OPNAST, MRNAST, KMRN","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"GAAL, MATUS (mg7002)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-458","issueKey":"JCJAMT-458","summary":"SPOJIT","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"JANOV, RICHARD (rj9872)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-459","issueKey":"JCJAMT-459","summary":"AI Foundry","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"PALUBOVA, KATARINA (kp0706)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-460","issueKey":"JCJAMT-460","summary":"NST- Wireline","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Joseph Croce","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-461","issueKey":"JCJAMT-461","summary":"NST-Fiber","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Joseph Croce","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-462","issueKey":"JCJAMT-462","summary":"NST-FPP","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Joseph Croce","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-463","issueKey":"JCJAMT-463","summary":"[Fresh Start] Cloud Crowd","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Jessica Eggers","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-464","issueKey":"JCJAMT-464","summary":"QFAC","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Jessica Eggers","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-465","issueKey":"JCJAMT-465","summary":"MCEP","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"SRI TELLAPUR","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-467","issueKey":"JCJAMT-467","summary":"NST-Wireless","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Ashish Sethi (as7139)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-469","issueKey":"JCJAMT-469","summary":"NST Wireless (NSTWS)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"SRI TELLAPUR","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-468","issueKey":"JCJAMT-468","summary":"NST E&O","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"SRI TELLAPUR","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-473","issueKey":"JCJAMT-473","summary":"Ticketing Cinema (ATTNOWIA)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"RAMACHANDRAN, BHARATHIMUTHU (br0549)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-474","issueKey":"JCJAMT-474","summary":"ABS Ticketing Transformation (Ticket2Ride)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"ROWAN, JOHN (jr1368)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-477","issueKey":"JCJAMT-477","summary":"CDO ATT Business Solutions","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"gp0225 & ma4386","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-476","issueKey":"JCJAMT-476","summary":"CDO Finance Transformation","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"mr2421","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-480","issueKey":"JCJAMT-480","summary":"TSS Now","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Karina Montatsky","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-478","issueKey":"JCJAMT-478","summary":"CDO Data Products","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"sp8226","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-823","issueKey":"JCJAMT-823","summary":"Wireline","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"RS2089/NK9543/AS328N/PM971A/AA2418","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-11","issueKey":"JCJAMT-11","summary":"Cloud Crowd - Blue","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Susana Barazza (se1363)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-482","issueKey":"JCJAMT-482","summary":"CDO Network & Field Ops","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"Nidhi Kapoor (nk1628)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-475","issueKey":"JCJAMT-475","summary":"CDO Quickstrike","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"rs1015","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-487","issueKey":"JCJAMT-487","summary":"Hybrid Cloud Connectivity","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Jennifer McLainNimish BuchJoseph Kufert","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-491","issueKey":"JCJAMT-491","summary":"1/2 DBOR Migration & Retirement","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Shashi Shanker (ss150g)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-479","issueKey":"JCJAMT-479","summary":"CDO Datalock","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"fz045f","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-481","issueKey":"JCJAMT-481","summary":"AT&T IT Private Cloud Transformation","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Aaron OfosuheneYanko Torralba","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-483","issueKey":"JCJAMT-483","summary":"Network Outage Management (NOM)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Maryann Germita (mg7923), Chun Jin (cj0519)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-485","issueKey":"JCJAMT-485","summary":"End-to-End Incident Management (EEIM)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Maryann Germita","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-488","issueKey":"JCJAMT-488","summary":"[Fresh Start] ISO Operations & Outage","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Jennifer McLainJoseph KufertJoseph Ortbals","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-489","issueKey":"JCJAMT-489","summary":"[Fresh Start] Fleet Management","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Jennifer McLainJustin MastersonJoseph Kufert","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-490","issueKey":"JCJAMT-490","summary":"[Fresh Start] ISO Database & Middleware","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Murali VasudevanJennifer MclainJoseph Kufert","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-486","issueKey":"JCJAMT-486","summary":"Network Hydration","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Tina Uy","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-641","issueKey":"JCJAMT-641","summary":"MCGW - Mobility and Cloud Gateway","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-643","issueKey":"JCJAMT-643","summary":"New Project One","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick, Lon Tanner and Jeong Min","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-648","issueKey":"JCJAMT-648","summary":"AINLM2","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-656","issueKey":"JCJAMT-656","summary":"DMP Product","tool":"ATT Jira","droPortfolio":"Cohen","programArtPoc":"KIESLING, ABBY ak8286; EVANS, NIKA de8176@att.com","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-658","issueKey":"JCJAMT-658","summary":"BTPD - GESS","tool":"ATT Jira","droPortfolio":"Cohen","programArtPoc":"Stephen Cardoza","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-659","issueKey":"JCJAMT-659","summary":"GTA","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Prudhvi Adapa, Aromal Suresh","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-668","issueKey":"JCJAMT-668","summary":"NCaaS Platform Capability","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Prudhvi Adapa/Aromal Suresh","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-670","issueKey":"JCJAMT-670","summary":"Network E2E","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-672","issueKey":"JCJAMT-672","summary":"[Fresh Start] FF37 - Neuron Express","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-681","issueKey":"JCJAMT-681","summary":"[Fresh Start] Mobility Core Capacity Optimizer","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Amita Vijay","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-684","issueKey":"JCJAMT-684","summary":"Royal Train - IOC","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-687","issueKey":"JCJAMT-687","summary":"MoW Access","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Corinna Van Der Veen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-692","issueKey":"JCJAMT-692","summary":"Legacy Transformation:  Labs","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-694","issueKey":"JCJAMT-694","summary":"Legacy DB Transformation","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Yisacc Demoz","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-695","issueKey":"JCJAMT-695","summary":"Fresh Start:  Connected Communities ToT with CC Team 1","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Tobi Coleman (tc7380)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-713","issueKey":"JCJAMT-713","summary":"[Fresh Start] - Remote Access-VPN: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Andrew Sengputa","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-710","issueKey":"JCJAMT-710","summary":"[Fresh Start] - Tanium for Secure The Labs: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Gloria Wang","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-709","issueKey":"JCJAMT-709","summary":"[Fresh Start] - Data Leakage Protection (DLP): Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Gloria Wang","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-698","issueKey":"JCJAMT-698","summary":"Flywheel","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"CHRISTIAN WALLACE (cw7047)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-699","issueKey":"JCJAMT-699","summary":"Access Hardware Replacement (AHR)","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"JONATHAN R SMITH (js2009)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-700","issueKey":"JCJAMT-700","summary":"Wireline Growth","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Michael Macintire","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-701","issueKey":"JCJAMT-701","summary":"Hybrid Cloud Connectivity","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Nimish Buch/Jennifer McClain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-702","issueKey":"JCJAMT-702","summary":"Hybrid Cloud Enablement","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-703","issueKey":"JCJAMT-703","summary":"[Fresh Start] Fleet Management- 2 additional teams","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Jennifer McLainJustin MastersonJoseph Kufert","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-706","issueKey":"JCJAMT-706","summary":"ACS 360","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-712","issueKey":"JCJAMT-712","summary":"[Fresh Start] - Corporate Wired NAC: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Rich Caiati","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-711","issueKey":"JCJAMT-711","summary":"[Fresh Start] - Mobile Device Security: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Danielle Repsher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-715","issueKey":"JCJAMT-715","summary":"CTX - ADO Strategy Teams","tool":"ADO Strategy","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-717","issueKey":"JCJAMT-717","summary":"Baich - Initial Fresh Start Team","tool":"Initial Fresh Start","droPortfolio":"Baich","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-718","issueKey":"JCJAMT-718","summary":"Cohen - Initial Fresh Start","tool":"Initial Fresh Start","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-719","issueKey":"JCJAMT-719","summary":"Elbaz - Initial Fresh Start","tool":"Initial Fresh Start","droPortfolio":"Elbaz","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-720","issueKey":"JCJAMT-720","summary":"Markus - Initial Fresh Start","tool":"Initial Fresh Start","droPortfolio":"Markus","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-721","issueKey":"JCJAMT-721","summary":"Press - Initial Fresh Start","tool":"Initial Fresh Start","droPortfolio":"Press","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-722","issueKey":"JCJAMT-722","summary":"Zilberstein - Initial Fresh Start","tool":"Initial Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-724","issueKey":"JCJAMT-724","summary":"UMNAST, OPNAST, MRNAST, KMRN","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"HOSSAIN, MOHAMMED (mh8989)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-725","issueKey":"JCJAMT-725","summary":"Zilberstein - ServiceNow","tool":"ServiceNow","droPortfolio":"Zilberstain","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-726","issueKey":"JCJAMT-726","summary":"Cohen - LeanKit","tool":"LeanKit","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-727","issueKey":"JCJAMT-727","summary":"Zilberstein - LeanKit","tool":"LeanKit","droPortfolio":"Zilberstain","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-728","issueKey":"JCJAMT-728","summary":"CDO Credit & Collections","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"hw1591","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-730","issueKey":"JCJAMT-730","summary":"CTX Infra Support","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Kelson Melekottu Kurian/Yissacc Demoz","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-731","issueKey":"JCJAMT-731","summary":"UBMCAT1 (UBM18653)","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Atul Moghe (am9262)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-732","issueKey":"JCJAMT-732","summary":"(FED) - GIOM","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-733","issueKey":"JCJAMT-733","summary":"2026 R3 Deprecation","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"dv6464","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-734","issueKey":"JCJAMT-734","summary":"OWS & Wholesale Billing and vendor governance","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"---","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-735","issueKey":"JCJAMT-735","summary":"ARIS","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"ee8291, mm6716","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-736","issueKey":"JCJAMT-736","summary":"HALO Deprecation - SAREA","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"dv6464","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-737","issueKey":"JCJAMT-737","summary":"FirstNet Security Compliance","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Erin Giesy","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-738","issueKey":"JCJAMT-738","summary":"FirstNet & NG911","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Sorin Netu","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-739","issueKey":"JCJAMT-739","summary":"24 Hour Internet","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Baviya Leelavinothan","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-740","issueKey":"JCJAMT-740","summary":"Network Service Assurance Systems (NSAS)","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Caridad Alejandro, SUBRAHMANYAM Jonnavittula","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-741","issueKey":"JCJAMT-741","summary":"Network Service Assurance Systems (NSAS)","tool":"Fresh Start","droPortfolio":"","programArtPoc":"SUBRAHMANYAM Jonnavittula","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-742","issueKey":"JCJAMT-742","summary":"Granite NIS Transformation","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Caridad Alejandro","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-743","issueKey":"JCJAMT-743","summary":"Nucleus","tool":"Fresh Start","droPortfolio":"","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-744","issueKey":"JCJAMT-744","summary":"Enterprise Release Management","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Nicole Slay Caldwell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-745","issueKey":"JCJAMT-745","summary":"NA","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Nancy Greenwell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-746","issueKey":"JCJAMT-746","summary":"Fleet Management","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Jennifer McLain, Justin Masterson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-747","issueKey":"JCJAMT-747","summary":"Fleet Management","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Jennifer McLainJustin Masterson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-748","issueKey":"JCJAMT-748","summary":"Fleet Management","tool":"Fresh Start","droPortfolio":"Green","programArtPoc":"Jennifer McLainJustin Masterson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-749","issueKey":"JCJAMT-749","summary":"N/A","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"FIRDAUS, NAILA (nf5538), SHEIKH, SOHAIB (ss156e), ASSASSI, YOUNESS (ya2257)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-750","issueKey":"JCJAMT-750","summary":"N/A","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"SHEIKH, SOHAIB (ss156e)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-751","issueKey":"JCJAMT-751","summary":"N/A","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"ASSASSI, YOUNESS (ya2257)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-752","issueKey":"JCJAMT-752","summary":"Field Enablement (TRIP)","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-753","issueKey":"JCJAMT-753","summary":"uDAS","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-754","issueKey":"JCJAMT-754","summary":"Field Enablement (ATLAS)","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-755","issueKey":"JCJAMT-755","summary":"Order fulfilment","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-756","issueKey":"JCJAMT-756","summary":"SAR / ICR","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-757","issueKey":"JCJAMT-757","summary":"CCP2 Apps/ITL","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-759","issueKey":"JCJAMT-759","summary":"Manual Device Testing","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-760","issueKey":"JCJAMT-760","summary":"Testing Automation","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-761","issueKey":"JCJAMT-761","summary":"Application Development","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-762","issueKey":"JCJAMT-762","summary":"Platform Operations - dup space name, get alternate?","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-763","issueKey":"JCJAMT-763","summary":"Platform Operations - dup space name, get alternate?","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-764","issueKey":"JCJAMT-764","summary":"Platform Operations - dup space name, get alternate?","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-765","issueKey":"JCJAMT-765","summary":"Platform Operations - dup space name, get alternate?","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-766","issueKey":"JCJAMT-766","summary":"Platform Operations - dup space name, get alternate?","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-783","issueKey":"JCJAMT-783","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-784","issueKey":"JCJAMT-784","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-785","issueKey":"JCJAMT-785","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-786","issueKey":"JCJAMT-786","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-787","issueKey":"JCJAMT-787","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-788","issueKey":"JCJAMT-788","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-789","issueKey":"JCJAMT-789","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-790","issueKey":"JCJAMT-790","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-791","issueKey":"JCJAMT-791","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-793","issueKey":"JCJAMT-793","summary":"FirstNet Security Compliance","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Erin Giesy","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-794","issueKey":"JCJAMT-794","summary":"Quantum Readiness and AI Innovation: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Jason Baird","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-795","issueKey":"JCJAMT-795","summary":"Quantum Readiness and AI Innovation: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Jason Baird","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-796","issueKey":"JCJAMT-796","summary":"CSRM 2026 Remediation-Audit-Architecture","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Jason Baird","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-797","issueKey":"JCJAMT-797","summary":"CSRM 2026 Remediation-Audit-Architecture","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Jason Baird","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-798","issueKey":"JCJAMT-798","summary":"CSRM 2026 Remediation-Audit-Architecture","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Jason Baird","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-799","issueKey":"JCJAMT-799","summary":"DSPM Platform Automation: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Mohan Periyasamy","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-802","issueKey":"JCJAMT-802","summary":"Data Leakage Protection (DLP): Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Yuliia Kolpakova","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-803","issueKey":"JCJAMT-803","summary":"Tanium for Secure The Labs: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Gloria Wang","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-804","issueKey":"JCJAMT-804","summary":"EDR/EPP – SentinelOne, Microsegmentation, Email Security & Firewall (FW) Operations: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Vinod Krishnankutty, Collin Quinn (Microseg), Sushil Rajput (Email), Gloria Wang (Firewall)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-805","issueKey":"JCJAMT-805","summary":"EDR/EPP – SentinelOne, Microsegmentation, Email Security & Firewall (FW) Operations: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Collin Quinn (Microseg)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-806","issueKey":"JCJAMT-806","summary":"EDR/EPP – SentinelOne, Microsegmentation, Email Security & Firewall (FW) Operations: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Sushil Rajput (Email)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-807","issueKey":"JCJAMT-807","summary":"EDR/EPP – SentinelOne, Microsegmentation, Email Security & Firewall (FW) Operations: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Gloria Wang (Firewall)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-808","issueKey":"JCJAMT-808","summary":"Mobile Device Security: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Danielle Repsher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-809","issueKey":"JCJAMT-809","summary":"Corporate Wired NAC: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Rich Caiati","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-810","issueKey":"JCJAMT-810","summary":"Remote Access-VPN: Jan-Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Andrew Sengputa","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-811","issueKey":"JCJAMT-811","summary":"CSO Cyber Threat Intelligence Operations","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Rajni Pandey","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-812","issueKey":"JCJAMT-812","summary":"CSO Cyber Threat Operations: Jan – Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Terrie Myerchin","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-813","issueKey":"JCJAMT-813","summary":"CSO Cyber Threat Operations: Jan – Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Alendian San","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-814","issueKey":"JCJAMT-814","summary":"CSO Cyber Threat Operations: Jan – Dec 2026","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Bill Storvoll","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-816","issueKey":"JCJAMT-816","summary":"CSO Akamai Security","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Patrick Burke, Stewart Darby","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-817","issueKey":"JCJAMT-817","summary":"Workforce","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Mona ShammaKevin Kim","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-818","issueKey":"JCJAMT-818","summary":"Consolidated Compliance","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Steve Tenhoor","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-819","issueKey":"JCJAMT-819","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-820","issueKey":"JCJAMT-820","summary":"P2B/P2O","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Pedro De Jesus/Lenajeannette Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-821","issueKey":"JCJAMT-821","summary":"ACPDSI-NRI-APISENTRY","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-822","issueKey":"JCJAMT-822","summary":"Integration Platforms","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-825","issueKey":"JCJAMT-825","summary":"5GINA","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-826","issueKey":"JCJAMT-826","summary":"Advanced Connectivity Platform","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-827","issueKey":"JCJAMT-827","summary":"ASE-ASEoD D1 CANOPI WIRELINE","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-828","issueKey":"JCJAMT-828","summary":"Bronze Train","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1039","issueKey":"JCJAMT-1039","summary":"BICON","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Abhishek Fnu","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-829","issueKey":"JCJAMT-829","summary":"Cloud Native Ericsson Network Manager","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Patrick Taylor","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-830","issueKey":"JCJAMT-830","summary":"IPLAT","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-831","issueKey":"JCJAMT-831","summary":"Orange Train","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Mark Thomson","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-832","issueKey":"JCJAMT-832","summary":"TBD - Culliton","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-833","issueKey":"JCJAMT-833","summary":"TBD - Culliton","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-834","issueKey":"JCJAMT-834","summary":"TBD - Culliton","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-835","issueKey":"JCJAMT-835","summary":"TBD - Culliton","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-836","issueKey":"JCJAMT-836","summary":"TBD-Culliton","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-837","issueKey":"JCJAMT-837","summary":"Hyperloop - ServiceNow","tool":"ServiceNow","droPortfolio":"Zilberstain","programArtPoc":"Vasundhara Akella","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-838","issueKey":"JCJAMT-838","summary":"Network Systems Provisioning Transformation","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Mark Thomson (mt527n)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-839","issueKey":"JCJAMT-839","summary":"New team and board request for AT&T IT Private Cloud Transformation program","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Aaron Ofosuhene","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-840","issueKey":"JCJAMT-840","summary":"[Fresh Start] Geo Modeler","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-841","issueKey":"JCJAMT-841","summary":"RAN-EIAP","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Carsten Lund","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-842","issueKey":"JCJAMT-842","summary":"RTTP","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Steven Hellstern","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-843","issueKey":"JCJAMT-843","summary":"Video Optimizer","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine and Lihsin Shih","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-844","issueKey":"JCJAMT-844","summary":"SLA Credit outage","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-845","issueKey":"JCJAMT-845","summary":"Deep Network Integration","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Andrea Mosca and Pearl Shreemal","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-846","issueKey":"JCJAMT-846","summary":"[CTX Migration] Wave 1: 40k Features","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-847","issueKey":"JCJAMT-847","summary":"[CTX Migration] Wave 2: 28k Stories","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-848","issueKey":"JCJAMT-848","summary":"[CTX Migration] Wave 3: 127k Stories","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-849","issueKey":"JCJAMT-849","summary":"[CTX Migration] Wave 4: 107k Stories","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-850","issueKey":"JCJAMT-850","summary":"[CTX Migration] Wave 5: 120k Stories and Comments","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-851","issueKey":"JCJAMT-851","summary":"SDBA Engineering","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Kevin Curameng","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-60","issueKey":"JCJAMT-60","summary":"BizOps Technology Services","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Mike Smith (ms8737)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-484","issueKey":"JCJAMT-484","summary":"Network Outage Management (NOM)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Maryann Germita","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-852","issueKey":"JCJAMT-852","summary":"Oracle DBA Windows","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Kevin Curameng","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-853","issueKey":"JCJAMT-853","summary":"GWS Transformation 2","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (AA158A)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-854","issueKey":"JCJAMT-854","summary":"L&D ART","tool":"","droPortfolio":"Zilberstain","programArtPoc":"SUSAN J O'DAY (sm6725)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-855","issueKey":"JCJAMT-855","summary":"CDO EEIM","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"nk1628","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-857","issueKey":"JCJAMT-857","summary":"Kur","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Kurt Brubaker (kb1945)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-858","issueKey":"JCJAMT-858","summary":"CDO Network & Field Ops","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"nk1628","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-859","issueKey":"JCJAMT-859","summary":"NEXTGENDC","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Kamil Pivovrnk","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-862","issueKey":"JCJAMT-862","summary":"OPSTOOLS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Jared Walton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-863","issueKey":"JCJAMT-863","summary":"AFOAP","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Micah Kralik","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-864","issueKey":"JCJAMT-864","summary":"Fiber Broadband","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-871","issueKey":"JCJAMT-871","summary":"Workflow test","tool":"","droPortfolio":"","programArtPoc":"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-872","issueKey":"JCJAMT-872","summary":"PI1 GCM Modernization (Wireless)","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Liz Pham lp182y@att.com (Liz P)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-873","issueKey":"JCJAMT-873","summary":"PI1 GCM Modernization (Wireline)","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Liz Pham (lp182y@att.com)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-875","issueKey":"JCJAMT-875","summary":"CDO EmpowHR Transformation","tool":"ADO Boards","droPortfolio":"Markus","programArtPoc":"sx2059 & aw0392","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-877","issueKey":"JCJAMT-877","summary":"CDO Testing Data & Rewards","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"dr4317","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-998","issueKey":"JCJAMT-998","summary":"CDO Snowflake Platform","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"jv1325 & ab821w","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-878","issueKey":"JCJAMT-878","summary":"CDO E2E","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dr4317","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-879","issueKey":"JCJAMT-879","summary":"CDO Cloud Services & Datalake Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"jv1325 & lm2430","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-897","issueKey":"JCJAMT-897","summary":"CDO Corporate","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"aw0392","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-880","issueKey":"JCJAMT-880","summary":"CDO Platform Architecture","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"UR9885","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-881","issueKey":"JCJAMT-881","summary":"CDO AI Governance Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"jv1325 & lm2430","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-882","issueKey":"JCJAMT-882","summary":"CDO Marketing AI","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"pb614a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-883","issueKey":"JCJAMT-883","summary":"CDO Connected Solutions","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"pb614a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-884","issueKey":"JCJAMT-884","summary":"CDO Pricing Engine","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"pb614a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-885","issueKey":"JCJAMT-885","summary":"CDO DSAIR","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"ec898v","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-886","issueKey":"JCJAMT-886","summary":"CDO Competitive Intelligence","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"vg4695","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-888","issueKey":"JCJAMT-888","summary":"CDO Document AI","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ec898v","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-889","issueKey":"JCJAMT-889","summary":"CDO ABS & Consumer Proserve","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"Ll889w & pb614a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-890","issueKey":"JCJAMT-890","summary":"CDO Ask Graph","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ar280d","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-891","issueKey":"JCJAMT-891","summary":"CDO Operational Excellence (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"nk1628","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-892","issueKey":"JCJAMT-892","summary":"CDO Intelligent Automation","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"gb1555","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-893","issueKey":"JCJAMT-893","summary":"CDO iQ Insights","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"nk1628","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-894","issueKey":"JCJAMT-894","summary":"CDO Mass Markets Care and Retail","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"as005d","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-895","issueKey":"JCJAMT-895","summary":"CDO Care & MM Transformation","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"as005d","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-896","issueKey":"JCJAMT-896","summary":"CDO FirstNet GenAI","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"as005d","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-898","issueKey":"JCJAMT-898","summary":"CDO GoldenGate","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"as005d","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-899","issueKey":"JCJAMT-899","summary":"CDO DNA","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"cw0078","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-900","issueKey":"JCJAMT-900","summary":"CDO Tax (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-901","issueKey":"JCJAMT-901","summary":"CDO MARS","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"cw0078","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-902","issueKey":"JCJAMT-902","summary":"CDO AART","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-903","issueKey":"JCJAMT-903","summary":"CDO RETFOR","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-904","issueKey":"JCJAMT-904","summary":"CDO ASIST","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-905","issueKey":"JCJAMT-905","summary":"CDO CONAI (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"cw0078","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-906","issueKey":"JCJAMT-906","summary":"CDO BIRT","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-907","issueKey":"JCJAMT-907","summary":"CDO CDE","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-908","issueKey":"JCJAMT-908","summary":"CDO IHX","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-909","issueKey":"JCJAMT-909","summary":"CDO BOP (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-910","issueKey":"JCJAMT-910","summary":"CDO SCAMPWEB","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dg777t","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-911","issueKey":"JCJAMT-911","summary":"CDO USAGE","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dg777t","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-912","issueKey":"JCJAMT-912","summary":"CDO DATAWEB","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dg777t","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-913","issueKey":"JCJAMT-913","summary":"CDO LDBoR","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dg777t","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-914","issueKey":"JCJAMT-914","summary":"CDO mASCERT (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dd160g","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-915","issueKey":"JCJAMT-915","summary":"CDO Legal Demand","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"gd2167","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-916","issueKey":"JCJAMT-916","summary":"CDO Snowflake ANT","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"jh2097","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-917","issueKey":"JCJAMT-917","summary":"CDO Prod Support (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"mb9322","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-918","issueKey":"JCJAMT-918","summary":"CDO DataMgr","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"bb1203","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-919","issueKey":"JCJAMT-919","summary":"CDO NULD","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"sl761m","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-920","issueKey":"JCJAMT-920","summary":"CDO RANWEB","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"sl761m","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-921","issueKey":"JCJAMT-921","summary":"CDO MXLegal Demands","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"sl761m","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-922","issueKey":"JCJAMT-922","summary":"Dynamic Defense","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Pavel Hruban","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-696","issueKey":"JCJAMT-696","summary":"OSDF IOS Depreciation","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"NITIN RAYKAR (nr009n)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-923","issueKey":"JCJAMT-923","summary":"[Fresh Start] CloudRunner-AI","tool":"Fresh Start","droPortfolio":"Press","programArtPoc":"Valarie Littles (vl1055@att.com)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-924","issueKey":"JCJAMT-924","summary":"ASEoD Oversubscription","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Susan Bond","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-925","issueKey":"JCJAMT-925","summary":"BTPD Engineering Platform Capabilities Program","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"BELOVIC, RADOSLAV (rb393f)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-926","issueKey":"JCJAMT-926","summary":"BTPD - Core Platform Common Capabilities ART","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"BELOVIC, RADOSLAV (rb393f)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1057","issueKey":"JCJAMT-1057","summary":"CDO Finance Reporting and Tools","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"pm8651","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-927","issueKey":"JCJAMT-927","summary":"TicketingOps","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (aa158a)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-928","issueKey":"JCJAMT-928","summary":"CDO MFR","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"ss6774","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-952","issueKey":"JCJAMT-952","summary":"New board request for  ISO Database & Middleware program","tool":"ATT Jira","droPortfolio":"Green","programArtPoc":"Murali Vasuvaden/Kurian, Kelson Melekoottu","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-953","issueKey":"JCJAMT-953","summary":"MVNO Modernization","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-954","issueKey":"JCJAMT-954","summary":"NCaaS Platform Capability","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Prudhvi Adapa/Aromal Suresh/Lalena Aria","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-955","issueKey":"JCJAMT-955","summary":"[Fresh Start] ACI_OSS_FIFA_TRAIN","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Tyrel Fitzpatrick","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-956","issueKey":"JCJAMT-956","summary":"Artemis-DevOps","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Prudhvi Adapa and Aromal Suresh","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-958","issueKey":"JCJAMT-958","summary":"BTPD - Business Customer Care","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Rod Simms, Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-959","issueKey":"JCJAMT-959","summary":"BTPD - Converged Experience","tool":"ADO Boards","droPortfolio":"Cohen","programArtPoc":"Kori Antuna","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-960","issueKey":"JCJAMT-960","summary":"BTPD - Wireline Connectivity","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Liezle Andico","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-961","issueKey":"JCJAMT-961","summary":"Wireless Transformation","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Bridget Smtih (bs0634)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-962","issueKey":"JCJAMT-962","summary":"BTPD - Business Development","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Jennifer Rosetta/Kris Klippel","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-963","issueKey":"JCJAMT-963","summary":"CSO CCDR","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Steve Tenhoor","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-77","issueKey":"JCJAMT-77","summary":"COMO","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Ginny Beyer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-391","issueKey":"JCJAMT-391","summary":"Hyperloop,HYLINC","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Sridhar Reddy","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-472","issueKey":"JCJAMT-472","summary":"Billing Dispute Assistant Program Board (GENAI)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"PATEL, BHAVINI (bp1348)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-420","issueKey":"JCJAMT-420","summary":"BBNMS-LS Program","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Prudhvi Adapa","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-174","issueKey":"JCJAMT-174","summary":"TITAN","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Tobi Coleman","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-860","issueKey":"JCJAMT-860","summary":"CDO BDMSP","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"fn050f & rw636j","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-729","issueKey":"JCJAMT-729","summary":"CDO SAM","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"nr7327 & mm616f","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-865","issueKey":"JCJAMT-865","summary":"CDO Ingest & Transform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lf7763","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-856","issueKey":"JCJAMT-856","summary":"CDO Cloud Services & Datalake Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"jv1325","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-887","issueKey":"JCJAMT-887","summary":"CDO Mobility Promotions","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ak0397","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-968","issueKey":"JCJAMT-968","summary":"Integration Platforms","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Culliton","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-970","issueKey":"JCJAMT-970","summary":"CDO Mobility Promotions v2","tool":"ATT Jira","droPortfolio":"Markus","programArtPoc":"ak0397","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-971","issueKey":"JCJAMT-971","summary":"CDO Finance Data Warehouse","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"gv2931","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-973","issueKey":"JCJAMT-973","summary":"DOMS","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Akella, Vasundhara","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-974","issueKey":"JCJAMT-974","summary":"NOVA NG","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Almog, Ela (ae1883)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-975","issueKey":"JCJAMT-975","summary":"NOVA NG","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"NOVA NG","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-980","issueKey":"JCJAMT-980","summary":"Tech M Middleware Delivery Leadership","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Shilpa Dutt","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-981","issueKey":"JCJAMT-981","summary":"Enhanced IP Flex Features Server","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Ashokkumar Mundru & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-982","issueKey":"JCJAMT-982","summary":"RAINBOW","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-983","issueKey":"JCJAMT-983","summary":"APISENTRY","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Anantha Kondaparthi & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-984","issueKey":"JCJAMT-984","summary":"Service Provisioning Platform","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Himagirish Hulamani & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-985","issueKey":"JCJAMT-985","summary":"SPP-DAS","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Himagirish Hulamani & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-986","issueKey":"JCJAMT-986","summary":"virtualized Network Services Portal","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Himagirish Hulamani & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-987","issueKey":"JCJAMT-987","summary":"Virtual International Routing Consolidation","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Veer Rajasekhar & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-988","issueKey":"JCJAMT-988","summary":"NEAM","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Anantha Kondaparthi & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1026","issueKey":"JCJAMT-1026","summary":"MCANS - JA only - no JC or connector","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Morgan Beaumier/Lena Camps/Branden Haisleys","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-989","issueKey":"JCJAMT-989","summary":"Service Activation System - Voice Provisioning","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Michelle Weickert & Randa Maher","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-990","issueKey":"JCJAMT-990","summary":"CSTEM Support","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Lorrie Joeschke","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-991","issueKey":"JCJAMT-991","summary":"AI Doctor","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Rafael Benitez","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-874","issueKey":"JCJAMT-874","summary":"CDO Sales & Marketing Data Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"kg1102 & rb9777 & aw146a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-120","issueKey":"JCJAMT-120","summary":"ARIA (Business Voice)","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Milan Guman","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-345","issueKey":"JCJAMT-345","summary":"Hybrid Cloud Connectivity","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Jennifer McClain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-384","issueKey":"JCJAMT-384","summary":"33890-Developer Experience","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Yaniv, Lilach","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-782","issueKey":"JCJAMT-782","summary":"CSO Permit Program","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Phil Campos/Pedro De Jesus","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-815","issueKey":"JCJAMT-815","summary":"CSO Edge Security","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Patrick Burke","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-800","issueKey":"JCJAMT-800","summary":"CSO Enterprise Cryptography & Secrets Platforms","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Ali Saqib","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-453","issueKey":"JCJAMT-453","summary":"OVALST (INLAP)","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Deepti Verma","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-992","issueKey":"JCJAMT-992","summary":"DOMS","tool":"ATT Jira","droPortfolio":"Zilberstain","programArtPoc":"Sudha L","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-993","issueKey":"JCJAMT-993","summary":"Platform Capabilities","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Hemanth Narayana","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-994","issueKey":"JCJAMT-994","summary":"IPTF Request Management","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Yamuna Ashok","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-995","issueKey":"JCJAMT-995","summary":"Mobile Applications (MCOE)","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga (vs3452)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-996","issueKey":"JCJAMT-996","summary":"Treasury-Payments","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga (vs3452)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-997","issueKey":"JCJAMT-997","summary":"CDO Core Rep","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"st639b","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1001","issueKey":"JCJAMT-1001","summary":"CSI CPSVC","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Shilpa Dutt","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1101","issueKey":"JCJAMT-1101","summary":"PINC","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sharon Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1002","issueKey":"JCJAMT-1002","summary":"LOCUS","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Andrew Eberhart","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1003","issueKey":"JCJAMT-1003","summary":"Converged Network Resource Orchestrator","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sharon Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1004","issueKey":"JCJAMT-1004","summary":"Business Mobility 5G","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1005","issueKey":"JCJAMT-1005","summary":"AIAB","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1006","issueKey":"JCJAMT-1006","summary":"Business Mobility KTLO","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1007","issueKey":"JCJAMT-1007","summary":"OWS Wholesale","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1008","issueKey":"JCJAMT-1008","summary":"ROME Deprecation","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1009","issueKey":"JCJAMT-1009","summary":"Business Billing","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1010","issueKey":"JCJAMT-1010","summary":"[Fresh Start w/LeanKit] GCM-Network to Network Interface/Enterprise Tail Migration","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Indrani Kodali","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1011","issueKey":"JCJAMT-1011","summary":"GCM-Ethernet Access Tail Migration (EATM)","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Indrani Kodali","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1013","issueKey":"JCJAMT-1013","summary":"CDO IQIPQTC","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"nk1628","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1014","issueKey":"JCJAMT-1014","summary":"Business Mobility Fixed Wireless","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1015","issueKey":"JCJAMT-1015","summary":"Business Mobility Wireless Growth","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1017","issueKey":"JCJAMT-1017","summary":"COLLECTIONS","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1018","issueKey":"JCJAMT-1018","summary":"CREDIT","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1019","issueKey":"JCJAMT-1019","summary":"DIGITAL","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Vinay Sanga","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1020","issueKey":"JCJAMT-1020","summary":"BTPD - ABS-on-BSSe","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Kelli Nguyen","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1021","issueKey":"JCJAMT-1021","summary":"CSO Insider Threat","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Lisa Francis, Kasturi Rangan Sankara Raman","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1022","issueKey":"JCJAMT-1022","summary":"Software Defined Networking Program","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Aromal Suresh, Lynn Rivera","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1023","issueKey":"JCJAMT-1023","summary":"Test Data Automation Framework","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"na4578","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1024","issueKey":"JCJAMT-1024","summary":"Jira iTrack Fresh Start Team - Main Frame Modernization","tool":"Fresh Start","droPortfolio":"Press","programArtPoc":"Keishawandra Wilson - KW5948","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1025","issueKey":"JCJAMT-1025","summary":"CDO Raptor (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"lb0402","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1027","issueKey":"JCJAMT-1027","summary":"A911PSP","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Tony Miller","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1028","issueKey":"JCJAMT-1028","summary":"CDO Core","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"st639b","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1029","issueKey":"JCJAMT-1029","summary":"ADW SAMS","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Beth Pruitt","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1030","issueKey":"JCJAMT-1030","summary":"CSO Revalidation","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Brittany Vassalo","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1031","issueKey":"JCJAMT-1031","summary":"CSO GTAC","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Duane Kimball, Hansi Rawat","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1034","issueKey":"JCJAMT-1034","summary":"CSO Kickstart","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Duane Kimball, Gary Boni","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1032","issueKey":"JCJAMT-1032","summary":"CSO BESS","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Duane Kimball, Hansi Rawat","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1033","issueKey":"JCJAMT-1033","summary":"CSO Security-Platforms","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Brittany Vassallo","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1036","issueKey":"JCJAMT-1036","summary":"CDO Network Data Science & Engg","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"vg4695","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1038","issueKey":"JCJAMT-1038","summary":"CDO ATS Proserve","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"vg4695","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1037","issueKey":"JCJAMT-1037","summary":"CDO Streaming Messages","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"vg4695 nk1628 gb1555","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1040","issueKey":"JCJAMT-1040","summary":"[FRESH START] [LABS] Inventory Data Federation","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sanjit Patnaik","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1041","issueKey":"JCJAMT-1041","summary":"CSO Cybersecurity and Technology Solutions","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Sarah A Williams","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1042","issueKey":"JCJAMT-1042","summary":"Physical Ericsson Network Manager","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Zachary Wyman, Lisa DonnaChaidh","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1043","issueKey":"JCJAMT-1043","summary":"[Fresh Start] Remote Desktop","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Patrick Taylor","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1044","issueKey":"JCJAMT-1044","summary":"[Fresh Start] Gateway2","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Aromal Suresh and Shailesh Patkar","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1045","issueKey":"JCJAMT-1045","summary":"Mainframe Transformation","tool":"ADO Boards","droPortfolio":"Green","programArtPoc":"Yumna Butterfield","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1046","issueKey":"JCJAMT-1046","summary":"[Fresh Start] RAN Validation Tool Requests","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Angela Morgan","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1047","issueKey":"JCJAMT-1047","summary":"Intelligent Automation","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Vishali Sutaria","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1049","issueKey":"JCJAMT-1049","summary":"[New] Network Service Assurance - Other","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1050","issueKey":"JCJAMT-1050","summary":"[New] CSO Abuse management Team","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"rp3862","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1051","issueKey":"JCJAMT-1051","summary":"[New] CSO AT&T Recon & Exploitation Specialists","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"rp3862","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1052","issueKey":"JCJAMT-1052","summary":"[New] CSO CTA - DAN","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Steve Pohlman, Vinod Krishnakutty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1053","issueKey":"JCJAMT-1053","summary":"[New] CSO CTA - Flood","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Steve Pohlman, Vinod Krishnakutty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1054","issueKey":"JCJAMT-1054","summary":"[New] CSO CTA - JCIS","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Steve Pohlman, Vinod Krishnakutty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1055","issueKey":"JCJAMT-1055","summary":"[New]  Logical Provisioning Platform","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"BIJU THOMAS","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1056","issueKey":"JCJAMT-1056","summary":"[New] BBND - Device Technology (JA Only.  No JC)","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"David Thomas","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1058","issueKey":"JCJAMT-1058","summary":"[New] CSO CTA - JDP","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Steve Pohlman","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1066","issueKey":"JCJAMT-1066","summary":"[New] AT&T Switched Ethernet and User Network Interface","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1064","issueKey":"JCJAMT-1064","summary":"[New] Artemis: Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1068","issueKey":"JCJAMT-1068","summary":"[New] GenAI: Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1060","issueKey":"JCJAMT-1060","summary":"[New] Systems Support: Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1067","issueKey":"JCJAMT-1067","summary":"[New] Billing: Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1059","issueKey":"JCJAMT-1059","summary":"[New] RAN:  Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1061","issueKey":"JCJAMT-1061","summary":"[New] Unauthorized Usage:  Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1062","issueKey":"JCJAMT-1062","summary":"[New] Web of AI:  Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1063","issueKey":"JCJAMT-1063","summary":"CSV - BTPD - RTB Wireless","tool":"LeanKit - net new","droPortfolio":"Cohen","programArtPoc":"Derrick Maxie","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1065","issueKey":"JCJAMT-1065","summary":"[New] 5G Simulator: Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1069","issueKey":"JCJAMT-1069","summary":"[New] Network Resiliency:  Labs","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sherry Harradine","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1070","issueKey":"JCJAMT-1070","summary":"[New] ICORE ABS","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"BIJU THOMAS","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1071","issueKey":"JCJAMT-1071","summary":"[New] Network Orchestration","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sharon Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1072","issueKey":"JCJAMT-1072","summary":"[New] PATH","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sharon Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1073","issueKey":"JCJAMT-1073","summary":"[New] PINC","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sharon Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1074","issueKey":"JCJAMT-1074","summary":"CDO Server Utilization Dashboard","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"as170c","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1075","issueKey":"JCJAMT-1075","summary":"CDO Watchtower AI","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ar7207","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1076","issueKey":"JCJAMT-1076","summary":"[New] LPP","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"BIJU THOMAS","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1077","issueKey":"JCJAMT-1077","summary":"[New] LPP Micro Services","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"BIJU THOMAS","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1078","issueKey":"JCJAMT-1078","summary":"[New] SDNWB GUI and MS","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"BIJU THOMAS","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1080","issueKey":"JCJAMT-1080","summary":"[CTX Migration] Wave 6: 40k Sub-Tasks","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1081","issueKey":"JCJAMT-1081","summary":"[CTX Migration] Wave 7: 360k Sub-Tasks","tool":"iTrack","droPortfolio":"Press","programArtPoc":"Dart Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1082","issueKey":"JCJAMT-1082","summary":"[New] RTNDT","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Abhishek Fnu","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1083","issueKey":"JCJAMT-1083","summary":"[New] Performance Engineering Group","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Danny James","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1084","issueKey":"JCJAMT-1084","summary":"[New] Broadband Gateway Architecture","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"David Thomas","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1085","issueKey":"JCJAMT-1085","summary":"CSV - BTPD - AT&T Business Ready (24hr Internet)","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"BTPD - AT&T Business Ready (24hr Internet)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1086","issueKey":"JCJAMT-1086","summary":"[New] NCAAS Finished Goods","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Kapil Singal","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1087","issueKey":"JCJAMT-1087","summary":"CDO DPG Finance","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"jm9359","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1088","issueKey":"JCJAMT-1088","summary":"[New] [Fresh Start] Granite NIS Transformation","tool":"ATT Jira","droPortfolio":"Elbaz","programArtPoc":"Ram Durvasula","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1089","issueKey":"JCJAMT-1089","summary":"BTPD - DevPlusSecureOps","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"KURUBA, SRIDHAR","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1091","issueKey":"JCJAMT-1091","summary":"[New] [FRESH START] GTP Certifications","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Anthony Hopkins","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1092","issueKey":"JCJAMT-1092","summary":"[New] 34324-CryptoUp","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"Aleksandra Korkus","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1093","issueKey":"JCJAMT-1093","summary":"[New] BBNMS Requests","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Suhas Jain","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1094","issueKey":"JCJAMT-1094","summary":"[New] Infra Requests","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Alexandria Curtis","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1095","issueKey":"JCJAMT-1095","summary":"CSV - Business Mobility KTLO","tool":"ATT Jira","droPortfolio":"Cohen","programArtPoc":"Existing - Business Mobility KTLO (In Process of being renamed to \"BTPD - Business Mobility KTLO\"","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1096","issueKey":"JCJAMT-1096","summary":"[New] Cohen Team - Clean-up post Migration Placeholder","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"TBD","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-140","issueKey":"JCJAMT-140","summary":"USH Security Governance","tool":"iTrack","droPortfolio":"Green","programArtPoc":"Michael Mcmanus; Justin Masterson; Kevin Cook","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1098","issueKey":"JCJAMT-1098","summary":"CSV - BTPD - Presale Platform Capabilities","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"BTPD - Presale Platform Capabilities","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1099","issueKey":"JCJAMT-1099","summary":"VAS - Dynamic Defense (additional teams)","tool":"ADO Boards","droPortfolio":"Zilberstain","programArtPoc":"aleksandra korkus","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1100","issueKey":"JCJAMT-1100","summary":"BTPD - Sentry","tool":"Fresh Start","droPortfolio":"Cohen","programArtPoc":"Andrew Lader","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1102","issueKey":"JCJAMT-1102","summary":"CDO Ask Platform and Proserve Support","tool":"Fresh Start","droPortfolio":"Markus","programArtPoc":"vg4695","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1105","issueKey":"JCJAMT-1105","summary":"CLONE - TicketingOps","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Adithya Atluri (aa158a)","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1106","issueKey":"JCJAMT-1106","summary":"CDO Defender","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"sk669a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1107","issueKey":"JCJAMT-1107","summary":"CDO Device Loss","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"sk669a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1108","issueKey":"JCJAMT-1108","summary":"CDO AvertackAI Platform","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"sk669a","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1109","issueKey":"JCJAMT-1109","summary":"CDO FAMLI","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dk2218","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1110","issueKey":"JCJAMT-1110","summary":"CDO Avertack","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"dk2218","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1111","issueKey":"JCJAMT-1111","summary":"CDO Deadshot","tool":"ADO Boards","droPortfolio":"Markus","programArtPoc":"dk2218","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1112","issueKey":"JCJAMT-1112","summary":"PRIME (aka Path)","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Sharon Smith","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1113","issueKey":"JCJAMT-1113","summary":"[FRESH START] Consumer-Business VOIP Services","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Anup Karnalkar","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1114","issueKey":"JCJAMT-1114","summary":"CLONE - CSO Cyber Threat Intelligence Operations","tool":"iTrack","droPortfolio":"Baich","programArtPoc":"Rajni Pandey","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1115","issueKey":"JCJAMT-1115","summary":"CSO Network Permit AI","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"BRITTANY VASSALLO, Vinod Krishnakutty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1116","issueKey":"JCJAMT-1116","summary":"CSO Digital Forensics Incident Response","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Rajni Pandey, Irving J Villanueva, Jamie Mellough","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1117","issueKey":"JCJAMT-1117","summary":"Network Data Inventory Tool (NetDB)","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Dorothy Chun-cromer","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1090","issueKey":"JCJAMT-1090","summary":"[New] Workflow Manager","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Vamsi Janga","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1118","issueKey":"JCJAMT-1118","summary":"PAI","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"erez hakim","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1119","issueKey":"JCJAMT-1119","summary":"CDO ELEAT","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"mr1613","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1120","issueKey":"JCJAMT-1120","summary":"CDO FELIX","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"mr1613","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1121","issueKey":"JCJAMT-1121","summary":"CDO MoWDR","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"mr1613","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1122","issueKey":"JCJAMT-1122","summary":"CDO MoWLI","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"mr1613","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1123","issueKey":"JCJAMT-1123","summary":"[FRESH START] Network Evolution Advanced Lab","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1124","issueKey":"JCJAMT-1124","summary":"CSO IAM - External (team move see description)","tool":"ADO Boards","droPortfolio":"Baich","programArtPoc":"MONA SHAMMA, KEVIN KIM, MOHAMMAD ASSAF","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1125","issueKey":"JCJAMT-1125","summary":"CSO IAM - Workforce (team move)","tool":"ADO Boards","droPortfolio":"Baich","programArtPoc":"Mona Shamma, Kevin Kim, Mohammad Assaf","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1126","issueKey":"JCJAMT-1126","summary":"CDO CCS","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"gd2167","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1127","issueKey":"JCJAMT-1127","summary":"CLONE - FASTWFBPM","tool":"iTrack","droPortfolio":"Zilberstain","programArtPoc":"Nazima Syed","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1128","issueKey":"JCJAMT-1128","summary":"Cloud RAN - CTO","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1129","issueKey":"JCJAMT-1129","summary":"RAN Transformation - Cloud RAN EIAP - CTO","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1130","issueKey":"JCJAMT-1130","summary":"EIAP and Cloud RAN Certification","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1035","issueKey":"JCJAMT-1035","summary":"MoW URL Blocking","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Richard Krajcik","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1131","issueKey":"JCJAMT-1131","summary":"CTO EIAP","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Vickie Mitchell","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1079","issueKey":"JCJAMT-1079","summary":"[New] [FRESH START] NDC","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Lydia San George","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1137","issueKey":"JCJAMT-1137","summary":"[Fresh Start] Integrated Fiber Planning","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Jackie Guelker","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1132","issueKey":"JCJAMT-1132","summary":"CSO Proxy Ops","tool":"Fresh Start","droPortfolio":"Baich","programArtPoc":"Jared Long, John Leboeuf, Vinod Krishnakutty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1133","issueKey":"JCJAMT-1133","summary":"Project Navis Transformation","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Lon Tanner","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1012","issueKey":"JCJAMT-1012","summary":"DDOS IP Security","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Alex Mendes","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1134","issueKey":"JCJAMT-1134","summary":"DDoS Address Family Exchange - AFX","tool":"iTrack","droPortfolio":"Elbaz","programArtPoc":"Richard Krajcik","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1135","issueKey":"JCJAMT-1135","summary":"CDO AI Compliance (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ec898v","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1136","issueKey":"JCJAMT-1136","summary":"CDO Proserv & Connected Solutions (Fresh Start)","tool":"iTrack","droPortfolio":"Markus","programArtPoc":"ts3335 & na502s","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1139","issueKey":"JCJAMT-1139","summary":"[Fresh Start] Enterprise IP Address Management","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Aromal Suresh","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1138","issueKey":"JCJAMT-1138","summary":"[Fresh Start] B2Bi Gateway","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Sarada Prusty","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1140","issueKey":"JCJAMT-1140","summary":"DBOR Migration & Retirement","tool":"Fresh Start","droPortfolio":"Zilberstain","programArtPoc":"Edward Hernandez","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1141","issueKey":"JCJAMT-1141","summary":"BTPD - LMCC","tool":"iTrack","droPortfolio":"Cohen","programArtPoc":"BTPD - LMCC","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1142","issueKey":"JCJAMT-1142","summary":"[FRESH START] IPLAT","tool":"Fresh Start","droPortfolio":"Elbaz","programArtPoc":"Erin Easter","stage":"Backlog","targetDate":"","assignee":"","droPo":""},{"id":"JCJAMT-1143","issueKey":"JCJAMT-1143","summary":"AON Automation","tool":"ADO Boards","droPortfolio":"Elbaz","programArtPoc":"Trevor Lovett","stage":"Backlog","targetDate":"","assignee":"","droPo":""}];
const SEED_VERSION = "2";   // bump when DEFAULT_TRACKER changes, to replace stale saved copies
let memTracker = null;
function loadTracker() {
  if (memTracker) return memTracker;
  try {
    const raw = localStorage.getItem("cc.tracker");
    const ver = localStorage.getItem("cc.tracker.seed");
    if (raw === null || ver !== SEED_VERSION) {
      memTracker = DEFAULT_TRACKER.slice();
      saveTracker(memTracker);
      try { localStorage.setItem("cc.tracker.seed", SEED_VERSION); } catch { /* ignore */ }
    } else memTracker = JSON.parse(raw);
  } catch { memTracker = DEFAULT_TRACKER.slice(); }
  return memTracker;
}
function saveTracker(items) { memTracker = items; try { localStorage.setItem("cc.tracker", JSON.stringify(items)); } catch { /* memory only */ } }

// Attach a generated report to the tracker ticket whose issueKey matches trackerId.
// Keeps only the latest auto-attached report per ticket (manual uploads are preserved).
function attachReportToTicket(trackerId, filename, dataUrl, sizeBytes) {
  const tid = (trackerId || "").trim();
  if (!tid) return "no-id";
  const items = loadTracker();
  let matched = false;
  const next = items.map((it) => {
    if (it.issueKey !== tid) return it;
    matched = true;
    const kept = (it.attachments || []).filter((a) => !a.auto);   // replace prior auto report
    kept.push({ name: filename, url: dataUrl, size: sizeBytes, ts: Date.now(), auto: true });
    return { ...it, attachments: kept };
  });
  if (!matched) return "no-match";
  memTracker = next;
  try { localStorage.setItem("cc.tracker", JSON.stringify(next)); return "ok"; }
  catch (e) { return "quota"; }   // kept in memory for this session
}

let memVR = {};
function bumpVersion(key) {
  let store;
  try { store = JSON.parse(localStorage.getItem("cc.versions") || "{}"); } catch { store = memVR; }
  const n = (store[key] || 0) + 1; store[key] = n; memVR = store;
  try { localStorage.setItem("cc.versions", JSON.stringify(store)); } catch { /* preview: memory only */ }
  return n;
}

function Clip() {
  return <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9"
    strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M21 8.5 12.5 17a4 4 0 0 1-5.7-5.7l8-8a2.5 2.5 0 0 1 3.5 3.5l-8 8a1 1 0 0 1-1.4-1.4l7.3-7.3" /></svg>;
}

function AggTable({ cols, rows, onRowClick }) {
  if (!rows.length) return <div style={{ padding: 24, textAlign: "center", color: C.muted, background: C.card, border: `1px solid ${C.border}`, borderRadius: 12 }}>No comparisons yet.</div>;
  const th = { textAlign: "left", padding: "9px 12px", fontSize: 11, fontWeight: 600, color: C.muted, textTransform: "uppercase", letterSpacing: "0.04em", background: C.head, borderBottom: `1px solid ${C.border}`, position: "sticky", top: 0 };
  const td = { padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}` };
  return (
    <div style={{ border: `1px solid ${C.border}`, borderRadius: 12, overflow: "auto", maxHeight: 380 }}>
      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 620 }}>
        <thead><tr>{cols.map((c) => <th key={c} style={th}>{c}</th>)}</tr></thead>
        <tbody>{rows.map((r, i) => (
          <tr key={i} onClick={() => onRowClick && onRowClick(r)} style={{ background: i % 2 ? C.rowAlt : "transparent", cursor: onRowClick ? "pointer" : "default" }}>
            {cols.map((c) => <td key={c} style={{ ...td, color: c.includes("Score") ? RAMP.green : c === "Program" && onRowClick ? C.accent : C.text }}>{r[c] ?? "—"}</td>)}
          </tr>))}
        </tbody>
      </table>
    </div>
  );
}

function StatCards({ items }) {
  return (
    <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
      {items.map(([label, value, color]) => (
        <div key={label} style={{ flex: 1, minWidth: 150, background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "14px 16px" }}>
          <div style={{ fontSize: 24, fontWeight: 700, color: color || C.text }}>{value}</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{label}</div>
        </div>
      ))}
    </div>
  );
}

function MetricCards({ items, onPick, active }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 12, marginBottom: 16 }}>
      {items.map(([label, value, color, field]) => (
        <button key={label} onClick={() => field && onPick({ label, field })} disabled={!field}
          style={{ textAlign: "left", background: C.card, border: `1.5px solid ${active === label ? C.accent : C.border}`, borderRadius: 12, padding: "14px 16px", cursor: field ? "pointer" : "default" }}>
          <div style={{ fontSize: 24, fontWeight: 700, color: color || C.text }}>{value}</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{label}{field ? " ›" : ""}</div>
        </button>
      ))}
    </div>
  );
}

function MigrationReportView() {
  const [runs, setRuns] = useState([]);
  const [mode, setMode] = useState("projects");     // projects | daily | weekly | program
  const [tab, setTab] = useState("current");     // current | previous
  const [selWeek, setSelWeek] = useState("");
  const [metric, setMetric] = useState(null);   // {label, field} for the clicked breakdown
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [q, setQ] = useState("");
  const [projF, setProjF] = useState("All");
  const [selProject, setSelProject] = useState(null);
  const [initF, setInitF] = useState("All");
  const [toolF, setToolF] = useState("All");
  const [scDetail, setScDetail] = useState(null);   // {type:'dro'|'tool', name} full detail view
  const [scSel, setScSel] = useState(null);         // {type,name} selected chip → summary strip
  const [scSort, setScSort] = useState("overall_desc");
  useEffect(() => { fetchRuns().then((r) => setRuns(r || [])); }, []);

  const isoOf = (ts) => new Date(ts).toISOString().slice(0, 10);
  const mmdd = (isoStr) => { const [y, m, d] = isoStr.split("-"); return `${m}/${d}/${y}`; };
  const wkStart = (isoStr) => { const d = new Date(isoStr + "T00:00:00"); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return d.toISOString().slice(0, 10); };
  const inRange = (isoStr) => (!from || isoStr >= from) && (!to || isoStr <= to);
  const stamp = (isoStr) => { if (!isoStr) return ""; const [y, m, d] = isoStr.split("-"); return m + d + y; };
  const rangeActive = !!(from || to);
  const sum = (arr, k) => arr.reduce((a, r) => a + (Number(r[k]) || 0), 0);
  const agg = (arr) => ({ comparisons: arr.length, users: new Set(arr.map((r) => r.user)).size, avgScore: arr.length ? Math.round(arr.reduce((a, r) => a + (Number(r.score) || 0), 0) / arr.length * 10) / 10 : 0, pending: sum(arr, "pending"), hierarchy: sum(arr, "hierarchy"), differences: sum(arr, "differences"), duplicates: sum(arr, "duplicates"), migrated: sum(arr, "migrated"), validated: sum(arr, "itrack") });

  const today = new Date().toISOString().slice(0, 10);
  const thisWeek = wkStart(today);
  const projectOptions = ["All", ...Array.from(new Set(runs.map((r) => r.targetProject).filter(Boolean))).sort()];
  const searchMatch = (r) => (projF === "All" || r.targetProject === projF) && (!q || ((r.trackerId || "") + " " + (r.targetProject || "") + " " + (r.source || "")).toLowerCase().includes(q.toLowerCase()));
  const fr = runs.filter(searchMatch);
  const todayRuns = fr.filter((r) => isoOf(r.ts) === today);

  const group = (keyFn, excludeKey) => {
    const g = {}; fr.forEach((r) => { const k = keyFn(isoOf(r.ts)); if (k !== excludeKey) (g[k] = g[k] || []).push(r); });
    return Object.keys(g).sort().reverse().map((k) => ({ k, ...agg(g[k]) }));
  };
  const prevDays = group((d) => d, today);
  const prevWeeks = group((d) => wkStart(d), thisWeek);

  const overall = agg(fr);
  const weekList = (() => {
    const set = new Set(fr.map((r) => wkStart(isoOf(r.ts)))); set.add(thisWeek);
    const sorted = [...set].sort(); let cur = sorted[0] || thisWeek, n = 1; const out = [];
    while (cur <= thisWeek && n < 520) { out.push({ num: n, monday: cur }); n++; const d = new Date(cur + "T00:00:00"); d.setDate(d.getDate() + 7); cur = d.toISOString().slice(0, 10); }
    return out;
  })();
  const curWeek = selWeek || thisWeek;
  const curWeekObj = weekList.find((w) => w.monday === curWeek) || { num: weekList.length, monday: curWeek };
  const prevMon = (isoStr) => { const d = new Date(isoStr + "T00:00:00"); d.setDate(d.getDate() - 7); return d.toISOString().slice(0, 10); };
  const selRuns = fr.filter((r) => wkStart(isoOf(r.ts)) === curWeek);
  const selAgg = agg(selRuns);
  const lastRuns = fr.filter((r) => wkStart(isoOf(r.ts)) === prevMon(curWeek));
  const delta = lastRuns.length ? Math.round((selAgg.avgScore - agg(lastRuns).avgScore) * 10) / 10 : null;
  const weekLabel = `Week ${curWeekObj.num} - ${mmdd(curWeek)}`;

  const runRows = (arr) => arr.slice().sort((a, b) => b.ts - a.ts).map((r) => ({ Time: new Date(r.ts).toLocaleString(), "Tracker ID": r.trackerId || "", Validator: r.user, Source: r.source, "Target program": r.targetProject, "Score %": r.score, "Missing in target": r.pending, "Hierarchy missing": r.hierarchy, "Field mismatches": r.differences, Duplicates: r.duplicates }));
  const dayRows = prevDays.filter((d) => inRange(d.k)).map((d) => ({ Date: mmdd(d.k), Comparisons: d.comparisons, Validators: d.users, "Avg score %": d.avgScore, "Missing in target": d.pending, "Hierarchy missing": d.hierarchy, "Field mismatches": d.differences, Duplicates: d.duplicates }));
  const weekRows = prevWeeks.filter((w) => inRange(w.k)).map((w) => ({ "Week of": mmdd(w.k), Comparisons: w.comparisons, Validators: w.users, "Avg score %": w.avgScore, "Missing in target": w.pending, "Hierarchy missing": w.hierarchy, "Field mismatches": w.differences, Duplicates: w.duplicates }));

  const exportReport = () => {
    const wb = XLSX.utils.book_new();
    const _d = new Date(), _p = (n) => String(n).padStart(2, "0");
    const stamp = `${_p(_d.getMonth() + 1)}${_p(_d.getDate())}${_d.getFullYear()}`;
    if (mode === "daily") {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(runRows(todayRuns).length ? runRows(todayRuns) : [{ "": "none" }]), "Today");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dayRows.length ? dayRows : [{ "": "none" }]), "Previous days");
    } else {
      const prog = [{ Metric: "Week", Value: weekLabel }, { Metric: "Overall avg score %", Value: overall.avgScore }, { Metric: "Total comparisons", Value: overall.comparisons }, { Metric: "Total validated", Value: overall.validated }, { Metric: "Missing in target", Value: overall.pending }, { Metric: "Hierarchy missing", Value: overall.hierarchy }, { Metric: "Field mismatches", Value: overall.differences }, { Metric: "Duplicates", Value: overall.duplicates }, { Metric: "Selected week avg score %", Value: selAgg.avgScore }, { Metric: "Δ vs previous week", Value: delta ?? "n/a" }];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prog), "Overall");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(runRows(selRuns).length ? runRows(selRuns) : [{ "": "none" }]), `Week ${curWeekObj.num}`);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(weekRows.length ? weekRows : [{ "": "none" }]), "Previous weeks");
    }
    try { XLSX.writeFile(wb, `Migration_Report_${mode === "daily" ? "Daily_" + stamp : "Week" + curWeekObj.num + "_" + (curWeek.split("-").slice(1).join("") + curWeek.split("-")[0])}.xlsx`); } catch (e) { /* preview */ }
  };

  const exportRange = () => {
    const runsIn = fr.filter((r) => inRange(isoOf(r.ts)));
    const wb = XLSX.utils.book_new();
    const g = {}; runsIn.forEach((r) => { const d = isoOf(r.ts); (g[d] = g[d] || []).push(r); });
    const daily = Object.keys(g).sort().reverse().map((d) => { const a = agg(g[d]); return { Date: mmdd(d), Comparisons: a.comparisons, Validators: a.users, "Avg score %": a.avgScore, "Missing in target": a.pending, "Hierarchy missing": a.hierarchy, "Field mismatches": a.differences, Duplicates: a.duplicates }; });
    const a = agg(runsIn);
    const summary = [{ Metric: "Range", Value: `${stamp(from) || "start"} – ${stamp(to) || "end"}` }, { Metric: "Comparisons", Value: a.comparisons }, { Metric: "Validators", Value: a.users }, { Metric: "Avg score %", Value: a.avgScore }, { Metric: "Validated", Value: a.validated }, { Metric: "Missing in target", Value: a.pending }, { Metric: "Hierarchy missing", Value: a.hierarchy }, { Metric: "Field mismatches", Value: a.differences }, { Metric: "Duplicates", Value: a.duplicates }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(daily.length ? daily : [{ "": "none" }]), "Daily summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(runRows(runsIn).length ? runRows(runsIn) : [{ "": "none" }]), "Runs");
    try { XLSX.writeFile(wb, `Migration_Report_${stamp(from) || "start"}_to_${stamp(to) || "end"}.xlsx`); } catch (e) { /* preview */ }
  };

  const allWkGroups = {}; fr.forEach((r) => { const w = wkStart(isoOf(r.ts)); (allWkGroups[w] = allWkGroups[w] || []).push(r); });
  const programWeekRows = Object.keys(allWkGroups).sort().reverse().map((w) => { const a = agg(allWkGroups[w]); const wo = weekList.find((x) => x.monday === w); return { Week: wo ? `Week ${wo.num}` : "", "Week of": mmdd(w), Comparisons: a.comparisons, "Avg score %": a.avgScore, "Missing in target": a.pending, "Hierarchy missing": a.hierarchy, "Field mismatches": a.differences, Duplicates: a.duplicates }; });

  const trackerAtt = {}; const trackerPortfolio = {}; const trackerTool = {};
  loadTracker().forEach((it) => { if (it.issueKey) { trackerAtt[it.issueKey] = it.attachments || []; trackerPortfolio[it.issueKey] = it.droPortfolio || ""; trackerTool[it.issueKey] = it.tool || ""; } });
  const projGroups = {}; fr.forEach((r) => { const p = r.targetProject || "(none)"; (projGroups[p] = projGroups[p] || []).push(r); });
  const projectRows = Object.keys(projGroups).sort().map((p) => {
    const a = agg(projGroups[p]); const latest = projGroups[p].slice().sort((x, y) => y.ts - x.ts)[0];
    const tids = new Set(projGroups[p].map((r) => r.trackerId).filter(Boolean));
    let reports = 0; tids.forEach((t) => { reports += (trackerAtt[t] || []).length; });
    return { Program: p, "Tracker ID": latest.trackerId || "", Comparisons: a.comparisons, "Latest score %": latest.score, "Missing in target": a.pending, "Hierarchy missing": a.hierarchy, "Field mismatches": a.differences, Duplicates: a.duplicates, Reports: reports, "Last run": new Date(latest.ts).toLocaleDateString() };
  });
  const projRuns = selProject ? fr.filter((r) => (r.targetProject || "(none)") === selProject) : [];
  const projAgg = agg(projRuns);
  const stampNow = () => { const _d = new Date(), _p = (n) => String(n).padStart(2, "0"); return _p(_d.getMonth() + 1) + _p(_d.getDate()) + _d.getFullYear(); };
  const safeName = (x) => String(x || "").replace(/[\\/:*?"<>|]/g, "_");
  const downloadRun = (r) => {
    const rows = [
      { Field: "Tracker ID", Value: r.trackerId || "" }, { Field: "Program", Value: r.targetProject }, { Field: "Source", Value: r.source },
      { Field: "Validator (ran by)", Value: r.user }, { Field: "Date", Value: new Date(r.ts).toLocaleString() }, { Field: "Overall score %", Value: r.score },
      { Field: "iTrack count", Value: r.itrack }, { Field: "Jira count", Value: r.jira }, { Field: "Migrated", Value: r.migrated },
      { Field: "Missing in target", Value: r.pending }, { Field: "Hierarchy missing", Value: r.hierarchy }, { Field: "Field mismatches", Value: r.differences },
      { Field: "Duplicates", Value: r.duplicates }, { Field: "Jira only", Value: r.jiraOnly }, { Field: "Matched", Value: r.matched },
    ];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Summary");
    try { XLSX.writeFile(wb, `Validation_Report_${safeName(r.trackerId || r.targetProject || "run")}_${stampNow()}.xlsx`); } catch (e) { /* preview */ }
  };
  const downloadProject = () => {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(runRows(projRuns).length ? runRows(projRuns) : [{ "": "none" }]), "Validation runs");
    try { XLSX.writeFile(wb, `Validation_Report_${safeName(selProject)}_${stampNow()}.xlsx`); } catch (e) { /* preview */ }
  };
  const exportGroup = () => {
    if (!scDetail) return;
    const summary = [
      { Metric: scDetail.type === "dro" ? "DRO Portfolio" : "Source tool", Value: scDetail.name },
      { Metric: "Comparisons (compared/total)", Value: grpCompared + "/" + grpTotal },
      { Metric: "Coverage %", Value: grpPct }, { Metric: "Avg score %", Value: gAgg.avgScore },
      { Metric: "Missing in target", Value: gAgg.pending }, { Metric: "Hierarchy missing", Value: gAgg.hierarchy },
      { Metric: "Field mismatches", Value: gAgg.differences }, { Metric: "Duplicates", Value: gAgg.duplicates },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(crossRows.length ? crossRows : [{ "": "none" }]), scDetail.type === "dro" ? "By source tool" : "By DRO");
    const trendRows = weekList.map((w, i) => ({ Week: "Week " + w.num, "Week of": mmdd(w.monday), "Avg score %": gTrend[i] == null ? "n/a" : gTrend[i] })).filter((r) => r["Avg score %"] !== "n/a");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trendRows.length ? trendRows : [{ "": "none" }]), "Weekly trend");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(gProjectRows.length ? gProjectRows : [{ "": "none" }]), "Projects");
    const _d = new Date(), _p = (n) => String(n).padStart(2, "0");
    try { XLSX.writeFile(wb, "Program_" + safeName((scDetail.type === "dro" ? "DRO_" : "Tool_") + scDetail.name) + "_" + _p(_d.getMonth() + 1) + _p(_d.getDate()) + _d.getFullYear() + ".xlsx"); } catch (e) { /* preview */ }
  };
  const exportProjects = () => {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(projectRows.length ? projectRows : [{ "": "none" }]), "Projects");
    const _d = new Date(), _p = (n) => String(n).padStart(2, "0");
    try { XLSX.writeFile(wb, "Migration_Report_Programs_" + _p(_d.getMonth() + 1) + _p(_d.getDate()) + _d.getFullYear() + ".xlsx"); } catch (e) { /* preview */ }
  };

  const exportProgram = () => {
    const wb = XLSX.utils.book_new();
    const prog = [{ Metric: "Overall avg score %", Value: overall.avgScore }, { Metric: "Total comparisons", Value: overall.comparisons }, { Metric: "Total validated", Value: overall.validated }, { Metric: "Missing in target", Value: overall.pending }, { Metric: "Hierarchy missing", Value: overall.hierarchy }, { Metric: "Field mismatches", Value: overall.differences }, { Metric: "Duplicates", Value: overall.duplicates }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prog), "Overall");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(programWeekRows.length ? programWeekRows : [{ "": "none" }]), "By week");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(runRows(fr).length ? runRows(fr) : [{ "": "none" }]), "All runs");
    const _d = new Date(), _p = (n) => String(n).padStart(2, "0");
    try { XLSX.writeFile(wb, `Migration_Report_Program_${_p(_d.getMonth() + 1)}${_p(_d.getDate())}${_d.getFullYear()}.xlsx`); } catch (e) { /* preview */ }
  };

  const breakdown = (scopeRuns) => metric ? (
    <div style={{ marginTop: 6 }}>
      <div style={{ fontSize: 13.5, fontWeight: 600, margin: "4px 0 8px" }}>{metric.label} — by comparison</div>
      <AggTable cols={["Time", "Source", "Target program", metric.label]}
        rows={scopeRuns.filter((r) => Number(r[metric.field]) > 0).sort((a, b) => b[metric.field] - a[metric.field]).map((r) => ({ Time: new Date(r.ts).toLocaleString(), Source: r.source, "Target program": r.targetProject, [metric.label]: r[metric.field] }))} />
      <div style={{ color: C.faint, fontSize: 12, marginTop: 6 }}>Per-issue detail is in each comparison's own report (Run comparison → tile → detail, or its downloaded Excel).</div>
    </div>
  ) : null;

  // ---- Program Health Scorecard model ----
  const CHECKS = [["status", "Status"], ["issue_type", "Type"], ["parent_key", "Parent"], ["assignee", "Assignee"], ["priority", "Priority"], ["story_points", "SP"], ["description", "Desc"], ["acceptance_criteria", "AC"], ["resolution", "Reso"], ["labels", "Labels"]];
  const scBand = (v) => v == null ? C.muted : v >= 75 ? RAMP.green : v >= 40 ? RAMP.amber : RAMP.red;
  const wkStartOf = (r) => wkStart(isoOf(r.ts));
  const scByProject = {}; fr.forEach((r) => { const p = r.targetProject || "(none)"; (scByProject[p] = scByProject[p] || []).push(r); });
  const latestUpTo = (runsArr, monday) => { const c = runsArr.filter((r) => wkStartOf(r) <= monday).sort((a, b) => b.ts - a.ts); return c[0] || null; };
  const initiatives = ["All", ...Array.from(new Set(Object.values(trackerPortfolio).filter(Boolean))).sort()];
  const wkIndex = Math.max(0, weekList.findIndex((w) => w.monday === curWeek));
  let scRows = Object.keys(scByProject).map((p) => {
    const runsArr = scByProject[p];
    const cur = latestUpTo(runsArr, curWeek);
    const tid = cur ? cur.trackerId : "";
    const portfolio = trackerPortfolio[tid] || "";
    const tool = trackerTool[tid] || (cur ? cur.source : "") || "";
    const trend = weekList.map((w) => { const rr = latestUpTo(runsArr, w.monday); return rr ? rr.score : null; });
    const prev = wkIndex > 0 ? (latestUpTo(runsArr, weekList[wkIndex - 1].monday)?.score ?? null) : null;
    return { project: p, tid, portfolio, tool, score: cur ? cur.score : null, fp: cur ? (cur.fieldPct || {}) : {}, trend, prev };
  });
  scRows = scRows.filter((r) => (initF === "All" || r.portfolio === initF) && (toolF === "All" || r.tool === toolF) && (!q || (r.project + " " + r.tid + " " + r.portfolio + " " + r.tool).toLowerCase().includes(q.toLowerCase())));
  if (scSort === "overall_desc") scRows.sort((a, b) => (b.score ?? -1) - (a.score ?? -1));
  else if (scSort === "overall_asc") scRows.sort((a, b) => (a.score ?? 999) - (b.score ?? 999));
  else if (scSort === "name_asc") scRows.sort((a, b) => a.project.localeCompare(b.project));
  else if (scSort === "initiative") scRows.sort((a, b) => (a.portfolio || "").localeCompare(b.portfolio || "") || (b.score ?? -1) - (a.score ?? -1));
  const scored = scRows.filter((r) => r.score != null);
  const scHealth = scored.length ? Math.round(scored.reduce((s, r) => s + r.score, 0) / scored.length) : 0;
  const onTrack = scored.filter((r) => r.score >= 75).length;
  const atRisk = scored.filter((r) => r.score >= 40 && r.score < 75).length;
  const offTrack = scored.filter((r) => r.score < 40).length;
  let totCells = 0, greenCells = 0;
  scRows.forEach((r) => CHECKS.forEach(([f]) => { const v = r.fp[f]; if (v != null) { totCells++; if (v >= 75) greenCells++; } }));
  const initAvg = (name) => { const rs = Object.keys(scByProject).map((p) => { const cur = latestUpTo(scByProject[p], curWeek); return cur && (trackerPortfolio[cur.trackerId] || "") === name ? cur.score : null; }).filter((x) => x != null); return { n: rs.length, avg: rs.length ? Math.round(rs.reduce((a, b) => a + b, 0) / rs.length) : 0 }; };
  const TOOL_ORDER = ["att jira", "itrack", "ado", "servicenow", "leankit", "fresh start"];
  const toolRank = (t) => { const tl = (t || "").toLowerCase(); const i = TOOL_ORDER.findIndex((k) => tl.includes(k)); return i < 0 ? TOOL_ORDER.length : i; };
  const toolsList = Array.from(new Set(Object.values(trackerTool).filter(Boolean))).sort((a, b) => toolRank(a) - toolRank(b) || a.localeCompare(b));
  const toolAvg = (name) => { const rs = Object.keys(scByProject).map((p) => { const cur = latestUpTo(scByProject[p], curWeek); return cur && (trackerTool[cur.trackerId] || cur.source || "") === name ? cur.score : null; }).filter((x) => x != null); return { n: rs.length, avg: rs.length ? Math.round(rs.reduce((a, b) => a + b, 0) / rs.length) : 0 }; };
  const groupStats = (sel) => { if (!sel) return null; const match = (r) => sel.type === "dro" ? (trackerPortfolio[r.trackerId] || "") === sel.name : (trackerTool[r.trackerId] || r.source || "") === sel.name; const runs = fr.filter(match); const tickets = loadTracker().filter((it) => sel.type === "dro" ? it.droPortfolio === sel.name : it.tool === sel.name); const keys = new Set(tickets.map((t) => t.issueKey)); const compared = new Set(runs.map((r) => r.trackerId).filter((t) => t && keys.has(t))).size; return { runs, a: agg(runs), total: tickets.length, compared, pct: tickets.length ? Math.round(compared / tickets.length * 100) : 0 }; };
  const selStats = groupStats(scSel);
  const detail = groupStats(scDetail);
  const groupRuns = detail ? detail.runs : [];
  const gAgg = detail ? detail.a : agg([]);
  const grpTotal = detail ? detail.total : 0;
  const grpCompared = detail ? detail.compared : 0;
  const grpPct = detail ? detail.pct : 0;
  const gByProj = {}; groupRuns.forEach((r) => { const p = r.targetProject || "(none)"; (gByProj[p] = gByProj[p] || []).push(r); });
  const gProjectRows = Object.keys(gByProj).sort().map((p) => { const a = agg(gByProj[p]); const latest = gByProj[p].slice().sort((x, y) => y.ts - x.ts)[0]; return { Program: p, "Tracker ID": latest.trackerId || "", Comparisons: a.comparisons, "Latest score %": latest.score, "Missing in target": a.pending, "Hierarchy missing": a.hierarchy, "Field mismatches": a.differences, Duplicates: a.duplicates, "Last run": new Date(latest.ts).toLocaleDateString() }; });
  const crossOf = (r) => scDetail ? (scDetail.type === "dro" ? (trackerTool[r.trackerId] || r.source || "—") : (trackerPortfolio[r.trackerId] || "—")) : "—";
  const crossLabel = scDetail && scDetail.type === "dro" ? "Source → target" : "DRO Portfolio";
  const crossGroups = {}; groupRuns.forEach((r) => { const k = crossOf(r); (crossGroups[k] = crossGroups[k] || []).push(r); });
  const crossRows = Object.keys(crossGroups).sort().map((k) => { const a = agg(crossGroups[k]); const nm = scDetail && scDetail.type === "dro" ? (k + " → Jira Cloud") : k; return { [crossLabel]: nm, Comparisons: a.comparisons, "Avg score %": a.avgScore, "Missing in target": a.pending, "Field mismatches": a.differences, Duplicates: a.duplicates }; });
  // group's average score per week (projects in the group, latest run up to each week)
  const detailProjects = detail ? Array.from(new Set(detail.runs.map((r) => r.targetProject || "(none)"))) : [];
  const gTrend = weekList.map((w) => {
    const scores = detailProjects.map((p) => { const rr = latestUpTo(scByProject[p] || [], w.monday); return rr ? rr.score : null; }).filter((x) => x != null);
    return scores.length ? Math.round(scores.reduce((a, b) => a + b, 0) / scores.length) : null;
  });
  const gTrendVals = gTrend.filter((x) => x != null);
  const gTrendCur = gTrendVals.length ? gTrendVals[gTrendVals.length - 1] : null;
  const gTrendPrev = gTrendVals.length > 1 ? gTrendVals[gTrendVals.length - 2] : null;
  const sparkPath = (trend) => {
    const w = 76, h = 24, bw = 10, gap = 2.4;
    return trend.map((s, i) => { if (s == null) return null; const bh = Math.max(2, s / 100 * h); return `<rect x="${i * (bw + gap)}" y="${h - bh}" width="${bw}" height="${bh}" rx="2" fill="${scBand(s)}" opacity="${i === wkIndex ? 1 : 0.4}"/>`; }).filter(Boolean).join("");
  };

  const seg = (setFn, k, cur, l) => (
    <button onClick={() => setFn(k)} style={{ padding: "7px 15px", borderRadius: 8, fontSize: 13.5, cursor: "pointer", fontWeight: cur === k ? 600 : 500, color: cur === k ? "#fff" : C.muted, background: cur === k ? C.accent : "transparent", border: "none" }}>{l}</button>
  );

  return (
    <div style={{ padding: "24px 28px", color: C.text, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
        <div>
          <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700 }}>Migration report</h1>
          <p style={{ margin: "6px 0 0", color: C.muted, fontSize: 14 }}>Roll-up of validation comparisons {RUNS_API ? "across all users" : "recorded on this machine"}.</p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          {mode === "daily" && (<>
            <span style={{ color: C.faint, fontSize: 12 }}>From</span>
            <input type="date" value={from} onChange={(e) => setFrom(e.target.value)}
              style={{ padding: "6px 9px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12.5, outline: "none", colorScheme: "dark" }} />
            <span style={{ color: C.faint, fontSize: 12 }}>To</span>
            <input type="date" value={to} onChange={(e) => setTo(e.target.value)}
              style={{ padding: "6px 9px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12.5, outline: "none", colorScheme: "dark" }} />
            {rangeActive && <button onClick={() => { setFrom(""); setTo(""); }} style={{ padding: "6px 10px", background: "transparent", color: C.muted, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12.5, cursor: "pointer" }}>Clear</button>}
          </>)}
          {mode === "weekly" && (<>
            <span style={{ color: C.faint, fontSize: 12 }}>Week</span>
            <select value={curWeek} onChange={(e) => setSelWeek(e.target.value)}
              style={{ padding: "7px 10px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, outline: "none", colorScheme: "dark" }}>
              {weekList.slice().reverse().map((w) => <option key={w.monday} value={w.monday}>{`Week ${w.num} - ${mmdd(w.monday)}`}</option>)}
            </select>
          </>)}
          <button onClick={mode === "daily" ? (rangeActive ? exportRange : exportReport) : mode === "weekly" ? exportReport : mode === "projects" ? exportProjects : exportProgram} style={{ padding: "8px 15px", background: C.accent + "18", color: C.accent, border: `1px solid ${C.accent}66`, borderRadius: 9, fontSize: 13.5, cursor: "pointer" }}>
            {mode === "daily" ? (rangeActive ? "Export range" : "Export daily") : mode === "weekly" ? `Export ${weekLabel}` : mode === "projects" ? "Export programs" : "Export overall"}
          </button>
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginTop: 14 }}>
        <span style={{ color: C.faint, fontSize: 12 }}>Program</span>
        <select value={projF} onChange={(e) => setProjF(e.target.value)} style={{ padding: "7px 10px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, outline: "none" }}>
          {projectOptions.map((p) => <option key={p} value={p}>{p}</option>)}
        </select>
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search Tracker ID or target program…" style={{ padding: "7px 11px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, outline: "none", minWidth: 260 }} />
        {(q || projF !== "All") && <button onClick={() => { setQ(""); setProjF("All"); }} style={{ padding: "6px 10px", background: "transparent", color: C.muted, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12.5, cursor: "pointer" }}>Clear</button>}
        <span style={{ color: C.faint, fontSize: 12.5 }}>{fr.length} of {runs.length} runs</span>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", margin: "18px 0" }}>
        <div style={{ display: "inline-flex", background: C.field, border: `1px solid ${C.border}`, borderRadius: 10, padding: 3 }}>
          {seg(setMode, "projects", mode, "Programs")}{seg(setMode, "daily", mode, "Daily")}{seg(setMode, "weekly", mode, "Weekly")}
        </div>
        <div style={{ display: "inline-flex", background: C.field, border: `1px solid ${C.border}`, borderRadius: 10, padding: 3 }}>
          {seg(setMode, "program", mode, "Overall")}
        </div>
      </div>

      {mode !== "program" && mode !== "projects" && (
        <div style={{ display: "flex", gap: 20, borderBottom: `1px solid ${C.border}`, marginBottom: 18 }}>
          {["current", "previous"].map((k) => (
            <button key={k} onClick={() => setTab(k)} style={{ background: "none", border: "none", cursor: "pointer", padding: "8px 2px", fontSize: 14, color: tab === k ? C.accent : C.muted, borderBottom: `2px solid ${tab === k ? C.accent : "transparent"}`, fontWeight: tab === k ? 600 : 500 }}>
              {mode === "daily" ? (k === "current" ? "Today" : "Previous days") : (k === "current" ? weekLabel : "Previous weeks")}
            </button>
          ))}
        </div>
      )}

      {rangeActive && mode === "daily" && (
        <div style={{ color: C.muted, fontSize: 12.5, marginBottom: 12 }}>
          Showing <b style={{ color: C.text }}>{stamp(from) || "start"} – {stamp(to) || "end"}</b> · applies to Previous days and the range export
        </div>
      )}

      {mode === "daily" && tab === "current" && (<>
        <MetricCards active={metric?.label} onPick={setMetric} items={[["Comparisons today", todayRuns.length, RAMP.blue], ["Avg score %", agg(todayRuns).avgScore, RAMP.green], ["Missing in target", agg(todayRuns).pending, RAMP.teal, "pending"], ["Hierarchy missing", agg(todayRuns).hierarchy, RAMP.indigo, "hierarchy"], ["Field mismatches", agg(todayRuns).differences, RAMP.amber, "differences"], ["Duplicates", agg(todayRuns).duplicates, RAMP.purple, "duplicates"]]} />
        {breakdown(todayRuns)}
        <AggTable cols={["Time", "Tracker ID", "Source", "Target program", "Score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates"]} rows={runRows(todayRuns)} />
      </>)}
      {mode === "daily" && tab === "previous" && (
        <AggTable cols={["Date", "Comparisons", "Avg score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates"]} rows={dayRows} />
      )}

      {mode === "weekly" && tab === "current" && (<>
        <MetricCards active={metric?.label} onPick={setMetric} items={[[`Comparisons (${weekLabel})`, selRuns.length, RAMP.blue], ["Avg score %", selAgg.avgScore, RAMP.green], ["Missing in target", selAgg.pending, RAMP.teal, "pending"], ["Hierarchy missing", selAgg.hierarchy, RAMP.indigo, "hierarchy"], ["Field mismatches", selAgg.differences, RAMP.amber, "differences"], ["Duplicates", selAgg.duplicates, RAMP.purple, "duplicates"]]} />
        {breakdown(selRuns)}
        <AggTable cols={["Time", "Tracker ID", "Source", "Target program", "Score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates"]} rows={runRows(selRuns)} />
      </>)}
      {mode === "weekly" && tab === "previous" && (
        <AggTable cols={["Week of", "Comparisons", "Avg score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates"]} rows={weekRows} />
      )}

      {mode === "program" && (scDetail ? (<>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <button onClick={() => setScDetail(null)} style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "6px 12px", background: C.card, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13, cursor: "pointer" }}>← Scorecard</button>
          <button onClick={exportGroup} style={{ padding: "8px 15px", background: C.accent + "18", color: C.accent, border: `1px solid ${C.accent}66`, borderRadius: 9, fontSize: 13.5, cursor: "pointer" }}>Export</button>
        </div>
        <div style={{ fontSize: 12, color: C.faint, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>{scDetail.type === "dro" ? "DRO Portfolio" : "Source tool"}</div>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 14 }}>{scDetail.name}</div>
        <MetricCards active={metric?.label} onPick={setMetric} items={[["Comparisons (" + grpPct + "%)", grpCompared + "/" + grpTotal, RAMP.indigo], ["Avg score %", gAgg.avgScore, RAMP.green], ["Missing in target", gAgg.pending, RAMP.teal, "pending"], ["Hierarchy missing", gAgg.hierarchy, RAMP.indigo, "hierarchy"], ["Field mismatches", gAgg.differences, RAMP.amber, "differences"], ["Duplicates", gAgg.duplicates, RAMP.purple, "duplicates"]]} />
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", marginBottom: 14, display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
          <div>
            <div style={{ fontSize: 11, color: C.muted, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Overall trend</div>
            <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginTop: 3 }}>
              <span style={{ fontSize: 22, fontWeight: 700, color: gTrendCur == null ? C.faint : scBand(gTrendCur) }}>{gTrendCur == null ? "—" : gTrendCur + "%"}</span>
              {gTrendPrev != null && gTrendCur != null && <span style={{ fontSize: 12, fontWeight: 700, color: gTrendCur > gTrendPrev ? RAMP.green : gTrendCur < gTrendPrev ? RAMP.red : C.faint }}>{gTrendCur > gTrendPrev ? "▲ " + (gTrendCur - gTrendPrev) : gTrendCur < gTrendPrev ? "▼ " + (gTrendPrev - gTrendCur) : "– 0"} <span style={{ color: C.faint, fontWeight: 400 }}>vs prior week</span></span>}
            </div>
          </div>
          <div style={{ flex: 1, minWidth: 220 }} dangerouslySetInnerHTML={{ __html: `<svg width="100%" height="46" viewBox="0 0 ${Math.max(60, weekList.length * 30)} 46" preserveAspectRatio="none">${weekList.map((w, i) => { const v = gTrend[i]; if (v == null) return ""; const bh = Math.max(3, v / 100 * 40); const x = i * 30; return `<rect x="${x + 4}" y="${44 - bh}" width="20" height="${bh}" rx="3" fill="${scBand(v)}" opacity="${i === weekList.length - 1 ? 1 : 0.45}"/><text x="${x + 14}" y="9" font-size="8" fill="${C.faint}" text-anchor="middle">W${w.num}</text>`; }).join("")}</svg>` }} />
          {gTrendVals.length === 0 && <span style={{ color: C.faint, fontSize: 12 }}>No validated weeks yet.</span>}
        </div>
        {breakdown(groupRuns)}
        <div style={{ fontSize: 13.5, fontWeight: 600, margin: "6px 0 8px" }}>{scDetail.type === "dro" ? "Source → target tool comparisons" : "By DRO Portfolio"}</div>
        <AggTable cols={[crossLabel, "Comparisons", "Avg score %", "Missing in target", "Field mismatches", "Duplicates"]} rows={crossRows} />
        <div style={{ fontSize: 13.5, fontWeight: 600, margin: "16px 0 8px" }}>Programs in {scDetail.name}</div>
        <AggTable cols={["Program", "Tracker ID", "Comparisons", "Latest score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates", "Last run"]} rows={gProjectRows} onRowClick={(r) => { setSelProject(r.Program); setMode("projects"); setScDetail(null); }} />
        <div style={{ color: C.faint, fontSize: 11.5, marginTop: 8 }}>Click a program to open its validation reports.</div>
      </>) : (<>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 14 }}>Overall health scorecard</div>

        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}>
          <span style={{ color: C.faint, fontSize: 12, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>DRO Portfolio</span>
          <select value={initF} onChange={(e) => setInitF(e.target.value)} style={{ padding: "8px 12px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, fontWeight: 600, outline: "none" }}>
            {initiatives.map((i) => <option key={i} value={i}>{i === "All" ? "All portfolios" : i}</option>)}
          </select>
          <span style={{ color: C.faint, fontSize: 12, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Source tool</span>
          <select value={toolF} onChange={(e) => setToolF(e.target.value)} style={{ padding: "8px 12px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, fontWeight: 600, outline: "none" }}>
            {["All", ...toolsList].map((t) => <option key={t} value={t}>{t === "All" ? "All tools" : t}</option>)}
          </select>
          <select value={scSort} onChange={(e) => setScSort(e.target.value)} style={{ padding: "8px 12px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, outline: "none" }}>
            <option value="overall_desc">Sort: Overall (high → low)</option>
            <option value="overall_asc">Sort: Overall (low → high)</option>
            <option value="name_asc">Sort: Program (A–Z)</option>
            <option value="initiative">Sort: Initiative</option>
          </select>
          <span style={{ color: C.faint, fontSize: 12.5 }}>{scRows.length} programs</span>
        </div>

        <MetricCards active={metric?.label} onPick={setMetric} items={[["Overall health", scHealth + "%", RAMP.blue], ["On track (≥75%)", onTrack, RAMP.green], ["At risk (40–74%)", atRisk, RAMP.amber], ["Off track (<40%)", offTrack, RAMP.red], ["Checks ≥75%", totCells ? Math.round(greenCells / totCells * 100) + "%" : "—", RAMP.teal]]} />

        <div style={{ fontSize: 12, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: "0.05em", margin: "4px 0 8px" }}>DRO Portfolio</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10, marginBottom: 8 }}>
          {initiatives.filter((i) => i !== "All").map((name) => { const a = initAvg(name); const sel = scSel && scSel.type === "dro" && scSel.name === name; return (
            <button key={name} onClick={() => { setScSel({ type: "dro", name }); setMetric(null); }} style={{ textAlign: "left", background: C.card, border: `1px solid ${sel ? C.accent : C.border}`, borderLeft: `4px solid ${C.accent}`, borderRadius: 8, padding: "10px 12px", cursor: "pointer", boxShadow: sel ? `0 0 0 1px ${C.accent}` : "none" }}>
              <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 6 }}>{name}</div>
              <div style={{ height: 5, background: C.field, borderRadius: 3, overflow: "hidden", marginBottom: 6 }}><div style={{ height: "100%", width: a.avg + "%", background: scBand(a.avg) }} /></div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10.5, color: C.muted, fontFamily: "IBM Plex Mono, ui-monospace, monospace" }}><span>{a.n} programs</span><span>{a.avg}%</span></div>
            </button>
          ); })}
        </div>
        <button onClick={() => scSel && scSel.type === "dro" && setScDetail(scSel)} disabled={!(scSel && scSel.type === "dro")}
          style={{ marginBottom: 16, padding: "6px 13px", background: scSel && scSel.type === "dro" ? C.accent + "18" : "transparent", color: scSel && scSel.type === "dro" ? C.accent : C.faint, border: `1px solid ${scSel && scSel.type === "dro" ? C.accent + "66" : C.border}`, borderRadius: 8, fontSize: 12.5, cursor: scSel && scSel.type === "dro" ? "pointer" : "default" }}>
          {scSel && scSel.type === "dro" ? `Click for details — ${scSel.name} ›` : "Select a DRO Portfolio above, then click for details"}
        </button>

        <div style={{ fontSize: 12, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: "0.05em", margin: "4px 0 8px" }}>Source tool</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10, marginBottom: 8 }}>
          {toolsList.map((name) => { const a = toolAvg(name); const sel = scSel && scSel.type === "tool" && scSel.name === name; return (
            <button key={name} onClick={() => { setScSel({ type: "tool", name }); setMetric(null); }} style={{ textAlign: "left", background: C.card, border: `1px solid ${sel ? RAMP.purple : C.border}`, borderLeft: `4px solid ${RAMP.purple}`, borderRadius: 8, padding: "10px 12px", cursor: "pointer", boxShadow: sel ? `0 0 0 1px ${RAMP.purple}` : "none" }}>
              <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 6 }}>{name}</div>
              <div style={{ height: 5, background: C.field, borderRadius: 3, overflow: "hidden", marginBottom: 6 }}><div style={{ height: "100%", width: a.avg + "%", background: scBand(a.avg) }} /></div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10.5, color: C.muted, fontFamily: "IBM Plex Mono, ui-monospace, monospace" }}><span>{a.n} programs</span><span>{a.avg}%</span></div>
            </button>
          ); })}
        </div>
        <button onClick={() => scSel && scSel.type === "tool" && setScDetail(scSel)} disabled={!(scSel && scSel.type === "tool")}
          style={{ marginBottom: 16, padding: "6px 13px", background: scSel && scSel.type === "tool" ? RAMP.purple + "22" : "transparent", color: scSel && scSel.type === "tool" ? RAMP.purple : C.faint, border: `1px solid ${scSel && scSel.type === "tool" ? RAMP.purple + "88" : C.border}`, borderRadius: 8, fontSize: 12.5, cursor: scSel && scSel.type === "tool" ? "pointer" : "default" }}>
          {scSel && scSel.type === "tool" ? `Click for details — ${scSel.name} ›` : "Select a Source tool above, then click for details"}
        </button>

        <div style={{ display: "flex", gap: 14, alignItems: "center", fontSize: 11.5, color: C.muted, marginBottom: 10, flexWrap: "wrap" }}>
          {[["≥75% On track", RAMP.green], ["40–74% At risk", RAMP.amber], ["<40% Off track", RAMP.red], ["N/A", C.muted]].map(([l, cc]) => <span key={l} style={{ display: "inline-flex", alignItems: "center", gap: 5 }}><span style={{ width: 12, height: 12, borderRadius: 4, background: cc + "33", border: `1px solid ${cc}` }} />{l}</span>)}
        </div>

        {scRows.length === 0 ? (
          <div style={{ padding: 30, textAlign: "center", color: C.muted, background: C.card, border: `1px solid ${C.border}`, borderRadius: 12 }}>No programs match the current filters. Run comparisons with a Tracker ID to populate the scorecard.</div>
        ) : (
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 12, overflow: "auto", maxHeight: 560 }}>
            <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 1100 }}>
              <thead><tr>
                <th style={{ position: "sticky", left: 0, top: 0, zIndex: 3, background: C.head, textAlign: "left", padding: "10px 14px", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.05em", color: C.muted, fontWeight: 700, borderBottom: `1px solid ${C.border}`, minWidth: 220 }}>Program</th>
                {CHECKS.map(([f, l]) => <th key={f} title={f} style={{ position: "sticky", top: 0, background: C.head, textAlign: "center", padding: "10px 6px", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.04em", color: C.muted, fontWeight: 700, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{l}</th>)}
                <th style={{ position: "sticky", top: 0, background: C.head, textAlign: "center", padding: "10px 10px", fontSize: 10.5, textTransform: "uppercase", color: C.muted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Overall</th>
              </tr></thead>
              <tbody>
                {scRows.map((r, i) => (
                  <tr key={i}>
                    <td style={{ position: "sticky", left: 0, background: i % 2 ? C.rowAlt : C.card, padding: "8px 14px", borderBottom: `1px solid ${C.border}`, borderRight: `1px solid ${C.border}` }}>
                      <div><span style={{ fontWeight: 600, fontSize: 13 }}>{r.project}</span> {r.tid && <span style={{ color: C.faint, fontFamily: "ui-monospace, monospace", fontSize: 11, marginLeft: 5 }}>{r.tid}</span>}</div>
                      {r.portfolio && <div style={{ marginTop: 4 }}><span style={{ fontSize: 9, fontWeight: 700, padding: "2px 6px", borderRadius: 4, color: "#fff", background: C.accent }}>{r.portfolio}</span></div>}
                    </td>
                    {CHECKS.map(([f]) => { const v = r.fp[f]; const cc = scBand(v); return (
                      <td key={f} style={{ textAlign: "center", padding: "7px 5px", borderBottom: `1px solid ${C.border}` }}>
                        <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 40, height: 24, borderRadius: 6, fontSize: 11.5, fontWeight: 700, fontFamily: "ui-monospace, monospace", color: cc, background: cc + "22" }}>{v == null ? "N/A" : v + "%"}</span>
                      </td>
                    ); })}
                    <td style={{ textAlign: "center", padding: "7px 8px", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>
                      <div style={{ display: "inline-flex", alignItems: "center", padding: "4px 9px", borderRadius: 100, fontFamily: "ui-monospace, monospace", fontWeight: 700, fontSize: 12.5, color: scBand(r.score), background: scBand(r.score) + "22" }}>{r.score == null ? "—" : r.score + "%"}</div>
                      {r.prev != null && r.score != null && <div style={{ fontSize: 10.5, fontWeight: 700, marginTop: 3, color: r.score > r.prev ? RAMP.green : r.score < r.prev ? RAMP.red : C.faint }}>{r.score > r.prev ? "▲ " + (r.score - r.prev) : r.score < r.prev ? "▼ " + (r.prev - r.score) : "– 0"}</div>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <div style={{ color: C.faint, fontSize: 11.5, marginTop: 10 }}>Checks are per-field match rates from each program's latest comparison up to the selected week. Initiatives map to DRO Portfolio via each run's Tracker ID. Cells show N/A until a comparison with per-field data has run.</div>

        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", marginTop: 16, display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
          <span style={{ color: C.faint, fontSize: 12, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Week</span>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {weekList.map((w) => {
              const rs = Object.keys(scByProject).map((p) => latestUpTo(scByProject[p], w.monday)).filter(Boolean).filter((r) => (initF === "All" || (trackerPortfolio[r.trackerId] || "") === initF) && (toolF === "All" || (trackerTool[r.trackerId] || r.source || "") === toolF));
              const avg = rs.length ? Math.round(rs.reduce((a, b) => a + b.score, 0) / rs.length) : 0;
              const active = w.monday === curWeek;
              return <button key={w.monday} onClick={() => setSelWeek(w.monday)} style={{ fontFamily: "IBM Plex Mono, ui-monospace, monospace", fontSize: 12, fontWeight: 600, padding: "7px 12px", borderRadius: 8, cursor: "pointer", border: `1px solid ${active ? C.accent : C.border}`, background: active ? C.accent : C.field, color: active ? "#fff" : C.muted, lineHeight: 1.3 }}>Wk {w.num}<span style={{ display: "block", fontSize: 9.5, opacity: 0.8 }}>{mmdd(w.monday)} · {avg}%</span></button>;
            })}
          </div>
        </div>
      </>))}

      {mode === "projects" && (selProject ? (<>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <button onClick={() => setSelProject(null)} style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "6px 12px", background: C.card, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13, cursor: "pointer" }}>← All programs</button>
          <button onClick={downloadProject} style={{ padding: "8px 15px", background: C.accent + "18", color: C.accent, border: `1px solid ${C.accent}66`, borderRadius: 9, fontSize: 13.5, cursor: "pointer" }}>Download program report</button>
        </div>
        <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 12 }}>{selProject}</div>
        <MetricCards active={metric?.label} onPick={setMetric} items={[["Comparisons", projAgg.comparisons, RAMP.blue], ["Avg score %", projAgg.avgScore, RAMP.green], ["Missing in target", projAgg.pending, RAMP.teal, "pending"], ["Hierarchy missing", projAgg.hierarchy, RAMP.indigo, "hierarchy"], ["Field mismatches", projAgg.differences, RAMP.amber, "differences"], ["Duplicates", projAgg.duplicates, RAMP.purple, "duplicates"]]} />
        {breakdown(projRuns)}
        <div style={{ fontSize: 13.5, fontWeight: 600, margin: "6px 0 8px" }}>Validation reports for {selProject}</div>
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 12, overflow: "auto", maxHeight: 400 }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
            <thead><tr>{["Date", "Tracker ID", "Ran by", "Score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates", "Report"].map((h) => <th key={h} style={{ textAlign: "left", padding: "9px 12px", fontSize: 11, fontWeight: 600, color: C.muted, textTransform: "uppercase", letterSpacing: "0.04em", background: C.head, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>)}</tr></thead>
            <tbody>
              {projRuns.slice().sort((a, b) => b.ts - a.ts).map((r, i) => (
                <tr key={i} style={{ background: i % 2 ? C.rowAlt : "transparent" }}>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{new Date(r.ts).toLocaleString()}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}`, color: C.accent, fontFamily: "ui-monospace, Menlo, monospace" }}>{r.trackerId || "—"}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}` }}>{r.user}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}`, color: RAMP.green }}>{r.score}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}` }}>{r.pending}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}` }}>{r.hierarchy}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}` }}>{r.differences}</td>
                  <td style={{ padding: "9px 12px", fontSize: 12.5, borderBottom: `1px solid ${C.border}` }}>{r.duplicates}</td>
                  <td style={{ padding: "7px 12px", borderBottom: `1px solid ${C.border}` }}>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
                      {(trackerAtt[r.trackerId] || []).map((a, ai) => (
                        <a key={ai} href={a.url} download={a.name} title={"Attached in Tracker: " + a.name}
                          style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "4px 9px", background: C.accent + "18", color: C.accent, border: `1px solid ${C.accent}66`, borderRadius: 7, fontSize: 11.5, textDecoration: "none", maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 8.5 12.5 17a4 4 0 0 1-5.7-5.7l8-8a2.5 2.5 0 0 1 3.5 3.5l-8 8a1 1 0 0 1-1.4-1.4l7.3-7.3" /></svg>{a.name}
                        </a>
                      ))}
                      <button onClick={() => downloadRun(r)} title="Summary built from stored metrics" style={{ padding: "4px 9px", background: C.card, color: C.muted, border: `1px solid ${C.border}`, borderRadius: 7, fontSize: 11.5, cursor: "pointer" }}>Summary</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </>) : (<>
        <MetricCards active={metric?.label} onPick={setMetric} items={[["Programs", projectRows.length, RAMP.blue], ["Comparisons", overall.comparisons, RAMP.indigo], ["Avg score %", overall.avgScore, RAMP.green], ["Missing in target", overall.pending, RAMP.teal, "pending"], ["Hierarchy missing", overall.hierarchy, RAMP.indigo, "hierarchy"], ["Field mismatches", overall.differences, RAMP.amber, "differences"], ["Duplicates", overall.duplicates, RAMP.purple, "duplicates"]]} />
        {breakdown(fr)}
        <div style={{ fontSize: 13.5, fontWeight: 600, margin: "6px 0 8px" }}>All programs · click a program to see its metrics and validation reports</div>
        <AggTable cols={["Program", "Tracker ID", "Comparisons", "Latest score %", "Missing in target", "Hierarchy missing", "Field mismatches", "Duplicates", "Reports", "Last run"]} rows={projectRows} onRowClick={(r) => setSelProject(r.Program)} />
      </>))}
    </div>
  );
}

function EditCell({ value, onChange, type = "text" }) {
  return <input type={type} value={value || ""} onChange={(e) => onChange(e.target.value)}
    style={{ width: "100%", boxSizing: "border-box", padding: "6px 8px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 6, fontSize: 12.5, outline: "none", colorScheme: "dark" }} />;
}

function TrackerView() {
  const [items, setItems] = useState(() => loadTracker().map((it, i) => ({ ...it, id: it.id || it.issueKey || "r" + i })));
  const [board, setBoard] = useState(true);        // board (Tool swimlanes) vs table
  const [portfolio, setPortfolio] = useState("All");
  const [q, setQ] = useState("");
  const [toolF, setToolF] = useState("All");
  const [stageF, setStageF] = useState("All");
  const [sortCol, setSortCol] = useState("");
  const [sortDir, setSortDir] = useState(1);
  const [assigneeF, setAssigneeF] = useState("All");
  const [droPoF, setDroPoF] = useState("All");
  const [dragId, setDragId] = useState(null);
  const [collapsed, setCollapsed] = useState(() => new Set());
  const [error, setError] = useState("");

  const commit = (next) => { setItems(next); saveTracker(next); };
  const edit = (id, field, val) => commit(items.map((it) => (it.id === id ? { ...it, [field]: val } : it)));
  const nextKey = () => {
    let max = 0;
    items.forEach((it) => { const m = /^JCJAMT-(\d+)$/i.exec(String(it.issueKey || "").trim()); if (m) max = Math.max(max, parseInt(m[1], 10)); });
    return "JCJAMT-" + (max + 1);
  };
  const addRow = () => { const key = nextKey(); commit([...items, { id: key, issueKey: key, summary: "", tool: "", droPortfolio: "", programArtPoc: "", stage: "Backlog", targetDate: "", assignee: "", droPo: "", attachments: [] }]); };
  const attach = (id, file) => {
    if (file.size > 3 * 1024 * 1024) { setError("File too large for browser storage (~3MB max). Use a shared backend for large reports."); return; }
    const r = new FileReader();
    r.onload = () => {
      const next = items.map((it) => it.id === id ? { ...it, attachments: [...(it.attachments || []), { name: file.name, url: r.result, size: file.size, ts: Date.now() }] } : it);
      memTracker = next; setItems(next);
      try { localStorage.setItem("cc.tracker", JSON.stringify(next)); setError(""); } catch (e) { setError("Browser storage is full — attachment kept for this session only. A shared backend is needed to persist many/large reports."); }
    };
    r.readAsDataURL(file);
  };
  const removeAtt = (id, idx) => commit(items.map((it) => it.id === id ? { ...it, attachments: (it.attachments || []).filter((_, j) => j !== idx) } : it));
  const removeRow = (id) => commit(items.filter((it) => it.id !== id));

  const pick = (row, ...names) => {
    const keys = Object.keys(row);
    for (const n of names) { const hit = keys.find((k) => norm(k) === norm(n)); if (hit) return String(row[hit] || "").trim(); }
    for (const n of names) { const hit = keys.find((k) => norm(k).includes(norm(n))); if (hit) return String(row[hit] || "").trim(); }
    return "";
  };
  const seedFromRows = (rows) => commit(rows.map((r, i) => ({
    id: pick(r, "Issue key", "Issue Key", "Key") || "r" + i,
    issueKey: pick(r, "Issue key", "Issue Key", "Key"), summary: pick(r, "Summary"),
    tool: pick(r, "Custom field (Tool)", "Tool"), droPortfolio: pick(r, "Custom field (DRO Portfolio)", "DRO Portfolio"),
    programArtPoc: pick(r, "Custom field (Program/Art POC)", "Program/Art POC"),
    stage: "Backlog",
    targetDate: pick(r, "Custom field (Target Launch Date)", "Target Date", "Target Launch Date"),
    assignee: pick(r, "Assignee"), droPo: "",
  })).filter((x) => x.issueKey || x.summary));

  const upload = (file) => {
    const r = new FileReader(), name = file.name.toLowerCase();
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      r.onload = () => { try { const wb = XLSX.read(r.result, { type: "array" }); seedFromRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" })); } catch (e) { setError("Could not read the tracker: " + e.message); } };
      r.readAsArrayBuffer(file);
    } else {
      r.onload = () => { const rows = parseCsv(r.result), hdr = rows[0] || []; seedFromRows(rows.slice(1).map((rr) => { const o = {}; hdr.forEach((h, i) => (o[h] = rr[i] || "")); return o; })); };
      r.readAsText(file);
    }
  };

  const exportTracker = () => {
    const rows = items.map((it) => ({ "Issue Key": it.issueKey, Summary: it.summary, "Custom field (Tool)": it.tool, "Custom field (DRO Portfolio)": it.droPortfolio, "Custom field (Program/Art POC)": it.programArtPoc, Status: it.stage, "Target Date": it.targetDate, Assignee: it.assignee, "DRO Product Owner": it.droPo }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ "": "none" }]), "Tracker");
    const _d = new Date(), _p = (n) => String(n).padStart(2, "0");
    try { XLSX.writeFile(wb, "Validation_Tracker_" + _p(_d.getMonth() + 1) + _p(_d.getDate()) + _d.getFullYear() + ".xlsx"); } catch (e) { setError("Export blocked in this preview; works when hosted."); }
  };

  const portfolios = ["All", ...Array.from(new Set(items.map((it) => it.droPortfolio).filter(Boolean))).sort()];
  const toolOptions = ["All", ...Array.from(new Set(items.map((it) => it.tool || "Unspecified"))).sort()];
  const assigneeOptions = ["All", ...Array.from(new Set(items.map((it) => it.assignee).filter(Boolean))).sort()];
  const droPoOptions = ["All", ...Array.from(new Set(items.map((it) => it.droPo).filter(Boolean))).sort()];
  const match = (it) => (portfolio === "All" || it.droPortfolio === portfolio)
    && (toolF === "All" || (it.tool || "Unspecified") === toolF)
    && (stageF === "All" || it.stage === stageF)
    && (assigneeF === "All" || it.assignee === assigneeF)
    && (droPoF === "All" || it.droPo === droPoF)
    && (!q || (it.summary + " " + it.issueKey + " " + it.programArtPoc + " " + it.assignee + " " + it.targetDate).toLowerCase().includes(q.toLowerCase()));
  const visible = items.filter(match);
  const sortedVisible = sortCol ? [...visible].sort((a, b) => { const x = String(a[sortCol] || "").toLowerCase(), y = String(b[sortCol] || "").toLowerCase(); return x < y ? -sortDir : x > y ? sortDir : 0; }) : visible;
  const toggleSort = (col) => { if (sortCol === col) setSortDir(-sortDir); else { setSortCol(col); setSortDir(1); } };
  const PREFERRED_TOOLS = ["ATT Jira", "iTrack", "ADO Boards"];
  const LAST_TOOLS = ["Fresh Start", "Initial Fresh Start"];
  const rawTools = Array.from(new Set(visible.map((it) => it.tool || "Unspecified")));
  const middle = rawTools.filter((t) => !PREFERRED_TOOLS.includes(t) && !LAST_TOOLS.includes(t) && t !== "Unspecified").sort();
  const tools = [
    ...PREFERRED_TOOLS.filter((t) => rawTools.includes(t)),
    ...middle,
    ...LAST_TOOLS.filter((t) => rawTools.includes(t)),
    ...(rawTools.includes("Unspecified") ? ["Unspecified"] : []),
  ];
  const stageColor = (st) => st === "Done" ? RAMP.green : st === "Cancelled" ? RAMP.red : st === "Backlog" ? RAMP.blue : st === "Launched" ? RAMP.teal : RAMP.amber;
  const btn = (bg, col, brd) => ({ padding: "8px 15px", background: bg, color: col, border: "1px solid " + brd, borderRadius: 9, fontSize: 13.5, cursor: "pointer" });

  const grid = { display: "grid", gridTemplateColumns: "150px repeat(" + STAGES.length + ", minmax(190px, 1fr))" };
  const gridS = { display: "grid", gridTemplateColumns: "repeat(" + STAGES.length + ", minmax(190px, 1fr))", minWidth: STAGES.length * 190 };
  const toggleTool = (t) => setCollapsed((prev) => { const n = new Set(prev); n.has(t) ? n.delete(t) : n.add(t); return n; });
  const colHead = { padding: "9px 10px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.03em", borderBottom: "1px solid " + C.border, whiteSpace: "nowrap" };

  const Card = ({ it }) => (
    <div draggable onDragStart={() => setDragId(it.id)} onDragEnd={() => setDragId(null)}
      style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 9, padding: "8px 10px", marginBottom: 8, cursor: "grab" }}>
      <div style={{ fontSize: 11.5, fontFamily: "ui-monospace, Menlo, monospace", color: C.accent }}>{it.issueKey || "—"}</div>
      <div style={{ fontSize: 12.5, margin: "3px 0", lineHeight: 1.3 }}>{it.summary || "(no summary)"}</div>
      <div style={{ fontSize: 11, color: C.faint }}>{[it.droPortfolio, it.droPo].filter(Boolean).join(" · ").slice(0, 40)}</div>
      {(it.attachments || []).length > 0 && <div style={{ fontSize: 10.5, color: C.accent, marginTop: 3 }}>{it.attachments.length} report{it.attachments.length > 1 ? "s" : ""} attached</div>}
    </div>
  );

  return (
    <div style={{ padding: "24px 28px", color: C.text, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
        <div>
          <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700 }}>Validation tracker</h1>
          <p style={{ margin: "6px 0 0", color: C.muted, fontSize: 14 }}>Board grouped by Tool. Drag a card across stages — the status is the stage.</p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <div style={{ display: "inline-flex", background: C.field, border: "1px solid " + C.border, borderRadius: 9, padding: 3 }}>
            {[["board", "Board"], ["table", "Table"]].map(([k, l]) => (
              <button key={k} onClick={() => setBoard(k === "board")} style={{ padding: "6px 13px", borderRadius: 7, fontSize: 13, cursor: "pointer", border: "none", fontWeight: (board === (k === "board")) ? 600 : 500, color: (board === (k === "board")) ? "#fff" : C.muted, background: (board === (k === "board")) ? C.accent : "transparent" }}>{l}</button>
            ))}
          </div>
          <button onClick={addRow} style={btn(C.card, C.text, C.border)}>+ Add</button>
          <button onClick={exportTracker} style={btn(C.accent + "18", C.accent, C.accent + "66")}>Export</button>
        </div>
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", margin: "16px 0" }}>
        {[["DRO Portfolio", portfolio, setPortfolio, portfolios], ["Tool", toolF, setToolF, toolOptions], ["Stage", stageF, setStageF, ["All", ...STAGES]], ["Assignee", assigneeF, setAssigneeF, assigneeOptions], ["DRO Product Owner", droPoF, setDroPoF, droPoOptions]].map(([label, val, setter, opts]) => (
          <span key={label} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
            <span style={{ color: C.faint, fontSize: 12 }}>{label}</span>
            <select value={val} onChange={(e) => setter(e.target.value)} style={{ padding: "7px 10px", background: C.field, color: C.text, border: "1px solid " + C.border, borderRadius: 8, fontSize: 13, outline: "none" }}>
              {opts.map((o) => <option key={o} value={o}>{o}</option>)}
            </select>
          </span>
        ))}
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search key, summary, POC, assignee…" style={{ padding: "7px 11px", background: C.field, color: C.text, border: "1px solid " + C.border, borderRadius: 8, fontSize: 13, outline: "none", minWidth: 240 }} />
        <button onClick={() => commit([])} style={{ padding: "7px 14px", background: "transparent", color: C.muted, border: "1px solid " + C.border, borderRadius: 8, fontSize: 13, cursor: "pointer" }}>Clear all</button>
        <span style={{ color: C.faint, fontSize: 12.5 }}>{visible.length} of {items.length} items</span>
      </div>
      {error && <div style={{ color: RAMP.red, fontSize: 13, marginBottom: 8 }}>{error}</div>}

      {items.length === 0 ? (
        <div style={{ padding: 34, textAlign: "center", color: C.muted, background: C.card, border: "1px solid " + C.border, borderRadius: 12 }}>No items yet — upload your Jira tracker or add a row.</div>
      ) : board ? (
        <div style={{ border: "1px solid " + C.border, borderRadius: 12, overflow: "auto", maxHeight: 600 }}>
          <div style={{ ...gridS, position: "sticky", top: 0, background: C.head, zIndex: 1 }}>
            {STAGES.map((st) => <div key={st} style={{ ...colHead, color: stageColor(st) }}>{st} · {visible.filter((it) => it.stage === st).length}</div>)}
          </div>
          {tools.map((tool) => {
            const open = !collapsed.has(tool);
            const count = visible.filter((it) => (it.tool || "Unspecified") === tool).length;
            return (
              <div key={tool} style={{ borderBottom: "1px solid " + C.border }}>
                <div onClick={() => toggleTool(tool)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 12px", cursor: "pointer", background: C.rowAlt, position: "sticky", left: 0 }}>
                  <span style={{ color: C.muted, fontSize: 12, width: 12 }}>{open ? "▾" : "▸"}</span>
                  <span style={{ fontWeight: 600, fontSize: 13 }}>{tool}</span>
                  <span style={{ color: C.faint, fontSize: 11 }}>{count} items</span>
                </div>
                {open && (
                  <div style={gridS}>
                    {STAGES.map((st) => (
                      <div key={st} onDragOver={(e) => e.preventDefault()}
                        onDrop={() => { if (dragId != null) { commit(items.map((it) => it.id === dragId ? { ...it, stage: st, tool: tool === "Unspecified" ? it.tool : tool } : it)); setDragId(null); } }}
                        style={{ padding: 8, borderRight: "1px solid " + C.border, minHeight: 60 }}>
                        {visible.filter((it) => (it.tool || "Unspecified") === tool && it.stage === st).map((it) => <Card key={it.id} it={it} />)}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        <div style={{ border: "1px solid " + C.border, borderRadius: 12, overflow: "auto", maxHeight: 560 }}>
          <table style={{ borderCollapse: "collapse", minWidth: 1180 }}>
            <thead><tr>{[["issueKey", "Issue Key"], ["summary", "Summary"], ["tool", "Tool"], ["droPortfolio", "DRO Portfolio"], ["programArtPoc", "Program/Art POC"], ["stage", "Status"], ["targetDate", "Target Date"], ["assignee", "Assignee"], ["droPo", "DRO Product Owner"], ["", "Attachments"], ["", ""]].map(([f, h], hi) => <th key={h + hi} onClick={() => f && toggleSort(f)} style={{ textAlign: "left", padding: "9px 10px", fontSize: 11, fontWeight: 600, color: C.muted, textTransform: "uppercase", letterSpacing: "0.04em", background: C.head, borderBottom: "1px solid " + C.border, whiteSpace: "nowrap", cursor: f ? "pointer" : "default" }}>{h}{f ? <span style={{ opacity: 0.5, marginLeft: 3 }}>{sortCol === f ? (sortDir > 0 ? "↑" : "↓") : "↕"}</span> : ""}</th>)}</tr></thead>
            <tbody>
              {sortedVisible.map((it) => (
                <tr key={it.id} style={{ borderBottom: "1px solid " + C.border }}>
                  <td style={{ padding: 5, minWidth: 120 }}><EditCell value={it.issueKey} onChange={(v) => edit(it.id, "issueKey", v)} /></td>
                  <td style={{ padding: 5, minWidth: 220 }}><EditCell value={it.summary} onChange={(v) => edit(it.id, "summary", v)} /></td>
                  <td style={{ padding: 5, minWidth: 130 }}><EditCell value={it.tool} onChange={(v) => edit(it.id, "tool", v)} /></td>
                  <td style={{ padding: 5, minWidth: 120 }}><EditCell value={it.droPortfolio} onChange={(v) => edit(it.id, "droPortfolio", v)} /></td>
                  <td style={{ padding: 5, minWidth: 150 }}><EditCell value={it.programArtPoc} onChange={(v) => edit(it.id, "programArtPoc", v)} /></td>
                  <td style={{ padding: 5, minWidth: 160 }}>
                    <select value={it.stage} onChange={(e) => edit(it.id, "stage", e.target.value)} style={{ width: "100%", padding: "6px 8px", background: C.field, color: stageColor(it.stage), fontWeight: 600, border: "1px solid " + C.border, borderRadius: 6, fontSize: 12.5, outline: "none" }}>
                      {STAGES.map((s) => <option key={s} value={s} style={{ color: C.text }}>{s}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: 5, minWidth: 140 }}><EditCell type="date" value={it.targetDate} onChange={(v) => edit(it.id, "targetDate", v)} /></td>
                  <td style={{ padding: 5, minWidth: 130 }}><EditCell value={it.assignee} onChange={(v) => edit(it.id, "assignee", v)} /></td>
                  <td style={{ padding: 5, minWidth: 110 }}><EditCell value={it.droPo} onChange={(v) => edit(it.id, "droPo", v)} /></td>
                  <td style={{ padding: "5px 8px", minWidth: 200 }}>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 5, alignItems: "center" }}>
                      {(it.attachments || []).map((a, ai) => (
                        <span key={ai} style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "3px 7px", background: C.card, border: "1px solid " + C.border, borderRadius: 7, fontSize: 11.5 }}>
                          <a href={a.url} download={a.name} title={a.name} style={{ color: C.accent, textDecoration: "none", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</a>
                          <button onClick={() => removeAtt(it.id, ai)} title="Remove" style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 13, lineHeight: 1, padding: 0 }}>×</button>
                        </span>
                      ))}
                      <label style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "4px 9px", background: C.accent + "18", color: C.accent, border: "1px solid " + C.accent + "66", borderRadius: 7, fontSize: 11.5, cursor: "pointer" }}>
                        ＋ Attach<input type="file" style={{ display: "none" }} onChange={(e) => { const f = e.target.files[0]; if (f) attach(it.id, f); e.target.value = ""; }} />
                      </label>
                    </div>
                  </td>
                  <td style={{ padding: "5px 8px" }}><button onClick={() => removeRow(it.id)} title="Remove" style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 16 }}>\u00d7</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}


function JiraMigrationV1() {
  const [theme, setThemeState] = useState(() => { try { return localStorage.getItem("cc.theme") || "dark"; } catch { return "dark"; } });
  C = PALETTES[theme] || PALETTES.dark;
  RAMP = RAMPS[theme] || RAMPS.dark;
  const setTheme = (t) => { setThemeState(t); try { localStorage.setItem("cc.theme", t); } catch { /* ignore */ } };
  const [source, setSource] = useState(null);     // null = show landing page
  const [view, setView] = useState("compare");     // compare | report
  const [srcFiles, setSrcFiles] = useState([]);   // [{name, text}]
  const [tgtFiles, setTgtFiles] = useState([]);
  const [srcProjects, setSrcProjects] = useState("");
  const [tgtProjects, setTgtProjects] = useState("");
  const [trackerId, setTrackerId] = useState("");
  const [mappings, setMappings] = useState({});         // { field: {normSource: cloudValue} }
  const [mappingNames, setMappingNames] = useState({}); // { field: filename }
  const [userMap, setUserMap] = useState(() => buildUserMap(DEFAULT_USER_ROWS));  // shared across sources
  const [userName, setUserName] = useState("Jeremy_Legg_ATS_Org_list_08282026.xlsx (sample)");
  const [refView, setRefView] = useState(null);         // {title, cols, rows} for the View modal
  const [autoAttach, setAutoAttach] = useState(() => { try { return localStorage.getItem("cc.autoAttachReport") !== "0"; } catch { return true; } });
  const [attachMsg, setAttachMsg] = useState("");
  const [res, setRes] = useState(null);
  const [bucket, setBucket] = useState("differences");
  const [error, setError] = useState("");
  const [validator, setValidator] = useState(getUser());

  const pickSource = (id) => {
    setSource(id); setSrcFiles([]); setTgtFiles([]); setSrcProjects(""); setTgtProjects(""); setTrackerId("");
    const dm = DEFAULT_MAPPINGS[id] || {};
    const rules = {}, names = {};
    Object.keys(dm).forEach((f) => { rules[f] = dm[f].rules; names[f] = dm[f].name + " (preloaded)"; });
    setMappings(rules); setMappingNames(names);
    // user list is common to all sources — leave it as-is
    setRes(null); setError("");
  };
  const srcName = SOURCES.find((x) => x.id === source)?.name || source;

  // Load the full bundled org list on start (falls back to the inlined sample).
  useEffect(() => {
    if (!REFERENCE_USERS_URL) return;
    const url = REFERENCE_USERS_URL, isJson = url.endsWith(".json");
    fetch(url).then((r) => (isJson ? r.json() : r.arrayBuffer())).then((data) => {
      if (isJson) setUserMap(data);
      else { const wb = XLSX.read(data, { type: "array" }); const ws = wb.Sheets[wb.SheetNames[0]]; setUserMap(buildUserMap(XLSX.utils.sheet_to_json(ws, { header: 1 }))); }
      setUserName("Jeremy_Legg_ATS_Org_list_08282026 (preloaded)");
    }).catch(() => {});
  }, []);

  const removeMapping = (f) => {
    setMappings((m) => { const n = { ...m }; delete n[f]; return n; });
    setMappingNames((n) => { const x = { ...n }; delete x[f]; return x; });
  };
  const uploadMapping = (f, file) => {
    const r = new FileReader();
    r.onload = () => { setMappings((m) => ({ ...m, [f]: parseMapping(r.result) })); setMappingNames((n) => ({ ...n, [f]: file.name + " (uploaded)" })); };
    r.readAsText(file);
  };
  const removeUser = () => { setUserMap(null); setUserName(null); };

  const cleanName = (n) => (n || "reference").replace(/\s*\((preloaded|sample|uploaded)\)\s*$/i, "");
  const mappingRows = (f) => Object.entries(mappings[f] || {}).map(([src, tgt]) => ({ "Source value": src, "Cloud value": Array.isArray(tgt) ? tgt.join(", ") : (tgt === "" ? "(unset)" : tgt) }));
  const userRows = () => Object.entries(userMap || {}).map(([alias, uid]) => ({ Alias: alias, ATTUID: uid }));
  const viewMapping = (f) => setRefView({ title: f.replace("_", " ").replace(/\b\w/, (c) => c.toUpperCase()) + " mapping", sub: mappingNames[f], cols: ["Source value", "Cloud value"], rows: mappingRows(f) });
  const viewUsers = () => setRefView({ title: "User list", sub: userName, cols: ["Alias", "ATTUID"], rows: userRows() });
  const downloadRows = (rows, base) => { try { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ "": "none" }]), "Rules"); XLSX.writeFile(wb, cleanName(base).replace(/\.(xlsx|xls|csv)$/i, "") + ".xlsx"); } catch (e) { /* preview */ } };

  const loadUserList = (file) => {
    const r = new FileReader();
    const name = file.name.toLowerCase();
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      r.onload = () => {
        try {
          const wb = XLSX.read(r.result, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          setUserMap(buildUserMap(XLSX.utils.sheet_to_json(ws, { header: 1 }))); setUserName(file.name + " (uploaded)");
        } catch (e) { setError("Could not read the user list: " + e.message); }
      };
      r.readAsArrayBuffer(file);
    } else {
      r.onload = () => { setUserMap(buildUserMap(parseCsv(r.result))); setUserName(file.name + " (uploaded)"); };
      r.readAsText(file);
    }
  };

  const run = () => {
    setError("");
    try {
      const merge = (files, side) => {
        const seen = new Set(), out = [];
        files.forEach((f) => toIssues(readJira(f.text), side, source).forEach((i) => {
          if (i.key && !seen.has(i.key)) { seen.add(i.key); out.push(i); }
        }));
        return out;
      };
      const filt = (issues, projects) => {
        const set = new Set(projects.split(",").map((p) => p.trim().toUpperCase()).filter(Boolean));
        return set.size ? issues.filter((i) => set.has((i.project || "").toUpperCase())) : issues;
      };
      const s = filt(merge(srcFiles, "source"), srcProjects);
      const t = filt(merge(tgtFiles, "target"), tgtProjects);
      if (!s.length || !t.length) { setError("No rows after parsing/filtering — check the files and program keys."); return; }
      const result = validate(s, t, mappings, userMap, SOURCE_COMPOSITE[source] || []);
      setRes(result);
      const pj = (tgtProjects.split(",")[0].trim() || "AllPrograms");
      const su = result.summary;
      saveRun({ ts: Date.now(), user: validator || "unknown", source: srcName, targetProject: pj, trackerId: trackerId.trim(), score: su.score,
        itrack: su.itrack_count, jira: su.jira_count, migrated: su.migrated, pending: su.pending_migration,
        differences: su.differences, duplicates: su.duplicates, jiraOnly: su.jira_only, matched: su.matched, hierarchy: su.hierarchy_missing, fieldPct: su.field_pct });
      // attach the report to its matching tracker ticket
      const tid = trackerId.trim();
      if (autoAttach && tid) {
        try {
          const b64 = XLSX.write(buildWorkbook(result), { bookType: "xlsx", type: "base64" });
          const url = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + b64;
          const st = attachReportToTicket(tid, reportBase() + ".xlsx", url, Math.round(b64.length * 0.75));
          setAttachMsg(
            st === "ok" ? `Report attached to ${tid} in the Tracker.` :
            st === "no-match" ? `Saved. No Tracker ticket matches "${tid}", so the report wasn't attached — add the ticket in the Tracker, or check the ID.` :
            st === "quota" ? `Report attached to ${tid} for this session, but browser storage is full — a shared backend is needed to persist it.` : ""
          );
        } catch (e) { setAttachMsg("Comparison saved, but the report couldn't be attached: " + e.message); }
      } else setAttachMsg("");
    } catch (e) { setError("Could not parse the files: " + e.message); }
  };

  const buildWorkbook = (r = res) => {
    const wb = XLSX.utils.book_new();
    const sc = [
      { Metric: "Source", Value: srcName }, { Metric: "Tracker ID", Value: trackerId || "" }, { Metric: "Source programs", Value: srcProjects || "all" },
      { Metric: "Target programs", Value: tgtProjects || "all" }, { Metric: "Run date", Value: runDate },
      { Metric: "iTrack count", Value: r.summary.itrack_count }, { Metric: "Jira count", Value: r.summary.jira_count },
      { Metric: "Migrated", Value: r.summary.migrated }, { Metric: "Pending migration", Value: r.summary.pending_migration },
      { Metric: "Differences", Value: r.summary.differences }, { Metric: "Duplicates", Value: r.summary.duplicates },
      { Metric: "Jira only", Value: r.summary.jira_only }, { Metric: "Matched (clean)", Value: r.summary.matched },
      { Metric: "Coverage %", Value: r.summary.coverage }, { Metric: "Field accuracy %", Value: r.summary.field_accuracy },
      { Metric: "Overall score %", Value: r.summary.score },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sc), "Scorecard");
    Object.entries(r.details).forEach(([k, rows]) => {
      const name = k.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()).slice(0, 31);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ "": "none" }]), name);
    });
    return wb;
  };

  const safe = (x) => String(x || "").replace(/[\\/:*?"<>|]/g, "_").trim();
  const _d = new Date(), _p = (n) => String(n).padStart(2, "0");
  const runDate = `${_p(_d.getMonth() + 1)}${_p(_d.getDate())}${_d.getFullYear()}`;   // MMDDYYYY
  const projName = safe(tgtProjects.split(",")[0].trim() || "AllPrograms");
  const reportBase = () => `${safe(srcName)}_${projName}_Validation_Report_${runDate}`;

  const exportXlsx = () => {
    if (!res) return;
    const n = bumpVersion(`${safe(srcName)}|${projName}`);           // V2, V3… on re-runs of the same source+project
    const fname = (n > 1 ? `V${n}_` : "") + reportBase() + ".xlsx";
    try { XLSX.writeFile(buildWorkbook(), fname); } catch (e) { setError("Export blocked in this preview; works when hosted."); }
  };

  const exportView = () => {
    if (!rows || !rows.length) return;
    try {
      const url = URL.createObjectURL(new Blob([csvOf(rows)], { type: "text/csv" }));
      const a = document.createElement("a"); a.href = url; a.download = `${bucket}.csv`; a.click(); URL.revokeObjectURL(url);
    } catch (e) { setError("Export blocked in this preview; works when hosted."); }
  };

  const s = res?.summary;
  const scoreColor = s ? (s.score >= SCORING.pass ? RAMP.green : s.score >= SCORING.warn ? RAMP.amber : RAMP.red) : C.muted;
  const rows = res ? res.details[bucket] : [];
  const cols = rows && rows[0] ? Object.keys(rows[0]) : [];

  const nav = (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 28px", borderBottom: `1px solid ${C.border}`, background: C.bg }}>
      <div style={{ display: "flex", gap: 24 }}>
        {[["compare", "Run comparison"], ["report", "Migration report"], ["tracker", "Tracker"]].map(([k, l]) => (
          <button key={k} onClick={() => setView(k)} style={{ background: "none", border: "none", cursor: "pointer", padding: "6px 2px", fontSize: 14, color: view === k ? C.accent : C.muted, borderBottom: `2px solid ${view === k ? C.accent : "transparent"}`, fontWeight: view === k ? 600 : 500 }}>{l}</button>
        ))}
      </div>
      <select value={theme} onChange={(e) => setTheme(e.target.value)} title="Theme"
        style={{ padding: "6px 10px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, outline: "none" }}>
        <option value="dark">Dark</option>
        <option value="light">Light</option>
        <option value="cb">Colour-blind</option>
      </select>
    </div>
  );

  if (view === "report") {
    return <div style={{ background: C.bg, minHeight: "100vh" }}>{nav}<MigrationReportView /></div>;
  }
  if (view === "tracker") {
    return <div style={{ background: C.bg, minHeight: "100vh" }}>{nav}<TrackerView /></div>;
  }

  if (!source) {
    return (
      <div style={{ background: C.bg, color: C.text, minHeight: "100vh", fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
        {nav}
        <div style={{ padding: "40px 28px" }}>
        <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700 }}>Select a source system</h1>
        <p style={{ margin: "6px 0 26px", color: C.muted, fontSize: 14 }}>Each source uses its own status / priority mapping files, so pick the one you're comparing against Jira Cloud.</p>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          {SOURCES.map((sx) => (
            <button key={sx.id} onClick={() => pickSource(sx.id)}
              style={{ flex: 1, minWidth: 220, textAlign: "left", background: C.card, border: `1px solid ${C.border}`,
                borderRadius: 14, padding: "22px", cursor: "pointer", color: C.text }}>
              <div style={{ fontSize: 19, fontWeight: 700, color: C.accent }}>{sx.name}</div>
              <div style={{ color: C.faint, fontSize: 12.5, marginTop: 16 }}>Compare {sx.name} → Jira Cloud →</div>
            </button>
          ))}
        </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ background: C.bg, color: C.text, minHeight: "100vh", fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
      {nav}
      <div style={{ padding: "26px 28px" }}>
      <button onClick={() => pickSource(null)}
        style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "6px 12px", background: C.card, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13, cursor: "pointer", marginBottom: 14 }}>‹ Change source</button>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700 }}>Migration report — dry run</h1>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ color: C.faint, fontSize: 12 }}>Validator</span>
          <input value={validator} onChange={(e) => { setValidator(e.target.value); setUserName(e.target.value); }} placeholder="your name"
            style={{ padding: "7px 11px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, width: 140, outline: "none" }} />
        </div>
      </div>
      <p style={{ margin: "6px 0 20px", color: C.muted, fontSize: 14 }}>Source: <b style={{ color: C.accent }}>{srcName}</b> → Jira Cloud. Everything runs in your browser — nothing is uploaded.</p>

      <div style={{ marginBottom: 16, maxWidth: 340 }}>
        <div style={{ color: C.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>Tracker ID</div>
        <input value={trackerId} onChange={(e) => setTrackerId(e.target.value)} placeholder="e.g. JCJAMT-1234"
          style={{ width: "100%", boxSizing: "border-box", padding: "11px 13px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13.5, outline: "none", fontFamily: "ui-monospace, Menlo, monospace" }} />
      </div>

      {/* project keys — top */}
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 16 }}>
        <div style={{ flex: 1, minWidth: 300 }}>
          <div style={{ color: C.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>{source === "ado" ? "Source Area Path" : "Source program key(s)"}</div>
          <input value={srcProjects} onChange={(e) => setSrcProjects(e.target.value)} placeholder={source === "ado" ? "e.g. Program\\Team — optional, comma-separated" : "e.g. COMS, PSL — optional, comma-separated"}
            style={{ width: "100%", boxSizing: "border-box", padding: "11px 13px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13.5, outline: "none" }} />
        </div>
        <div style={{ flex: 1, minWidth: 300 }}>
          <div style={{ color: C.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>Target program key(s)</div>
          <input value={tgtProjects} onChange={(e) => setTgtProjects(e.target.value)} placeholder="e.g. CPRI, PRESALE — optional, comma-separated"
            style={{ width: "100%", boxSizing: "border-box", padding: "11px 13px", background: C.field, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13.5, outline: "none" }} />
        </div>
      </div>

      {/* export files */}
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 16 }}>
        <MultiFileDrop label={`Source — ${srcName} export(s)`} files={srcFiles}
          onAdd={(fs) => setSrcFiles((p) => [...p, ...fs])} onRemove={(i) => setSrcFiles((p) => p.filter((_, j) => j !== i))} />
        <MultiFileDrop label="Target — Jira Cloud export(s)" files={tgtFiles}
          onAdd={(fs) => setTgtFiles((p) => [...p, ...fs])} onRemove={(i) => setTgtFiles((p) => p.filter((_, j) => j !== i))} />
      </div>

      {/* per-source value mappings + shared user list, shown as attachments */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ color: C.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>
          Reference files <span style={{ textTransform: "none", letterSpacing: 0, color: C.faint }}>· mappings preloaded for {srcName} · user list common to all sources</span>
        </div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          {MAPPABLE.map((f) => (
            <div key={f} style={{ flex: 1, minWidth: 210 }}>
              {mappings[f] ? (
                <div style={{ border: `1px solid ${C.accent}`, borderRadius: 10, padding: "12px 14px", background: C.field }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontSize: 13, fontWeight: 600, textTransform: "capitalize" }}>{f.replace("_", " ")} mapping</span>
                    <button onClick={() => removeMapping(f)} title="Remove" aria-label="Remove"
                      style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 16, lineHeight: 1, padding: 0 }}>×</button>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, color: C.accent, fontSize: 12, marginTop: 5 }}>
                    <Clip /> {mappingNames[f]} · {Object.keys(mappings[f]).length} rules
                  </div>
                  <div style={{ display: "flex", gap: 12, marginTop: 8 }}>
                    <button onClick={() => viewMapping(f)} style={{ background: "none", border: "none", color: C.accent, cursor: "pointer", fontSize: 12, padding: 0, textDecoration: "underline" }}>View</button>
                    <button onClick={() => downloadRows(mappingRows(f), mappingNames[f])} style={{ background: "none", border: "none", color: C.accent, cursor: "pointer", fontSize: 12, padding: 0, textDecoration: "underline" }}>Download</button>
                  </div>
                </div>
              ) : (
                <label style={{ display: "block", cursor: "pointer" }}>
                  <div style={{ border: `1px dashed ${C.border}`, borderRadius: 10, padding: "12px 14px", background: C.field, color: C.muted }}>
                    <div style={{ fontSize: 13, fontWeight: 600, textTransform: "capitalize", color: C.text }}>{f.replace("_", " ")} mapping</div>
                    <div style={{ fontSize: 12, marginTop: 4 }}>Not loaded — click to upload CSV</div>
                  </div>
                  <input type="file" accept=".csv" style={{ display: "none" }}
                    onChange={(e) => { const file = e.target.files[0]; if (file) uploadMapping(f, file); e.target.value = ""; }} />
                </label>
              )}
            </div>
          ))}
          <div style={{ flex: 1, minWidth: 210 }}>
            {userMap ? (
              <div style={{ border: `1px solid ${C.accent}`, borderRadius: 10, padding: "12px 14px", background: C.field }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 13, fontWeight: 600 }}>User list <span style={{ color: C.faint, fontWeight: 400 }}>(shared)</span></span>
                  <button onClick={removeUser} title="Remove" aria-label="Remove"
                    style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 16, lineHeight: 1, padding: 0 }}>×</button>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 6, color: C.accent, fontSize: 12, marginTop: 5 }}>
                  <Clip /> {userName} · {new Set(Object.values(userMap)).size} users
                </div>
                <div style={{ display: "flex", gap: 12, marginTop: 8 }}>
                  <button onClick={viewUsers} style={{ background: "none", border: "none", color: C.accent, cursor: "pointer", fontSize: 12, padding: 0, textDecoration: "underline" }}>View</button>
                  <button onClick={() => downloadRows(userRows(), userName)} style={{ background: "none", border: "none", color: C.accent, cursor: "pointer", fontSize: 12, padding: 0, textDecoration: "underline" }}>Download</button>
                </div>
              </div>
            ) : (
              <label style={{ display: "block", cursor: "pointer" }}>
                <div style={{ border: `1px dashed ${C.border}`, borderRadius: 10, padding: "12px 14px", background: C.field, color: C.muted }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.text }}>User list</div>
                  <div style={{ fontSize: 12, marginTop: 4 }}>Not loaded — click to upload Excel/CSV</div>
                </div>
                <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
                  onChange={(e) => { const file = e.target.files[0]; if (file) loadUserList(file); e.target.value = ""; }} />
              </label>
            )}
          </div>
        </div>
      </div>

      {refView && (
        <div onClick={() => setRefView(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 14, width: "min(560px, 100%)", maxHeight: "80vh", display: "flex", flexDirection: "column", overflow: "hidden" }}>
            <div style={{ padding: "14px 18px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
              <div>
                <div style={{ fontSize: 15, fontWeight: 700 }}>{refView.title}</div>
                <div style={{ fontSize: 11.5, color: C.faint, marginTop: 2 }}>{refView.sub} · {refView.rows.length} {refView.cols[0] === "Alias" ? "users" : "rules"}</div>
              </div>
              <button onClick={() => setRefView(null)} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 20, lineHeight: 1, padding: 0 }}>×</button>
            </div>
            <div style={{ overflow: "auto" }}>
              <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12.5 }}>
                <thead><tr>{refView.cols.map((c) => <th key={c} style={{ position: "sticky", top: 0, background: C.head, textAlign: "left", padding: "9px 14px", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.04em", color: C.muted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{c}</th>)}</tr></thead>
                <tbody>
                  {refView.rows.map((r, i) => <tr key={i}>{refView.cols.map((c) => <td key={c} style={{ padding: "7px 14px", borderBottom: `1px solid ${C.border}`, background: i % 2 ? C.rowAlt : C.card, fontFamily: c === "ATTUID" || c === "Source value" ? "ui-monospace, monospace" : "inherit" }}>{r[c]}</td>)}</tr>)}
                </tbody>
              </table>
            </div>
            <div style={{ padding: "12px 18px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <button onClick={() => downloadRows(refView.rows, refView.sub || refView.title)} style={{ padding: "8px 15px", background: C.accent + "18", color: C.accent, border: `1px solid ${C.accent}66`, borderRadius: 9, fontSize: 13, cursor: "pointer" }}>Download .xlsx</button>
              <button onClick={() => setRefView(null)} style={{ padding: "8px 15px", background: C.card, color: C.text, border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13, cursor: "pointer" }}>Close</button>
            </div>
          </div>
        </div>
      )}

      {/* run */}
      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 8 }}>
        <button onClick={run} disabled={!srcFiles.length || !tgtFiles.length}
          style={{ padding: "10px 24px", background: (srcFiles.length && tgtFiles.length) ? C.accent : C.card, color: "#fff", border: "none",
            borderRadius: 9, fontSize: 14, fontWeight: 600, cursor: (srcFiles.length && tgtFiles.length) ? "pointer" : "default", opacity: (srcFiles.length && tgtFiles.length) ? 1 : 0.6 }}>
          Run comparison</button>
        {res && <button onClick={exportXlsx}
          style={{ display: "flex", alignItems: "center", gap: 7, padding: "10px 20px", background: C.accent + "18", color: C.accent, border: `1px solid ${C.accent}66`, borderRadius: 9, fontSize: 13.5, cursor: "pointer" }}>
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round"><path d="M12 3v12m0 0 4-4m-4 4-4-4M4 21h16" /></svg>
          Download validation report</button>}
        <label style={{ display: "inline-flex", alignItems: "center", gap: 7, color: C.muted, fontSize: 12.5, cursor: "pointer", marginLeft: "auto" }}>
          <input type="checkbox" checked={autoAttach} onChange={(e) => { setAutoAttach(e.target.checked); try { localStorage.setItem("cc.autoAttachReport", e.target.checked ? "1" : "0"); } catch { /* ignore */ } }} />
          Attach report to its Tracker ticket
        </label>
      </div>
      {attachMsg && <div style={{ color: attachMsg.startsWith("Report attached") ? RAMP.green : C.muted, fontSize: 12.5, marginBottom: 8 }}>{attachMsg}</div>}
      {res && (
        <div style={{ color: C.faint, fontSize: 12.5, marginBottom: 8 }}>
          File: <span style={{ color: C.muted, fontFamily: "ui-monospace, Menlo, monospace" }}>{reportBase()}.xlsx</span> · re-runs of the same source + program add V2, V3…
        </div>
      )}
      {error && <div style={{ color: RAMP.red, fontSize: 13, marginTop: 6 }}>{error}</div>}

      {s && (
        <div style={{ marginTop: 22 }}>
          {/* score + rectification */}
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 16 }}>
            <div style={{ minWidth: 200, background: C.card, border: `1px solid ${scoreColor}55`, borderRadius: 14, padding: "18px 22px" }}>
              <div style={{ color: C.muted, fontSize: 13 }}>Overall migration score</div>
              <div style={{ fontSize: 44, fontWeight: 800, color: scoreColor, lineHeight: 1.1 }}>{s.score}%</div>
              <div style={{ color: C.faint, fontSize: 12 }}>coverage {s.coverage}% · field accuracy {s.field_accuracy}%</div>
            </div>
            <div style={{ flex: 1, minWidth: 300, background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "16px 20px" }}>
              <div style={{ fontWeight: 600, marginBottom: 8 }}>To reach 100%</div>
              <ul style={{ margin: 0, paddingLeft: 18, color: C.muted, fontSize: 13.5, lineHeight: 1.7 }}>
                <li>Migrate <b style={{ color: C.text }}>{s.pending_migration}</b> pending item(s)</li>
                <li>Resolve field differences on <b style={{ color: C.text }}>{s.differences}</b> issue(s)</li>
                <li>De-duplicate <b style={{ color: C.text }}>{s.duplicates}</b> conflicting Cloud issue(s)</li>
                <li>Review <b style={{ color: C.text }}>{s.jira_only}</b> Jira-only issue(s)</li>
              </ul>
            </div>
          </div>

          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 16 }}>
            <Card title="iTrack count" value={s.itrack_count} color={RAMP.blue} desc="Issues in the source export" />
            <Card title="Jira count" value={s.jira_count} color={RAMP.purple} desc="Issues in the target export" />
            <Card title="Migrated" value={s.migrated} color={RAMP.green} desc="Source issues matched in target" />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))", gap: 12, marginBottom: 18 }}>
            {TILES.map(([k, label, ck]) => (
              <button key={k} onClick={() => setBucket(k)}
                style={{ textAlign: "left", background: C.card, border: `1.5px solid ${bucket === k ? C.accent : C.border}`,
                  borderRadius: 12, padding: "13px 15px", cursor: "pointer" }}>
                <div style={{ fontSize: 22, fontWeight: 700, color: RAMP[ck] }}>{s[k].toLocaleString()}</div>
                <div style={{ fontSize: 12.5, color: C.muted, marginTop: 2 }}>{label}</div>
              </button>
            ))}
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontSize: 15, fontWeight: 600 }}>{TILES.find((t) => t[0] === bucket)?.[1]} · {rows.length} rows</div>
            {rows.length > 0 && (
              <button onClick={exportView}
                style={{ display: "flex", alignItems: "center", gap: 7, padding: "8px 15px", background: C.card, color: C.text,
                  border: `1px solid ${C.border}`, borderRadius: 9, fontSize: 13, cursor: "pointer" }}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round"><path d="M12 3v12m0 0 4-4m-4 4-4-4M4 21h16" /></svg>
                Export this view
              </button>
            )}
          </div>
          {rows.length === 0 ? (
            <div style={{ padding: 30, textAlign: "center", color: C.muted, background: C.card, border: `1px solid ${C.border}`, borderRadius: 12 }}>Nothing here — clean.</div>
          ) : (
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 12, overflow: "auto", maxHeight: 420 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 620 }}>
                <thead><tr>{cols.map((c) => (
                  <th key={c} style={{ textAlign: "left", padding: "10px 13px", fontSize: 11.5, fontWeight: 600, color: C.muted,
                    textTransform: "uppercase", letterSpacing: "0.04em", background: C.head, borderBottom: `1px solid ${C.border}`, position: "sticky", top: 0 }}>{c}</th>
                ))}</tr></thead>
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={i} style={{ background: i % 2 ? C.rowAlt : "transparent" }}>
                      {cols.map((c) => {
                        const v = r[c];
                        const isKey = c.includes("key") || c === "Matched with";
                        const isExpected = c.startsWith("Expected") || c === "Target";
                        const color = isExpected ? RAMP.green : isKey ? C.accent : (v === "" || v == null ? C.faint : C.text);
                        return <td key={c} style={{ padding: "10px 13px", fontSize: 12.5, borderBottom: `1px solid ${C.border}`,
                          color, fontFamily: isKey ? "ui-monospace, Menlo, monospace" : "inherit" }}>{v === "" || v == null ? "—" : v}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
      </div>
    </div>
  );
}


ReactDOM.createRoot(document.getElementById("root")).render(<JiraMigrationV1 />);
