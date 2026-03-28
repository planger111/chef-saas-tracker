// Chef SaaS Motion Tracker - SharePoint CSV file backend
// Assignments read from assignments.json, engagements written to per-play CSV files
// Only needs Sites.ReadWrite.All (already granted)

let _siteId = null;
let _driveId = null;

async function _graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : (CONFIG.graphBaseUrl + path);
  const resp = await fetch(url, { ...options, headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", ...(options.headers || {}) } });
  if (!resp.ok) { const t = await resp.text(); throw new Error("Graph " + resp.status + ": " + t); }
  if (resp.status === 204) return null;
  return resp.json();
}

async function getSiteId() {
  if (_siteId) return _siteId;
  const url = new URL(CONFIG.sharepointSiteUrl);
  const data = await _graphFetch("/sites/" + url.hostname + ":" + url.pathname);
  _siteId = data.id; return _siteId;
}

async function getDriveId() {
  if (_driveId) return _driveId;
  const siteId = await getSiteId();
  const data = await _graphFetch("/sites/" + siteId + "/drive");
  _driveId = data.id; return _driveId;
}

async function putFile(filePath, content, contentType) {
  const driveId = await getDriveId();
  const token = await getAccessToken();
  const resp = await fetch(CONFIG.graphBaseUrl + "/drives/" + driveId + "/root:/" + filePath + ":/content", {
    method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": contentType || "text/plain" }, body: content
  });
  if (!resp.ok) { const t = await resp.text(); throw new Error("Upload " + resp.status + ": " + t); }
  return resp.json();
}

async function getFileText(filePath) {
  const driveId = await getDriveId();
  const token = await getAccessToken();
  try {
    const meta = await _graphFetch("/drives/" + driveId + "/root:/" + filePath);
    const resp = await fetch(meta["@microsoft.graph.downloadUrl"], { headers: { Authorization: "Bearer " + token } });
    if (!resp.ok) return null;
    return resp.text();
  } catch(e) { return null; }
}

async function writeJsonFile(filePath, data) {
  return putFile("ChefSaaS/" + filePath, JSON.stringify(data, null, 2), "application/json");
}

// ─── Assignments ──────────────────────────────────────────────────────────

async function getPlayAssignments(repEmail) {
  const text = await getFileText("ChefSaaS/assignments.json");
  if (!text) return [];
  const data = JSON.parse(text);
  const list = (data.assignments || []).filter(a => a.active_flag !== false);
  if (!repEmail) return list;
  // Try exact match first
  const matched = list.filter(a => (a.rep_email || "").toLowerCase() === repEmail.toLowerCase());
  console.log("[ChefSaaS] Signed in as: " + repEmail + " — matched " + matched.length + " accounts");
  return matched;
}

async function getRepAccounts(repEmail) { return getPlayAssignments(repEmail); }

async function getPlays() {
  try {
    const all = await getPlayAssignments(null);
    const seen = new Set();
    return all.filter(a => { if (!a.play_id || seen.has(a.play_id)) return false; seen.add(a.play_id); return true; })
              .map(a => ({ play_id: a.play_id, play_name: a.play_name || a.play_id }));
  } catch(e) { return [{ play_id: "chef-saas", play_name: "Chef SaaS" }]; }
}

// ─── Response Options ─────────────────────────────────────────────────────

async function getResponseOptions(responseSetId, outcomeType) {
  const text = await getFileText("ChefSaaS/response-options.json");
  if (!text) return [];
  const data = JSON.parse(text);
  return (data.options || [])
    .filter(o => o.active_flag !== false &&
      (!responseSetId || (o.response_set_id || "").toUpperCase() === responseSetId.toUpperCase()) &&
      (!outcomeType || o.outcome_type === outcomeType))
    .sort((a, b) => (a.sort_order || 0) - (b.sort_order || 0))
    .map(o => o.reason_label);
}

// ─── Points ───────────────────────────────────────────────────────────────

function calculatePoints(outcome, nextStepType) {
  if (outcome === "Not Interested") return 15;
  if (outcome === "Interested") return nextStepType ? 45 : 30;
  return 0;
}

// ─── CSV helpers ──────────────────────────────────────────────────────────

const CSV_HEADERS = ["id","submitted_at","play_id","play_name","account_id","account_name","rep_email","rep_name","interaction_type","outcome","reason_label","next_step_type","timing","contact_engaged","opportunity_id","pitch_confidence","short_reaction","notes","points_earned"];

function csvEscape(val) {
  val = val !== undefined && val !== null ? String(val) : "";
  if (val.includes(",") || val.includes('"') || val.includes("\n")) val = '"' + val.replace(/"/g, '""') + '"';
  return val;
}

function toCsvRow(fields) { return CSV_HEADERS.map(h => csvEscape(fields[h])).join(","); }

function parseCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map(h => h.trim());
  return lines.slice(1).map(line => {
    const vals = [];
    let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"' && !inQ) { inQ = true; }
      else if (c === '"' && inQ && line[i+1] === '"') { cur += '"'; i++; }
      else if (c === '"' && inQ) { inQ = false; }
      else if (c === ',' && !inQ) { vals.push(cur); cur = ""; }
      else cur += c;
    }
    vals.push(cur);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || "").trim(); });
    return obj;
  });
}

// ─── Log engagement (append row to play CSV) ──────────────────────────────

async function logEngagement(fields) {
  const pts = calculatePoints(fields.outcome, fields.next_step_type);
  const entry = { ...fields, id: Date.now().toString(), submitted_at: new Date().toISOString(), points_earned: pts };
  const playId = (fields.play_id || "engagements").replace(/[^a-z0-9]/gi, "_");
  const filePath = "ChefSaaS/" + playId + "_engagements.csv";
  const existing = await getFileText(filePath);
  const csv = (existing && existing.trim()) ? existing.trimEnd() + "\n" + toCsvRow(entry) : CSV_HEADERS.join(",") + "\n" + toCsvRow(entry);
  await putFile(filePath, csv, "text/csv");
  return entry;
}

// ─── Read engagements ─────────────────────────────────────────────────────

async function _getAllLogsForRep(repEmail) {
  const plays = await getPlays();
  let allLogs = [];
  for (const play of plays) {
    const playId = play.play_id.replace(/[^a-z0-9]/gi, "_");
    const text = await getFileText("ChefSaaS/" + playId + "_engagements.csv");
    if (!text) continue;
    const rows = parseCSV(text).filter(r => !repEmail || (r.rep_email || "").toLowerCase() === repEmail.toLowerCase());
    allLogs = allLogs.concat(rows);
  }
  return allLogs;
}

async function getLatestLogsForRep(repEmail, accountIds) {
  const logs = await _getAllLogsForRep(repEmail);
  const latest = {};
  logs.forEach(log => {
    const key = log.account_id;
    if (!key) return;
    if (!latest[key] || log.submitted_at > (latest[key].submitted_at || "")) latest[key] = log;
  });
  if (accountIds) { const r = {}; accountIds.forEach(id => { r[id] = latest[id] || null; }); return r; }
  return latest;
}

async function getLatestExecutionPerAccount(playId, accountIds) {
  const user = getCurrentUser();
  return getLatestLogsForRep(user ? user.email : null, accountIds);
}

async function getRepPoints(repEmail) {
  const logs = await _getAllLogsForRep(repEmail);
  return logs.reduce((sum, l) => sum + (Number(l.points_earned) || 0), 0);
}

async function getLeaderboard() {
  try {
    const plays = await getPlays();
    const repMap = {};
    for (const play of plays) {
      const playId = play.play_id.replace(/[^a-z0-9]/gi, "_");
      const text = await getFileText("ChefSaaS/" + playId + "_engagements.csv");
      if (!text) continue;
      parseCSV(text).forEach(log => {
        const key = log.rep_email; if (!key) return;
        if (!repMap[key]) repMap[key] = { email: key, repName: log.rep_name || key, totalPoints: 0, engagedAccounts: new Set(), interestedAccounts: new Set() };
        repMap[key].totalPoints += Number(log.points_earned) || 0;
        if (log.account_id) { repMap[key].engagedAccounts.add(log.account_id); if (log.outcome === "Interested") repMap[key].interestedAccounts.add(log.account_id); }
      });
    }
    return Object.values(repMap).map(r => ({ ...r, engagedAccounts: r.engagedAccounts.size, interestedAccounts: r.interestedAccounts.size }))
                                .sort((a, b) => b.totalPoints - a.totalPoints);
  } catch(e) { return []; }
}
