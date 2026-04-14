// INFRA SalesPlay Tracker - SharePoint CSV file backend
// Assignments read from assignments.json, engagements written to per-play CSV files
// Only needs Sites.ReadWrite.All (already granted)

// Root folder inside SharePoint Documents — matches the OneDrive-synced local folder
const SP_ROOT = "Chef SaaS Tracker/ChefSaaS";

let _siteId = null;
let _driveId = null;

// ─── Section 2: Identity ──────────────────────────────────────────────────
// rep-identity.json lives in SharePoint only — never in the GitHub repo.
// Schema per rep: { oid, upn, display_name, aliases[], role, active, last_seen }

let _repIdentityMap = null;

async function getRepIdentityMap() {
  if (_repIdentityMap) return _repIdentityMap;
  try {
    const text = await getFileText(SP_ROOT + "/rep-identity.json");
    if (text) _repIdentityMap = JSON.parse(text);
  } catch(e) {
    console.warn("[Identity] Could not load rep-identity.json:", e.message);
    _repIdentityMap = { reps: [], pending: [] };
  }
  return _repIdentityMap;
}

// Match an email address against all aliases[] in the identity map.
// Returns the rep entry or null.
async function resolveRepByEmail(email) {
  if (!email) return null;
  const map = await getRepIdentityMap();
  const lower = email.toLowerCase().trim();
  return (map.reps || []).find(r => (r.aliases || []).some(a => a.toLowerCase() === lower)) || null;
}

// Match a display name using last-name-first scoring.
// Returns { match: repEntry|null, ambiguous: bool, candidates: [] }
async function resolveRepByName(fullName) {
  if (!fullName) return { match: null, ambiguous: false, candidates: [] };
  const map = await getRepIdentityMap();
  const parts = fullName.trim().toLowerCase().split(/\s+/);
  const lastName = parts[parts.length - 1];
  const firstName = parts[0];

  // Step 1: filter to last-name matches
  const lastMatches = (map.reps || []).filter(r => {
    const rParts = r.display_name.toLowerCase().split(/\s+/);
    return rParts[rParts.length - 1] === lastName;
  });
  if (lastMatches.length === 0) return { match: null, ambiguous: false, candidates: [] };

  // Step 2: score first name similarity
  function firstScore(rep) {
    const rFirst = rep.display_name.toLowerCase().split(/\s+/)[0];
    if (rFirst === firstName) return 100;
    if (rFirst.startsWith(firstName) || firstName.startsWith(rFirst)) return 80;
    if (rFirst.slice(0, 3) === firstName.slice(0, 3)) return 60;
    return 0;
  }
  const scored = lastMatches.map(r => ({ rep: r, score: firstScore(r) })).filter(s => s.score >= 60);
  if (scored.length === 0) return { match: null, ambiguous: false, candidates: lastMatches };
  if (scored.length === 1) return { match: scored[0].rep, ambiguous: false, candidates: [] };
  // Multiple above threshold — ambiguous, admin must pick
  return { match: null, ambiguous: true, candidates: scored.map(s => s.rep) };
}

// Write the Entra OID to a rep's identity map entry after first login.
// Read-check-write to avoid race: only writes if oid is currently empty.
async function registerOID(upn, oid) {
  if (!upn || !oid) return;
  try {
    const map = await getRepIdentityMap();
    const rep = (map.reps || []).find(r => (r.aliases || []).some(a => a.toLowerCase() === upn.toLowerCase()));
    if (!rep) return; // not in map — write to pending[] handled separately
    if (rep.oid && rep.oid === oid) return; // already registered, no-op
    rep.oid = oid;
    rep.last_seen = new Date().toISOString();
    await writeJsonFile("rep-identity.json", map);
    _repIdentityMap = map; // update cache
  } catch(e) {
    console.warn("[Identity] registerOID failed (non-blocking):", e.message);
  }
}

// ─── Section 3: Activity Logging ─────────────────────────────────────────
// Appends to activity_log.csv in SharePoint. Always non-blocking — never throws.

const ACTIVITY_LOG_HEADERS = ["timestamp","user_email","user_name","action_type","detail"];

async function logActivity(user, actionType, detail) {
  try {
    if (!user || !user.username) return;
    const row = [
      new Date().toISOString(),
      user.username,
      user.name || user.username,
      actionType || '',
      detail || ''
    ].map(v => csvEscape(String(v))).join(',');
    const filePath = SP_ROOT + "/activity_log.csv";
    let existing = '';
    try { existing = await getFileText(filePath) || ''; } catch(e) {}
    const csv = existing.trim()
      ? existing.trimEnd() + '\n' + row
      : ACTIVITY_LOG_HEADERS.join(',') + '\n' + row;
    await putFile(filePath, csv, 'text/csv');
  } catch(e) {
    console.warn("[Activity] logActivity failed (non-blocking):", e.message);
  }
}




async function _graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : (CONFIG.graphBaseUrl + path);
  const resp = await fetch(url, { ...options, headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", ...(options.headers || {}) } });
  if (resp.status === 401) {
    // Token rejected by Graph — force a fresh token then retry once
    _driveId = null; _siteId = null;
    const freshToken = await getAccessToken();
    const retry = await fetch(url, { ...options, headers: { Authorization: "Bearer " + freshToken, "Content-Type": "application/json", ...(options.headers || {}) } });
    if (!retry.ok) { const t = await retry.text(); throw new Error("Graph " + retry.status + ": " + t); }
    if (retry.status === 204) return null;
    return retry.json();
  }
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
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    const resp = await fetch(CONFIG.graphBaseUrl + "/drives/" + driveId + "/root:/" + filePath + ":/content", {
      method: "PUT", signal: controller.signal,
      headers: { Authorization: "Bearer " + token, "Content-Type": contentType || "text/plain" }, body: content
    });
    if (!resp.ok) { const t = await resp.text(); throw new Error("Upload " + resp.status + ": " + t); }
    return resp.json();
  } catch(e) {
    if (e.name === 'AbortError') throw new Error("Save timed out — check your connection and try again.");
    throw e;
  } finally {
    clearTimeout(timeout);
  }
}

async function getFileText(filePath) {
  const driveId = await getDriveId();
  const token = await getAccessToken();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const meta = await fetch(CONFIG.graphBaseUrl + "/drives/" + driveId + "/root:/" + filePath, {
      signal: controller.signal,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }
    });
    if (!meta.ok) return null;
    const metaJson = await meta.json();
    const resp = await fetch(metaJson["@microsoft.graph.downloadUrl"], {
      signal: controller.signal,
      headers: { Authorization: "Bearer " + token }
    });
    if (!resp.ok) return null;
    return resp.text();
  } catch(e) { return null; } finally { clearTimeout(timeout); }
}

async function writeJsonFile(filePath, data) {
  return putFile(SP_ROOT + "/" + filePath, JSON.stringify(data, null, 2), "application/json");
}

// ─── Assignments ──────────────────────────────────────────────────────────

async function getAllAssignedReps() {
  const text = await getFileText(SP_ROOT + "/assignments.json");
  if (!text) return [];
  const data = JSON.parse(text);
  const map = {};
  (data.assignments || []).filter(a => a.active_flag !== false).forEach(a => {
    const email = (a.rep_sso_login || a.rep_email || '').toLowerCase();
    if (email && !map[email]) map[email] = { email, repName: a.rep_name || email };
  });
  return Object.values(map);
}

async function getPlayAssignments(repEmail) {
  const text = await getFileText(SP_ROOT + "/assignments.json");
  if (!text) return [];
  const data = JSON.parse(text);
  const list = (data.assignments || []).filter(a => a.active_flag !== false);
  if (!repEmail) return list;

  const emailLower = repEmail.toLowerCase();

  // Tier 1: match by registered OID
  const user = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
  if (user && user.oid) {
    const byOid = list.filter(a => a.rep_oid && a.rep_oid === user.oid);
    if (byOid.length > 0) {
      console.log("[ChefSaaS] Matched " + byOid.length + " accounts via OID");
      return byOid;
    }
  }

  // Tier 2: resolve via identity map aliases
  try {
    const repEntry = await resolveRepByEmail(emailLower);
    if (repEntry) {
      const aliases = (repEntry.aliases || []).map(a => a.toLowerCase());
      const byAlias = list.filter(a =>
        aliases.includes((a.rep_sso_login || '').toLowerCase()) ||
        aliases.includes((a.rep_email || '').toLowerCase())
      );
      if (byAlias.length > 0) {
        console.log("[ChefSaaS] Matched " + byAlias.length + " accounts via identity map for " + repEntry.display_name);
        return byAlias;
      }
    }
  } catch(e) {
    console.warn("[ChefSaaS] Identity map lookup failed, falling back:", e.message);
  }

  // Tier 3: direct match — try ALL email candidates from the token
  const user3 = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
  const emailCandidates = (user3?.emails?.length) ? user3.emails : [emailLower];
  let matched = list.filter(a => {
    const repLogin = (a.rep_sso_login || '').toLowerCase();
    const repEmail = (a.rep_email    || '').toLowerCase();
    return emailCandidates.some(e => e === repLogin || e === repEmail);
  });
  console.log("[ChefSaaS] Direct match candidates:", emailCandidates, "→", matched.length, "accounts");
  return matched;
}

async function getRepAccounts(repEmail) { return getPlayAssignments(repEmail); }

// ─── Plays config (plays.json) ────────────────────────────────────────────

async function getPlaysConfig() {
  try {
    const text = await getFileText(SP_ROOT + "/plays.json");
    if (text) return JSON.parse(text);
  } catch(e) {}
  // Fall back to deriving plays from assignments
  const all = await getPlayAssignments(null);
  const seen = new Set();
  const plays = [];
  all.filter(a => a.play_id).forEach(a => {
    if (!seen.has(a.play_id)) {
      seen.add(a.play_id);
      plays.push({ play_id: a.play_id, play_name: a.play_name || a.play_id, goal: "qualify_in", icp: "", target_persona: "", elevator_pitch: a.elevator_pitch || "", active: true });
    }
  });
  return { plays };
}

async function savePlaysConfig(config) {
  await writeJsonFile("plays.json", config);
}

async function getPlays() {
  try {
    const config = await getPlaysConfig();
    if (config && config.plays && config.plays.length > 0) return config.plays.filter(p => p.active !== false);
    // fallback: derive from assignments
    const all = await getPlayAssignments(null);
    const seen = new Set();
    return all.filter(a => { if (!a.play_id || seen.has(a.play_id)) return false; seen.add(a.play_id); return true; })
              .map(a => ({ play_id: a.play_id, play_name: a.play_name || a.play_id }));
  } catch(e) { return [{ play_id: "chef-saas", play_name: "Chef SaaS" }]; }
}

// ─── Response Options ─────────────────────────────────────────────────────

async function getResponseOptions(responseSetId, outcomeType) {
  const text = await getFileText(SP_ROOT + "/response-options.json");
  if (!text) return [];
  const data = JSON.parse(text);
  return (data.options || [])
    .filter(o => o.active_flag !== false &&
      (!responseSetId || (o.response_set_id || "").toUpperCase() === responseSetId.toUpperCase()) &&
      (!outcomeType || o.outcome_type === outcomeType))
    .sort((a, b) => (a.sort_order || 0) - (b.sort_order || 0))
    .map(o => o.reason_label);
}

const _DEFAULT_FORM_CONFIG = {
  interaction_types: ["Call", "Email", "Text", "In-Person Meeting"],
  next_step_types:   ["Schedule follow-up", "Send information", "Introduce stakeholder", "Create opportunity", "Revisit later"],
  timing_options:    ["Now", "This month", "Next quarter", "Later"]
};

async function getFormConfig() {
  try {
    const text = await getFileText(SP_ROOT + "/response-options.json");
    if (!text) return _DEFAULT_FORM_CONFIG;
    const data = JSON.parse(text);
    return {
      interaction_types: (data.form_config && data.form_config.interaction_types) || _DEFAULT_FORM_CONFIG.interaction_types,
      next_step_types:   (data.form_config && data.form_config.next_step_types)   || _DEFAULT_FORM_CONFIG.next_step_types,
      timing_options:    (data.form_config && data.form_config.timing_options)     || _DEFAULT_FORM_CONFIG.timing_options,
    };
  } catch(e) { return _DEFAULT_FORM_CONFIG; }
}

async function getFullResponseConfig() {
  try {
    const text = await getFileText(SP_ROOT + "/response-options.json");
    return text ? JSON.parse(text) : { options: [], form_config: _DEFAULT_FORM_CONFIG };
  } catch(e) { return { options: [], form_config: _DEFAULT_FORM_CONFIG }; }
}

async function saveFullResponseConfig(config) {
  await writeJsonFile("response-options.json", config);
}

// ─── Section 5: Scoring ───────────────────────────────────────────────────

// Score notes based on content signals, not character count.
// Returns bonus points 0–40. Never exposed in UI.
function scoreNotes(notes) {
  if (!notes || notes.trim().length < 10) return 0;
  const text = notes.toLowerCase();
  let score = 0;

  const competitors = ['ansible','puppet','terraform','servicenow','saltstack','jenkins',
                       'harness','octopus','github actions','gitlab','azure devops'];
  const competitorHits = competitors.filter(c => text.includes(c));
  score += competitorHits.length * 8;

  const budgetSignals = ['budget','funded','approved','allocated','q1','q2','q3','q4','fiscal'];
  if (budgetSignals.some(s => text.includes(s))) score += 8;

  const stakeholderSignals = ['cto','ciso','vp ','director','executive','sponsor','champion'];
  if (stakeholderSignals.some(s => text.includes(s))) score += 6;

  const urgencySignals = ['urgent','asap','priority','this quarter','next quarter','eoy'];
  if (urgencySignals.some(s => text.includes(s))) score += 5;

  const riskSignals = ['concern','risk','blocker','hesitant','pushback','compliance'];
  if (riskSignals.some(s => text.includes(s))) score += 5;

  // Minimum reward for substantive notes with no detected signals
  if (score === 0 && notes.trim().length >= 30) score += 2;

  return score;
}

function calculatePoints(outcome, nextStepType, data) {
  // Base points
  let base = 0;
  if (outcome === 'Not Interested') base = 15;
  else if (outcome === 'Interested') base = nextStepType ? 45 : 30;

  // Hidden bonus layer
  let bonus = 0;
  if (data) {
    const notes = (data.notes || '').trim();
    const reasonLabel = (data.reason_label || '').toLowerCase();
    const nextStep = (data.next_step_type || '').toLowerCase();

    // Notes: intelligence-driven scoring (replaces flat +5 for length)
    bonus += scoreNotes(notes);

    // Strong interest signals
    if (reasonLabel.includes('active budget') || reasonLabel.includes('executive sponsor')) bonus += 10;
    // Strong next step
    if (nextStep.includes('schedule follow-up')) bonus += 10;
    // Competitive intel in reason
    if (reasonLabel.includes('already have solution') || reasonLabel.includes('already solved')) bonus += 8;
  }

  return base + bonus;
}

// ─── CSV helpers ──────────────────────────────────────────────────────────

const CSV_HEADERS = ["id","submitted_at","play_id","play_name","account_id","account_name","rep_email","rep_name","interaction_type","outcome","reason_label","reason_code","next_step_type","timing","contact_level","notes","points_earned"];

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
  const pts = calculatePoints(fields.outcome, fields.next_step_type, fields);
  const entry = { ...fields, id: Date.now().toString(), submitted_at: new Date().toISOString(), points_earned: pts };
  const playId = (fields.play_id || "engagements").replace(/[^a-z0-9]/gi, "_");
  const filePath = SP_ROOT + "/" + playId + "_engagements.csv";
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
    const text = await getFileText(SP_ROOT + "/" + playId + "_engagements.csv");
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
      const text = await getFileText(SP_ROOT + "/" + playId + "_engagements.csv");
      if (!text) continue;
      parseCSV(text).forEach(log => {
        const key = log.rep_email; if (!key) return;
        if (!repMap[key]) repMap[key] = { email: key, repName: log.rep_name || key, totalPoints: 0, engagedAccounts: new Set(), interestedAccounts: new Set() };
        repMap[key].totalPoints += Number(log.points_earned) || 0;
        if (log.account_id) { repMap[key].engagedAccounts.add(log.account_id); if (log.outcome === "Interested") repMap[key].interestedAccounts.add(log.account_id); }
      });
    }
    return Object.values(repMap).map(r => ({ ...r, completedAccounts: r.engagedAccounts.size, engagedAccounts: r.engagedAccounts.size, interestedAccounts: r.interestedAccounts.size }))
                                .sort((a, b) => b.totalPoints - a.totalPoints);
  } catch(e) { return []; }
}

// ─── All engagements (admin view) ────────────────────────────────────────────

async function getMotionLog(filters) {
  const plays = await getPlays();
  let allLogs = [];
  for (const play of plays) {
    const playId = play.play_id.replace(/[^a-z0-9]/gi, "_");
    const text = await getFileText(SP_ROOT + "/" + playId + "_engagements.csv");
    if (!text) continue;
    allLogs = allLogs.concat(parseCSV(text));
  }
  // Apply filters
  if (filters) {
    if (filters.rep_email) allLogs = allLogs.filter(r => (r.rep_email||"").toLowerCase() === filters.rep_email.toLowerCase());
    if (filters.play_id)   allLogs = allLogs.filter(r => r.play_id === filters.play_id);
    if (filters.outcome)   allLogs = allLogs.filter(r => r.outcome === filters.outcome);
  }
  return allLogs.sort((a, b) => (b.submitted_at || "").localeCompare(a.submitted_at || ""));
}

// ─── Section 7: Snapshot System ──────────────────────────────────────────────

const SNAPSHOT_ROOT = SP_ROOT + "/snapshots";

async function createSnapshot({ label = '', type = 'manual', triggeredBy = '' } = {}) {
  const now = new Date();
  const pad = n => String(n).padStart(2,'0');
  const ts = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
  const snapshotId = `snapshot-${ts}`;
  const snapshotPath = SNAPSHOT_ROOT + '/' + snapshotId;

  const manifest = {
    snapshot_id: snapshotId,
    timestamp: now.toISOString(),
    triggered_by: triggeredBy,
    label: label || '',
    type: type,
    files: [],
    record_counts: {}
  };

  const coreFiles = [
    { name: 'plays.json',            path: SP_ROOT + '/plays.json' },
    { name: 'assignments.json',      path: SP_ROOT + '/assignments.json' },
    { name: 'response-options.json', path: SP_ROOT + '/response-options.json' },
    { name: 'rep-identity.json',     path: SP_ROOT + '/rep-identity.json' },
    { name: 'activity_log.csv',      path: SP_ROOT + '/activity_log.csv' },
  ];

  for (const f of coreFiles) {
    try {
      const content = await getFileText(f.path);
      if (!content) continue;
      const ct = f.name.endsWith('.json') ? 'application/json' : 'text/csv';
      await putFile(snapshotPath + '/' + f.name, content, ct);
      manifest.files.push(f.name);
      try {
        if (f.name.endsWith('.json')) {
          const parsed = JSON.parse(content);
          const arrKey = Object.keys(parsed).find(k => Array.isArray(parsed[k]));
          if (arrKey) manifest.record_counts[f.name] = parsed[arrKey].length;
        } else {
          manifest.record_counts[f.name] = Math.max(0, content.trim().split('\n').length - 1);
        }
      } catch(e) {}
    } catch(e) {
      // File may not exist yet — skip silently
    }
  }

  // Per-play engagement CSVs
  try {
    const playsText = await getFileText(SP_ROOT + '/plays.json');
    if (playsText) {
      const playsData = JSON.parse(playsText);
      const plays = playsData.plays || [];
      for (const play of plays) {
        const playId = (play.play_id || '').replace(/[^a-z0-9]/gi, '_');
        if (!playId) continue;
        const engPath = SP_ROOT + '/' + playId + '_engagements.csv';
        try {
          const content = await getFileText(engPath);
          if (content) {
            const fname = playId + '_engagements.csv';
            await putFile(snapshotPath + '/' + fname, content, 'text/csv');
            manifest.files.push(fname);
            manifest.record_counts[fname] = Math.max(0, content.trim().split('\n').length - 1);
          }
        } catch(e) {}
      }
    }
  } catch(e) {}

  await putFile(snapshotPath + '/manifest.json', JSON.stringify(manifest, null, 2), 'application/json');

  return manifest;
}

async function listSnapshots() {
  const driveId = await getDriveId();
  const token = await getAccessToken();

  try {
    const encodedPath = encodeURIComponent(SNAPSHOT_ROOT);
    const resp = await fetch(
      `${CONFIG.graphBaseUrl}/drives/${driveId}/root:/${SNAPSHOT_ROOT}:/children`,
      { headers: { Authorization: 'Bearer ' + token } }
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    const folders = (data.value || []).filter(item => item.folder && item.name.startsWith('snapshot-'));

    const manifests = await Promise.all(folders.map(async folder => {
      try {
        const text = await getFileText(SNAPSHOT_ROOT + '/' + folder.name + '/manifest.json');
        return JSON.parse(text);
      } catch(e) {
        return {
          snapshot_id: folder.name,
          timestamp: '',
          label: '',
          type: 'unknown',
          files: [],
          record_counts: {}
        };
      }
    }));

    return manifests.sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || ''));
  } catch(e) {
    return [];
  }
}

async function restoreSnapshot(snapshotId, scope = 'full') {
  const snapshotPath = SNAPSHOT_ROOT + '/' + snapshotId;

  const scopeFiles = {
    full:  ['plays.json', 'assignments.json', 'response-options.json', 'rep-identity.json', 'activity_log.csv'],
    plays: ['plays.json', 'assignments.json', 'response-options.json'],
    logs:  ['activity_log.csv'],
  };
  const filesToRestore = scopeFiles[scope] || scopeFiles.full;

  const results = { restored: [], failed: [], skipped: [] };

  for (const fname of filesToRestore) {
    try {
      const content = await getFileText(snapshotPath + '/' + fname);
      if (!content) { results.skipped.push(fname); continue; }
      const ct = fname.endsWith('.json') ? 'application/json' : 'text/csv';
      await putFile(SP_ROOT + '/' + fname, content, ct);
      results.restored.push(fname);
    } catch(e) {
      results.failed.push(fname + ': ' + e.message);
    }
  }

  if (scope === 'full') {
    try {
      const manifestText = await getFileText(snapshotPath + '/manifest.json');
      const manifest = JSON.parse(manifestText);
      const engFiles = (manifest.files || []).filter(f => f.endsWith('_engagements.csv'));
      for (const fname of engFiles) {
        try {
          const content = await getFileText(snapshotPath + '/' + fname);
          if (content) {
            await putFile(SP_ROOT + '/' + fname, content, 'text/csv');
            results.restored.push(fname);
          }
        } catch(e) {
          results.failed.push(fname + ': ' + e.message);
        }
      }
    } catch(e) {}
  }

  return results;
}

async function getLastPreResetSnapshot() {
  const snapshots = await listSnapshots();
  return snapshots.find(s => s.type === 'pre-reset') || null;
}
