// INFRA SalesPlay Tracker - SharePoint CSV file backend
// Assignments read from assignments.json, engagements written to per-play CSV files
// Only needs Sites.ReadWrite.All (already granted)

// Root folder inside SharePoint Documents — matches the OneDrive-synced local folder
const SP_ROOT = "Chef SaaS Tracker/ChefSaaS";

let _siteId = null;
let _driveId = null;

// ─── Static alias map (module-level, shared by all matching functions) ────────
// Maps each known email variant to its siblings. Built from rep-identity.json.
// Used in both getPlayAssignments (Tier 3) and _getAllLogsForRep for consistent matching.
const STATIC_ALIASES = {
  'klesel@progress.com': ['micah.klesel@progress.com'],
  'micah.klesel@progress.com': ['klesel@progress.com'],
  'capparel@progress.com': ['richard.capparelli@progress.com'],
  'richard.capparelli@progress.com': ['capparel@progress.com'],
  'robert.wosneski@progress.com': ['wosneski@progress.com'],
  'wosneski@progress.com': ['robert.wosneski@progress.com'],
  'dandria.alston-popovic@progress.com': ['dpopovic@progress.com'],
  'dpopovic@progress.com': ['dandria.alston-popovic@progress.com'],
  'fatemah.shirzad@progress.com': ['shirzad@progress.com'],
  'shirzad@progress.com': ['fatemah.shirzad@progress.com'],
  'jakub.krzyzak@progress.com': ['krzyzak@progress.com'],
  'krzyzak@progress.com': ['jakub.krzyzak@progress.com'],
  'jorobert@progress.com': ['joshua.roberts@progress.com'],
  'joshua.roberts@progress.com': ['jorobert@progress.com'],
  'gould@progress.com': ['matt.gould@progress.com'],
  'matt.gould@progress.com': ['gould@progress.com'],
  'comrie@progress.com': ['rob.comrie@progress.com'],
  'rob.comrie@progress.com': ['comrie@progress.com'],
  'lemlein@progress.com': ['ryan.lemlein@progress.com'],
  'ryan.lemlein@progress.com': ['lemlein@progress.com'],
  'stinson.fernandes@progress.com': ['stinsonf@progress.com'],
  'stinsonf@progress.com': ['stinson.fernandes@progress.com'],
  'somukher@progress.com': ['souvik.mukherjee@progress.com'],
  'souvik.mukherjee@progress.com': ['somukher@progress.com'],
  'preeti.kumari2@progress.com': ['prkumari@progress.com'],
  'prkumari@progress.com': ['preeti.kumari2@progress.com'],
  'pprabhu@progress.com': ['prakash.prabhu@progress.com'],
  'prakash.prabhu@progress.com': ['pprabhu@progress.com'],
  'sak@progress.com': ['sampath.acharyak@progress.com'],
  'sampath.acharyak@progress.com': ['sak@progress.com'],
  'adam.weir@progress.com': ['adweir@progress.com'],
  'adweir@progress.com': ['adam.weir@progress.com'],
  'joshua.sands@progress.com': ['sands@progress.com'],
  'sands@progress.com': ['joshua.sands@progress.com'],
  'kerry.lafferty@progress.com': ['lafferty@progress.com'],
  'lafferty@progress.com': ['kerry.lafferty@progress.com'],
  'octon@progress.com': ['tim.octon@progress.com'],
  'tim.octon@progress.com': ['octon@progress.com'],
  'ana.valeva@progress.com': ['valeva@progress.com'],
  'valeva@progress.com': ['ana.valeva@progress.com'],
  'friedrich.scharz@progress.com': ['scharz@progress.com'],
  'scharz@progress.com': ['friedrich.scharz@progress.com'],
  'gbeblein@progress.com': ['georgiana.beblein@progress.com'],
  'georgiana.beblein@progress.com': ['gbeblein@progress.com'],
  'zhaneta.radeva@progress.com': ['zradeva@progress.com'],
  'zradeva@progress.com': ['zhaneta.radeva@progress.com'],
  'zkolev@progress.com': ['zlatomir.kolev@progress.com'],
  'zlatomir.kolev@progress.com': ['zkolev@progress.com'],
  'dimitar.kachorev@progress.com': ['kachorev@progress.com'],
  'kachorev@progress.com': ['dimitar.kachorev@progress.com'],
  'cacciama@progress.com': ['nicolo.cacciamano@progress.com'],
  'nicolo.cacciamano@progress.com': ['cacciama@progress.com'],
  'plaangel@progress.com': ['plamena.angelova@progress.com'],
  'plamena.angelova@progress.com': ['plaangel@progress.com'],
  'genov@progress.com': ['todor.genov@progress.com'],
  'todor.genov@progress.com': ['genov@progress.com'],
  'lulcheva@progress.com': ['vyara.lulcheva@progress.com'],
  'vyara.lulcheva@progress.com': ['lulcheva@progress.com'],
  'zdinev@progress.com': ['zdravko.dinev@progress.com'],
  'zdravko.dinev@progress.com': ['zdinev@progress.com'],
  'filip.gieci@progress.com': ['gieci@progress.com'],
  'gieci@progress.com': ['filip.gieci@progress.com'],
  'jakub.andrzejewski@progress.com': ['jandrzej@progress.com'],
  'jandrzej@progress.com': ['jakub.andrzejewski@progress.com'],
  'lrausche@progress.com': ['lukas.rauscher@progress.com'],
  'lukas.rauscher@progress.com': ['lrausche@progress.com'],
  'marek.machalek@progress.com': ['mmachale@progress.com'],
  'mmachale@progress.com': ['marek.machalek@progress.com'],
  'pavla.sehnalova@progress.com': ['psehnalo@progress.com'],
  'psehnalo@progress.com': ['pavla.sehnalova@progress.com'],
  'martinek@progress.com': ['vojtech.martinek@progress.com'],
  'vojtech.martinek@progress.com': ['martinek@progress.com'],
  'arjun.sp@progress.com': ['arjunsp@progress.com'],
  'arjunsp@progress.com': ['arjun.sp@progress.com'],
  'ashneeja.mp@progress.com': ['ashnmp@progress.com'],
  'ashnmp@progress.com': ['ashneeja.mp@progress.com'],
  'jtiwari@progress.com': ['jyoti.tiwari@progress.com'],
  'jyoti.tiwari@progress.com': ['jtiwari@progress.com'],
  'mmuheeb@progress.com': ['mohammed.muheeb@progress.com'],
  'mohammed.muheeb@progress.com': ['mmuheeb@progress.com'],
  'neha.katare@progress.com': ['nkatare@progress.com'],
  'nkatare@progress.com': ['neha.katare@progress.com'],
  'sourav.das@progress.com': ['souravda@progress.com'],
  'souravda@progress.com': ['sourav.das@progress.com'],
  'aakansha.nishu@progress.com': ['aakansha@progress.com'],
  'aakansha@progress.com': ['aakansha.nishu@progress.com'],
  'achowdar@progress.com': ['apathiakhil.chowdary@progress.com'],
  'apathiakhil.chowdary@progress.com': ['achowdar@progress.com'],
  'aganguly@progress.com': ['arka.ganguly@progress.com'],
  'arka.ganguly@progress.com': ['aganguly@progress.com'],
  'balamuru@progress.com': ['varshaa.balamurugan@progress.com'],
  'varshaa.balamurugan@progress.com': ['balamuru@progress.com'],
  'manusha.p@progress.com': ['manusha@progress.com'],
  'manusha@progress.com': ['manusha.p@progress.com'],
  'anumu@progress.com': ['anushree.mulimani@progress.com'],
  'anushree.mulimani@progress.com': ['anumu@progress.com'],
  'keerthana.s@progress.com': ['keerths@progress.com'],
  'keerths@progress.com': ['keerthana.s@progress.com'],
  'gracer@progress.com': ['mona.gracer@progress.com'],
  'mona.gracer@progress.com': ['gracer@progress.com'],
  'rcleetus@progress.com': ['riyamary.cleetus@progress.com'],
  'riyamary.cleetus@progress.com': ['rcleetus@progress.com'],
  'rudrabhavin.trivedi@progress.com': ['trivedi@progress.com'],
  'trivedi@progress.com': ['rudrabhavin.trivedi@progress.com'],
  'shyamkumar.ny@progress.com': ['skumary@progress.com'],
  'skumary@progress.com': ['shyamkumar.ny@progress.com'],
  'anusha.seethepalli@progress.com': ['aseethe@progress.com'],
  'aseethe@progress.com': ['anusha.seethepalli@progress.com'],
  'jyoti.sharma@progress.com': ['jysharma@progress.com'],
  'jysharma@progress.com': ['jyoti.sharma@progress.com'],
  'mdukhand@progress.com': ['mehul.dukhande@progress.com'],
  'mehul.dukhande@progress.com': ['mdukhand@progress.com'],
  'rahul.honawad@progress.com': ['rhonawad@progress.com'],
  'rhonawad@progress.com': ['rahul.honawad@progress.com'],
  'simanta.saha@progress.com': ['sisaha@progress.com'],
  'sisaha@progress.com': ['simanta.saha@progress.com'],
  'snagappa@progress.com': ['subashini.nagappan@progress.com'],
  'subashini.nagappan@progress.com': ['snagappa@progress.com'],
  'anirudh.v@progress.com': ['anirudhv@progress.com'],
  'anirudhv@progress.com': ['anirudh.v@progress.com'],
  'j.abhiram.varma@progress.com': ['jvarma@progress.com'],
  'jvarma@progress.com': ['j.abhiram.varma@progress.com'],
  'assadi.khader@progress.com': ['mabdul@progress.com'],
  'mabdul@progress.com': ['assadi.khader@progress.com'],
  'pkaarthi@progress.com': ['pnkoushal.kaarthiek@progress.com'],
  'pnkoushal.kaarthiek@progress.com': ['pkaarthi@progress.com'],
  'carrie.yuan@progress.com': ['yuan@progress.com'],
  'yuan@progress.com': ['carrie.yuan@progress.com'],
  'amardeep.singh@progress.com': ['amarsing@progress.com'],
  'amarsing@progress.com': ['amardeep.singh@progress.com'],
  'langer@progress.com': ['philip.langer@progress.com', 'planger@progress.com', 'phil.langer@progress.com'],
  'philip.langer@progress.com': ['langer@progress.com', 'planger@progress.com', 'phil.langer@progress.com'],
  'planger@progress.com': ['langer@progress.com', 'philip.langer@progress.com', 'phil.langer@progress.com'],
  'phil.langer@progress.com': ['langer@progress.com', 'philip.langer@progress.com', 'planger@progress.com'],
  'mae.witcher@progress.com': ['witcher@progress.com'],
  'witcher@progress.com': ['mae.witcher@progress.com'],
  'ichaudha@progress.com': ['isha.chaudhary@progress.com'],
  'isha.chaudhary@progress.com': ['ichaudha@progress.com'],
  'jaro.stusak@progress.com': ['jstusak@progress.com'],
  'jstusak@progress.com': ['jaro.stusak@progress.com'],
  'bergsma@progress.com': ['shawn.bergsma@progress.com'],
  'shawn.bergsma@progress.com': ['bergsma@progress.com'],
  'joseph.kuderer@progress.com': ['kuderer@progress.com'],
  'kuderer@progress.com': ['joseph.kuderer@progress.com'],
  'mcgowan@progress.com': ['stephen.mcgowan@progress.com'],
  'stephen.mcgowan@progress.com': ['mcgowan@progress.com'],
  'faria@progress.com': ['kathleen.faria@progress.com', 'kfaria@progress.com', 'k.faria@progress.com', 'kathy.faria@progress.com'],
  'kathleen.faria@progress.com': ['faria@progress.com', 'kfaria@progress.com', 'k.faria@progress.com', 'kathy.faria@progress.com'],
  'kfaria@progress.com': ['faria@progress.com', 'kathleen.faria@progress.com', 'k.faria@progress.com', 'kathy.faria@progress.com'],
  'k.faria@progress.com': ['faria@progress.com', 'kathleen.faria@progress.com', 'kfaria@progress.com', 'kathy.faria@progress.com'],
  'kathy.faria@progress.com': ['faria@progress.com', 'kathleen.faria@progress.com', 'kfaria@progress.com', 'k.faria@progress.com'],
  'cfranke@progress.com': ['courtney.franke@progress.com'],
  'courtney.franke@progress.com': ['cfranke@progress.com'],
  'mihail.hristov@progress.com': ['mihhrist@progress.com'],
  'mihhrist@progress.com': ['mihail.hristov@progress.com'],
  'denikolo@progress.com': ['desislava.nikolova@progress.com'],
  'desislava.nikolova@progress.com': ['denikolo@progress.com'],
  'jiri.mazal@progress.com': ['jmazal@progress.com'],
  'jmazal@progress.com': ['jiri.mazal@progress.com'],
  'schwarz@progress.com': ['thomas.schwarz@progress.com'],
  'thomas.schwarz@progress.com': ['schwarz@progress.com'],
  'laszlo.vanya@progress.com': ['lvanya@progress.com'],
  'lvanya@progress.com': ['laszlo.vanya@progress.com'],
  'alisha.wilson@progress.com': ['awilson@progress.com'],
  'awilson@progress.com': ['alisha.wilson@progress.com'],
  'chguerre@progress.com': ['christopher.guerrero@progress.com'],
  'christopher.guerrero@progress.com': ['chguerre@progress.com'],
  'lauren.dipresso@progress.com': ['ldipress@progress.com'],
  'ldipress@progress.com': ['lauren.dipresso@progress.com'],
  'asahay@progress.com': ['ayush.sahay@progress.com'],
  'ayush.sahay@progress.com': ['asahay@progress.com'],
  'sageorge@progress.com': ['sathish.george@progress.com'],
  'sathish.george@progress.com': ['sageorge@progress.com'],
  'bvivek@progress.com': ['vivek.bolde@progress.com'],
  'vivek.bolde@progress.com': ['bvivek@progress.com'],
  'coronado@progress.com': ['david.coronado@progress.com'],
  'david.coronado@progress.com': ['coronado@progress.com'],
  'lmihaylo@progress.com': ['ludmila.mihaylovitch@progress.com'],
  'ludmila.mihaylovitch@progress.com': ['lmihaylo@progress.com'],
  'mohammad.sameer@progress.com': ['msameer@progress.com'],
  'msameer@progress.com': ['mohammad.sameer@progress.com'],
  'vijetha.p@progress.com': ['vijetp@progress.com'],
  'vijetp@progress.com': ['vijetha.p@progress.com'],
  'shylesh.kb@progress.com': ['shylkb@progress.com'],
  'shylkb@progress.com': ['shylesh.kb@progress.com'],
  'sophiya.shaheen@progress.com': ['sshahe@progress.com'],
  'sshahe@progress.com': ['sophiya.shaheen@progress.com'],
};

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
    for (let attempt = 0; attempt < 3; attempt++) {
      const { text: existing, eTag } = await getFileWithETag(filePath);
      const csv = (existing || '').trim()
        ? (existing || '').trimEnd() + '\n' + row
        : ACTIVITY_LOG_HEADERS.join(',') + '\n' + row;
      const ok = await putFileWithETag(filePath, csv, 'text/csv', eTag);
      if (ok) return;
      await new Promise(r => setTimeout(r, 200 + Math.random() * 500));
    }
    // Final fallback — force write
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

// ─── Section 3a: Assignment Corrections ──────────────────────────────────
// updateAssignment()        — finds a row by account_id+play_id, applies changes, writes back.
// logAssignmentCorrection() — appends one row to assignment_corrections_log.csv (fire-and-forget).
// getAssignmentCorrectionsLog() — reads the full corrections audit log.

const CORRECTION_LOG_PATH = SP_ROOT + '/assignment_corrections_log.csv';
const CORRECTION_LOG_HEADERS = [
  'timestamp', 'admin_email', 'admin_name',
  'account_id', 'account_name', 'play_id', 'play_name',
  'old_rep_email', 'new_rep_email', 'action_type', 'notes',
];

async function updateAssignment(accountId, playId, changes, adminEmail, adminName) {
  const text = await getFileText(SP_ROOT + '/assignments.json');
  if (!text) throw new Error('assignments.json not found');
  const data = JSON.parse(text);
  const rows = data.assignments || [];
  const idx = rows.findIndex(a => a.account_id === accountId && a.play_id === playId);
  if (idx === -1) throw new Error('Assignment not found: ' + accountId + ' / ' + playId);
  rows[idx] = {
    ...rows[idx],
    ...changes,
    updated_at: new Date().toISOString(),
    updated_by: adminEmail || '',
  };
  data.assignments = rows;
  await putFile(SP_ROOT + '/assignments.json', JSON.stringify(data, null, 2), 'application/json');
  return rows[idx];
}

async function logAssignmentCorrection({
  adminEmail = '', adminName = '',
  accountId = '', accountName = '',
  playId = '', playName = '',
  oldRepEmail = '', newRepEmail = '',
  actionType = '', notes = '',
} = {}) {
  try {
    const row = [
      new Date().toISOString(), adminEmail, adminName,
      accountId, accountName, playId, playName,
      oldRepEmail, newRepEmail, actionType, notes,
    ].map(v => csvEscape(String(v))).join(',');
    let existing = '';
    try { existing = await getFileText(CORRECTION_LOG_PATH) || ''; } catch(e) {}
    const csv = existing.trim()
      ? existing.trimEnd() + '\n' + row
      : CORRECTION_LOG_HEADERS.join(',') + '\n' + row;
    await putFile(CORRECTION_LOG_PATH, csv, 'text/csv');
  } catch(e) {
    console.warn('[CorrectionLog] Failed (non-blocking):', e.message);
  }
}

async function getAssignmentCorrectionsLog() {
  try {
    const text = await getFileText(CORRECTION_LOG_PATH);
    if (!text) return [];
    const lines = text.trim().split('\n');
    if (lines.length < 2) return [];
    const headers = lines[0].split(',').map(h => h.replace(/^"|"$/g, '').trim());
    return lines.slice(1).map(line => {
      const vals = line.match(/("(?:[^"]|"")*"|[^,]*)/g) || [];
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = (vals[i] || '').replace(/^"|"$/g, '').replace(/""/g, '"').trim();
      });
      return obj;
    });
  } catch(e) { return []; }
}

// ─── Section 3b: Identity Event Logging ──────────────────────────────────
// Separate CSV from activity_log. Records every login attempt, identity
// resolution result, and assignment mismatch. Always fire-and-forget.
//
// Schema: identity_log.csv
//   timestamp          ISO-8601
//   event_type         login_attempt | login_success | identity_mismatch | identity_match
//   rep_display_name   Display name from token/Graph
//   raw_identity       The email/UPN directly from the MSAL token
//   resolved_identity  The email that actually matched an assignment row (or '' if none)
//   match_type         OID | identity_map | fallback | none
//   candidate_identities  JSON array of all emails checked during resolution
//   assignment_count   Number of assignment rows matched
//   notes              Optional free-text debug info

const IDENTITY_LOG_HEADERS = [
  'timestamp', 'event_type', 'rep_display_name', 'raw_identity',
  'resolved_identity', 'match_type', 'candidate_identities',
  'assignment_count', 'notes',
];
const IDENTITY_LOG_PATH = SP_ROOT + '/identity_log.csv';

async function logIdentityEvent({
  eventType       = '',
  repDisplayName  = '',
  rawIdentity     = '',
  resolvedIdentity = '',
  matchType       = '',
  candidates      = [],
  assignmentCount = '',
  notes           = '',
} = {}) {
  try {
    const row = [
      new Date().toISOString(),
      eventType,
      repDisplayName,
      rawIdentity,
      resolvedIdentity,
      matchType,
      Array.isArray(candidates) ? JSON.stringify(candidates) : String(candidates),
      assignmentCount != null ? String(assignmentCount) : '',
      notes,
    ].map(v => csvEscape(String(v))).join(',');

    let existing = '';
    try { existing = await getFileText(IDENTITY_LOG_PATH) || ''; } catch(e) {}
    const csv = existing.trim()
      ? existing.trimEnd() + '\n' + row
      : IDENTITY_LOG_HEADERS.join(',') + '\n' + row;
    await putFile(IDENTITY_LOG_PATH, csv, 'text/csv');
  } catch(e) {
    console.warn('[IdentityLog] logIdentityEvent failed (non-blocking):', e.message);
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
    if (!retry.ok) {
      const t = await retry.text();
      const err = new Error("Graph " + retry.status + ": " + t);
      if (retry.status === 403) { err.isSharePointAccessDenied = true; err.failingEndpoint = url; }
      throw err;
    }
    if (retry.status === 204) return null;
    return retry.json();
  }
  if (!resp.ok) {
    const t = await resp.text();
    const err = new Error("Graph " + resp.status + ": " + t);
    if (resp.status === 403) { err.isSharePointAccessDenied = true; err.failingEndpoint = url; }
    throw err;
  }
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

// Read file with eTag for optimistic concurrency
async function getFileWithETag(filePath) {
  const driveId = await getDriveId();
  const token = await getAccessToken();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const meta = await fetch(CONFIG.graphBaseUrl + "/drives/" + driveId + "/root:/" + filePath, {
      signal: controller.signal,
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }
    });
    if (!meta.ok) return { text: null, eTag: null };
    const metaJson = await meta.json();
    const eTag = metaJson.eTag || null;
    const resp = await fetch(metaJson["@microsoft.graph.downloadUrl"], {
      signal: controller.signal,
      headers: { Authorization: "Bearer " + token }
    });
    if (!resp.ok) return { text: null, eTag: null };
    const text = await resp.text();
    return { text, eTag };
  } catch(e) { return { text: null, eTag: null }; } finally { clearTimeout(timeout); }
}

// Write file with optional eTag check (returns false on conflict)
async function putFileWithETag(filePath, content, contentType, eTag) {
  const driveId = await getDriveId();
  const token = await getAccessToken();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    const headers = { Authorization: "Bearer " + token, "Content-Type": contentType || "text/plain" };
    if (eTag) headers["If-Match"] = eTag;
    const resp = await fetch(CONFIG.graphBaseUrl + "/drives/" + driveId + "/root:/" + filePath + ":/content", {
      method: "PUT", signal: controller.signal, headers, body: content
    });
    if (resp.status === 412) return false; // conflict — someone else modified the file
    if (!resp.ok) { const t = await resp.text(); throw new Error("Upload " + resp.status + ": " + t); }
    return true;
  } catch(e) {
    if (e.name === 'AbortError') throw new Error("Save timed out — check your connection and try again.");
    throw e;
  } finally { clearTimeout(timeout); }
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

// ─── Section: Rep Identity Enrichment ─────────────────────────────────────
// Fetches /me from Microsoft Graph to get the rep's primary mail attribute.
// This fills the gap where the MSAL token only carries the UPN/preferred_username
// but the spreadsheet was uploaded with the primary mail (or vice-versa).
//
// Gracefully returns {} on any failure — callers treat missing fields as absent.

let _graphUserProfile = null;

async function fetchGraphUserProfile() {
  if (_graphUserProfile !== null) return _graphUserProfile;
  try {
    const data = await _graphFetch('/me?$select=mail,userPrincipalName,otherMails,displayName');
    _graphUserProfile = data || {};
    console.log('[Identity] /me profile:', _graphUserProfile.mail, '/', _graphUserProfile.userPrincipalName);
  } catch(e) {
    console.warn('[Identity] /me fetch failed (non-critical, falling back to token claims):', e.message);
    _graphUserProfile = {};
  }
  return _graphUserProfile;
}

// ─── Email Lookup table (from "EMAIL LookUP for SSO.xlsx") ────────────────
// This file is the authoritative rep identity map maintained by the admin.
// Each entry has { primary_email, upn (optional), name, notes }.
//
// Storage hierarchy (in priority order):
//   1. SharePoint: Chef SaaS Tracker/ChefSaaS/email-lookup.json
//      → Admin uploads this via the Identity tab. Updates take effect within
//        24 hours for all reps without any code redeploy.
//   2. localStorage cache (TTL: 24 hours)
//      → Avoids a SharePoint round-trip on every page load.
//   3. Static /email-lookup.json (deployed with the app)
//      → Always-available fallback. Updated only via redeploy.
//
// Admin tools: clearEmailLookupCache() / refreshEmailLookup() (exposed globally).

const EMAIL_LOOKUP_CACHE_KEY = 'chef_email_lookup_v2';
const EMAIL_LOOKUP_TS_KEY    = 'chef_email_lookup_ts_v2';
const EMAIL_LOOKUP_TTL_MS    = 24 * 60 * 60 * 1000; // 24 hours

let _emailLookup    = null;  // null = not yet loaded; [] = loaded (possibly empty)
let _emailLookupMap = null;  // Map<lowercase-email, entry>

function _buildEmailLookupMap(lookup) {
  const map = new Map();
  for (const entry of lookup) {
    if (entry.primary_email) map.set(entry.primary_email.toLowerCase().trim(), entry);
    if (entry.upn)           map.set(entry.upn.toLowerCase().trim(), entry);
  }
  return map;
}

// Clear the localStorage cache so the next loadEmailLookup() fetches fresh data.
// Called by the admin UI after uploading a new file.
function clearEmailLookupCache() {
  try {
    localStorage.removeItem(EMAIL_LOOKUP_CACHE_KEY);
    localStorage.removeItem(EMAIL_LOOKUP_TS_KEY);
  } catch(e) {}
  _emailLookup = null;
  _emailLookupMap = null;
  console.log('[Identity] Email lookup cache cleared — next load will fetch from SharePoint');
}

// Force a fresh fetch from SharePoint, bypassing the cache.
async function refreshEmailLookup() {
  clearEmailLookupCache();
  return loadEmailLookup();
}

async function loadEmailLookup() {
  // Return cached in-memory copy immediately if available this session
  if (_emailLookup !== null) return _emailLookup;

  // ── Check localStorage cache (24h TTL) ───────────────────────────────────
  try {
    const cached = localStorage.getItem(EMAIL_LOOKUP_CACHE_KEY);
    const ts     = localStorage.getItem(EMAIL_LOOKUP_TS_KEY);
    if (cached && ts) {
      const ageMs = Date.now() - new Date(ts).getTime();
      if (ageMs < EMAIL_LOOKUP_TTL_MS) {
        _emailLookup    = JSON.parse(cached);
        _emailLookupMap = _buildEmailLookupMap(_emailLookup);
        console.log(`[Identity] email-lookup from localStorage cache (${Math.round(ageMs / 60000)}m old, ${_emailLookup.length} reps)`);
        return _emailLookup;
      }
      console.log('[Identity] email-lookup cache expired — fetching fresh');
    }
  } catch(e) { /* corrupt cache — proceed to fetch */ }

  // ── Tier 1: SharePoint (updateable without code redeploy) ────────────────
  let loaded = null;
  let source = null;
  try {
    const text = await getFileText(SP_ROOT + '/email-lookup.json');
    if (text) { loaded = JSON.parse(text); source = 'SharePoint'; }
  } catch(e) {
    console.warn('[Identity] SharePoint email-lookup.json not found — will try static fallback:', e.message);
  }

  // ── Tier 2: Static asset (deployed with the app) ─────────────────────────
  if (!loaded) {
    try {
      const res = await fetch('/email-lookup.json');
      if (res.ok) { loaded = await res.json(); source = 'static (deploy)'; }
    } catch(e) {
      console.warn('[Identity] Static email-lookup.json not available:', e.message);
    }
  }

  _emailLookup    = loaded || [];
  _emailLookupMap = _buildEmailLookupMap(_emailLookup);

  // ── Store in localStorage cache ───────────────────────────────────────────
  if (_emailLookup.length > 0) {
    try {
      localStorage.setItem(EMAIL_LOOKUP_CACHE_KEY, JSON.stringify(_emailLookup));
      localStorage.setItem(EMAIL_LOOKUP_TS_KEY, new Date().toISOString());
    } catch(e) { /* localStorage full or blocked — non-fatal */ }
  }

  console.log(`[Identity] email-lookup loaded from ${source || 'nowhere'}: ${_emailLookup.length} reps`);
  return _emailLookup;
}

// Given any known email string, resolve the matching entry from the lookup table.
// Returns the entry object ({ name, primary_email, upn, notes }) or null.
function _lookupRepEntry(emailLower) {
  if (!_emailLookupMap) return null;
  return _emailLookupMap.get(emailLower) || null;
}

// ─── Canonical email candidate builder ────────────────────────────────────
// Returns a Set of all lowercase, trimmed email-like identities for the
// signed-in rep. Every matching function should use this as the single
// source of truth for "who is this person?"
//
// Resolution layers (highest to lowest confidence):
//   0. email-lookup.json (authoritative Excel-derived identity map)
//      — for each token/Graph email, look up the row and add BOTH the
//        primary_email AND the upn so either form matches assignments
//   1. Graph /me: mail, userPrincipalName, otherMails
//   2. MSAL token claims: preferred_username, email, upn, unique_name, username
//   3. rep-identity.json aliases (SharePoint-hosted, admin-managed)
//   4. STATIC_ALIASES hardcoded map (legacy fallback; now superseded by layer 0)

async function getEmailCandidates(primaryEmail) {
  const candidates = new Set();
  const add = v => { if (v && typeof v === 'string') candidates.add(v.toLowerCase().trim()); };

  // Seed from primaryEmail argument
  if (primaryEmail) add(primaryEmail);

  // Layer 1: MSAL token claims (synchronous)
  try {
    const user = typeof getCurrentUser === 'function' ? getCurrentUser() : null;
    if (user?.emails?.length) user.emails.forEach(add);
  } catch(e) {}

  // Layer 2: Graph /me profile (async, cached after first call)
  try {
    const profile = await fetchGraphUserProfile();
    add(profile.mail);
    add(profile.userPrincipalName);
    (profile.otherMails || []).forEach(add);
  } catch(e) {}

  // Layer 0 (applied after layers 1+2 so we have the full token picture):
  // Resolve every known candidate against the authoritative email-lookup table.
  // If any candidate matches a row (by primary_email OR upn), add BOTH values
  // so the rep can be found regardless of which form is in the assignment spreadsheet.
  try {
    await loadEmailLookup();
    const snapshot0 = [...candidates];
    for (const e of snapshot0) {
      const entry = _lookupRepEntry(e);
      if (entry) {
        if (entry.primary_email) add(entry.primary_email);
        if (entry.upn)           add(entry.upn);
      }
    }
  } catch(e) {}

  // Layer 3: rep-identity.json aliases (SharePoint-managed overrides)
  try {
    const seed = primaryEmail || [...candidates][0];
    if (seed) {
      const repEntry = await resolveRepByEmail(seed);
      if (repEntry) (repEntry.aliases || []).forEach(add);
    }
  } catch(e) {}

  // Layer 4: STATIC_ALIASES expansion (legacy fallback — kept for backward compat
  // in case a rep isn't in email-lookup.json yet)
  const snapshot4 = [...candidates];
  snapshot4.forEach(e => (STATIC_ALIASES[e] || []).forEach(add));

  return candidates;
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

async function getPlayAssignments(repEmail, meta = null) {
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
      if (meta) {
        meta.matchType        = 'OID';
        meta.resolvedIdentity = emailLower;
        meta.candidates       = [emailLower];
      }
      return byOid;
    }
  }

  // Tier 2: resolve via rep-identity.json alias overrides
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
        if (meta) {
          meta.matchType        = 'identity_map';
          meta.resolvedIdentity = aliases[0] || emailLower;
          meta.candidates       = aliases;
        }
        return byAlias;
      }
    }
  } catch(e) {
    console.warn("[ChefSaaS] Identity map lookup failed, falling back:", e.message);
  }

  // Tier 3: direct match using enriched candidate set
  // (token claims + Graph /me + email-lookup.json + STATIC_ALIASES)
  const allCandidates = await getEmailCandidates(emailLower);

  const matched = list.filter(a => {
    const repLogin = (a.rep_sso_login || '').toLowerCase().trim();
    const repEmail  = (a.rep_email    || '').toLowerCase().trim();
    return [...allCandidates].some(e => e === repLogin || e === repEmail);
  });

  if (meta) {
    // Did any candidate come from email-lookup.json? (→ identity_map) or only from
    // STATIC_ALIASES / token? (→ fallback)
    const lookupHit = [...allCandidates].some(c => _lookupRepEntry(c) !== null);
    meta.matchType = matched.length > 0 ? (lookupHit ? 'identity_map' : 'fallback') : 'none';

    if (matched.length > 0) {
      // Identify the specific candidate email that hit a row
      const assignedEmails = new Set(
        list.flatMap(a => [
          (a.rep_sso_login || '').toLowerCase().trim(),
          (a.rep_email     || '').toLowerCase().trim(),
        ].filter(Boolean))
      );
      meta.resolvedIdentity = [...allCandidates].find(c => assignedEmails.has(c)) || emailLower;
    } else {
      meta.resolvedIdentity = null;
    }
    meta.candidates = [...allCandidates];
  }

  console.log('[ChefSaaS] Candidates:', [...allCandidates], '→', matched.length, 'accounts');
  return matched;
}

async function getRepAccounts(repEmail) { return getPlayAssignments(repEmail); }

// ─── Plays config (plays.json) ────────────────────────────────────────────

let _playsCache = null;

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
  _playsCache = null; // invalidate cache on write
  await writeJsonFile("plays.json", config);
}

async function getPlays() {
  if (_playsCache) return _playsCache;
  try {
    const config = await getPlaysConfig();
    if (config && config.plays && config.plays.length > 0) {
      _playsCache = config.plays.filter(p => p.active !== false);
      return _playsCache;
    }
    // fallback: derive from assignments
    const all = await getPlayAssignments(null);
    const seen = new Set();
    _playsCache = all.filter(a => { if (!a.play_id || seen.has(a.play_id)) return false; seen.add(a.play_id); return true; })
              .map(a => ({ play_id: a.play_id, play_name: a.play_name || a.play_id }));
    return _playsCache;
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

  // Speed bonus — reward reps who finish before the May 21 deadline
  const deadline = new Date('2025-05-21T23:59:59');
  const logDate = new Date(data && (data.call_date || data.submitted_at) || Date.now());
  if (!isNaN(logDate) && logDate < deadline) {
    const daysLeft = Math.floor((deadline - logDate) / 86400000);
    if (daysLeft >= 14) bonus += 15;
    else if (daysLeft >= 7) bonus += 10;
    else if (daysLeft >= 1) bonus += 5;
  }

  return base + bonus;
}

// ─── CSV helpers ──────────────────────────────────────────────────────────

const CSV_HEADERS = ["id","submitted_at","play_id","play_name","account_id","account_name","rep_email","rep_name","interaction_type","outcome","reason_label","reason_code","next_step_type","timing","contact_level","notes","call_date","points_earned"];

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

  // Retry loop with eTag to prevent lost updates from concurrent writes
  const MAX_RETRIES = 4;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    const { text: existing, eTag } = await getFileWithETag(filePath);
    const csv = (existing && existing.trim())
      ? existing.trimEnd() + "\n" + toCsvRow(entry)
      : CSV_HEADERS.join(",") + "\n" + toCsvRow(entry);

    const ok = await putFileWithETag(filePath, csv, "text/csv", eTag);
    if (ok) return entry;

    // Conflict — wait briefly then retry with fresh data
    await new Promise(r => setTimeout(r, 300 + Math.random() * 700));
  }
  // Final attempt without eTag (force write rather than lose the entry)
  const { text: existing } = await getFileWithETag(filePath);
  const csv = (existing && existing.trim())
    ? existing.trimEnd() + "\n" + toCsvRow(entry)
    : CSV_HEADERS.join(",") + "\n" + toCsvRow(entry);
  await putFile(filePath, csv, "text/csv");
  return entry;
}

// ─── Read engagements ─────────────────────────────────────────────────────

async function _getAllLogsForRep(repEmail) {
  // Build full candidate set using the same resolver as getPlayAssignments()
  let emailCandidates = null;
  if (repEmail) {
    emailCandidates = [...(await getEmailCandidates(repEmail.toLowerCase()))];
  }

  const plays = await getPlays();
  let allLogs = [];
  for (const play of plays) {
    const playId = play.play_id.replace(/[^a-z0-9]/gi, "_");
    const text = await getFileText(SP_ROOT + "/" + playId + "_engagements.csv");
    if (!text) continue;
    const rows = parseCSV(text).filter(r => {
      if (!emailCandidates) return true;
      const rowEmail = (r.rep_email || "").toLowerCase();
      return emailCandidates.some(e => e === rowEmail);
    });
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
  return logs.reduce((sum, l) => sum + _recalcPoints(l), 0);
}

// Recalculate points from raw row data — source of truth is the CSV fields, not stored points_earned
function _recalcPoints(row) {
  return calculatePoints(row.outcome, row.next_step_type, row);
}

// Normalize a rep email to a canonical key using STATIC_ALIASES
function _canonicalEmail(email) {
  const e = (email || '').toLowerCase();
  // Pick the alphabetically-first alias as canonical so both variants map to the same key
  const aliases = STATIC_ALIASES[e];
  if (aliases && aliases.length) {
    const all = [e, ...aliases].sort();
    return all[0];
  }
  return e;
}

async function getLeaderboard(playId) {
  try {
    let playIds;
    if (playId) {
      playIds = [playId];
    } else {
      const plays = await getPlays();
      playIds = plays.map(p => p.play_id);
    }

    // Load assignments to get assigned account counts per rep
    const assText = await getFileText(SP_ROOT + "/assignments.json");
    const allAssignments = assText ? (JSON.parse(assText).assignments || []).filter(a => a.active_flag !== false) : [];
    const repAssignedCounts = {};
    allAssignments.forEach(a => {
      if (playId && a.play_id !== playId) return;
      const email = _canonicalEmail((a.rep_sso_login || a.rep_email || '').toLowerCase());
      if (!email) return;
      if (!repAssignedCounts[email]) repAssignedCounts[email] = new Set();
      if (a.account_id) repAssignedCounts[email].add(a.account_id);
    });

    const repMap = {};
    for (const pid of playIds) {
      const safePid = pid.replace(/[^a-z0-9]/gi, "_");
      const text = await getFileText(SP_ROOT + "/" + safePid + "_engagements.csv");
      if (!text) continue;
      parseCSV(text).forEach(log => {
        const rawEmail = (log.rep_email || '').toLowerCase();
        if (!rawEmail) return;
        const key = _canonicalEmail(rawEmail);
        if (!repMap[key]) repMap[key] = {
          email: rawEmail, repName: log.rep_name || rawEmail,
          totalPoints: 0, engagedAccounts: new Set(), interestedAccounts: new Set(),
          totalEngagements: 0, lastActivity: ''
        };
        if ((log.rep_name || '').length > (repMap[key].repName || '').length) {
          repMap[key].repName = log.rep_name;
        }
        repMap[key].totalPoints += _recalcPoints(log);
        repMap[key].totalEngagements += 1;
        if (log.account_id) {
          repMap[key].engagedAccounts.add(log.account_id);
          if (log.outcome === 'Interested') repMap[key].interestedAccounts.add(log.account_id);
        }
        const ts = log.submitted_at || log.call_date || '';
        if (ts > repMap[key].lastActivity) repMap[key].lastActivity = ts;
      });
    }

    // Build results from reps who have engagements
    const results = Object.values(repMap).map(r => {
      const assigned = repAssignedCounts[_canonicalEmail(r.email)] ? repAssignedCounts[_canonicalEmail(r.email)].size : 0;
      const engaged = r.engagedAccounts.size;
      let pts = r.totalPoints;
      if (assigned > 0 && engaged >= assigned) pts += 50;
      const ppa = assigned > 0 ? Math.round(pts / assigned) : 0;
      return {
        ...r,
        totalPoints: pts,
        completedAccounts: engaged,
        engagedAccounts: engaged,
        interestedAccounts: r.interestedAccounts.size,
        totalEngagements: r.totalEngagements,
        lastActivity: r.lastActivity,
        assignedAccounts: assigned,
        pointsPerAccount: ppa,
        isComplete: assigned > 0 && engaged >= assigned,
      };
    });

    // Add unstarted reps who have assigned accounts but zero engagements
    const seenEmails = new Set(Object.keys(repMap));
    const assReps = {};
    allAssignments.forEach(a => {
      if (playId && a.play_id !== playId) return;
      const email = _canonicalEmail((a.rep_sso_login || a.rep_email || '').toLowerCase());
      if (!email || seenEmails.has(email)) return;
      if (!assReps[email]) assReps[email] = { name: a.rep_name || email, count: 0 };
      if (a.account_id) assReps[email].count++;
    });
    Object.entries(assReps).forEach(([email, info]) => {
      if (info.count > 0) {
        results.push({
          email, repName: info.name, totalPoints: 0,
          completedAccounts: 0, engagedAccounts: 0, interestedAccounts: 0,
          totalEngagements: 0, lastActivity: '',
          assignedAccounts: info.count, pointsPerAccount: 0, isComplete: false,
        });
      }
    });

    return results.sort((a, b) => b.pointsPerAccount - a.pointsPerAccount);
  } catch(e) { return []; }
}

async function getRepsForPlay(playId) {
  const text = await getFileText(SP_ROOT + "/assignments.json");
  if (!text) return [];
  const data = JSON.parse(text);
  const map = {};
  (data.assignments || [])
    .filter(a => a.active_flag !== false && a.play_id === playId)
    .forEach(a => {
      const email = (a.rep_sso_login || a.rep_email || '').toLowerCase();
      if (email && !map[email]) map[email] = { email, repName: a.rep_name || email };
    });
  return Object.values(map);
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
