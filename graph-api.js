// Chef SaaS Motion Tracker — Graph API (v3 engagement model)

let _siteId = null;

async function _graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : `${CONFIG.graphBaseUrl}${path}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph API error ${response.status}: ${errorText}`);
  }
  if (response.status === 204) return null;
  return response.json();
}

async function getSiteId() {
  if (_siteId) return _siteId;
  const url = new URL(CONFIG.sharepointSiteUrl);
  const data = await _graphFetch(`/sites/${url.hostname}:${url.pathname}`);
  _siteId = data.id;
  return _siteId;
}

async function getListItems(listName, selectFields, filterQuery) {
  const siteId = await getSiteId();
  let url = `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?$expand=fields&$top=999`;
  if (filterQuery) url += `&$filter=${encodeURIComponent(filterQuery)}`;
  if (selectFields) url += `&$select=${encodeURIComponent(selectFields)}`;

  const allItems = [];
  let nextUrl = url;
  while (nextUrl) {
    const data = await _graphFetch(nextUrl);
    if (data.value) allItems.push(...data.value);
    nextUrl = data["@odata.nextLink"] || null;
  }
  return allItems;
}

async function createListItem(listName, fields) {
  const siteId = await getSiteId();
  const url = `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`;
  return _graphFetch(url, { method: "POST", body: JSON.stringify({ fields }) });
}

async function createList(listDisplayName, columns) {
  const siteId = await getSiteId();
  const columnDefs = columns.map((col) => {
    const def = { name: col.name };
    if (col.type === "text") def.text = col.multiline ? { allowMultipleLines: true } : {};
    else if (col.type === "number") def.number = {};
    else if (col.type === "boolean") def.boolean = {};
    else def.text = {};
    return def;
  });
  return _graphFetch(`/sites/${siteId}/lists`, {
    method: "POST",
    body: JSON.stringify({
      displayName: listDisplayName,
      columns: columnDefs,
      list: { template: "genericList" },
    }),
  });
}

// ─── Points & Levels ──────────────────────────────────────────────────────────

const POINTS = {
  "Not Interested":         15,   // learning value — reason required
  "Interested":             30,
  "Interested + Next Step": 45,
};

function calculatePoints(outcome, nextStepType) {
  if (outcome === "Interested" && nextStepType) return POINTS["Interested + Next Step"];
  if (outcome === "Interested") return POINTS["Interested"];
  if (outcome === "Not Interested") return POINTS["Not Interested"];
  return 0;
}

const LEVELS = [
  { min: 500, name: "Legend",   badge: "🌟" },
  { min: 300, name: "Champion", badge: "🏆" },
  { min: 150, name: "Pro",      badge: "🎯" },
  { min: 50,  name: "Pitcher",  badge: "⚾" },
  { min: 0,   name: "Rookie",   badge: "🌱" },
];

function getLevel(totalPoints) {
  return LEVELS.find((l) => totalPoints >= l.min) || LEVELS[LEVELS.length - 1];
}

// ─── Play Assignments ─────────────────────────────────────────────────────────

async function getPlayAssignments(repEmail) {
  const items = await getListItems("sales_play_assignment");
  return items.map(i => i.fields).filter(f =>
    f.active_flag !== false &&
    (repEmail ? (f.rep_email || "").toLowerCase() === repEmail.toLowerCase() : true)
  );
}

// ─── Response Options ─────────────────────────────────────────────────────────

async function getResponseOptions(responseSetId) {
  const items = await getListItems("response_option");
  return items.map(i => i.fields).filter(f =>
    f.active_flag !== false &&
    (!responseSetId || (f.response_set_id || "").toUpperCase() === responseSetId.toUpperCase())
  ).sort((a, b) => (a.sort_order || 0) - (b.sort_order || 0));
}

// ─── Executions ───────────────────────────────────────────────────────────────

async function logEngagement(fields) {
  const pts = calculatePoints(fields.outcome, fields.next_step_type);
  return createListItem("sales_play_execution", {
    ...fields,
    points_earned: pts,
    submitted_at: new Date().toISOString(),
    source: "web-app-v3",
  });
}

async function getExecutions(filters = {}) {
  const items = await getListItems("sales_play_execution");
  let results = items.map(i => i.fields);
  if (filters.repEmail) results = results.filter(r => r.rep_email === filters.repEmail);
  if (filters.playId)   results = results.filter(r => r.play_id   === filters.playId);
  return results.sort((a, b) => (b.submitted_at || "").localeCompare(a.submitted_at || ""));
}

async function getLatestExecutionPerAccount(playId, accountIds) {
  const items = await getListItems("sales_play_execution");
  const logs = items.map(i => i.fields).filter(f => !playId || f.play_id === playId);

  const latest = {};
  for (const log of logs) {
    const key = log.account_id;
    if (!key) continue;
    if (!latest[key] || (log.submitted_at || "") > (latest[key].submitted_at || "")) {
      latest[key] = log;
    }
  }
  if (accountIds) {
    const result = {};
    for (const id of accountIds) result[id] = latest[id] || null;
    return result;
  }
  return latest;
}

async function getRepPoints(repEmail) {
  const items = await getListItems("sales_play_execution");
  const logs = items.map(i => i.fields).filter(f => f.rep_email === repEmail || f.submitted_by === repEmail);
  return logs.reduce((sum, l) => sum + (Number(l.points_earned) || 0), 0);
}

async function getLeaderboard() {
  const items = await getListItems("sales_play_execution");
  const logs = items.map(i => i.fields);

  const repMap = {};
  for (const log of logs) {
    const key = log.rep_email || log.submitted_by;
    if (!key) continue;
    if (!repMap[key]) {
      repMap[key] = {
        email: key,
        repName: log.rep_name || key,
        totalPoints: 0,
        engagedAccounts: new Set(),
        interestedAccounts: new Set(),
      };
    }
    repMap[key].totalPoints += Number(log.points_earned) || 0;
    if (log.account_id) {
      repMap[key].engagedAccounts.add(log.account_id);
      if (log.outcome === "Interested") repMap[key].interestedAccounts.add(log.account_id);
    }
  }

  return Object.values(repMap)
    .map(r => ({
      ...r,
      engagedAccounts: r.engagedAccounts.size,
      interestedAccounts: r.interestedAccounts.size,
    }))
    .sort((a, b) => b.totalPoints - a.totalPoints);
}
