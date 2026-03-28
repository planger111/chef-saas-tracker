// Chef SaaS Motion Tracker — Export Functions

const EXPORT_COLUMNS = [
  "submitted_at",
  "rep_name",
  "rep_email",
  "account_id",
  "account_name",
  "play_id",
  "play_name",
  "interaction_type",
  "outcome",
  "reason_label",
  "next_step_type",
  "timing",
  "contact_engaged",
  "pitch_confidence",
  "short_reaction",
  "notes",
  "points_earned",
  "id",
];

const EXPORT_HEADERS = {
  submitted_at:     "Date / Time",
  rep_name:         "Rep Name",
  rep_email:        "Rep Email",
  account_id:       "Account ID",
  account_name:     "Account Name",
  play_id:          "Play ID",
  play_name:        "Play Name",
  interaction_type: "Interaction Type",
  outcome:          "Outcome",
  reason_label:     "Reason",
  next_step_type:   "Next Step",
  timing:           "Timing",
  contact_engaged:  "Contact Engaged",
  pitch_confidence: "Pitch Confidence",
  short_reaction:   "Short Reaction",
  notes:            "Notes",
  points_earned:    "Points Earned",
  id:               "Record ID",
};

function _toRows(data) {
  return data.map((item) =>
    EXPORT_COLUMNS.reduce((row, col) => {
      row[col] = item[col] !== undefined ? item[col] : "";
      return row;
    }, {})
  );
}

function _escapeCsvValue(value) {
  const str = String(value === null || value === undefined ? "" : value);
  if (str.includes(",") || str.includes('"') || str.includes("\n")) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

function exportToCSV(data, filename) {
  const rows = _toRows(data);
  const headers = EXPORT_COLUMNS.map((c) => EXPORT_HEADERS[c] || c);
  const lines = [
    headers.join(","),
    ...rows.map((row) =>
      EXPORT_COLUMNS.map((col) => _escapeCsvValue(row[col])).join(",")
    ),
  ];

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  _triggerDownload(blob, filename || "engagements-export.csv");
}

function exportToExcel(data, filename) {
  if (!window.XLSX) {
    alert("Excel export library not loaded. Please check your internet connection.");
    return;
  }

  const rows = _toRows(data);
  const sheetData = [
    EXPORT_COLUMNS.map((c) => EXPORT_HEADERS[c] || c),
    ...rows.map((row) => EXPORT_COLUMNS.map((col) => row[col])),
  ];

  const workbook = window.XLSX.utils.book_new();
  const worksheet = window.XLSX.utils.aoa_to_sheet(sheetData);

  // Auto-size columns
  const colWidths = EXPORT_COLUMNS.map((col, i) => ({
    wch: Math.max(EXPORT_HEADERS[col].length, ...sheetData.slice(1).map(r => String(r[i]||"").length), 10)
  }));
  worksheet["!cols"] = colWidths;

  window.XLSX.utils.book_append_sheet(workbook, worksheet, "Engagements");
  window.XLSX.writeFile(workbook, filename || "engagements-export.xlsx");
}

function _triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
