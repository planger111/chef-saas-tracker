// Chef SaaS Motion Tracker — Export Functions

const EXPORT_COLUMNS = [
  "log_id",
  "account_id",
  "account_name",
  "rep_name",
  "csm_name",
  "sales_play",
  "motion_type",
  "answer_status",
  "answer_why_not_pitched",
  "answer_customer_need",
  "answer_blocker",
  "answer_date_last_pitch",
  "answer_date_next_step",
  "notes",
  "submitted_at",
  "submitted_by",
  "source",
];

const EXPORT_HEADERS = {
  log_id: "Log ID",
  account_id: "Account ID",
  account_name: "Account Name",
  rep_name: "Rep Name",
  csm_name: "CSM Name",
  sales_play: "Sales Play",
  motion_type: "Motion Type",
  answer_status: "Status",
  answer_why_not_pitched: "Why Not Pitched",
  answer_customer_need: "Customer Need",
  answer_blocker: "Blocker",
  answer_date_last_pitch: "Last Pitch Date",
  answer_date_next_step: "Next Step Date",
  notes: "Notes",
  submitted_at: "Submitted At",
  submitted_by: "Submitted By",
  source: "Source",
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
  _triggerDownload(blob, filename || "motion-log.csv");
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
  window.XLSX.utils.book_append_sheet(workbook, worksheet, "Motion Log");
  window.XLSX.writeFile(workbook, filename || "motion-log.xlsx");
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
