/**
 * CSAT DASHBOARD - BACKEND (Code.gs)
 */

const CONFIG = {
  EXTERNAL_SHEET_ID: "1EN8zM6wdVeAaokTXi27TyENfZu-e1s3yRaR20V4VTYI",
  EXCLUDED_SHEETS: ["Ticketing", "Onboarding Data", "Master", "Tasks", "NSDC"], 
  TIMEZONE: Session.getScriptTimeZone()
};

function doGet(e) {
  if (e && e.parameter.page === 'ticket') {
    return HtmlService.createHtmlOutputFromFile("ticket")
      .setTitle("Raise a Support Ticket")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("CSAT Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ------------------ SESSIONS ------------------ */
function getSessions() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(s => s.getName())
    .filter(name => !CONFIG.EXCLUDED_SHEETS.includes(name));
}

/* ------------------ MAIN CALCULATION (CSAT) ------------------ */
function calculateAll(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let stats = { promoters: 0, passives: 0, detractors: 0, total: 0 };
  let feedbacks = [];

  const startLimit = filters.start ? new Date(filters.start).getTime() : null;
  const endLimit = filters.end ? new Date(filters.end).setHours(23,59,59,999) : null;
  const targetRating = filters.rating !== "All" ? Number(filters.rating) : null;

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (CONFIG.EXCLUDED_SHEETS.includes(name)) return;
    if (filters.session && filters.session !== "All" && name !== filters.session) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    
    const idx = {
      date: headers.indexOf("timestamp"),
      rate: headers.findIndex(h => h.includes("rate") || h.includes("score") || h.includes("recommend")),
      mob: headers.findIndex(h => h.includes("mobile") || h.includes("phone") || h.includes("contact")),
      name: headers.findIndex(h => h.includes("name")),
      feed: headers.findIndex(h => h.includes("feedback") || h.includes("remark") || h.includes("comment"))
    };

    if (idx.date === -1 || idx.rate === -1) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[idx.date];
      if (!rawDate || row[idx.rate] === "") continue;

      const rowTime = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();

      if (startLimit && rowTime < startLimit) continue;
      if (endLimit && rowTime > endLimit) continue;

      const rating = parseInt(row[idx.rate].toString().replace(/\D/g, ''));
      if (isNaN(rating)) continue;
      if (targetRating && rating !== targetRating) continue;

      if (rating >= 4) stats.promoters++;
      else if (rating === 3) stats.passives++;
      else if (rating >= 1) stats.detractors++;

      if (feedbacks.length < 300) {
        let mobileNo = "N/A";
        if (idx.mob !== -1 && row[idx.mob]) { mobileNo = row[idx.mob].toString(); } 
        else { let possiblePhone = row.find(cell => { let str = String(cell).replace(/\D/g, ''); return str.length >= 10 && str.length <= 15; }); if (possiblePhone) mobileNo = String(possiblePhone).trim(); }

        feedbacks.push({
          date: Utilities.formatDate(new Date(rowTime), CONFIG.TIMEZONE, "dd-MM-yyyy"),
          session: name,
          name: idx.name !== -1 ? (row[idx.name] || "Anonymous") : "Anonymous",
          mobile: mobileNo,
          rating: rating,
          comment: idx.feed !== -1 ? (row[idx.feed] || "-") : "-"
        });
      }
    }
  });

  stats.total = stats.promoters + stats.passives + stats.detractors;
  const nps = stats.total === 0 ? 0 : Math.round(((stats.promoters - stats.detractors) / stats.total) * 100);

  return { stats: { ...stats, nps }, feedbacks: feedbacks.reverse() };
}

/* ------------------ HELPER FUNCTION FOR TAT ------------------ */
function calculateTAT(startDate, endDate) {
  if (!startDate || !endDate) return "N/A";
  const start = new Date(startDate).getTime();
  const end = new Date(endDate).getTime();
  if (isNaN(start) || isNaN(end)) return "N/A";
  
  const diffInMs = end - start;
  if (diffInMs < 0) return "N/A";
  
  const diffInHours = Math.floor(diffInMs / (1000 * 60 * 60));
  const diffInDays = Math.floor(diffInHours / 24);
  
  if (diffInDays > 0) return `${diffInDays} Day${diffInDays > 1 ? 's' : ''}`;
  else if (diffInHours > 0) return `${diffInHours} Hr${diffInHours > 1 ? 's' : ''}`;
  else { const diffInMins = Math.floor(diffInMs / (1000 * 60)); return `${diffInMins} Min${diffInMins > 1 ? 's' : ''}`; }
}

/* ------------------ TICKETING ------------------ */
function saveTicket(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Ticketing") || ss.insertSheet("Ticketing");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date Raised", "Name", "Phone", "Email", "Issue", "Remark", "Status", "Course", "Source", "Assigned To", "Priority", "Date Resolved", "TAT"]);
    sheet.getRange("A1:M1").setFontWeight("bold").setBackground("#f3f3f3");
  }

  let finalPriority = data.priority ? data.priority : "Urgent";
  let finalSource = data.source ? data.source : "Web Form";

  sheet.appendRow([ new Date(), data.name, data.phone || "-", data.email || "-", data.issue, data.remark || "-", "Untouched", data.course || "N/A", finalSource, data.assigned || "Unassigned", finalPriority, "", "" ]);
  return "Success";
}

function getTickets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ticketing");
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getDataRange().getValues().slice(1).map((row, i) => ({
    row: i + 2, dateRaised: row[0] instanceof Date ? Utilities.formatDate(row[0], CONFIG.TIMEZONE, "dd-MMM HH:mm") : (row[0] || "-"),
    name: row[1] || "-", phone: row[2] || "-", issue: row[4] || "-", remark: row[5] || "-", status: row[6] || "Untouched",
    course: row[7] || "N/A", source: row[8] || "N/A", assigned: row[9] || "Unassigned", priority: row[10] || "Medium",
    dateResolved: row[11] instanceof Date ? Utilities.formatDate(row[11], CONFIG.TIMEZONE, "dd-MMM HH:mm") : (row[11] || "-"), tat: row[12] || "-"
  })).reverse();
}

function updateTicketStatus(row, status, remark) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ticketing");
  sheet.getRange(row, 7).setValue(status);
  
  if (status === "Solved") {
    const resolvedDate = new Date();
    sheet.getRange(row, 12).setValue(resolvedDate);
    const raisedDate = sheet.getRange(row, 1).getValue();
    sheet.getRange(row, 13).setValue(calculateTAT(raisedDate, resolvedDate));
  } else {
    sheet.getRange(row, 12).setValue("");
    sheet.getRange(row, 13).setValue("");
  }
  if (remark) { const cell = sheet.getRange(row, 6); cell.setValue(cell.getValue() + "\n[Res]: " + remark); }
}

function editTicketDetails(row, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ticketing");
  sheet.getRange(row, 8).setValue(data.course);    
  sheet.getRange(row, 9).setValue(data.source);    
  sheet.getRange(row, 10).setValue(data.assigned); 
  sheet.getRange(row, 11).setValue(data.priority); 
  return "Success";
}

/* ------------------ ONBOARDING DATA (INTERNAL TAB) ------------------ */
function saveEnrollmentData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Onboarding Data") || ss.insertSheet("Onboarding Data");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date Added", "Course", "Month", "Count", "Onboarded By"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
  }

  sheet.appendRow([new Date(), data.course, data.month, data.count, data.by]);
  return "Success";
}

// Safe Fetch for Internal Enrollments Table
function getEnrollmentData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding Data");
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    let result = [];
    const tz = Session.getScriptTimeZone() || "GMT";

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (!row[0]) continue; 

      let dateStr = "";
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], tz, "dd-MMM-yyyy HH:mm");
      } else {
        dateStr = String(row[0]);
      }

      let monthStr = "";
      if (row[2] instanceof Date) {
         monthStr = Utilities.formatDate(row[2], tz, "MMM-yyyy");
      } else {
         monthStr = String(row[2] || "-");
      }

      result.push({
        date: dateStr,
        course: String(row[1] || "-"),
        month: monthStr,
        count: String(row[3] || "0"),
        by: String(row[4] || "-")
      });
    }
    return result.reverse();
    
  } catch(error) {
    throw new Error("Backend Error: " + error.toString());
  }
}

/* ------------------ DASHBOARD KPI: ENROLLED COUNT (NEW) ------------------ */
function getEnrolledStats() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding Data");
    if (!sheet || sheet.getLastRow() < 2) return 0;
    
    const data = sheet.getDataRange().getValues().slice(1);
    let totalCount = 0;
    
    for (let i = 0; i < data.length; i++) {
      let val = parseInt(data[i][3]); // Count is in Col D (Index 3)
      if (!isNaN(val)) {
        totalCount += val;
      }
    }
    return totalCount;
  } catch(e) {
    return 0;
  }
}

/* ------------------ DASHBOARD CHART DATA: ONBOARDING GRAPH ------------------ */
function getOnboardingChartData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding Data");
    if (!sheet || sheet.getLastRow() < 2) return { labels: [], data: [] };

    const data = sheet.getDataRange().getValues().slice(1);
    const aggregated = {};
    const tz = Session.getScriptTimeZone() || "GMT";

    data.forEach(row => {
      let dateVal = row[0];
      let count = parseInt(row[3]) || 0; // Count column index 3
      if (!dateVal || count === 0) return;
      
      let dateStr = "";
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, tz, "dd-MMM");
      } else {
        dateStr = new Date(dateVal).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
      }

      if (aggregated[dateStr]) {
        aggregated[dateStr] += count;
      } else {
        aggregated[dateStr] = count;
      }
    });

    return {
      labels: Object.keys(aggregated),
      data: Object.values(aggregated)
    };
  } catch(e) {
    return { labels: [], data: [] };
  }
}

/* ------------------ ONBOARDING STATS & TABLE (External) ------------------ */
function getOnboardingStats() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.EXTERNAL_SHEET_ID).getSheetByName("Registered in APP");
    return Math.max(sheet.getLastRow() - 1, 0);
  } catch(e) { 
    return 0; 
  }
}

function getOnboardingTable() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.EXTERNAL_SHEET_ID);
    const sheet = ss.getSheetByName("Registered in APP");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    const fNameIdx = headers.indexOf("Learner First Name");
    const lNameIdx = headers.indexOf("Learner Last Name");
    const phoneIdx = headers.indexOf("Mobile Number");
    const emailIdx = headers.indexOf("Email ID");
    const programIdx = headers.indexOf("Selected Program");
    const sourceIdx = headers.indexOf("Source Tab");

    return data.slice(1).reverse().slice(0, 300).map(row => ({
      name: ((row[fNameIdx] || "") + " " + (row[lNameIdx] || "")).trim() || "N/A",
      phone: row[phoneIdx] || "N/A",
      email: row[emailIdx] || "N/A",
      program: row[programIdx] || "N/A",
      source: row[sourceIdx] || "N/A"
    }));
  } catch(e) { return []; }
}

/* ------------------ DAILY TASKS ------------------ */
function saveTask(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Tasks") || ss.insertSheet("Tasks");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Task Name", "Executive", "Status"]);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  sheet.appendRow([new Date(), data.task, data.exec, data.status]);
  return "Success";
}

function getTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map((row, i) => ({
    row: i + 2, date: row[0] instanceof Date ? Utilities.formatDate(row[0], CONFIG.TIMEZONE, "dd-MMM-yyyy HH:mm") : row[0],
    task: row[1], executive: row[2], status: row[3]
  })).reverse();
}

function updateTaskStatusRow(row, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
  sheet.getRange(row, 4).setValue(status);
}

/* ------------------ NSDC / CERTIFICATION ------------------ */
function saveNsdcData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("NSDC") || ss.insertSheet("NSDC");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Action Type", "Batch Number", "Count"]);
    sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  sheet.appendRow([new Date(), data.type, data.batch, data.count]);
  return "Success";
}

function getNsdcData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NSDC");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    date: row[0] instanceof Date ? Utilities.formatDate(row[0], CONFIG.TIMEZONE, "dd-MMM-yyyy HH:mm") : row[0],
    type: row[1], batch: row[2], count: row[3]
  })).reverse();
}
