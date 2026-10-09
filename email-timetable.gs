var LOGO_FILE_ID = "1EOwCeyW2VxeeY5OIiD-FnHCWIyNblo9C";

function sendExamDutyEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  if (!sheet) {
    showAlert("Error: 'Master' sheet not found.");
    return;
  }

  const trackerSheetName = "Email_Tracker";
  const tz = Session.getScriptTimeZone();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 4 || lastCol < 7) {
    showAlert("No data or date columns found to process.");
    return;
  }

  const COL_NAME = 1;
  const COL_CONTACT = 2;
  const COL_EMAIL = 3;
  const COL_DEPT = 4;
  const COL_SHIFT = 5;
  const COL_DATE_START = 6;

  const headers = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(4, 1, lastRow - 3, lastCol).getValues();

  let logoBlob = null;
  try {
    logoBlob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
  } catch (e) {
    Logger.log("Logo file could not be loaded: " + e.message);
  }

  let trackerSheet = ss.getSheetByName(trackerSheetName);
  if (!trackerSheet) {
    trackerSheet = ss.insertSheet(trackerSheetName);
    trackerSheet.appendRow([
      "Email",
      "Name",
      "Department",
      "Shift(s)",
      "Duties Count",
      "Schedule Signature",
      "Last Sent Timestamp"
    ]);
    trackerSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e2e8f0");
  }

  const trackerLastCol = Math.max(trackerSheet.getLastColumn(), 1);
  const trackerHeaderRow = trackerSheet.getRange(1, 1, 1, trackerLastCol).getValues()[0];

  let emailColIdx = 0;
  let sigColIdx = -1;
  let lastSentColIdx = -1;

  trackerHeaderRow.forEach((colName, idx) => {
    const h = String(colName || "").trim().toLowerCase();
    if (h.includes("email")) emailColIdx = idx;
    if (h.includes("signature")) sigColIdx = idx;
    if (h.includes("last sent") || h.includes("timestamp") || h.includes("sent")) lastSentColIdx = idx;
  });

  if (sigColIdx === -1) {
    sigColIdx = trackerHeaderRow.length >= 6 ? 5 : 3;
  }
  if (lastSentColIdx === -1) {
    lastSentColIdx = sigColIdx + 1;
  }

  const sigCol1Based = sigColIdx + 1;
  const lastSentCol1Based = lastSentColIdx + 1;

  const trackerData = trackerSheet.getDataRange().getValues();
  const trackerMap = {};

  trackerData.slice(1).forEach((row, i) => {
    const trackedEmail = String(row[emailColIdx] || "").trim().toLowerCase();
    if (trackedEmail) {
      trackerMap[trackedEmail] = {
        rowIndex: i + 2,
        signature: String(row[sigColIdx] || "").trim()
      };
    }
  });

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let currentDept = "";
  const facultyMap = {};

  data.forEach((row) => {
    const rawDept = String(row[COL_DEPT] || "").trim();
    if (rawDept) {
      currentDept = rawDept;
    }
    const dept = currentDept;

    const name = String(row[COL_NAME] || "").trim();
    const contact = String(row[COL_CONTACT] || "").trim();
    const email = String(row[COL_EMAIL] || "").trim();
    const rawShift = row[COL_SHIFT];

    if (!email || !email.includes("@")) return;

    const emailKey = email.toLowerCase();
    const shiftInfo = parseShiftDetails(rawShift);

    if (!facultyMap[emailKey]) {
      facultyMap[emailKey] = {
        name: name,
        email: email,
        contact: contact,
        shifts: new Set(),
        departments: new Set(),
        duties: []
      };
    } else {
      if (!facultyMap[emailKey].name && name) {
        facultyMap[emailKey].name = name;
      }
    }

    if (dept) {
      facultyMap[emailKey].departments.add(dept);
    }

    if (shiftInfo) {
      facultyMap[emailKey].shifts.add(shiftInfo.shift);
    }

    for (let c = COL_DATE_START; c < lastCol; c++) {
      const cellVal = String(row[c] || "").trim().toLowerCase();
      if (cellVal !== "yes") continue;

      const headerVal = headers[c];
      const parsedDate = parseDateHeader(headerVal, tz);
      if (!parsedDate) continue;

      if (parsedDate.dateObj < today) {
        continue;
      }

      if (!shiftInfo) {
        Logger.log(`Duty on ${parsedDate.dateKey} for ${email} in ${dept} skipped: Shift not set in Column F.`);
        continue;
      }

      facultyMap[emailKey].duties.push({
        dateObj: parsedDate.dateObj,
        dateKey: parsedDate.dateKey,
        formattedDate: Utilities.formatDate(parsedDate.dateObj, tz, "dd MMM yyyy"),
        dayOfWeek: Utilities.formatDate(parsedDate.dateObj, tz, "EEEE"),
        shift: shiftInfo.shift,
        reportingTime: shiftInfo.reportingTime,
        dept: dept
      });
    }
  });

  let emailsSent = 0;
  let emailsSkipped = 0;
  let skippedNoShift = 0;
  const facultyKeys = Object.keys(facultyMap);

  facultyKeys.forEach((emailKey) => {
    const faculty = facultyMap[emailKey];

    if (faculty.shifts.size === 0) {
      Logger.log(`Skipped ${faculty.name} (${faculty.email}): No Shift assigned in Column F.`);
      skippedNoShift++;
      return;
    }

    const uniqueDutiesMap = {};
    faculty.duties.forEach((d) => {
      const dutyKey = `${d.dateKey}_${d.shift}_${d.dept}`;
      if (!uniqueDutiesMap[dutyKey]) {
        uniqueDutiesMap[dutyKey] = d;
      }
    });

    const consolidatedDuties = Object.values(uniqueDutiesMap);

    consolidatedDuties.sort((a, b) => {
      const diff = a.dateObj.getTime() - b.dateObj.getTime();
      if (diff !== 0) return diff;
      const rankA = a.shift.includes("First") ? 1 : 2;
      const rankB = b.shift.includes("First") ? 1 : 2;
      return rankA - rankB;
    });

    if (consolidatedDuties.length === 0) {
      return;
    }

    const signatureParts = consolidatedDuties.map(
      (d) => `${d.dateKey}_${d.shift}_${d.dept}`
    );
    const signature = signatureParts.join("|");

    let shouldSend = false;
    if (!trackerMap[emailKey]) {
      shouldSend = true;
    } else if (trackerMap[emailKey].signature !== signature) {
      shouldSend = true;
    }

    if (!shouldSend) {
      emailsSkipped++;
      return;
    }

    const deptString = Array.from(faculty.departments).join(", ") || "General";
    const shiftsArray = Array.from(faculty.shifts);
    const shiftsString = shiftsArray.join(" & ");

    const subject = `Examination Duty Timetable - ${faculty.name}`;

    const htmlBody = buildTimetableEmailHTML(
      faculty.name,
      deptString,
      shiftsArray,
      consolidatedDuties
    );

    const recipient = TEST_MODE ? TEST_EMAIL : faculty.email;

    const emailOptions = {
      htmlBody: htmlBody,
      name: "Examination Cell - SBJITMR"
    };

    if (logoBlob) {
      emailOptions.inlineImages = { logoImage: logoBlob };
    }

    try {
      GmailApp.sendEmail(recipient, subject, "", emailOptions);
      emailsSent++;

      if (trackerMap[emailKey]) {
        const rIndex = trackerMap[emailKey].rowIndex;
        trackerSheet.getRange(rIndex, 2).setValue(faculty.name);
        trackerSheet.getRange(rIndex, 3).setValue(deptString);
        if (trackerHeaderRow.length >= 6) {
          trackerSheet.getRange(rIndex, 4).setValue(shiftsString);
          trackerSheet.getRange(rIndex, 5).setValue(consolidatedDuties.length);
        }
        trackerSheet.getRange(rIndex, sigCol1Based).setValue(signature);
        trackerSheet.getRange(rIndex, lastSentCol1Based).setValue(new Date());
        trackerMap[emailKey].signature = signature;
      } else {
        if (trackerHeaderRow.length <= 5) {
          trackerSheet.appendRow([
            faculty.email,
            faculty.name,
            deptString,
            signature,
            new Date()
          ]);
        } else {
          trackerSheet.appendRow([
            faculty.email,
            faculty.name,
            deptString,
            shiftsString,
            consolidatedDuties.length,
            signature,
            new Date()
          ]);
        }
        trackerMap[emailKey] = {
          rowIndex: trackerSheet.getLastRow(),
          signature: signature
        };
      }
    } catch (err) {
      Logger.log(`Failed to send email to ${faculty.email}: ${err.message}`);
    }
  });

  const msg = TEST_MODE
    ? `[TEST MODE ACTIVE]\nSent: ${emailsSent} timetable email(s) to ${TEST_EMAIL}.\nSkipped (Already up to date): ${emailsSkipped}.\nSkipped (Shift not set): ${skippedNoShift}.`
    : `[LIVE MODE]\nSent: ${emailsSent} timetable email(s).\nSkipped (Already up to date): ${emailsSkipped}.\nSkipped (Shift not set): ${skippedNoShift}.`;

  showAlert(msg);
  Logger.log(msg);
}

function buildTimetableEmailHTML(name, dept, shiftsArray, duties) {
  let tableRows = "";

  duties.forEach((d, index) => {
    const rowBg = index % 2 === 0 ? "#ffffff" : "#f8fafc";
    const isFirstShift = d.shift.includes("First");
    const shiftBadgeBg = isFirstShift ? "#e0e7ff" : "#fef3c7";
    const shiftBadgeColor = isFirstShift ? "#3730a3" : "#92400e";
    const shiftBorder = isFirstShift ? "#c7d2fe" : "#fde68a";

    tableRows += `
      <tr style="background-color:${rowBg};">
        <td style="padding:12px 14px;text-align:center;border-bottom:1px solid #e2e8f0;font-weight:600;color:#64748b;font-size:13px;">
          ${index + 1}
        </td>
        <td style="padding:12px 14px;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;white-space:nowrap;">
          ${d.formattedDate}
        </td>
        <td style="padding:12px 14px;border-bottom:1px solid #e2e8f0;color:#475569;font-size:14px;">
          ${d.dayOfWeek}
        </td>
        <td style="padding:12px 14px;border-bottom:1px solid #e2e8f0;">
          <span style="display:inline-block;padding:3px 10px;background:${shiftBadgeBg};color:${shiftBadgeColor};border:1px solid ${shiftBorder};border-radius:12px;font-size:12px;font-weight:700;">
            ${d.shift}
          </span>
          <div style="font-size:11px;color:#64748b;margin-top:2px;font-weight:600;">
            Report: ${d.reportingTime}
          </div>
        </td>
        <td style="padding:12px 14px;border-bottom:1px solid #e2e8f0;color:#1e293b;font-size:13px;">
          <span style="display:inline-block;padding:3px 10px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:6px;font-weight:600;color:#334155;">
            ${d.dept}
          </span>
        </td>
      </tr>
    `;
  });

  let shiftBadgesHTML = "";
  let reportingTimesHTML = "";

  const hasFirst = shiftsArray.some((s) => s.includes("First"));
  const hasSecond = shiftsArray.some((s) => s.includes("Second"));

  if (hasFirst && hasSecond) {
    shiftBadgesHTML = `
      <span style="display:inline-block;padding:4px 12px;background:#e0e7ff;color:#3730a3;border-radius:20px;font-size:12px;font-weight:700;margin-right:6px;">
        First Shift
      </span>
      <span style="display:inline-block;padding:4px 12px;background:#fef3c7;color:#92400e;border-radius:20px;font-size:12px;font-weight:700;">
        Second Shift
      </span>
    `;
    reportingTimesHTML = `
      <div style="font-size:13px;color:#166534;font-weight:700;line-height:1.6;">
        First Shift: <b>10:30 AM sharp</b><br>
        Second Shift: <b>04:00 PM sharp</b>
      </div>
    `;
  } else if (hasSecond) {
    shiftBadgesHTML = `
      <span style="display:inline-block;padding:4px 12px;background:#fef3c7;color:#92400e;border-radius:20px;font-size:13px;font-weight:700;">
        Second Shift
      </span>
    `;
    reportingTimesHTML = `
      <span style="display:inline-block;padding:4px 14px;background:#dcfce7;color:#166534;border-radius:20px;font-size:13px;font-weight:800;border:1px solid #bbf7d0;">
        04:00 PM sharp
      </span>
    `;
  } else {
    shiftBadgesHTML = `
      <span style="display:inline-block;padding:4px 12px;background:#e0e7ff;color:#3730a3;border-radius:20px;font-size:13px;font-weight:700;">
        First Shift
      </span>
    `;
    reportingTimesHTML = `
      <span style="display:inline-block;padding:4px 14px;background:#dcfce7;color:#166534;border-radius:20px;font-size:13px;font-weight:800;border:1px solid #bbf7d0;">
        10:30 AM sharp
      </span>
    `;
  }

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body style="margin:0;padding:0;background-color:#f1f5f9;font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased;">
    
    <div style="max-width:680px;margin:24px auto;padding:0 12px;">

      <div style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 10px 25px -5px rgba(0,0,0,0.06),0 8px 10px -6px rgba(0,0,0,0.04);border:1px solid #cbd5e1;">

        <div style="background:linear-gradient(135deg, #ffffff 0%, #f1f5f9 22%, #e2e8f0 42%, #ffffff 58%, #cbd5e1 78%, #f8fafc 100%);padding:28px 24px 22px;text-align:center;border-bottom:2px solid #cbd5e1;border-top:5px solid #1e293b;">
          
          <div style="display:inline-block;padding:8px 18px;background:rgba(255,255,255,0.7);border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.06);border:1px solid rgba(255,255,255,0.95);margin-bottom:12px;">
            <img src="cid:logoImage" alt="SBJITMR Logo" style="height:76px;max-width:92%;display:block;margin:auto;" />
          </div>

          <h1 style="margin:6px 0 0;font-size:23px;font-weight:800;letter-spacing:-0.3px;color:#0f172a;">
            Examination Duty Timetable
          </h1>
          <p style="margin:5px 0 0;font-size:13px;color:#475569;font-weight:600;">
            Examination Cell &bull; S. B. Jain Institute of Technology, Management & Research, Nagpur
          </p>
        </div>

        <div style="padding:28px 24px;">

          <p style="font-size:16px;margin:0 0 12px;color:#0f172a;">
            Respected <b>${name}</b>,
          </p>

          <p style="font-size:14px;line-height:1.6;color:#475569;margin:0 0 20px;">
            Please find below your finalized invigilation duties for the upcoming examinations. Kindly verify your schedule and ensure punctual attendance.
          </p>

          <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:18px 20px;margin-bottom:24px;">
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
              <tr>
                <td style="padding:6px 0;color:#64748b;width:32%;">Department(s):</td>
                <td style="padding:6px 0;color:#0f172a;font-weight:600;font-size:14px;">${dept}</td>
              </tr>
              <tr>
                <td style="padding:6px 0;color:#64748b;">Assigned Shift(s):</td>
                <td style="padding:6px 0;">
                  ${shiftBadgesHTML}
                </td>
              </tr>
              <tr>
                <td style="padding:6px 0;color:#64748b;vertical-align:top;">Reporting Time:</td>
                <td style="padding:6px 0;">
                  ${reportingTimesHTML}
                </td>
              </tr>
            </table>
          </div>

          <div style="overflow-x:auto;border:1px solid #e2e8f0;border-radius:12px;margin-bottom:24px;">
            <table style="width:100%;border-collapse:collapse;text-align:left;font-size:13px;">
              <thead>
                <tr style="background:#0f172a;color:#ffffff;">
                  <th style="padding:12px 14px;text-align:center;font-weight:600;font-size:12px;letter-spacing:0.5px;text-transform:uppercase;">#</th>
                  <th style="padding:12px 14px;font-weight:600;font-size:12px;letter-spacing:0.5px;text-transform:uppercase;">Date</th>
                  <th style="padding:12px 14px;font-weight:600;font-size:12px;letter-spacing:0.5px;text-transform:uppercase;">Day</th>
                  <th style="padding:12px 14px;font-weight:600;font-size:12px;letter-spacing:0.5px;text-transform:uppercase;">Shift</th>
                  <th style="padding:12px 14px;font-weight:600;font-size:12px;letter-spacing:0.5px;text-transform:uppercase;">Department</th>
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>
          </div>

          <div style="background:#fffbeb;border:1px solid #fde68a;border-left:4px solid #f59e0b;border-radius:8px;padding:18px 20px;margin-bottom:24px;">
            <div style="font-weight:700;color:#92400e;font-size:14px;margin-bottom:10px;">
                Important Invigilation Guidelines:
            </div>
            <ul style="margin:0;padding-left:18px;color:#78350f;font-size:13px;line-height:1.7;">
              <li>
                <b>Punctuality:</b> Kindly report to the Examination Control Room by <b>10:30 AM sharp</b> for First Shift and <b>04:00 PM sharp</b> for Second Shift for receiving answer booklets and question papers.
              </li>
              <li>
                <b>Substitute Protocol:</b> In case of any unavoidable circumstances or emergency, arrange an official substitute in advance and inform your Departmental Time-Table Incharge / HoD at least <b>one day prior</b>.
              </li>
              <li>
                <b>Electronic Gadgets:</b> The use of mobile phones and smart watches inside the examination hall is strictly prohibited.
              </li>
              <li>
                <b>Guidelines Document:</b> For detailed rules, code of conduct, and examination protocols: 
                <a href="https://drive.google.com/file/d/1DTcoiKGcywojvppiU8r1Y7bdXy3URzrv/view?usp=sharing" target="_blank" style="color:#2563eb;text-decoration:underline;font-weight:600;">
                  Click here to download instructions PDF
                </a>.
              </li>
            </ul>
          </div>

          <div style="margin-top:28px;padding-top:16px;border-top:1px solid #f1f5f9;color:#334155;font-size:13px;line-height:1.5;">
            Warm regards,<br>
            <b style="font-size:14px;color:#0f172a;">Exam Officer In-charge</b><br>
            Examination Cell<br>
            S. B. Jain Institute of Technology, Management & Research, Nagpur
          </div>

        </div>

        <div style="background:#f8fafc;padding:14px 20px;text-align:center;font-size:11px;color:#64748b;border-top:1px solid #e2e8f0;letter-spacing:0.2px;">
          This is an automated system notification from SBJITMR Examination Cell. Please do not reply directly to this email.
        </div>

      </div>

    </div>

  </body>
  </html>
  `;
}

function parseShiftDetails(rawShift) {
  const str = String(rawShift || "").trim().toLowerCase();
  if (str.includes("second") || str.includes("2nd") || str.includes("ii")) {
    return {
      shift: "Second Shift",
      reportingTime: "04:00 PM"
    };
  }
  if (str.includes("first") || str.includes("1st") || str.includes("i")) {
    return {
      shift: "First Shift",
      reportingTime: "10:30 AM"
    };
  }
  return null;
}

function parseDateHeader(val, tz) {
  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    return {
      dateObj: val,
      dateKey: Utilities.formatDate(val, tz, "dd-MM-yyyy")
    };
  }

  const str = String(val).trim();
  const match = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (match) {
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const year = parseInt(match[3], 10);
    const d = new Date(year, month, day, 0, 0, 0);
    return {
      dateObj: d,
      dateKey: Utilities.formatDate(d, tz, "dd-MM-yyyy")
    };
  }

  const fallback = new Date(str);
  if (!isNaN(fallback.getTime())) {
    return {
      dateObj: fallback,
      dateKey: Utilities.formatDate(fallback, tz, "dd-MM-yyyy")
    };
  }

  return null;
}

function showAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}