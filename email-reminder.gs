var REMINDER_LOGO_FILE_ID = "1EOwCeyW2VxeeY5OIiD-FnHCWIyNblo9C";

function sendReminderEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  if (!sheet) {
    Logger.log("Error: 'Master' sheet not found.");
    return;
  }

  const tz = Session.getScriptTimeZone();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 4 || lastCol < 7) {
    Logger.log("No data or date columns found to process.");
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
    logoBlob = DriveApp.getFileById(REMINDER_LOGO_FILE_ID).getBlob();
  } catch (e) {
    Logger.log("Reminder Logo file could not be loaded: " + e.message);
  }

  let targetDateKey = "";
  let targetDateObj = null;

  if (TEST_MODE) {
    targetDateKey = typeof TEST_DATE !== "undefined" && TEST_DATE ? TEST_DATE : "12-10-2026";
    const testMatch = targetDateKey.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
    if (testMatch) {
      targetDateObj = new Date(
        parseInt(testMatch[3], 10),
        parseInt(testMatch[2], 10) - 1,
        parseInt(testMatch[1], 10)
      );
    } else {
      targetDateObj = new Date();
    }
  } else {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0);
    targetDateObj = tomorrow;
    targetDateKey = Utilities.formatDate(tomorrow, tz, "dd-MM-yyyy");
  }

  const matchingDateCols = [];

  for (let c = COL_DATE_START; c < lastCol; c++) {
    const headerVal = headers[c];
    const parsed = parseReminderDateHeader(headerVal, tz);
    if (parsed && parsed.dateKey === targetDateKey) {
      matchingDateCols.push({
        colIndex: c,
        dateObj: parsed.dateObj,
        dateKey: parsed.dateKey
      });
      targetDateObj = parsed.dateObj;
    }
  }

  if (matchingDateCols.length === 0) {
    Logger.log(`No exam date column found matching target date: ${targetDateKey}. No reminders to send.`);
    return;
  }

  let currentDept = "";
  const scheduledFacultyMap = {};

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
    const shiftInfo = parseReminderShiftDetails(rawShift);

    let hasDutyOnTargetDate = false;
    matchingDateCols.forEach((colInfo) => {
      const cellVal = String(row[colInfo.colIndex] || "").trim().toLowerCase();
      if (cellVal === "yes") {
        hasDutyOnTargetDate = true;
      }
    });

    if (!hasDutyOnTargetDate) return;

    if (!shiftInfo) {
      Logger.log(`Reminder skipped for ${email} in ${dept}: Shift not set in Column F.`);
      return;
    }

    if (!scheduledFacultyMap[emailKey]) {
      scheduledFacultyMap[emailKey] = {
        name: name,
        email: email,
        contact: contact,
        shifts: new Set([shiftInfo.shift]),
        departments: new Set([dept]),
        duties: [
          {
            dept: dept,
            shift: shiftInfo.shift,
            reportingTime: shiftInfo.reportingTime
          }
        ],
        dateObj: targetDateObj
      };
    } else {
      if (dept) {
        scheduledFacultyMap[emailKey].departments.add(dept);
      }
      if (!scheduledFacultyMap[emailKey].name && name) {
        scheduledFacultyMap[emailKey].name = name;
      }
      scheduledFacultyMap[emailKey].shifts.add(shiftInfo.shift);

      const isDuplicate = scheduledFacultyMap[emailKey].duties.some(
        (d) => d.dept === dept && d.shift === shiftInfo.shift
      );
      if (!isDuplicate) {
        scheduledFacultyMap[emailKey].duties.push({
          dept: dept,
          shift: shiftInfo.shift,
          reportingTime: shiftInfo.reportingTime
        });
      }
    }
  });

  let emailsSent = 0;
  let skippedNoShift = 0;
  const facultyKeys = Object.keys(scheduledFacultyMap);

  if (facultyKeys.length === 0) {
    Logger.log(`No faculty assigned duties for ${targetDateKey}.`);
    return;
  }

  const formattedDate = Utilities.formatDate(targetDateObj, tz, "dd MMM yyyy");
  const dayOfWeek = Utilities.formatDate(targetDateObj, tz, "EEEE");

  facultyKeys.forEach((emailKey) => {
    const faculty = scheduledFacultyMap[emailKey];

    if (faculty.duties.length === 0) {
      skippedNoShift++;
      return;
    }

    faculty.duties.sort((a, b) => {
      const rankA = a.shift.includes("First") ? 1 : 2;
      const rankB = b.shift.includes("First") ? 1 : 2;
      return rankA - rankB;
    });

    const deptString = Array.from(faculty.departments).join(", ") || "General";
    const shiftsString = Array.from(faculty.shifts).join(" & ");

    const subject = `Reminder: Exam Duty Tomorrow - ${formattedDate} (${shiftsString})`;

    const htmlBody = buildReminderEmailHTML(
      faculty.name,
      deptString,
      faculty.duties,
      formattedDate,
      dayOfWeek
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
      Logger.log(`Reminder sent to: ${recipient} (Faculty: ${faculty.name}, Duties: ${faculty.duties.length})`);
    } catch (err) {
      Logger.log(`Failed reminder for ${faculty.email}: ${err.message}`);
    }
  });

  Logger.log(`Reminder process finished. Sent: ${emailsSent}, Skipped (No Shift): ${skippedNoShift} for ${targetDateKey}.`);
}

function buildReminderEmailHTML(name, dept, duties, formattedDate, dayOfWeek) {
  let dutyRowsHTML = "";

  duties.forEach((d) => {
    const isFirstShift = d.shift.includes("First");
    const shiftBadgeBg = isFirstShift ? "#e0e7ff" : "#fef3c7";
    const shiftBadgeColor = isFirstShift ? "#3730a3" : "#92400e";
    const shiftBorder = isFirstShift ? "#c7d2fe" : "#fde68a";

    dutyRowsHTML += `
      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px;margin-bottom:12px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
          <span style="display:inline-block;padding:3px 12px;background:${shiftBadgeBg};color:${shiftBadgeColor};border:1px solid ${shiftBorder};border-radius:20px;font-size:12px;font-weight:700;">
            ${d.shift}
          </span>
          <span style="display:inline-block;padding:3px 10px;background:#dcfce7;color:#166534;border:1px solid #bbf7d0;border-radius:20px;font-size:12px;font-weight:800;">
            Report: ${d.reportingTime} sharp
          </span>
        </div>
        <div style="font-size:13px;color:#475569;">
          Department: <b style="color:#0f172a;">${d.dept}</b>
        </div>
      </div>
    `;
  });

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body style="margin:0;padding:0;background-color:#f1f5f9;font-family:'Segoe UI',-apple-system,BlinkMacSystemFont,Roboto,Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased;">
    
    <div style="max-width:640px;margin:24px auto;padding:0 12px;">

      <div style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 10px 25px -5px rgba(0,0,0,0.06),0 8px 10px -6px rgba(0,0,0,0.04);border:1px solid #cbd5e1;">

        <div style="background:linear-gradient(135deg, #ffffff 0%, #f1f5f9 22%, #e2e8f0 42%, #ffffff 58%, #cbd5e1 78%, #f8fafc 100%);padding:28px 24px 22px;text-align:center;border-bottom:2px solid #cbd5e1;border-top:5px solid #1e293b;">
          
          <div style="display:inline-block;padding:8px 18px;background:rgba(255,255,255,0.7);border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.06);border:1px solid rgba(255,255,255,0.95);margin-bottom:12px;">
            <img src="cid:logoImage" alt="SBJITMR Logo" style="height:76px;max-width:92%;display:block;margin:auto;" />
          </div>

          <div>
            <span style="display:inline-block;background:#e2e8f0;color:#1e293b;padding:4px 14px;border-radius:20px;font-size:12px;font-weight:700;letter-spacing:0.5px;margin-bottom:6px;text-transform:uppercase;border:1px solid #cbd5e1;">
              ⏰ Official Duty Reminder
            </span>
          </div>

          <h1 style="margin:4px 0 0;font-size:23px;font-weight:800;color:#0f172a;letter-spacing:-0.3px;">
            Exam Duty Scheduled Tomorrow
          </h1>
          <p style="margin:5px 0 0;font-size:13px;color:#475569;font-weight:600;">
            Examination Cell &bull; S. B. Jain Institute of Technology, Management & Research, Nagpur
          </p>
        </div>

        <div style="padding:28px 24px;">

          <p style="font-size:16px;margin:0 0 10px;color:#0f172a;">
            Respected <b>${name}</b>,
          </p>

          <p style="font-size:14px;line-height:1.6;color:#475569;margin:0 0 24px;">
            This is a gentle reminder that you are scheduled for examination invigilation duty <b>tomorrow</b>. Please review your duty schedule below:
          </p>

          <div style="background:linear-gradient(180deg,#f8fafc 0%,#f1f5f9 100%);border:2px solid #e0e7ff;border-radius:14px;padding:22px 20px;margin-bottom:24px;">

            <div style="text-align:center;padding-bottom:16px;border-bottom:1px solid #e2e8f0;margin-bottom:16px;">
              <div style="font-size:12px;font-weight:700;color:#4f46e5;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">
                Exam Date & Day
              </div>
              <div style="font-size:22px;font-weight:800;color:#0f172a;letter-spacing:-0.3px;">
                ${formattedDate}
              </div>
              <div style="font-size:14px;font-weight:600;color:#64748b;margin-top:2px;">
                ${dayOfWeek}
              </div>
            </div>

            <div style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">
              Tomorrow's Assigned Duty / Shift(s):
            </div>
            ${dutyRowsHTML}

          </div>

          <div style="background:#fff7ed;border:1px solid #fed7aa;border-left:4px solid #f97316;border-radius:8px;padding:16px 18px;margin-bottom:24px;">
            <div style="font-weight:700;color:#9a3412;font-size:14px;margin-bottom:6px;">
              ⚠️ Reporting & Adjustment Protocol:
            </div>
            <p style="margin:0;font-size:13px;line-height:1.6;color:#7c2d12;">
              Kindly report to the Examination Control Room by <b>10:30 AM sharp</b> for First Shift and <b>04:00 PM sharp</b> for Second Shift. If you require an emergency exchange or substitute, please notify your Departmental Timetable In-charge or HoD by <b>2:00 PM today</b>. Adjustments may not be feasible after this time.
            </p>
          </div>

          <div style="margin-top:28px;padding-top:16px;border-top:1px solid #f1f5f9;color:#334155;font-size:13px;line-height:1.5;">
            Warm regards,<br>
            <b style="font-size:14px;color:#0f172a;">Exam Officer In-charge</b><br>
            Examination Cell<br>
            S. B. Jain Institute of Technology, Management & Research, Nagpur
          </div>

        </div>

        <div style="background:#f8fafc;padding:14px 20px;text-align:center;font-size:11px;color:#64748b;border-top:1px solid #e2e8f0;letter-spacing:0.2px;">
          This is an automated system reminder from SBJITMR Examination Cell. Please do not reply directly to this email.
        </div>

      </div>

    </div>

  </body>
  </html>
  `;
}

function parseReminderShiftDetails(rawShift) {
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

function parseReminderDateHeader(val, tz) {
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

function installDailyReminderTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach((trig) => {
    if (trig.getHandlerFunction() === "sendReminderEmails") {
      ScriptApp.deleteTrigger(trig);
    }
  });

  ScriptApp.newTrigger("sendReminderEmails")
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  Logger.log("Daily Reminder Trigger installed successfully for 6:00 PM.");
}