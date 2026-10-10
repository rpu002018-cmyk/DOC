/***************************************************************
 * ADMISSION DOCUMENT ACKNOWLEDGEMENT SYSTEM
 * Backend: Google Apps Script (code.gs)
 *
 * Sheets:
 * 1. Students  : Register No, Application No, Name, Section,
 *                Sub Section, Email
 * 2. Documents : Application No, Student Name, Document Name,
 *                Original, Xerox, Status, Submitted Date,
 *                Submitted By, Received By, Last Updated, Updated By
 * 3. Operators : Operator ID, Password, Operator Name, Role
 * 4. Email Log : Timestamp, Application No, Student Name,
 *                Email, Operator Name
 *
 * Added:
 * - Admin/Operator roles
 * - Server-verified admin-only pending report
 * - Pending report data for PDF generation in index.html
 * - Preservation of existing document receipt details
 ***************************************************************/

/***********************
 * CONFIGURATION
 ***********************/
const SHEET_ID = "1TFnZhXGQMszla0Xsaa6raof9zylLoEtuu1WIE_7X8uw";
const DATE_FORMAT = "dd-MM-yyyy";
const TZ = Session.getScriptTimeZone() || "Asia/Kolkata";

const SHEETS = {
  STUDENTS: "Students",
  DOCUMENTS: "Documents",
  OPERATORS: "Operators",
  EMAIL_LOG: "Email Log"
};

const REQUIRED_DOCUMENTS = [
  "10TH Hall Ticket Xerox",
  "10TH marks sheet Original",
  "10TH marks sheet Xerox",
  "Transfer Certificate Original",
  "Transfer Certificate Xerox",
  "Migration Certificate Original",
  "Migration Certificate Xerox",
  "Student Photo",
  "Caste Certificate Xerox",
  "Aadhar Card Student Xerox",
  "Aadhar Card Mother Xerox",
  "Aadhar Card Father Xerox",
  "Study Certificate Xerox"
];

/***********************
 * COMMON HELPERS
 ***********************/
function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheetOrThrow(sheetName) {
  const sheet = getSS().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Required sheet "' + sheetName + '" was not found.');
  }
  return sheet;
}

function toBool(value) {
  if (value === true) return true;
  if (value === false || value == null || value === "") return false;

  const s = String(value).trim().toLowerCase();
  return ["true", "yes", "y", "1", "received"].includes(s);
}

function cleanText(value) {
  return String(value == null ? "" : value).trim();
}

function formatDate(value) {
  if (!value) return "";
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, TZ, DATE_FORMAT);
}

function normalizeKey(value) {
  return cleanText(value).toLowerCase();
}

function ensureHeader(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1, 1, 1, headers.length)
    .getDisplayValues()[0];

  const isBlank = current.every(v => !cleanText(v));

  if (isBlank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

/***********************
 * WEB APP ENTRY
 ***********************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Admission Document Acknowledgement")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/***********************
 * OPERATOR LOGIN
 *
 * Operators sheet:
 * A = Operator ID
 * B = Password
 * C = Operator Name
 * D = Role (Admin / Operator)
 ***********************/
function validateOperator(id, pwd) {
  try {
    const sh = getSheetOrThrow(SHEETS.OPERATORS);
    const data = sh.getDataRange().getDisplayValues();

    const userId = cleanText(id);
    const userPassword = cleanText(pwd);

    if (!userId || !userPassword) {
      return {
        success: false,
        message: "Enter Operator ID and password."
      };
    }

    for (let i = 1; i < data.length; i++) {
      const rowId = cleanText(data[i][0]);
      const rowPassword = cleanText(data[i][1]);

      if (rowId === userId && rowPassword === userPassword) {
        const name = cleanText(data[i][2]) || userId;

        // Existing rows without a role default to Operator.
        // Admin access must be explicitly assigned in column D.
        const role = cleanText(data[i][3]) || "Operator";

        return {
          success: true,
          name: name,
          operatorId: userId,
          role: role
        };
      }
    }

    return {
      success: false,
      message: "Invalid Operator ID or password."
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || "Login validation failed."
    };
  }
}

/***********************
 * STUDENT REGISTRATION
 *
 * student:
 * {
 *   registerNo, applicationNo, name,
 *   section, subSection, email
 * }
 ***********************/
function addStudent(student) {
  try {
    const sh = getSheetOrThrow(SHEETS.STUDENTS);

    if (
      !student ||
      !cleanText(student.registerNo) ||
      !cleanText(student.applicationNo) ||
      !cleanText(student.name)
    ) {
      return {
        success: false,
        error: "Register No, Application No and Name are required."
      };
    }

    const registerNo = cleanText(student.registerNo);
    const applicationNo = cleanText(student.applicationNo);
    const name = cleanText(student.name);

    const data = sh.getDataRange().getDisplayValues();

    for (let i = 1; i < data.length; i++) {
      const existingRegisterNo = cleanText(data[i][0]);
      const existingApplicationNo = cleanText(data[i][1]);

      if (
        normalizeKey(existingRegisterNo) === normalizeKey(registerNo) ||
        normalizeKey(existingApplicationNo) === normalizeKey(applicationNo)
      ) {
        return {
          success: false,
          error: "A student with the same Register No or Application No already exists."
        };
      }
    }

    sh.appendRow([
      registerNo,
      applicationNo,
      name,
      cleanText(student.section),
      cleanText(student.subSection),
      cleanText(student.email)
    ]);

    return { success: true };
  } catch (error) {
    return {
      success: false,
      error: error.message || "Unable to register student."
    };
  }
}

/***********************
 * STUDENT SEARCH
 * Search by Register No or Application No.
 ***********************/
function searchStudent(value) {
  try {
    const sh = getSheetOrThrow(SHEETS.STUDENTS);
    const data = sh.getDataRange().getDisplayValues();
    const searchValue = cleanText(value);

    if (!searchValue) return null;

    for (let i = 1; i < data.length; i++) {
      const registerNo = cleanText(data[i][0]);
      const applicationNo = cleanText(data[i][1]);

      if (
        normalizeKey(registerNo) === normalizeKey(searchValue) ||
        normalizeKey(applicationNo) === normalizeKey(searchValue)
      ) {
        return {
          registerNo: data[i][0],
          applicationNo: data[i][1],
          name: data[i][2],
          section: data[i][3],
          subSection: data[i][4],
          email: data[i][5]
        };
      }
    }

    return null;
  } catch (error) {
    throw new Error(error.message || "Student search failed.");
  }
}

/***********************
 * LOAD EXISTING DOCUMENTS
 ***********************/
function loadExistingDocuments(applicationNo) {
  const sh = getSheetOrThrow(SHEETS.DOCUMENTS);
  const data = sh.getDataRange().getValues();
  const appNo = cleanText(applicationNo);

  return data
    .filter((row, index) =>
      index > 0 &&
      cleanText(row[0]) === appNo
    )
    .map(row => ({
      name: row[2],
      original: toBool(row[3]),
      xerox: toBool(row[4]),
      status: row[5] ||
        ((toBool(row[3]) || toBool(row[4])) ? "Submitted" : "Pending"),
      submittedDate: formatDate(row[6]),
      submittedBy: row[7] || "",
      receivedBy: row[8] || ""
    }));
}

/***********************
 * SAVE DOCUMENTS
 *
 * Existing submitted flags cannot be removed.
 * Existing submitted date/by/received by are preserved.
 * Received By is assigned only if it was previously blank.
 ***********************/
function saveDocuments(payload) {
  const lock = LockService.getScriptLock();

  try {
    if (!payload || !cleanText(payload.applicationNo)) {
      return {
        success: false,
        error: "Application No is required."
      };
    }

    if (!Array.isArray(payload.documents)) {
      return {
        success: false,
        error: "Document list is missing."
      };
    }

    lock.waitLock(20000);

    const sh = getSheetOrThrow(SHEETS.DOCUMENTS);
    const now = new Date();

    payload.documents.forEach(function(doc) {
      if (!doc || !cleanText(doc.name)) return;

      const appNo = cleanText(payload.applicationNo);
      const docName = cleanText(doc.name);

      // Re-read data so each document finds the latest row.
      const data = sh.getDataRange().getValues();
      let rowIndex = -1;
      let existing = null;

      for (let i = 1; i < data.length; i++) {
        if (
          cleanText(data[i][0]) === appNo &&
          cleanText(data[i][2]) === docName
        ) {
          rowIndex = i + 1;
          existing = data[i];
          break;
        }
      }

      const existingOriginal = existing ? toBool(existing[3]) : false;
      const existingXerox = existing ? toBool(existing[4]) : false;

      const wantsOriginal = toBool(doc.original);
      const wantsXerox = toBool(doc.xerox);

      const newOriginalSubmitted = wantsOriginal && !existingOriginal;
      const newXeroxSubmitted = wantsXerox && !existingXerox;
      const isNewSubmission = newOriginalSubmitted || newXeroxSubmitted;

      const finalOriginal = existingOriginal || wantsOriginal;
      const finalXerox = existingXerox || wantsXerox;

      // Keep already received information unchanged if no new flag was added.
      if (existing && !isNewSubmission) {
        return;
      }

      let submittedDate = existing ? existing[6] : "";
      let submittedBy = existing ? existing[7] : "";
      let receivedBy = existing ? existing[8] : "";

      if (isNewSubmission) {
        submittedDate = now;
        submittedBy = cleanText(doc.submittedBy) || submittedBy || "";

        if (!cleanText(receivedBy)) {
          receivedBy = cleanText(payload.operatorName);
        }
      }

      const status = (finalOriginal || finalXerox)
        ? "Submitted"
        : "Pending";

      const row = [
        appNo,                                  // A Application No
        cleanText(payload.studentName),         // B Student Name
        docName,                                // C Document Name
        finalOriginal,                          // D Original
        finalXerox,                             // E Xerox
        status,                                 // F Status
        submittedDate,                          // G Submitted Date
        submittedBy,                            // H Submitted By
        receivedBy,                             // I Received By
        now,                                    // J Last Updated
        cleanText(payload.operatorId)           // K Updated By
      ];

      if (rowIndex > 0) {
        sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
      }
    });

    if (payload.email) {
      try {
        sendAcknowledgementEmail(payload);
      } catch (emailError) {
        console.error("Acknowledgement email failed: " + emailError.message);
        return {
          success: true,
          emailSent: false,
          message: "Documents saved, but the acknowledgement email could not be sent."
        };
      }
    }

    return { success: true };
  } catch (error) {
    return {
      success: false,
      error: error.message || "Unable to save documents."
    };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/***********************
 * ADMIN-ONLY DOCUMENT PENDING REPORT
 *
 * Called by index.html:
 * getDocumentPendingReportForAdmin(operatorId, password)
 *
 * Report includes each student and each required document
 * that has no received/submitted Original or Xerox flag.
 *
 * A4 PDF is generated in index.html using this returned data.
 * No sheet data is modified by this function.
 ***********************/
function getDocumentPendingReportForAdmin(operatorId, password) {
  try {
    const userId = cleanText(operatorId);
    const userPassword = cleanText(password);

    if (!userId || !userPassword) {
      return {
        success: false,
        message: "Administrator login is required."
      };
    }

    const operatorsSheet = getSheetOrThrow(SHEETS.OPERATORS);
    const studentsSheet = getSheetOrThrow(SHEETS.STUDENTS);
    const documentsSheet = getSheetOrThrow(SHEETS.DOCUMENTS);

    // Verify administrator credentials directly against the sheet.
    // Never rely only on a role value supplied by the browser.
    const operatorData = operatorsSheet.getDataRange().getDisplayValues();
    let adminFound = false;

    for (let i = 1; i < operatorData.length; i++) {
      const rowId = cleanText(operatorData[i][0]);
      const rowPassword = cleanText(operatorData[i][1]);
      const rowRole = cleanText(operatorData[i][3]).toLowerCase();

      if (
        rowId === userId &&
        rowPassword === userPassword &&
        rowRole === "admin"
      ) {
        adminFound = true;
        break;
      }
    }

    if (!adminFound) {
      return {
        success: false,
        message: "Access denied. Only an administrator can generate this report."
      };
    }

    const studentData = studentsSheet.getDataRange().getDisplayValues();
    const documentData = documentsSheet.getDataRange().getValues();

    // Map only documents that have at least one received flag.
    const receivedMap = {};

    for (let i = 1; i < documentData.length; i++) {
      const row = documentData[i];
      const appNo = cleanText(row[0]);
      const docName = normalizeKey(row[2]);

      if (!appNo || !docName) continue;

      const received =
        toBool(row[3]) ||
        toBool(row[4]);

      if (received) {
        receivedMap[appNo + "||" + docName] = true;
      }
    }

    const pendingRows = [];
    const studentsSeen = {};
    let validStudentCount = 0;

    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];

      const registerNo = cleanText(row[0]);
      const applicationNo = cleanText(row[1]);
      const studentName = cleanText(row[2]);
      const section = cleanText(row[3]);
      const subSection = cleanText(row[4]);
      const email = cleanText(row[5]);

      if (!registerNo && !applicationNo && !studentName) continue;

      validStudentCount++;

      const studentKey = applicationNo ||
        registerNo ||
        ("ROW-" + (i + 1));

      if (studentsSeen[studentKey]) continue;
      studentsSeen[studentKey] = true;

      REQUIRED_DOCUMENTS.forEach(function(documentName) {
        const key = applicationNo + "||" + normalizeKey(documentName);

        if (!receivedMap[key]) {
          pendingRows.push({
            registerNo: registerNo,
            applicationNo: applicationNo,
            studentName: studentName,
            section: section,
            subSection: subSection,
            email: email,
            documentName: documentName
          });
        }
      });
    }

    return {
      success: true,
      generatedAt: Utilities.formatDate(
        new Date(),
        TZ,
        "dd-MM-yyyy HH:mm"
      ),
      totalStudents: validStudentCount,
      totalPending: pendingRows.length,
      rows: pendingRows
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || "Unable to generate pending report."
    };
  }
}

/***********************
 * EMAIL ACKNOWLEDGEMENT
 ***********************/
function sendAcknowledgementEmail(payload) {
  if (!payload || !cleanText(payload.email)) return;

  const subject =
    "Admission Document Acknowledgement - " +
    cleanText(payload.applicationNo);

  const escapeHtml = function(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  };

  let table = `
    <table border="1" cellpadding="6" cellspacing="0"
      style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:12px;">
      <tr style="background:#e8eef8;">
        <th>SL.NO</th>
        <th>Document</th>
        <th>Original</th>
        <th>Xerox</th>
        <th>Submitted By</th>
        <th>Date</th>
        <th>Received By</th>
      </tr>`;

  let sl = 1;
  const docs = Array.isArray(payload.documents) ? payload.documents : [];

  docs.forEach(function(doc) {
    if (toBool(doc.original) || toBool(doc.xerox)) {
      const dateText = doc.submittedDate ||
        formatDate(new Date());

      table += `
        <tr>
          <td>${sl++}</td>
          <td>${escapeHtml(doc.name)}</td>
          <td style="text-align:center">${toBool(doc.original) ? "Yes" : ""}</td>
          <td style="text-align:center">${toBool(doc.xerox) ? "Yes" : ""}</td>
          <td>${escapeHtml(doc.submittedBy || "")}</td>
          <td>${escapeHtml(dateText)}</td>
          <td>${escapeHtml(payload.operatorName || "")}</td>
        </tr>`;
    }
  });

  table += "</table>";

  const body = `
    <div style="font-family:Arial,sans-serif;color:#26344d;line-height:1.6;">
      <p>Dear Student / Parent,</p>

      <p>
        This is to acknowledge the receipt of the documents submitted by the student.
        This acknowledgement is issued for college record purposes.
      </p>

      <p><strong>Student Details</strong></p>
      <p>
        Name: ${escapeHtml(payload.studentName)}<br>
        Application No: ${escapeHtml(payload.applicationNo)}<br>
        Section: ${escapeHtml(payload.section || "-")}<br>
        Sub-Section: ${escapeHtml(payload.subSection || "-")}
      </p>

      ${table}

      <p>Yours faithfully,<br>
      Office<br>
      <strong>REVA Independent PU College</strong><br>
      Rukmini Knowledge Park,<br>
      Kattigenahalli, Yelahanka,<br>
      Bangalore – 560064.<br><br>
      Phone: 080-46966966 (Ext: 184)<br>
      Email: <a href="mailto:ripu.k@reva.edu.in">ripu.k@reva.edu.in</a>
      </p>
    </div>`;

  MailApp.sendEmail({
    to: cleanText(payload.email),
    subject: subject,
    htmlBody: body
  });

  logEmail(payload);
}

/***********************
 * EMAIL LOG
 ***********************/
function logEmail(payload) {
  const ss = getSS();
  let sh = ss.getSheetByName(SHEETS.EMAIL_LOG);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.EMAIL_LOG);
    sh.appendRow([
      "Timestamp",
      "Application No",
      "Student Name",
      "Email",
      "Operator Name"
    ]);
  }

  sh.appendRow([
    new Date(),
    cleanText(payload.applicationNo),
    cleanText(payload.studentName),
    cleanText(payload.email),
    cleanText(payload.operatorName)
  ]);
}

/***********************
 * ADMIN-ONLY PENDING DOCUMENT EMAIL REMINDER
 ***********************/
function sendPendingDocumentReminderForAdmin(operatorId, password, applicationNo) {
  try {
    const userId = cleanText(operatorId);
    const userPassword = cleanText(password);
    const appNo = cleanText(applicationNo);
    if (!userId || !userPassword || !appNo) {
      return { success: false, message: "Admin login and Application No are required." };
    }

    const operatorData = getSheetOrThrow(SHEETS.OPERATORS).getDataRange().getDisplayValues();
    const isAdmin = operatorData.slice(1).some(function(row) {
      return cleanText(row[0]) === userId &&
        cleanText(row[1]) === userPassword &&
        cleanText(row[3]).toLowerCase() === "admin";
    });
    if (!isAdmin) return { success: false, message: "Access denied. Admin only." };

    const studentData = getSheetOrThrow(SHEETS.STUDENTS).getDataRange().getDisplayValues();
    let student = null;
    for (let i = 1; i < studentData.length; i++) {
      if (cleanText(studentData[i][1]) === appNo) {
        student = {
          registerNo: cleanText(studentData[i][0]),
          applicationNo: cleanText(studentData[i][1]),
          name: cleanText(studentData[i][2]),
          section: cleanText(studentData[i][3]),
          subSection: cleanText(studentData[i][4]),
          email: cleanText(studentData[i][5])
        };
        break;
      }
    }
    if (!student) return { success: false, message: "Student record not found." };
    if (!student.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(student.email)) {
      return { success: false, message: "No valid email address is recorded for this student." };
    }

    const docData = getSheetOrThrow(SHEETS.DOCUMENTS).getDataRange().getValues();
    const receivedMap = {};
    for (let i = 1; i < docData.length; i++) {
      if (cleanText(docData[i][0]) === appNo &&
          (toBool(docData[i][3]) || toBool(docData[i][4]))) {
        receivedMap[normalizeKey(docData[i][2])] = true;
      }
    }
    const pending = REQUIRED_DOCUMENTS.filter(function(name) {
      return !receivedMap[normalizeKey(name)];
    });
    if (!pending.length) return { success: false, message: "This student has no pending documents." };

    const esc = function(v) {
      return String(v == null ? "" : v).replace(/&/g,"&amp;").replace(/</g,"&lt;")
        .replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
    };
    const subject = "Reminder: Pending Admission Documents - " + student.applicationNo;
    const list = pending.map(function(name, i) {
      return "<li>" + esc(name) + "</li>";
    }).join("");
    const html = '<div style="font-family:Arial,sans-serif;line-height:1.6;color:#24344d">' +
      '<p>Dear ' + esc(student.name || "Student / Parent") + ',</p>' +
      '<p>This is a reminder that the following admission documents are still pending. Kindly submit them to the college office.</p>' +
      '<p><b>Register No:</b> ' + esc(student.registerNo) + '<br><b>Application No:</b> ' + esc(student.applicationNo) +
      '<br><b>Section:</b> ' + esc([student.section,student.subSection].filter(Boolean).join(" / ") || "-") + '</p>' +
      '<ol>' + list + '</ol><p>For assistance, please contact the college office.</p>' +
      '<p>Yours faithfully,<br>Office<br><b>REVA Independent PU College</b><br>' +
      'Phone: 080-46966966 (Ext: 184)<br>Email: ripu.k@reva.edu.in</p></div>';

    MailApp.sendEmail({ to: student.email, subject: subject, htmlBody: html });
    logEmail({
      applicationNo: student.applicationNo,
      studentName: student.name,
      email: student.email,
      operatorName: "Admin reminder: " + userId
    });
    return { success: true, message: "Reminder email sent to " + student.email, pendingCount: pending.length };
  } catch (error) {
    return { success: false, message: error.message || "Unable to send reminder email." };
  }
}
