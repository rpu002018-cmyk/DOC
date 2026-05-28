/***********************
 * CONFIGURATION
 ***********************/
const SHEET_ID = "1uP4cJ8eke-UHTDV1qiJ_Fm0SlwmJRr6CkkuMqpKAyRU";
const DATE_FORMAT = "dd-MM-yyyy";
const TZ = Session.getScriptTimeZone();

const SHEETS = {
  STUDENTS: "Students",
  DOCUMENTS: "Documents",
  OPERATORS: "Operators",
  EMAIL_LOG: "Email Log",
  ISSUED_DOCUMENTS: "Issued Documents"
};

const RECEIVED_DOCUMENTS = [
  "SSLC Marks Card",
  "Transfer Certificate",
  "Study Certificate",
  "Aadhar Card",
  "Caste Certificate",
  "Income Certificate",
  "Passport Size Photos"
];

const ISSUED_DOCUMENTS_MASTER = [
  "Bonafide Certificate",
  "ID Card",
  "Fee Receipt",
  "Migration Certificate"
];

/***********************
 * COMMON
 ***********************/
function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function toBool(v) {
  if (v === true) return true;
  if (v === false) return false;
  if (v == null) return false;

  const s = String(v).trim().toLowerCase();

  return (
    s === "true" ||
    s === "yes" ||
    s === "y" ||
    s === "1"
  );
}

/***********************
 * WEB APP ENTRY
 ***********************/
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle(
      "Admission Document Acknowledgement"
    );
}

/***********************
 * DOCUMENT MASTER
 ***********************/
function getReceivedDocumentsList() {
  return RECEIVED_DOCUMENTS;
}

function getIssuedDocumentsList() {
  return ISSUED_DOCUMENTS_MASTER;
}

/***********************
 * LOGIN
 ***********************/
function validateOperator(id,pwd){

 const sh=
 getSS()
 .getSheetByName(
 SHEETS.OPERATORS
 );

 if(!sh)
 throw new Error(
 "Operators sheet missing"
 );

 const data=
 sh.getDataRange()
 .getValues();

 const uid=
 String(id||"").trim();

 const upwd=
 String(pwd||"").trim();

 for(
 let i=1;
 i<data.length;
 i++
 ){

 const rowId=
 String(
 data[i][0]||""
 ).trim();

 const rowPwd=
 String(
 data[i][1]||""
 ).trim();

 if(
 rowId===uid &&
 rowPwd===upwd
 ){

 return{
 success:true,
 name:data[i][2]
 };

 }

 }

 return{
 success:false
 };

}

/***********************
 * STUDENTS
 ***********************/
function addStudent(student){

 const sh=
 getSS()
 .getSheetByName(
 SHEETS.STUDENTS
 );

 if(!sh)
 throw new Error(
 "Students sheet missing"
 );

 const data=
 sh.getDataRange()
 .getValues();

 for(
 let i=1;
 i<data.length;
 i++
 ){

 if(
 String(data[i][0])
 ===student.registerNo ||

 String(data[i][1])
 ===student.applicationNo
 ){

 return{
 success:false,
 error:
 "Duplicate student"
 };

 }

 }

 sh.appendRow([

 student.registerNo,

 student.applicationNo,

 student.name,

 student.section||"",

 student.subSection||"",

 student.email||""

 ]);

 return{
 success:true
 };

}

function searchStudent(value){

 const sh=
 getSS()
 .getSheetByName(
 SHEETS.STUDENTS
 );

 if(!sh)
 throw new Error(
 "Students sheet missing"
 );

 const data=
 sh.getDataRange()
 .getValues();

 const v=
 String(
 value||""
 ).trim();

 for(
 let i=1;
 i<data.length;
 i++
 ){

 if(

 String(
 data[i][0]
 ).trim()===v ||

 String(
 data[i][1]
 ).trim()===v

 ){

 return{

 registerNo:
 data[i][0],

 applicationNo:
 data[i][1],

 name:
 data[i][2],

 section:
 data[i][3],

 subSection:
 data[i][4],

 email:
 data[i][5]||""

 };

 }

 }

 return null;

}

/***********************
 * EXISTING DOCUMENTS
 ***********************/
function loadExistingDocuments(
 applicationNo
){

 const sh=
 getSS()
 .getSheetByName(
 SHEETS.DOCUMENTS
 );

 if(!sh)
 return [];

 const data=
 sh.getDataRange()
 .getValues();

 return data
 .filter((r,i)=>

 i>0 &&

 String(r[0])
 ===String(
 applicationNo
 )

 )

 .map(r=>({

 name:r[2],

 original:
 toBool(r[3]),

 xerox:
 toBool(r[4]),

 status:
 r[5] ||

 (
 toBool(r[3])
 ||
 toBool(r[4])
 ?
 "Submitted"
 :
 "Pending"
 ),

 submittedDate:

 r[6]
 ?
 Utilities
 .formatDate(
 new Date(r[6]),
 TZ,
 DATE_FORMAT
 )
 :"",

 submittedBy:
 r[7]||"",

 receivedBy:
 r[8]||""

 }));

}

/***********************
 * SAVE RECEIVED DOCS
 ***********************/
function saveDocuments(payload){

 const lock=
 LockService
 .getScriptLock();

 lock.waitLock(
 30000
 );

 try{

 const sh=
 getSS()
 .getSheetByName(
 SHEETS.DOCUMENTS
 );

 const data=
 sh.getDataRange()
 .getValues();

 const now=
 new Date();

 payload.documents
 .forEach(doc=>{

 let rowIndex=-1;
 let existing=null;

 for(
 let i=1;
 i<data.length;
 i++
 ){

 if(

 String(
 data[i][0]
 ).trim()
 ===
 String(
 payload.applicationNo
 ).trim()

 &&

 String(
 data[i][2]
 ).trim()
 ===
 String(
 doc.name
 ).trim()

 ){

 rowIndex=i+1;

 existing=data[i];

 break;

 }

 }

 const existingOriginal=
 existing
 ?toBool(
 existing[3]
 )
 :false;

 const existingXerox=
 existing
 ?toBool(
 existing[4]
 )
 :false;

 const finalOriginal=
 existingOriginal
 ||
 doc.original;

 const finalXerox=
 existingXerox
 ||
 doc.xerox;

 const row=[

 payload.applicationNo,

 payload.studentName,

 doc.name,

 finalOriginal,

 finalXerox,

 (
 finalOriginal||
 finalXerox
 )
 ?
 "Submitted"
 :
 "Pending",

 now,

 doc.submittedBy||"",

 payload.operatorName,

 now,

 payload.operatorId

 ];

 if(
 rowIndex>0
 ){

 sh
 .getRange(
 rowIndex,
 1,
 1,
 row.length
 )
 .setValues([row]);

 }

 else{

 sh.appendRow(
 row
 );

 }

 });

 if(
 payload.email
 ){

 sendAcknowledgementEmail(
 payload
 );

 }

 return{
 success:true
 };

 }

 catch(err){

 return{

 success:false,

 error:
 err.message

 };

 }

 finally{

 lock.releaseLock();

 }

}

/***********************
 * RECEIVED DOCUMENTS EMAIL
 ***********************/
function sendAcknowledgementEmail(payload) {
  if (!payload.email || String(payload.email).trim() === "") {
    return { success: true };
  }

  try {
    const studentName = payload.studentName || "Student";
    const applicationNo = payload.applicationNo || "N/A";
    const submissionDate = Utilities.formatDate(new Date(), TZ, DATE_FORMAT);
    
    let documentsList = "";

    payload.documents.forEach(d => {
      const docTypes = [];
      if (d.original) docTypes.push("Original");
      if (d.xerox) docTypes.push("Xerox");
      documentsList += `
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${d.name}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${docTypes.join(", ") || "-"}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${d.submittedBy || "-"}</td>
        </tr>
      `;
    });

    const htmlBody = `
      <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; color: #333; }
            .container { max-width: 600px; margin: 0 auto; padding: 20px; }
            .header { background-color: #2196F3; color: white; padding: 15px; border-radius: 5px; }
            .content { padding: 20px 0; }
            .table-container { margin: 20px 0; }
            table { width: 100%; border-collapse: collapse; }
            th { background-color: #f2f2f2; padding: 10px; text-align: left; border-bottom: 2px solid #ddd; }
            .footer { color: #666; font-size: 12px; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
            .info-box { background-color: #f9f9f9; padding: 10px; border-left: 4px solid #2196F3; margin: 10px 0; }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="header">
              <h2>Document Receipt Acknowledgement</h2>
            </div>
            
            <div class="content">
              <p>Dear <strong>${studentName}</strong>,</p>
              
              <p>This is to acknowledge receipt of the following documents submitted for your admission.</p>
              
              <div class="info-box">
                <p><strong>Application No:</strong> ${applicationNo}</p>
                <p><strong>Submission Date:</strong> ${submissionDate}</p>
                <p><strong>Received By:</strong> ${payload.operatorName || "-"}</p>
              </div>
              
              <div class="table-container">
                <h3>Submitted Documents:</h3>
                <table>
                  <thead>
                    <tr>
                      <th>Document Name</th>
                      <th>Type</th>
                      <th>Submitted By</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${documentsList}
                  </tbody>
                </table>
              </div>
              
              <p>If you have any questions or need to submit additional documents, please contact the admission office immediately.</p>
              
              <p>Thank you!</p>
            </div>
            
            <div class="footer">
              <p>This is an automated email. Please do not reply to this message.</p>
              <p>For assistance, contact the admission office.</p>
            </div>
          </div>
        </body>
      </html>
    `;

    MailApp.sendEmail({
      to: payload.email,
      subject: `Document Receipt Acknowledgement - ${applicationNo}`,
      htmlBody: htmlBody
    });

    logEmail(payload);
    
    return { success: true };

  } catch (err) {
    Logger.log("Error in sendAcknowledgementEmail: " + err.message);
    return {
      success: false,
      error: err.message
    };
  }
}

/***********************
 * ISSUED DOCUMENTS - ENHANCED
 ***********************/

/**
 * Validates issued document payload
 * @param {Object} payload - Document payload
 * @returns {Object} - {valid: boolean, error: string}
 */
function validateIssuedDocumentsPayload(payload) {
  if (!payload.applicationNo || String(payload.applicationNo).trim() === "") {
    return { valid: false, error: "Application number is required" };
  }
  
  if (!payload.studentName || String(payload.studentName).trim() === "") {
    return { valid: false, error: "Student name is required" };
  }
  
  if (!Array.isArray(payload.documents) || payload.documents.length === 0) {
    return { valid: false, error: "At least one document must be selected" };
  }
  
  // Validate each document
  for (let doc of payload.documents) {
    if (!doc.name || String(doc.name).trim() === "") {
      return { valid: false, error: "Document name cannot be empty" };
    }
    
    if (!doc.quantity || doc.quantity < 1) {
      return { valid: false, error: `Invalid quantity for ${doc.name}` };
    }
    
    if (!doc.receivedBy || String(doc.receivedBy).trim() === "") {
      return { valid: false, error: `Received by information required for ${doc.name}` };
    }
  }
  
  if (!payload.operatorName || String(payload.operatorName).trim() === "") {
    return { valid: false, error: "Operator name is required" };
  }
  
  if (!payload.operatorId || String(payload.operatorId).trim() === "") {
    return { valid: false, error: "Operator ID is required" };
  }
  
  return { valid: true };
}

/**
 * Enhanced saveIssuedDocuments with validation and error handling
 * @param {Object} payload - Contains applicationNo, studentName, documents, operatorName, operatorId, email
 * @returns {Object} - {success: boolean, error?: string, documentCount?: number}
 */
function saveIssuedDocuments(payload) {
  // Validate payload
  const validation = validateIssuedDocumentsPayload(payload);
  if (!validation.valid) {
    return {
      success: false,
      error: validation.error
    };
  }

  try {
    let documentCount = 0;

    payload.documents.forEach(doc => {
      const result = addIssuedDocument({
        applicationNo: payload.applicationNo,
        studentName: payload.studentName,
        documentName: doc.name,
        quantity: doc.quantity,
        issuedDate: new Date(),
        issuedBy: payload.operatorName,
        purpose: doc.receivedBy,
        notes: doc.remarks || "",
        issuedTo: doc.issuedTo || ""
      });
      
      if (result.success) {
        documentCount++;
      }
    });

    if (payload.email) {
      const emailResult = sendIssuedEmail(payload);
      if (!emailResult.success) {
        Logger.log("Email sending failed: " + emailResult.error);
      }
    }

    return {
      success: true,
      documentCount: documentCount
    };

  } catch (err) {
    Logger.log("Error in saveIssuedDocuments: " + err.message);
    return {
      success: false,
      error: err.message
    };
  }
}

/**
 * Enhanced addIssuedDocument with new fields and validation
 * @param {Object} issuedDoc - Document details with new fields
 * @returns {Object} - {success: boolean, id?: string, error?: string}
 */
function addIssuedDocument(issuedDoc) {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const sh = getSS().getSheetByName(SHEETS.ISSUED_DOCUMENTS);

    if (!sh) {
      throw new Error("Issued Documents sheet missing");
    }

    // Validation
    if (!issuedDoc.applicationNo || !issuedDoc.documentName) {
      throw new Error("Application number and document name are required");
    }

    const id = Utilities.getUuid();
    const now = new Date();
    const operatorEmail = Session.getEffectiveUser().getEmail();

    // Enhanced row with new fields
    const row = [
      id,                                    // Column 0: ID
      issuedDoc.applicationNo,               // Column 1: Application No
      issuedDoc.studentName,                 // Column 2: Student Name
      issuedDoc.documentName,                // Column 3: Document Name
      issuedDoc.quantity || 1,               // Column 4: Quantity
      issuedDoc.issuedDate || now,           // Column 5: Issued Date
      issuedDoc.issuedBy || "",              // Column 6: Issued By
      issuedDoc.purpose || "",               // Column 7: Purpose/Received By
      issuedDoc.notes || "",                 // Column 8: Notes/Remarks
      issuedDoc.issuedTo || "",              // Column 9: Issued To (NEW)
      now,                                   // Column 10: Created Date
      operatorEmail,                         // Column 11: Created By
      "",                                    // Column 12: Last Updated (for future use)
      ""                                     // Column 13: Status (for future use)
    ];

    sh.appendRow(row);

    return {
      success: true,
      id: id
    };

  } catch (err) {
    Logger.log("Error in addIssuedDocument: " + err.message);
    return {
      success: false,
      error: err.message
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Enhanced loadIssuedDocuments with better filtering and formatting
 * @param {string} applicationNo - Application number
 * @returns {Array} - Array of issued documents
 */
function loadIssuedDocuments(applicationNo) {
  try {
    const sh = getSS().getSheetByName(SHEETS.ISSUED_DOCUMENTS);

    if (!sh) {
      return [];
    }

    const data = sh.getDataRange().getValues();

    return data
      .filter((r, i) => i > 0 && String(r[1]).trim() === String(applicationNo).trim())
      .map(r => ({
        id: r[0],
        applicationNo: r[1],
        studentName: r[2],
        documentName: r[3],
        quantity: r[4],
        issuedDate: r[5] 
          ? Utilities.formatDate(new Date(r[5]), TZ, DATE_FORMAT)
          : "",
        issuedBy: r[6],
        purpose: r[7],
        notes: r[8],
        issuedTo: r[9] || "",  // NEW field
        createdDate: r[10]
          ? Utilities.formatDate(new Date(r[10]), TZ, DATE_FORMAT)
          : "",
        createdBy: r[11]
      }))
      .sort((a, b) => new Date(b.issuedDate) - new Date(a.issuedDate)); // Sort by date, newest first
      
  } catch (err) {
    Logger.log("Error in loadIssuedDocuments: " + err.message);
    return [];
  }
}

/**
 * Get count of issued documents for a student
 * @param {string} applicationNo - Application number
 * @returns {number} - Total quantity of documents issued
 */
function getIssuedDocumentsCount(applicationNo) {
  const documents = loadIssuedDocuments(applicationNo);
  return documents.reduce((total, doc) => total + (doc.quantity || 0), 0);
}

/**
 * Enhanced email sending with better formatting and HTML template
 * @param {Object} payload - Document payload with email info
 * @returns {Object} - {success: boolean, error?: string}
 */
function sendIssuedEmail(payload) {
  if (!payload.email || String(payload.email).trim() === "") {
    return { success: true }; // Email not required
  }

  try {
    const studentName = payload.studentName || "Student";
    const applicationNo = payload.applicationNo || "N/A";
    const issuedDate = Utilities.formatDate(new Date(), TZ, DATE_FORMAT);
    
    let documentsList = "";
    let totalQuantity = 0;

    payload.documents.forEach(d => {
      const qty = d.quantity || 1;
      totalQuantity += qty;
      documentsList += `
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${d.name}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">${qty}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${d.receivedBy || "-"}</td>
        </tr>
      `;
    });

    const htmlBody = `
      <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; color: #333; }
            .container { max-width: 600px; margin: 0 auto; padding: 20px; }
            .header { background-color: #4CAF50; color: white; padding: 15px; border-radius: 5px; }
            .content { padding: 20px 0; }
            .table-container { margin: 20px 0; }
            table { width: 100%; border-collapse: collapse; }
            th { background-color: #f2f2f2; padding: 10px; text-align: left; border-bottom: 2px solid #ddd; }
            .footer { color: #666; font-size: 12px; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
            .info-box { background-color: #f9f9f9; padding: 10px; border-left: 4px solid #4CAF50; margin: 10px 0; }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="header">
              <h2>Document Issued Acknowledgement</h2>
            </div>
            
            <div class="content">
              <p>Dear <strong>${studentName}</strong>,</p>
              
              <p>This is to confirm that the following documents have been issued to you.</p>
              
              <div class="info-box">
                <p><strong>Application No:</strong> ${applicationNo}</p>
                <p><strong>Issued Date:</strong> ${issuedDate}</p>
                <p><strong>Issued By:</strong> ${payload.operatorName || "-"}</p>
              </div>
              
              <div class="table-container">
                <h3>Issued Documents Details:</h3>
                <table>
                  <thead>
                    <tr>
                      <th>Document Name</th>
                      <th style="text-align: center;">Quantity</th>
                      <th>Purpose/Received By</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${documentsList}
                    <tr style="background-color: #f2f2f2; font-weight: bold;">
                      <td colspan="1" style="padding: 8px;">Total</td>
                      <td style="padding: 8px; text-align: center;">${totalQuantity}</td>
                      <td style="padding: 8px;"></td>
                    </tr>
                  </tbody>
                </table>
              </div>
              
              <p>If you have any questions or discrepancies, please contact the admission office immediately.</p>
              
              <p>Thank you!</p>
            </div>
            
            <div class="footer">
              <p>This is an automated email. Please do not reply to this message.</p>
              <p>For assistance, contact the admission office.</p>
            </div>
          </div>
        </body>
      </html>
    `;

    MailApp.sendEmail({
      to: payload.email,
      subject: `Document Issued Acknowledgement - ${applicationNo}`,
      htmlBody: htmlBody
    });

    logEmail(payload);
    
    return { success: true };

  } catch (err) {
    Logger.log("Error in sendIssuedEmail: " + err.message);
    return {
      success: false,
      error: err.message
    };
  }
}

/***********************
 * EMAIL LOG
 ***********************/
function logEmail(payload){

 const sh=
 getSS()
 .getSheetByName(
 SHEETS.EMAIL_LOG
 );

 if(!sh)
 return;

 sh.appendRow([

 new Date(),

 payload.applicationNo,

 payload.studentName,

 payload.email,

 payload.operatorName

 ]);

}
