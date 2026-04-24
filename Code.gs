/** * WEBSERVER SETUP */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('KSRCT Student Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** * HELPER: CONVERT SEMESTER NUMBER TO ROMAN */
function getRoman(sem) {
  const map = { "1": "I", "2": "II", "3": "III", "4": "IV", "5": "V", "6": "VI", "7": "VII", "8": "VIII" };
  return map[sem];
}


/** * LOGIN LOGIC WITH BIRTHDAY CHECK, LOGGING */
/** * LOGIN LOGIC WITH ACADEMIC DATA FETCHING */
function processLogin(regNo, dob, deviceInfo = "Unknown Device") {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Database"); 
    const masterSheet = ss.getSheetByName("Data Sheet"); // We need this for the summary
    
    // 1. Admin Bypass
    if (dob === "ADMIN_BYPASS") {
      const data = dbSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === regNo.toString()) {
          return { status: "success", name: data[i][2], reg: data[i][0] };
        }
      }
      return { status: "fail", message: "Student Registration Number not found." };
    }

    // 2. Maintenance Check
    if (dbSheet.getRange("F1").getValue() === "ON") {
      return { status: "fail", message: "Portal is under maintenance." };
    }

    const data = dbSheet.getRange(1, 1, dbSheet.getLastRow(), 3).getDisplayValues();
    let authSuccess = false;
    let studentData = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].trim() === regNo.trim() && data[i][2].trim() === dob.trim()) {
        
        // --- BIRTHDAY LOGIC ---
        const today = new Date();
        const tDay = today.getDate();
        const tMonth = today.getMonth() + 1;
        const dobValue = data[i][2].toString().trim();
        const dobParts = dobValue.split(/\D/); 
        const isBirthday = parseInt(dobParts[2]) === tDay && parseInt(dobParts[1]) === tMonth;
        
        // --- NEW: FETCH ACADEMIC SUMMARY DATA ---
        let sgpaData = { s1: "0.00", s2: "0.00", s3: "0.00", s4: "0.00", s5: "0.00", cgpa: "0.00" };
        const masterData = masterSheet.getDataRange().getValues();
        
        for (let m = 0; m < masterData.length; m++) {
          if (masterData[m][1].toString().trim() === regNo.trim()) {
            sgpaData.s1 = masterData[m][6] || "0.00"; // Column G
            sgpaData.s2 = masterData[m][7] || "0.00"; // Column H
            sgpaData.s3 = masterData[m][8] || "0.00"; // Column I
            sgpaData.s4 = masterData[m][9] || "0.00"; // Column J
            sgpaData.s5 = masterData[m][10] || "0.00"; // Column K
            sgpaData.cgpa = masterData[m][21] || "0.00"; // Column V
            break;
          }
        }

        authSuccess = true;
        studentData = { 
          status: "success", 
          name: data[i][1], 
          reg: regNo,
          isBirthday: isBirthday,
          // Sending the summary data back to Index.html
          sgpa1: sgpaData.s1,
          sgpa2: sgpaData.s2,
          sgpa3: sgpaData.s3,
          sgpa4: sgpaData.s4,
          sgpa5: sgpaData.s5,
          cgpa: sgpaData.cgpa
        };
        break;
      }
    }

    recordLoginLog(regNo, authSuccess ? "SUCCESS" : "FAILED", deviceInfo);
    return authSuccess ? studentData : { status: "fail", message: "Invalid Credentials." };

  } catch (e) { 
    return { status: "error", message: "Server Error: " + e.toString() }; 
  }
}

/** * FETCH DATA */
function getLiveGrades(semester, mode, regNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roman = getRoman(semester);
    const targetSheet = ss.getSheetByName(roman + "-" + mode);
    const masterSheet = ss.getSheetByName("Data Sheet");
    const dbSheet = ss.getSheetByName("Database");
    const notice = dbSheet.getRange("E1").getValue();

    if (!targetSheet) throw new Error("Sheet not found.");

    const nameData = targetSheet.getRange("D151:D170").getValues(); 
    let subjectNames = [];
    for (let i = 0; i < nameData.length; i++) {
      let val = nameData[i][0].toString().trim();
      if (val !== "") subjectNames.push(val); else break; 
    }

    const fullData = targetSheet.getDataRange().getValues();
    let studentGrades = [];
    for (let j = 3; j < fullData.length; j++) { 
      if (fullData[j][1].toString().trim() === regNo.trim()) {
        for (let k = 0; k < subjectNames.length; k++) {
          studentGrades.push(fullData[j][4 + (k * 2)] || "");
        }
        break;
      }
    }

    const masterData = masterSheet.getDataRange().getValues();
    let sgpa = "0.00", cgpa = "0.00";
    for (let m = 0; m < masterData.length; m++) {
       if (masterData[m][1].toString().trim() === regNo.trim()) {
         sgpa = masterData[m][5 + parseInt(semester)] || "0.00"; 
         cgpa = masterData[m][21] || "0.00"; 
         break;
       }
    }
    return { status: "success", subjects: subjectNames, grades: studentGrades, sgpa: sgpa, cgpa: cgpa, notice: notice };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/** * UPDATE GRADES */
/** * UPDATE GRADES & RETURN REFRESHED SUMMARY */
function submitGrades(semester, mode, gradesArray, regNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(getRoman(semester) + "-" + mode);
    const masterSheet = ss.getSheetByName("Data Sheet");
    
    // 1. Update the grades in the specific Semester Sheet
    const data = targetSheet.getDataRange().getValues();
    for (let i = 3; i < data.length; i++) {
      if (data[i][1].toString().trim() === regNo.trim()) {
        gradesArray.forEach((grade, idx) => {
          targetSheet.getRange(i + 1, 5 + (idx * 2)).setValue(grade.toUpperCase());
        });
        break;
      }
    }

    // 2. Spreadsheet needs a moment to recalculate formulas
    SpreadsheetApp.flush(); 

    // 3. Fetch the updated SGPA/CGPA from "Data Sheet" to refresh the UI
    let updatedSummary = { sgpa1: "0.00", sgpa2: "0.00", sgpa3: "0.00", sgpa4: "0.00", sgpa5: "0.00", cgpa: "0.00" };
    const masterData = masterSheet.getDataRange().getValues();
    
    for (let m = 0; m < masterData.length; m++) {
       if (masterData[m][1].toString().trim() === regNo.trim()) {
         updatedSummary.sgpa1 = masterData[m][6] || "0.00"; 
         updatedSummary.sgpa2 = masterData[m][7] || "0.00";
         updatedSummary.sgpa3 = masterData[m][8] || "0.00";
         updatedSummary.sgpa4 = masterData[m][9] || "0.00";
         updatedSummary.sgpa5 = masterData[m][10] || "0.00";
         updatedSummary.cgpa = masterData[m][21] || "0.00"; 
         break;
       }
    }

    return { 
      status: "success", 
      updatedData: updatedSummary 
    };
  } catch (e) { 
    return { status: "error", message: e.toString() }; 
  }
}

function getProfileData(regNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data Sheet");
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1].toString().trim() === regNo.trim()) {
        return { status: "success", phone: data[i][19] || "", email: data[i][20] || "" };
      }
    }
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function updateProfileData(regNo, phone, email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data Sheet");
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1].toString().trim() === regNo.trim()) {
        sheet.getRange(i + 1, 20).setValue(phone); 
        sheet.getRange(i + 1, 21).setValue(email); 
        return { status: "success" };
      }
    }
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/** * UPDATED FEEDBACK FUNCTION 
 * Now fetches the student's name from the Database sheet before saving.
 */
function submitFeedback(regNo, feedbackText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Database");
    const fbSheet = ss.getSheetByName("Feedback");
    
    if (!fbSheet) throw new Error("Feedback sheet not found.");

    // 1. Find the Student's Name in the Database sheet
    let studentName = "Unknown";
    const dbData = dbSheet.getRange(1, 1, dbSheet.getLastRow(), 2).getValues();
    
    for (let i = 1; i < dbData.length; i++) {
      if (dbData[i][0].toString().trim() === regNo.toString().trim()) {
        studentName = dbData[i][1]; // Get Name from Column B
        break;
      }
    }

    // 2. Append to Feedback sheet with the Name
    // Format: [Timestamp, Reg No, Name, Feedback Message]
    fbSheet.appendRow([new Date(), regNo, studentName, feedbackText]);
    
    return { status: "success" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/** --- ADMIN BACKEND --- **/
function getAdminData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Database");
    const fbSheet = ss.getSheetByName("Feedback");
    
    // Fetch the current status from your sheet
    const maintenance = dbSheet.getRange("F1").getValue(); 
    const freeze = dbSheet.getRange("G1").getValue();      
    
    const studentCount = dbSheet.getLastRow() - 1; 
    const notice = dbSheet.getRange("E1").getValue();
    
    let feedbacks = [];
    if (fbSheet && fbSheet.getLastRow() > 1) {
      feedbacks = fbSheet.getRange(2, 1, fbSheet.getLastRow() - 1, 4).getDisplayValues().reverse();
    }
    
    // Check line 201 area - make sure commas are present!
    return {
      status: "success",
      studentCount: studentCount,
      notice: notice,
      feedbacks: feedbacks.slice(0, 20),
      maintenance: maintenance, // <--- Ensure comma is here
      freeze: freeze             // <--- Ensure variable name is correct
    };
  } catch (e) { 
    return { status: "error", message: e.toString() }; 
  }
}

function updatePortalNotice(newText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Database");
    dbSheet.getRange("E1").setValue(newText);
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

/** GET SYSTEM SETTINGS & STUDENT DATA **/
function getAdminPortalStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName("Database");
  return {
    maintenance: db.getRange("F1").getValue() === "ON",
    freeze: db.getRange("G1").getValue() === "ON"
  };
}

/** TOGGLE SETTINGS **/
function toggleSetting(type, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName("Database");
  const cell = (type === 'maintenance') ? "F1" : "G1";
  db.getRange(cell).setValue(value ? "ON" : "OFF");
  return { status: "success" };
}

/** SEARCH STUDENT FOR DIRECT EDIT **/
function adminSearchStudent(regNo) {
  return processLogin(regNo, ""); // Reuses your login logic but skips DOB check for admin
}


/** ADMIN HELPER: FETCH STUDENT GRADES WITHOUT LOGIN **/
function getStudentGrades(regNo, sem) {
  try {
    // If sem is undefined, default to 5
    const selectedSem = sem ? sem.toString() : "5";
    const currentMode = "Current Status";
    
    // This calls your existing logic that pulls from the spreadsheet
    const res = getLiveGrades(selectedSem, currentMode, regNo);
    
    if (res && res.status === "success") {
      return res;
    } else {
      return { status: "error", message: "No records found for Semester " + selectedSem };
    }
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}


/** * INTERNAL: RECORD LOGIN TO LOGS SHEET */
function recordLoginLog(regNo, status, deviceInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Logs");
  
  if (!logSheet) {
    logSheet = ss.insertSheet("Logs");
    logSheet.appendRow(["Timestamp", "Reg No", "Status", "Device Type", "User Agent"]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  let deviceType = "Desktop";
  if (deviceInfo.indexOf("Mobi") > -1) deviceType = "Mobile";
  if (deviceInfo.indexOf("Tablet") > -1) deviceType = "Tablet";

  logSheet.appendRow([new Date(), regNo, status, deviceType, deviceInfo]);
}

/** * ADMIN: FETCH RAW DATA FOR THE LOGS */
function getLogsFromServer() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    // Get last 50 logs, columns A to D
    const startRow = Math.max(2, lastRow - 49);
    const numRows = lastRow - startRow + 1;
    
    return sheet.getRange(startRow, 1, numRows, 4).getDisplayValues();
  } catch (e) {
    return [];
  }
}

function ping() {
  return true; // Just a handshake to prove the script is alive
}

/** * AUTOMATIC BIRTHDAY MAILER - RUNS DAILY */
function sendBirthdayWishes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Database");
    if (!dbSheet) return;

    // Get today's Day and Month
    const today = new Date();
    const tDay = today.getDate();
    const tMonth = today.getMonth() + 1;

    // Get all data from Database (RegNo, Name, DOB, Email)
    const data = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 4).getDisplayValues();
    let sentCount = 0;

    data.forEach(row => {
      const regNo = row[0];
      const name = row[1];
      const dobValue = row[2]; // e.g. "2026-02-11"
      const email = row[3];    // Column D

      if (dobValue && email && email.includes("@")) {
        const dobParts = dobValue.split(/\D/); 
        
        // Match Day (parts[2]) and Month (parts[1]) for YYYY-MM-DD format
        if (parseInt(dobParts[2]) === tDay && parseInt(dobParts[1]) === tMonth) {
          
          const cakeImg = "https://fonts.gstatic.com/s/e/notoemoji/latest/1f382/512.gif";
          const subject = "Happy Birthday, " + name + "!"; 
          const htmlMessage = `
            <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; max-width: 600px; border: 1px solid #eee; padding: 20px; border-radius: 10px;">
              <h2 style="color: #F37021; text-align: center;">Happy Birthday, ${name}!</h2>
              <div style="text-align: center; margin: 20px 0;">
                <img src="${cakeImg}" alt="Birthday Cake" width="100" height="100">
              </div>
              <p>Dear <strong>${name}</strong> (${regNo}),</p>
              <p>Wishing you a very <b>Happy Birthday</b> from the KSRCT Student Portal team! May your year ahead be filled with success and happiness.</p>
              <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
              <p style="font-size: 0.9em; color: #777;">
                Best Regards,<br>
                <strong>Admin Team</strong><br>
                KSRCT Student Portal
              </p>
            </div>
          `;

          GmailApp.sendEmail(email, subject, "", { htmlBody: htmlMessage });
          sentCount++; // FIX 1: Added this to track success
        }
      }
    }); // FIX 2: Added closing bracket for forEach

    console.log("Birthday emails sent today: " + sentCount);
  } catch (e) {
    console.log("Mailer Error: " + e.toString());
  }
} // FIX 2: Added closing bracket for try/catch
