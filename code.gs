function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const type = rowData[3]; // Identify if it's Leave or OD
  const uniqueToken = generateUniqueToken();
  const statusColumn = 27;
  const tokenColumn = 28;

 
  if (type === "Leave") {
    processLeaveRequest(sheet, lastRow, rowData, uniqueToken, statusColumn, tokenColumn);
  } else if (type === "OD") {
    processODRequest(sheet, lastRow, rowData, uniqueToken, statusColumn, tokenColumn);
  }
}

function processLeaveRequest(sheet, lastRow, rowData, uniqueToken, statusColumn, tokenColumn) {
  const employeeEmail = rowData[2];
  const hodEmail = rowData[15];
  const requestNumber = generateRequestNumber(lastRow);

  sheet.getRange(lastRow, 2).setValue(requestNumber);
  sheet.getRange(lastRow, statusColumn).setValue("Pending");
  sheet.getRange(lastRow, tokenColumn).setValue(uniqueToken);

  sendApprovalRequestToHOD("Leave", rowData, requestNumber, uniqueToken, hodEmail);
  notifyEmployeeLeave(employeeEmail, "Leave", requestNumber);
}

function processODRequest(sheet, lastRow, rowData, uniqueToken, statusColumn, tokenColumn) {
  const hodEmail = rowData[25];
  const employeeEmail = rowData[2];
  const requestNumber = generateRequestNumber(lastRow);

  sheet.getRange(lastRow, 2).setValue(requestNumber);
  sheet.getRange(lastRow, statusColumn).setValue("Pending");
  sheet.getRange(lastRow, tokenColumn).setValue(uniqueToken);

  sendApprovalRequestToHOD("OD", rowData,  requestNumber, uniqueToken, hodEmail);
  notifyEmployeeOD(employeeEmail, "OD", requestNumber);
}

// Function to generate unique token for security
function generateUniqueToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  for (let i = 0; i < 20; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// Function to generate unique request number
function generateRequestNumber(rowNumber) {
  const prefix = "REQ";
  const date = new Date();
  const year = date.getFullYear();
  const month = ("0" + (date.getMonth() + 1)).slice(-2);
  const day = ("0" + date.getDate()).slice(-2);
  const uniqueId = `${prefix}-${year}${month}${day}-${rowNumber}`;
  return uniqueId;
}

function sendApprovalRequestToHOD(type, rowData, requestNumber, uniqueToken, hodEmail) {
 const formatDate = (dateValue) => {
    if (!dateValue || isNaN(new Date(dateValue).getTime())) {
        return "Invalid Date"; // Handle invalid dates gracefully
    }
    return Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), "dd/MM/yyyy");
};
function formatTime(timeString) {
  if (!timeString) return ""; // Handle empty times
  return Utilities.formatDate(new Date(timeString), Session.getScriptTimeZone(), "hh:mm a"); // 12-hour format
}

  // Subject Handling
  const subject = (type === "Leave")
    ? `Approval Request: Leave - ${rowData[7]} (${rowData[6]}) -  ${requestNumber}`
    : `Approval Request: OD - ${rowData[16]} (${rowData[17]}) - ${requestNumber}`;

  // Employee Details Formatting
  const employeeDetails = (type === "Leave")
  ? `<strong>${rowData[7]}</strong> (<strong>${rowData[6]}</strong>) <br>from <strong>${rowData[8]}</strong>`
  : `<strong>${rowData[16]}</strong> (<strong>${rowData[17]}</strong>) <br>from <strong>${rowData[18]}
  </strong>`;



    // Approval & Rejection Links
  const approveUrl = `${ScriptApp.getService().getUrl()}?action=approve&requestNumber=${requestNumber}&token=${uniqueToken}`;
  const rejectUrl = `${ScriptApp.getService().getUrl()}?action=reject&requestNumber=${requestNumber}&token=${uniqueToken}`;
 

  let details = "<table border='1' cellpadding='10' cellspacing='0'width='100%' style='border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;'>";
  if (type === "Leave") {
    details += `<tr style="background-color: #f2f2f2;">
                 <th colspan="2" style="text-align: center; font-size: 18px; padding: 12px;">Leave Request Details:</th>
                </tr>
        <tr><td style="font-weight: bold; width: 30%;">Employee Name:</td><td>${rowData[7]}</td></tr>
        <tr><td style="font-weight: bold;">Employee ID:</td><td>${rowData[6]}</td></tr>
        <tr><td style="font-weight: bold;">Email:</td><td>${rowData[2]}</td></tr>
        <tr><td style="font-weight: bold;">Location:</td><td>${rowData[4]}</td></tr>
        <tr><td style="font-weight: bold;">Plant Name:</td><td>${rowData[5]}</td></tr>
        <tr><td style="font-weight: bold;">Department:</td><td>${rowData[8]}</td></tr>
        <tr><td style="font-weight: bold;">Nature of Leave:</td><td>${rowData[10]}</td></tr>
        <tr><td style="font-weight: bold;">Leave From:</td><td>${formatDate(rowData[11])}</td></tr>
        <tr><td style="font-weight: bold;">Leave Till:</td><td>${formatDate(rowData[12])}</td></tr>
        <tr><td style="font-weight: bold;">Reason:</td><td>${rowData[14]}</td></tr>`;
  } else {
    details += `<table border="1" cellpadding="10" cellspacing="0" width="100%" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
        <tr style="background-color: #f2f2f2;">
            <th colspan="2" style="text-align: center; font-size: 18px; padding: 12px;">On-Duty (OD) Request Details:</th>
        </tr>
        <tr><td style="font-weight: bold; width: 30%;">Employee Name:</td><td>${rowData[16]}</td></tr>
        <tr><td style="font-weight: bold;">Employee ID:</td><td>${rowData[17]}</td></tr>
        <tr><td style="font-weight: bold;">Department:</td><td>${rowData[18]}</td></tr>
        <tr><td style="font-weight: bold;">Location:</td><td>${rowData[4]}</td></tr>
        <tr><td style="font-weight: bold;">Plant Name:</td><td>${rowData[5]}</td></tr>
        <tr><td style="font-weight: bold;">From Date:</td><td>${formatDate(rowData[19])}</td></tr>
        <tr><td style="font-weight: bold;">Out Time:</td><td>${formatTime(rowData[20])}</td></tr>
        <tr><td style="font-weight: bold;">To Date:</td><td>${formatDate(rowData[21])}</td></tr>
        <tr><td style="font-weight: bold;">Return Time:</td><td>${formatTime(rowData[22])}</td></tr>
        <tr><td style="font-weight: bold;">Places to Visit:</td><td>${rowData[23]}</td></tr>
        <tr><td style="font-weight: bold;">Purpose:</td><td>${rowData[24]}</td></tr>`;
  }
  details += "</table>";

  // Email Body
  const body = `<div style='font-family: Arial, sans-serif; line-height: 1.6; color: #333;'>
   <h2 style="color: #4CAF50;">Approval Request for ${type}</h2>
    <h2>Dear HOD,<h2>
    <p>${employeeDetails} has requested a ${type}</strong> with the following details:</p>
    ${details}
    <p>
      <a href='${approveUrl}' style='background-color: green; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;'>Approve</a>
      &nbsp;
      <a href='${rejectUrl}' style='background-color: red; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;'>Reject</a>
    </p>
  </div>`;

  // email for hod
  GmailApp.sendEmail(hodEmail, subject, "", {
    htmlBody: body,
    name: 'Approval System'
  });
}



function notifyEmployeeLeave(employeeEmail, type, requestNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const row = data.find(row => row[1] === requestNumber);

  const formatDate = (dateValue) => {
    if (!dateValue || isNaN(new Date(dateValue).getTime())) return "Invalid Date";
    return Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), "dd/MM/yyyy");
  };

  const details = `
    <table border='1' cellpadding='10' cellspacing='0' width='100%' style='border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;'>
      <tr style="background-color: #f2f2f2;">
        <th colspan="2" style="text-align: center; font-size: 18px; padding: 12px;">Leave Request Details:</th>
      </tr>
      <tr><td style="font-weight: bold;">Employee Name:</td><td>${row[7]}</td></tr>
      <tr><td style="font-weight: bold;">Employee ID:</td><td>${row[6]}</td></tr>
      <tr><td style="font-weight: bold;">Email:</td><td>${row[2]}</td></tr>
      <tr><td style="font-weight: bold;">Location:</td><td>${row[4]}</td></tr>
      <tr><td style="font-weight: bold;">Plant Name:</td><td>${row[5]}</td></tr>
      <tr><td style="font-weight: bold;">Department:</td><td>${row[8]}</td></tr>
      <tr><td style="font-weight: bold;">Nature of Leave:</td><td>${row[10]}</td></tr>
      <tr><td style="font-weight: bold;">Leave From:</td><td>${formatDate(row[11])}</td></tr>
      <tr><td style="font-weight: bold;">Leave Till:</td><td>${formatDate(row[12])}</td></tr>
      <tr><td style="font-weight: bold;">Reason:</td><td>${row[14]}</td></tr>
    </table>`;

  const subject = `Your Leave Request is Under Review - ${requestNumber}`;
  const body = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <h2 style="color: #4CAF50;">Request Submitted Successfully</h2>
      <p>Dear Employee,</p>
      <p>Your request for <strong>${type}</strong> has been submitted and is under review. Below are the details:</p>
      ${details}
    </div>`;

  GmailApp.sendEmail(employeeEmail, subject, "", { htmlBody: body });
}



function notifyEmployeeOD(employeeEmail, type, requestNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const row = data.find(row => row[1] === requestNumber);

  const formatDate = (dateValue) => {
    if (!dateValue || isNaN(new Date(dateValue).getTime())) return "Invalid Date";
    return Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), "dd/MM/yyyy");
  };

  const formatTime = (timeValue) => {
    if (!timeValue) return "";
    return Utilities.formatDate(new Date(timeValue), Session.getScriptTimeZone(), "hh:mm a");
  };

  const details = `
    <table border='1' cellpadding='10' cellspacing='0' width='100%' style='border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;'>
      <tr style="background-color: #f2f2f2;">
        <th colspan="2" style="text-align: center; font-size: 18px; padding: 12px;">On-Duty (OD) Request Details:</th>
      </tr>
      <tr><td style="font-weight: bold;">Employee Name:</td><td>${row[16]}</td></tr>
      <tr><td style="font-weight: bold;">Employee ID:</td><td>${row[17]}</td></tr>
      <tr><td style="font-weight: bold;">Department:</td><td>${row[18]}</td></tr>
      <tr><td style="font-weight: bold;">Location:</td><td>${row[4]}</td></tr>
      <tr><td style="font-weight: bold;">Plant Name:</td><td>${row[5]}</td></tr>
      <tr><td style="font-weight: bold;">From Date:</td><td>${formatDate(row[19])}</td></tr>
      <tr><td style="font-weight: bold;">Out Time:</td><td>${formatTime(row[20])}</td></tr>
      <tr><td style="font-weight: bold;">To Date:</td><td>${formatDate(row[21])}</td></tr>
      <tr><td style="font-weight: bold;">Return Time:</td><td>${formatTime(row[22])}</td></tr>
      <tr><td style="font-weight: bold;">Places to Visit:</td><td>${row[23]}</td></tr>
      <tr><td style="font-weight: bold;">Purpose:</td><td>${row[24]}</td></tr>
    </table>`;

  const subject = `Your OD Request is Under Review -  ${requestNumber}`;
  const body = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <h2 style="color: #4CAF50;">Request Submitted Successfully</h2>
      <p>Dear Employee,</p>
      <p>Your request for <strong>${type}</strong> has been submitted and is under review. Below are the details:</p>
      ${details}
    </div>`;

  GmailApp.sendEmail(employeeEmail, subject, "", { htmlBody: body });
}




function doGet(e) {
  const action = e.parameter.action;
  const requestNumber = e.parameter.requestNumber;
  const token = e.parameter.token;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let rowToUpdate;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === requestNumber && data[i][26] === "Pending" && data[i][27] === token) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate) {
    const employeeEmail = data[rowToUpdate - 1][2];
    const type = data[rowToUpdate - 1][3];

    // Update approval status
    const decisionText = action === 'approve' ? "Approved" : "Rejected";
    sheet.getRange(rowToUpdate, 27).setValue(decisionText);

    notifyEmployeeDecision(employeeEmail, type, requestNumber, decisionText);

    return ContentService.createTextOutput(`Request ${decisionText} Successfully.`);
  }

  return ContentService.createTextOutput("Invalid or Expired Request.");
}

function doGet(e) {
  const action = e.parameter.action;
  const requestNumber = e.parameter.requestNumber;
  const token = e.parameter.token;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let rowToUpdate;

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === requestNumber && data[i][26] === "Pending" && data[i][27] === token) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate) {
    const employeeEmail = data[rowToUpdate - 1][2];
    const type = data[rowToUpdate - 1][3];

    // Update approval status
    const decisionText = action === 'approve' ? "Approved" : "Rejected";
    sheet.getRange(rowToUpdate, 27).setValue(decisionText);

    notifyEmployeeDecision(employeeEmail, type, requestNumber, decisionText);

    return ContentService.createTextOutput(`Request ${decisionText} Successfully.`);
  }

  return ContentService.createTextOutput("Invalid or Expired Request.");
}

// Function to notify the employee about HOD's decision
function notifyEmployeeDecision(employeeEmail, type, requestNumber, decision) {
  decision = decision.toLowerCase();
  decision = decision.charAt(0).toUpperCase() + decision.slice(1); // Ensure capitalization

  // Subject without emojis
  const subject = `Your ${type} Request has been ${decision}`;

  const body = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <h2 style="color: ${decision === 'Approved' ? '#4CAF50' : '#f44336'};">
        ${type} Request ${decision}
      </h2>
      <p>Dear Employee,</p>
      <p>Your request for <strong>${type}</strong> (Request Number: <strong>${requestNumber || "N/A"}</strong>)
      has been <strong>${decision}</strong> by your HOD.</p>
      <hr style="border: none; border-top: 1px solid #ccc;">
      <p style="font-size: 0.9em; color: #777;">
        If you have any questions, please contact your HOD or the HR department.
      </p>
      <p>Regards,<br><strong>Approval System</strong></p>
    </div>
  `;

  // Send email with proper HTML formatting
  GmailApp.sendEmail(employeeEmail, subject, "", {
    htmlBody: body,
    name: 'Approval System'
  });
}
