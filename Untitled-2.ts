var SS_ID = "1_jzaRlHMsYNYZOcrttDRGC0meQuJQH7TyDzhQKMcIfc";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Acquisition CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkLogin(username, password) {
  const users = { "admin": "1234", "negin": "snapp123", "ezt": "123" };
  const u = username ? username.toLowerCase().trim() : "";
  const p = password ? password.toString().trim() : "";
  if (users[u] && users[u] === p) {
    logActivity(u, 'ورود به سیستم', '');
    return { success: true, name: u, isAdmin: u === "admin" };
  }
  return { success: false };
}

function getNextInboxId() {
  const props = PropertiesService.getScriptProperties();
  let lastId = props.getProperty('LAST_ID_NUMBER') || 0;
  let nextNumber = parseInt(lastId) + 1;
  props.setProperty('LAST_ID_NUMBER', nextNumber.toString());
  return "INB-" + nextNumber.toString().padStart(6, '0');
}

function getAllLeadsForAdmin() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("Leads_Inbox");
  if (!sheet) return { headers: [], rows: [] };
  const data = sheet.getDataRange().getDisplayValues();
  return { headers: data[0], rows: data.slice(1) };
}

function updateLeadStatus(rowIndex, status, priority, owner) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sourceSheet = ss.getSheetByName("Leads_Inbox");
  const targetSheet = ss.getSheetByName("Negotiation Stage");

  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const actualRow = parseInt(rowIndex) + 2;

  // به‌روزرسانی فقط ستون‌هایی که وجود دارن
  const statusColIdx = sourceHeaders.indexOf("Manager_Status");
  if (statusColIdx !== -1) {
    sourceSheet.getRange(actualRow, statusColIdx + 1).setValue(status);
  }
  
  const priorityColIdx = sourceHeaders.indexOf("Manager_Priority");
  if (priorityColIdx !== -1) {
    sourceSheet.getRange(actualRow, priorityColIdx + 1).setValue(priority);
  }
  
  const ownerColIdx = sourceHeaders.indexOf("Assigned_Owner");
  if (ownerColIdx !== -1) {
    sourceSheet.getRange(actualRow, ownerColIdx + 1).setValue(owner);
  }

  // خواندن INBOX_ID
  const inboxIdColIdx = sourceHeaders.indexOf("INBOX_ID");
  let inboxId = "";
  if (inboxIdColIdx !== -1) {
    inboxId = sourceSheet.getRange(actualRow, inboxIdColIdx + 1).getValue().toString();
  }

  if (status === "Approved") {
    const pmerColIdx = sourceHeaders.indexOf("Linked_PMER_ID");
    if (pmerColIdx !== -1) {
      if (!sourceSheet.getRange(actualRow, pmerColIdx + 1).getValue()) {
        const createdDateColIdx = sourceHeaders.indexOf("Created_Date");
        if (createdDateColIdx !== -1) {
          sourceSheet.getRange(actualRow, createdDateColIdx + 1).setValue(new Date());
        }
        
        const props = PropertiesService.getScriptProperties();
        let lastPmer = parseInt(props.getProperty('LAST_PMER_NUM')) || 4547;
        let nextPmer = lastPmer + 1;
        const pmerId = "PMER-" + nextPmer.toString().padStart(6, '0');
        sourceSheet.getRange(actualRow, pmerColIdx + 1).setValue(pmerId);
        props.setProperty('LAST_PMER_NUM', nextPmer.toString());
      }
    }

    if (targetSheet) {
      const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
      const sourceValues = sourceSheet.getRange(actualRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

      // انتقال فقط ستون‌هایی که در target sheet وجود دارن
      const alignedData = [];
      for (let i = 0; i < targetHeaders.length; i++) {
        const targetHeader = targetHeaders[i];
        const sourceIdx = sourceHeaders.indexOf(targetHeader);
        if (sourceIdx !== -1 && sourceIdx < sourceValues.length) {
          alignedData.push(sourceValues[sourceIdx]);
        } else {
          alignedData.push(""); // ستون وجود نداره، خالی میذاریم
        }
      }

      let foundRow = -1;
      const targetData = targetSheet.getDataRange().getValues();
      const targetInboxIdx = targetHeaders.indexOf("INBOX_ID");
      
      if (targetData.length > 1 && targetInboxIdx !== -1) {
        for (let i = 1; i < targetData.length; i++) {
          if (targetData[i][targetInboxIdx] == inboxId) {
            foundRow = i + 1;
            break;
          }
        }
      }

      if (foundRow > 0) {
        // به‌روزرسانی ردیف موجود - فقط ستون‌هایی که وجود دارن
        for (let i = 0; i < alignedData.length && i < targetHeaders.length; i++) {
          targetSheet.getRange(foundRow, i + 1).setValue(alignedData[i]);
        }
      } else {
        // اضافه کردن ردیف جدید
        if (alignedData.length > 0) {
          targetSheet.appendRow(alignedData);
        }
      }
    }
  }
  return "Success";
}

function searchInContractStage(biName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("Contract Stage");
  if (!sheet) return { headers: [], rows: [] };
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const biIdx = headers.indexOf("BI_Name");
  if (biIdx === -1) return { headers: [], rows: [] };
  const filteredRows = data.slice(1).filter(row => row[biIdx] && row[biIdx].toString().toLowerCase() === biName.toLowerCase());
  return { headers: headers, rows: filteredRows };
}

function deleteLeadRow(rowIndex) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("Leads_Inbox");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const actualRow = parseInt(rowIndex) + 2;
  const inboxId = sheet.getRange(actualRow, headers.indexOf("INBOX_ID") + 1).getValue().toString();
  sheet.deleteRow(actualRow);
  const targetSheet = ss.getSheetByName("Negotiation Stage");
  if (targetSheet) {
    const targetData = targetSheet.getDataRange().getValues();
    const idCol = targetData[0].indexOf("INBOX_ID");
    for (let i = 1; i < targetData.length; i++) {
      if (targetData[i][idCol] && targetData[i][idCol].toString() === inboxId) {
        targetSheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return "Deleted Successfully";
}

function submitToSheet(payload, targetSheetName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(targetSheetName);

  if (!sheet) {
    throw new Error("Target sheet not found: " + targetSheetName);
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => h.toString().trim());

  const now = new Date();
  const submittedAt = Utilities.formatDate(
    now,
    "Asia/Tehran",
    "dd/MM/yyyy HH:mm:ss"
  );

  const newRowData = headers.map(h => {
    if (h === "Manager_Status") return "Not Checked";

    if (h === "Submitted_At") {
      return submittedAt;
    }

    if (h === "Submitter") {
      return payload.__currentUser || "";
    }

    return payload[h] !== undefined ? payload[h] : "";
  });

  sheet.appendRow(newRowData);
  return "Success";
}

function getDropdownOptions() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("helper");
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const options = {};
  
  headers.forEach((header, colIdx) => {
    options[header] = data.slice(1)
      .map(row => row[colIdx])
      .filter(v => v !== "");
  });
  
  return options;
}

function getAllNegotiationLeads() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("Negotiation Stage");
  if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);

  const createdIdx = headers.indexOf("Created_Date");
  const acqIdx = headers.indexOf("Acq_Stage");
  const pmerIdx = headers.indexOf("Linked_PMER_ID");
  const inboxIdx = headers.indexOf("INBOX_ID");

  const cutoffDate = new Date('2025-12-29');

  const filteredRows = rows.filter(row => {
    if (createdIdx === -1) return false;
    const created = row[createdIdx];
    if (!created) return false;
    const createdDate = new Date(created);
    const isNew = createdDate >= cutoffDate;
    const isActive = acqIdx === -1 || row[acqIdx] !== "Sent for contract adjustment";
    return isNew && isActive;
  });

  const rowsWithIds = filteredRows.map(row => {
    return {
      data: row,
      pmerId: pmerIdx !== -1 ? row[pmerIdx] : "",
      inboxId: inboxIdx !== -1 ? row[inboxIdx] : ""
    };
  });

  return { headers: headers, rows: rowsWithIds };
}

function updateNegotiationRow(identifier, payload) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName("Negotiation Stage");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    
    const pmerColIdx = headers.indexOf("Linked_PMER_ID");
    const inboxColIdx = headers.indexOf("INBOX_ID");
    
    const allData = sheet.getDataRange().getValues();
    let actualRow = -1;
    
    for (let i = 1; i < allData.length; i++) {
      const rowPmerId = pmerColIdx !== -1 ? allData[i][pmerColIdx] : "";
      const rowInboxId = inboxColIdx !== -1 ? allData[i][inboxColIdx] : "";
      
      if ((identifier.pmerId && rowPmerId === identifier.pmerId) || 
          (identifier.inboxId && rowInboxId === identifier.inboxId)) {
        actualRow = i + 1;
        break;
      }
    }
    
    if (actualRow === -1) {
      throw new Error("ردیف مورد نظر یافت نشد");
    }
    
    // به‌روزرسانی فیلدهای payload - فقط ستون‌هایی که وجود دارن
    for (let key in payload) {
      const colIdx = headers.indexOf(key);
      if (colIdx !== -1) {
        // ستون وجود داره، مقدار رو میذاریم
        sheet.getRange(actualRow, colIdx + 1).setValue(payload[key]);
      }
      // اگه ستون وجود نداشته باشه، skip می‌کنیم (خطا نمیده)
    }
    
    // به‌روزرسانی تاریخ - فقط اگه ستون وجود داشته باشه
    const updatedColIdx = headers.indexOf("Acq_Stage_Updated_At");
    if (updatedColIdx !== -1) {
      sheet.getRange(actualRow, updatedColIdx + 1).setValue(new Date());
    }
    
    return "Success";
  } catch (error) {
    Logger.log("Error in updateNegotiationRow: " + error.message);
    throw error;
  }
}

function setSentForAdjustment(identifier) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sourceSheet = ss.getSheetByName("Negotiation Stage");
    const targetSheet = ss.getSheetByName("Contract Stage");
    
    if (!targetSheet) {
      throw new Error("Sheet Contract Stage یافت نشد");
    }
    
    const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    
    const pmerColIdx = sourceHeaders.indexOf("Linked_PMER_ID");
    const inboxColIdx = sourceHeaders.indexOf("INBOX_ID");
    
    const allData = sourceSheet.getDataRange().getValues();
    let actualRow = -1;
    
    for (let i = 1; i < allData.length; i++) {
      const rowPmerId = pmerColIdx !== -1 ? allData[i][pmerColIdx] : "";
      const rowInboxId = inboxColIdx !== -1 ? allData[i][inboxColIdx] : "";
      
      if ((identifier.pmerId && rowPmerId === identifier.pmerId) || 
          (identifier.inboxId && rowInboxId === identifier.inboxId)) {
        actualRow = i + 1;
        break;
      }
    }
    
    if (actualRow === -1 || actualRow <= 0) {
      throw new Error("ردیف مورد نظر یافت نشد");
    }
    
    // به‌روزرسانی ستون‌های source sheet - فقط اگه وجود داشته باشن
    const acqStageColIdx = sourceHeaders.indexOf("Acq_Stage");
    if (acqStageColIdx !== -1) {
      sourceSheet.getRange(actualRow, acqStageColIdx + 1).setValue("Sent for contract adjustment");
    }
    
    const lastActionColIdx = sourceHeaders.indexOf("Acq_Last_Action_At");
    if (lastActionColIdx !== -1) {
      sourceSheet.getRange(actualRow, lastActionColIdx + 1).setValue(new Date());
    }
    
    const finalStatusColIdx = sourceHeaders.indexOf("Final Acq_Status");
    if (finalStatusColIdx !== -1) {
      sourceSheet.getRange(actualRow, finalStatusColIdx + 1).setValue("Sent for contract adjustment");
    }

    // خواندن داده‌های source
    const sourceLastCol = sourceSheet.getLastColumn();
    if (sourceLastCol <= 0) {
      throw new Error("تعداد ستون‌های source sheet نامعتبر است");
    }
    
    const sourceData = sourceSheet.getRange(actualRow, 1, 1, sourceLastCol).getDisplayValues()[0];
    
    // انتقال بر اساس ستون‌های موجود در target sheet - فقط ستون‌هایی که وجود دارن
    const alignedData = [];
    for (let i = 0; i < targetHeaders.length; i++) {
      const targetHeader = targetHeaders[i];
      const sourceIdx = sourceHeaders.indexOf(targetHeader);
      if (sourceIdx !== -1 && sourceIdx < sourceData.length) {
        alignedData.push(sourceData[sourceIdx]);
      } else {
        alignedData.push(""); // ستون وجود نداره، خالی میذاریم
      }
    }

    if (alignedData.length === 0) {
      throw new Error("هیچ داده‌ای برای انتقال یافت نشد");
    }

    // پیدا کردن یا ساخت ردیف در target sheet
    let foundRow = -1;
    const targetData = targetSheet.getDataRange().getValues();
    const targetPmerIdx = targetHeaders.indexOf("Linked_PMER_ID");
    const targetInboxIdx = targetHeaders.indexOf("INBOX_ID");
    
    if (targetData.length > 1) {
      for (let i = 1; i < targetData.length; i++) {
        const rowPmerId = targetPmerIdx !== -1 ? targetData[i][targetPmerIdx] : "";
        const rowInboxId = targetInboxIdx !== -1 ? targetData[i][targetInboxIdx] : "";
        
        if ((identifier.pmerId && rowPmerId === identifier.pmerId) || 
            (identifier.inboxId && rowInboxId === identifier.inboxId)) {
          foundRow = i + 1;
          break;
        }
      }
    }

    // نوشتن داده‌ها - فقط اگه داده وجود داشته باشه
    if (foundRow > 0) {
      // به‌روزرسانی ردیف موجود
      for (let i = 0; i < alignedData.length && i < targetHeaders.length; i++) {
        targetSheet.getRange(foundRow, i + 1).setValue(alignedData[i]);
      }
    } else {
      // اضافه کردن ردیف جدید
      targetSheet.appendRow(alignedData);
    }
    
    return "Success";
  } catch (error) {
    Logger.log("Error in setSentForAdjustment: " + error.message);
    throw error;
  }
}

function getAllContractStageLeads() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("Contract Stage");
  if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);

  const createdIdx = headers.indexOf("Created_Date");
  const finalStatusIdx = headers.indexOf("Final_Contract_Status");
  const pmerIdx = headers.indexOf("Linked_PMER_ID");
  const inboxIdx = headers.indexOf("INBOX_ID");

  const cutoffDate = new Date('2025-12-29');

  // فیلتر: فقط قراردادهای جدید که نهایی نشده‌اند
  const filteredRows = rows.filter(row => {
    // فقط قراردادهایی که نهایی نشده‌اند
    const finalStatus = finalStatusIdx !== -1 ? row[finalStatusIdx] : "";
    if (finalStatus === "Finalized") return false;
    
    // فقط قراردادهای جدید (از تاریخ cutoff به بعد)
    if (createdIdx !== -1) {
      const created = row[createdIdx];
      if (created) {
        try {
          const createdDate = new Date(created);
          if (createdDate < cutoffDate) return false;
        } catch (e) {
          return false;
        }
      } else {
        return false;
      }
    }
    
    return true;
  });

  const rowsWithIds = filteredRows.map(row => {
    return {
      data: row,
      pmerId: pmerIdx !== -1 ? row[pmerIdx] : "",
      inboxId: inboxIdx !== -1 ? row[inboxIdx] : ""
    };
  });

  return { headers: headers, rows: rowsWithIds };
}

function updateContractStageRow(identifier, payload, isFinal) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName("Contract Stage");
    if (!sheet) {
      throw new Error("Sheet Contract Stage یافت نشد");
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    
    const pmerColIdx = headers.indexOf("Linked_PMER_ID");
    const inboxColIdx = headers.indexOf("INBOX_ID");
    
    const allData = sheet.getDataRange().getValues();
    let actualRow = -1;
    
    for (let i = 1; i < allData.length; i++) {
      const rowPmerId = pmerColIdx !== -1 ? allData[i][pmerColIdx] : "";
      const rowInboxId = inboxColIdx !== -1 ? allData[i][inboxColIdx] : "";
      
      if ((identifier.pmerId && rowPmerId === identifier.pmerId) || 
          (identifier.inboxId && rowInboxId === identifier.inboxId)) {
        actualRow = i + 1;
        break;
      }
    }
    
    if (actualRow === -1) {
      throw new Error("ردیف مورد نظر یافت نشد");
    }
    
    // به‌روزرسانی فیلدهای payload - فقط ستون‌هایی که وجود دارن
    for (let key in payload) {
      const colIdx = headers.indexOf(key);
      if (colIdx !== -1) {
        // ستون وجود داره، مقدار رو میذاریم
        sheet.getRange(actualRow, colIdx + 1).setValue(payload[key]);
      }
      // اگه ستون وجود نداشته باشه، skip می‌کنیم (خطا نمیده)
    }
    
    // به‌روزرسانی تاریخ - فقط اگه ستون وجود داشته باشه
    const updatedColIdx = headers.indexOf("Contract_Updated_At");
    if (updatedColIdx !== -1) {
      sheet.getRange(actualRow, updatedColIdx + 1).setValue(new Date());
    }
    
    // ذخیره نهایی - فقط اگه ستون وجود داشته باشه
    if (isFinal) {
      const finalStatusColIdx = headers.indexOf("Final_Contract_Status");
      if (finalStatusColIdx !== -1) {
        sheet.getRange(actualRow, finalStatusColIdx + 1).setValue("Finalized");
      }
    }
    
    return "Success";
  } catch (error) {
    Logger.log("Error in updateContractStageRow: " + error.message);
    throw error;
  }
}

function copyToContractStage(sourceRowIdx) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sourceSheet = ss.getSheetByName("Negotiation Stage");
  const targetSheet = ss.getSheetByName("Contract Stage");
  if (!targetSheet) return;

  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const sourceData = sourceSheet.getRange(sourceRowIdx, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0];

  const rowToWrite = targetHeaders.map(th => {
    const idx = sourceHeaders.indexOf(th);
    return idx !== -1 ? sourceData[idx] : "";
  });

  const startRow = 399;
  const colAValues = targetSheet.getRange(startRow, 1, targetSheet.getMaxRows() - startRow + 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < colAValues.length; i++) {
    if (colAValues[i][0] === "" || colAValues[i][0] === null) {
      targetRow = startRow + i;
      break;
    }
  }
  if (targetRow === -1) targetRow = targetSheet.getLastRow() + 1;

  targetSheet.getRange(targetRow, 1, 1, rowToWrite.length).setValues([rowToWrite]);

  const createdColIdx = targetHeaders.indexOf("Created_Date") + 1;
  if (createdColIdx > 0) {
    targetSheet.getRange(targetRow, createdColIdx).setValue(new Date());
  }
}

function insertIntoFirstEmptyRow(newRowData) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName("Contract Stage");
  if (!sheet) return "Error";

  const startRow = 399;
  const colAValues = sheet.getRange(startRow, 1, sheet.getMaxRows() - startRow + 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < colAValues.length; i++) {
    if (colAValues[i][0] === "" || colAValues[i][0] === null) {
      targetRow = startRow + i;
      break;
    }
  }
  if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

  sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const createdColIdx = headers.indexOf("Created_Date") + 1;
  if (createdColIdx > 0) {
    sheet.getRange(targetRow, createdColIdx).setValue(new Date());
  }
  return "Success";
}

function logActivity(user, activity, details) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let logSheet = ss.getSheetByName("Activity_Log");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("Activity_Log");
      logSheet.appendRow(["User", "Time", "Activity", "Details"]);
    }
    
    const timestamp = new Date();
    logSheet.appendRow([user, timestamp, activity, details]);
    
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1000) {
      logSheet.deleteRows(2, lastRow - 1000);
    }
    
    return "Success";
  } catch (error) {
    Logger.log("Error in logActivity: " + error.message);
    return "Error";
  }
}

function getLogs() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const logSheet = ss.getSheetByName("Activity_Log");
    
    if (!logSheet || logSheet.getLastRow() < 2) {
      return [];
    }
    
    const data = logSheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    return rows.slice(-100).reverse().map(row => {
      return {
        user: row[0] || "",
        time: row[1] ? new Date(row[1]).toLocaleString('fa-IR') : "",
        activity: row[2] || "",
        details: row[3] || ""
      };
    });
  } catch (error) {
    Logger.log("Error in getLogs: " + error.message);
    return [];
  }
}