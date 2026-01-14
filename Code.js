/**
 * 🏢 Project MASTER SCRIPT V61.0 - THE FORTRESS EDITION (PART 1/3)
 * ===================================================================
 * A Zero-Trust System for RealEstateProject Allottees Association.
 * VERIFIED: 2026-01-13 | STATUS: PRODUCTION READY
 * * * CORE MODULES IN THIS BLOCK:
 * 1. 🔒 GATEKEEPER: Validates Flat & Email Identity on Form Submit (No Auto-Create).
 * 2. 🛡️ AUDITOR: Priority Queue + Deep API Scan + Ledger Math + Path Walker.
 */

// --- 1. CONFIGURATION (USER MUST UPDATE THESE) ---
// 🔒 SECURITY NOTE: Replace these placeholders with your actual IDs before running.
// DO NOT commit your real IDs to GitHub/Public repositories.
const ROOT_FOLDER_ID = 'ENTER_YOUR_ROOT_FOLDER_ID_HERE'; 
const MNGR_ADMIN_EMAIL = 'admin@example.com'; 
const IGNORED_EMAILS = ['admin@example.com', 'committee@example.com']; 
const BACKUP_FOLDER_ID = 'ENTER_YOUR_BACKUP_FOLDER_ID_HERE';

// ⚠️ PASTE YOUR GOOGLE DOC ID BELOW.
const RECEIPT_TEMPLATE_ID = 'ENTER_YOUR_TEMPLATE_ID_HERE'; 

const BANK_DETAILS = {
  vpa: "your-association@bank",
  bank: "Your Bank Name"
};
// ==========================================
// 🚀 ENGINE 1: THE RECEPTIONIST (ON FORM SUBMIT)
// ==========================================

function onFormSubmitTrigger(e) {
  if (!e) return;
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    processRow(e.range.getRow(), ss); 
    updateNCLTWatchlist(); 
  } catch (err) {
    MailApp.sendEmail(MNGR_ADMIN_EMAIL, "🚨 Project SCRIPT ERROR", "Row " + e.range.getRow() + ": " + err.message);
  } finally { lock.releaseLock(); }
}

function processRow(rowIndex, ss) {
  const respSheet = ss.getSheetByName("Responded");
  const structSheet = ss.getSheetByName("Structure");
  const headers = structSheet.getRange(1, 1, 1, structSheet.getLastColumn()).getValues()[0];
  const colIdx = (name) => headers.indexOf(name);
  
  const rowData = respSheet.getRange(rowIndex, 1, 1, respSheet.getLastColumn()).getValues()[0];
  const timestamp = rowData[0];
  const submitterEmail = String(rowData[1]).trim().toLowerCase(); 
  const rawFlat = String(rowData[2]).trim(); 
  const ownerName = rowData[4]; 
  const phoneNumber = String(rowData[5]).trim();
  const fileUrls = rowData[6]; 

  respSheet.getRange(rowIndex, respSheet.getLastColumn()).setValue("Processing...").setBackground("#fff2cc");

  const structData = structSheet.getDataRange().getValues();
  let folderId = null, folderRow = -1, folderUrl = "", masterEmail = "";
  const targetKey = rawFlat.toLowerCase().replace(/\s+/g, '');

  for (let i = 1; i < structData.length; i++) {
    if (String(structData[i][colIdx("Flat")]).toLowerCase().replace(/\s+/g, '') === targetKey) { 
      masterEmail = String(structData[i][colIdx("Email")] || "").trim().toLowerCase();
      folderRow = i + 1; 
      folderUrl = structData[i][colIdx("Folder_Link")]; 
      folderId = (folderUrl || "").match(/[-\w]{25,}/)?.[0];
      break;
    }
  }

  // 🔒 SECURITY GATE 1: UNKNOWN FLAT
  if (folderRow === -1) {
    respSheet.getRange(rowIndex, respSheet.getLastColumn()).setValue("⏳ Pending Review").setBackground("#f4cccc");
    sendTemplateEmail(ss, "AUTO_UNKNOWN", {Flat: rawFlat, Name: ownerName, Email: submitterEmail});
    return; // STOP.
  }

  // 🔒 SECURITY GATE 2: IDENTITY MISMATCH (True Zero-Trust)
  if (!masterEmail || submitterEmail !== masterEmail) {
    const reason = !masterEmail ? "Email not pre-configured" : `User ${submitterEmail} != Owner ${masterEmail}`;
    respSheet.getRange(rowIndex, respSheet.getLastColumn()).setValue("⛔ Mismatch").setBackground("#ea9999");
    logToAudit(ss, "IDENTITY_MISMATCH", rawFlat, "Form Submission", "⛔ BLOCKED", reason, "Manual Intervention Required");
    sendTemplateEmail(ss, "AUTO_MISMATCH", {Flat: rawFlat, Name: ownerName, Email: submitterEmail});
    return; // STOP.
  }

  // 3. FILE SYNC (Only if Gates Passed)
  if (fileUrls) {
    if (!folderId) {
        respSheet.getRange(rowIndex, respSheet.getLastColumn()).setValue("⚠️ No Folder").setBackground("#ffe599");
        return;
    }

    const targetFolder = DriveApp.getFolderById(folderId);
    const ids = fileUrls.match(/[-\w]{25,}/g) || [];
    
    // Ledger Check
    let uploadedStr = String(structData[folderRow-1][colIdx("Uploaded IDs")] || "");
    let presentStr = String(structData[folderRow-1][colIdx("Present IDs")] || "");
    let currentSize = parseFloat(structData[folderRow-1][colIdx("Total Size (MB)")] || 0);
    let currentFound = parseInt(structData[folderRow-1][colIdx("Files Found")] || 0);
    let newSuccess = 0, addedSize = 0;

    ids.forEach(id => {
      if (uploadedStr.includes(id)) return; // Skip duplicates

      try {
        const file = DriveApp.getFileById(id);
        if (file.isTrashed()) return;

        const sizeMB = file.getSize() / (1024 * 1024);
        file.setName(getSanitizedName(file.getName())); // Sanitize Name
        file.moveTo(targetFolder); 
        
        uploadedStr += (uploadedStr ? "," : "") + id;
        presentStr += (presentStr ? "," : "") + id;
        newSuccess++;
        addedSize += sizeMB;
        
        logToAudit(ss, "SYNC", rawFlat, file.getName(), "✅ SUCCESS", "Moved to Vault", "Ledger Updated");
      } catch (err) { console.error("ID Err: " + id); }
    });

    // ⚡ LIVE DASHBOARD UPDATE
    structSheet.getRange(folderRow, colIdx("Uploaded IDs") + 1).setValue(uploadedStr);
    structSheet.getRange(folderRow, colIdx("Present IDs") + 1).setValue(presentStr);
    structSheet.getRange(folderRow, colIdx("Files Found") + 1).setValue(currentFound + newSuccess);
    structSheet.getRange(folderRow, colIdx("Total Size (MB)") + 1).setValue((currentSize + addedSize).toFixed(2));
    structSheet.getRange(folderRow, colIdx("Last Synced") + 1).setValue(new Date());

    respSheet.getRange(rowIndex, respSheet.getLastColumn()).setValue("Done").setBackground("#d9ead3");
    
    if (newSuccess > 0) {
      sendTemplateEmail(ss, "AUTO_RECEIPT", {Flat: rawFlat, Name: ownerName, Email: submitterEmail, Link: folderUrl});
    }
  }
  
  // 4. UPDATE MNGR SHEET
  updateCommitteeSheet(ss, timestamp, rawFlat, submitterEmail, ownerName, phoneNumber, folderUrl);
}

// ==========================================
// 🛡️ ENGINE 2: THE AUDITOR (PRIORITY + ADVANCED API)
// ==========================================

function provisionAndAuditDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Structure");
  const headers = sheet.getDataRange().getValues()[0];
  const col = (name) => headers.indexOf(name);
  const data = sheet.getDataRange().getValues();
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  
  const updates = []; 
  const timestamp = new Date();
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 280 * 1000; 

  // 1. BUILD PRIORITY QUEUE
  let workQueue = [];
  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const folderLink = String(rowData[col("Folder_Link")] || "");
    const fullPath = String(rowData[col("Full_Path")] || "");
    const access = String(rowData[col("Access")] || "");
    const lastCheck = rowData[col("Last Checked")];
    
    let tier = 3; 
    if (folderLink === "") tier = 1; // TIER 1: New Onboarding
    else if (fullPath === "" || access === "" || access.includes("Error")) tier = 2; // TIER 2: Repairs
    
    workQueue.push({
      rowIndex: i + 1,
      rowData: rowData,
      tier: tier,
      lastCheckTime: lastCheck ? new Date(lastCheck).getTime() : 0
    });
  }

  workQueue.sort((a, b) => a.tier !== b.tier ? a.tier - b.tier : a.lastCheckTime - b.lastCheckTime);

  // 2. PROCESS QUEUE
  for (const item of workQueue) {
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) break;

    const row = item.rowIndex;
    const flat = item.rowData[col("Flat")];
    const email = String(item.rowData[col("Email")] || "").trim().toLowerCase();
    if (!flat) continue;

    try {
      let folderId = (item.rowData[col("Folder_Link")] || "").match(/[-\w]{25,}/)?.[0];
      let folder;
      
      // A. PROVISIONING
      if (!folderId) {
        folder = root.createFolder(flat);
        folderId = folder.getId();
        updates.push({row: row, col: col("Folder_Link")+1, val: folder.getUrl()});
      } else {
        folder = DriveApp.getFolderById(folderId);
      }

      // B. RECURSIVE PATH CHECK
      const actualPath = getFullPath(folder, ROOT_FOLDER_ID);
      updates.push({row: row, col: col("Full_Path")+1, val: actualPath});

      // C. DEEP PERMISSION AUDIT
      let accessStatus = "✅ OK";
      let issueNote = "";
      try {
        const filePerms = Drive.Files.get(folderId, { fields: 'permissions' }).permissions || [];
        const sharedEmails = filePerms
          .filter(p => p.role !== 'owner' && p.type === 'user')
          .map(p => p.emailAddress?.toLowerCase())
          .filter(e => e && !IGNORED_EMAILS.includes(e));

        if (email && email.includes("@") && !sharedEmails.includes(email)) {
          try { 
            folder.addViewer(email); 
            accessStatus = "✅ Fixed (Added Member)"; 
          } catch (e) { accessStatus = "❌ Error: Non-Compatible Domain (Yahoo/Corporate)"; }
        }

        const strangers = sharedEmails.filter(e => e !== email);
        if (strangers.length > 0) issueNote = `❗ Unexpected: ${strangers.join(", ")}`;
      } catch (e) { accessStatus = "⚠️ API Access Error"; }
      
      updates.push({row: row, col: col("Access")+1, val: accessStatus});
      updates.push({row: row, col: col("Access Issues")+1, val: issueNote || "✅ OK"});

      // D. LEDGER MATH
      const files = folder.getFiles();
      let currentIDs = [];
      let totalBytes = 0;
      while (files.hasNext()) {
        const f = files.next();
        currentIDs.push(f.getId());
        totalBytes += f.getSize();
      }

      let uploadedStr = String(item.rowData[col("Uploaded IDs")] || "").trim();
      let uploadedArr = uploadedStr ? uploadedStr.split(",").map(s => s.trim()) : [];
      
      const manualUploads = currentIDs.filter(id => !uploadedArr.includes(id));
      const deletedFiles = uploadedArr.filter(id => !currentIDs.includes(id));

      let health = (totalBytes > 52428800) ? "⚠️ HEAVY STORAGE" : "✅ Normal";
      let issues = "✅ Synced";
      if (manualUploads.length > 0) issues = `⚠️ Manual Upload (+${manualUploads.length})`;
      if (deletedFiles.length > 0) issues = (issues.includes("⚠️") ? issues + " | " : "") + `❌ Missing (${deletedFiles.length} deleted)`;

      updates.push({row: row, col: col("Present IDs")+1, val: currentIDs.join(",")});
      updates.push({row: row, col: col("Deleted IDs")+1, val: deletedFiles.join(",")});
      updates.push({row: row, col: col("Issues/Comments")+1, val: issues});
      updates.push({row: row, col: col("Health Status")+1, val: health});
      updates.push({row: row, col: col("Total Size (MB)")+1, val: (totalBytes/1048576).toFixed(2)});
      updates.push({row: row, col: col("Last Checked")+1, val: timestamp});

    } catch (e) {
      updates.push({row: row, col: col("Issues/Comments")+1, val: "Crit Error: " + e.message});
    }
  }
  batchWriteUpdates(sheet, updates);
}

// ==========================================
// ⚖️ ENGINE 3: NCLT COMPLIANCE (DYNAMIC SCANNER)
// ==========================================

function updateNCLTWatchlist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const commSheet = ss.getSheetByName("MNGR_Worksheet");
  const ncltSheet = ss.getSheetByName("NCLT");
  
  if (!commSheet || !ncltSheet) return;

  const commDataRaw = commSheet.getDataRange().getValues();
  const commHeaders = commDataRaw[0];
  const col = (name) => commHeaders.indexOf(name);
  
  const ncltTargetCol = col("NCLT");
  if (ncltTargetCol === -1) { console.error("❌ 'NCLT' column missing."); return; }

  // 1. BUILD WATCHLIST SET (O(1) Speed)
  const ncltSet = new Set(
    ncltSheet.getDataRange().getValues().flat()
    .map(v => String(v).toLowerCase().trim())
    .filter(v => v.length > 2)
  );

  // 2. DYNAMIC LOOKUP (Only check Flat, Name, Phone)
  const flatIdx = col("Flat");
  const nameIdx = col("Owner Name");
  const phoneIdx = col("Phone Number");

  const results = commDataRaw.slice(1).map(row => {
    const fieldsToCheck = [
      flatIdx !== -1 ? row[flatIdx] : "",
      nameIdx !== -1 ? row[nameIdx] : "",
      phoneIdx !== -1 ? row[phoneIdx] : ""
    ];
    const isMatch = fieldsToCheck.some(field => field && ncltSet.has(String(field).toLowerCase().trim()));
    return isMatch ? ["⚖️ NCLT MATCH"] : [""];
  });

  commSheet.getRange(2, ncltTargetCol + 1, results.length, 1).setValues(results);
}

// ==========================================
// 💳 ENGINE 4: FINANCE & QR (UPI SAFE + NOTES)
// ==========================================

function sendPaymentRequest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== "MNGR_Worksheet") { SpreadsheetApp.getUi().alert("⚠️ Switch to 'MNGR_Worksheet'."); return; }
   
  const row = sheet.getActiveCell().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("⚠️ Select a member row."); return; }
   
  const fullRange = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIdx = (name) => headers.indexOf(name);
   
  const rawFlat = String(fullRange[colIdx("Flat")] || "").trim(); 
  const cleanFlat = rawFlat.replace(/[^a-zA-Z0-9]/g, ""); // Hyphen-Free for UPI
  const amount = fullRange[colIdx("Payment Amount")];

  if (!cleanFlat || !amount || !BANK_DETAILS.vpa) { SpreadsheetApp.getUi().alert("⛔ Missing Flat/Amount/UPI ID."); return; }

  // GENERATE UPI (tr = Ref, tn = Note)
  const upiString = `upi://pay?pa=${BANK_DETAILS.vpa}&pn=ProjectAOA&am=${amount}&cu=INR&tr=${cleanFlat}&tn=${cleanFlat}`;
  const qrImageUrl = `https://quickchart.io/qr?text=${encodeURIComponent(upiString)}&size=350&ecLevel=Q&margin=2`;
  const qrCol = colIdx("QR Code Link");
  if (qrCol > -1) sheet.getRange(row, qrCol + 1).setValue(qrImageUrl);

  const emailData = {
    Flat: cleanFlat, // Shows as "A10405"
    OriginalFlat: rawFlat, 
    Name: fullRange[colIdx("Owner Name")], 
    Email: fullRange[colIdx("Email")], 
    Amount: amount,
    BankDetails: `UPI ID: ${BANK_DETAILS.vpa}<br>Bank: ${BANK_DETAILS.bank}<br><strong>Ref/Note: ${cleanFlat}</strong>`,
    QRCode: `<img src="${qrImageUrl}" alt="Scan to Pay" width="250">`
  };

  sendTemplateEmail(ss, "PAYMENT_REQ", emailData);
  SpreadsheetApp.getUi().alert(`✅ Payment Request Sent to ${rawFlat}`);
}

// ==========================================
// 🧾 ENGINE 6: RECEIPTS (VAULT + CLEANUP)
// ==========================================

function sendPaymentReceipts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MNGR_Worksheet");
  const ui = SpreadsheetApp.getUi();

  // SAFETY LOCK
  if (!RECEIPT_TEMPLATE_ID || RECEIPT_TEMPLATE_ID.includes("PASTE_YOUR")) {
    ui.alert("⛔ STOP: Update Line 17 (RECEIPT_TEMPLATE_ID) with your Google Doc ID.");
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert("⚠️ Select a member row."); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name);
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
   
  const invoiceNo = data[col("Invoice Number")];
  const folderLink = data[col("Folder_Link")];
  if (!invoiceNo) { ui.alert("⛔ Missing Invoice Number."); return; }

  let rawStatus = String(data[col("Payment Status")] || "");
  let cleanRef = rawStatus.match(/\b\d{12}\b/)?.[0] || rawStatus.match(/\(([^)]+)\)$/)?.[1] || rawStatus;
  
  // DATE FIX: Use Timestamp/Payment Date if available, else Today
  const paymentDateRaw = data[col("Timestamp")] || new Date(); 
  const formattedDate = Utilities.formatDate(new Date(paymentDateRaw), "GMT+5:30", "dd/MM/yyyy");

  const cleanFlat = String(data[col("Flat")]).replace(/[^a-zA-Z0-9]/g, "");
  const replacements = {
    '{{Name}}': data[col("Owner Name")],
    '{{Flat}}': data[col("Flat")],
    '{{Amount}}': data[col("Payment Amount")],
    '{{InvoiceNo}}': invoiceNo,
    '{{MembershipID}}': "MEM-" + cleanFlat,
    '{{Date}}': formattedDate, 
    '{{RefNo}}': cleanRef
  };

  let tempFileId = null;
  try {
    const templateFile = DriveApp.getFileById(RECEIPT_TEMPLATE_ID);
    const tempCopy = templateFile.makeCopy("TEMP_RECEIPT_" + invoiceNo);
    tempFileId = tempCopy.getId();
    const doc = DocumentApp.openById(tempFileId);
    const body = doc.getBody();
    for (const [key, value] of Object.entries(replacements)) body.replaceText(key, String(value));
    doc.saveAndClose();

    const pdfBlob = tempCopy.getAs(MimeType.PDF).setName(`Receipt_${invoiceNo}.pdf`);

    // SAVE TO VAULT FIRST (Security)
    const folderId = (folderLink || "").match(/[-\w]{25,}/)?.[0];
    if (folderId) {
        try { DriveApp.getFolderById(folderId).createFile(pdfBlob); }
        catch (e) { console.error("Vault Save Failed: " + e.message); }
    }

    // EMAIL
    MailApp.sendEmail({
      to: data[col("Email")],
      subject: `✅ Project Payment Receipt: ${replacements['{{Flat}}']}`,
      htmlBody: `Dear ${replacements['{{Name}}']},<br><br>Payment Confirmed. Receipt attached.<br><br>Project Committee`,
      attachments: [pdfBlob]
    });

    ui.alert(`✅ Receipt Sent to ${replacements['{{Name}}']}`);
  } catch (err) { ui.alert("❌ Error: " + err.message); } 
  finally { if (tempFileId) try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch(e){} }
}

// ==========================================
// 🤝 ENGINE 7: WELCOME PACK (AUDITED)
// ==========================================

function sendWelcomePack() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MNGR_Worksheet");
  const ui = SpreadsheetApp.getUi();

  const row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert("⚠️ Select a Member Row first."); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name);
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const folderLink = data[col("Folder_Link")];
  // 1. LINK VALIDATION
  if (!folderLink || folderLink === "") {
    ui.alert("⛔ STOP: No Vault Link found.\n\n👉 Run 'Security Audit' first.");
    return;
  }

  const rawFlat = String(data[col("Flat")]);
  const memID = "MEM-" + rawFlat.replace(/[^a-zA-Z0-9]/g, ""); 
  
  const emailData = {
    Name: data[col("Owner Name")],
    Flat: rawFlat,
    Email: data[col("Email")],
    MembershipID: memID,
    Folder_Link: folderLink
  };

  try {
    sendTemplateEmail(ss, "WELCOME_PACK", emailData);
    
    // 2. AUDIT TRAIL
    const remarksCol = col("Remarks"); 
    if (remarksCol > -1) {
      const currentNote = data[remarksCol];
      const newNote = (currentNote ? currentNote + " | " : "") + `Welcome Sent: ${Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM")}`;
      sheet.getRange(row, remarksCol + 1).setValue(newNote);
    }
    ui.alert(`✅ Welcome Pack sent to ${rawFlat}`);
  } catch (e) { ui.alert("❌ Error: " + e.message); }
}

// ==========================================
// 📝 ENGINE 5: TEMPLATE SYSTEM (ATTACHMENT READY)
// ==========================================

function sendTemplateEmail(ss, templateKey, data, attachments = []) {
  const tmplSheet = ss.getSheetByName("Email_Templates");
  if (!tmplSheet) return;
  const templates = tmplSheet.getDataRange().getValues();
  let subject = "", body = "";

  // Case-Insensitive Lookup
  for (let i = 1; i < templates.length; i++) {
    if (String(templates[i][0]).trim().toUpperCase() === templateKey.toUpperCase()) {
      subject = templates[i][1]; body = templates[i][2]; break;
    }
  }

  if (!subject) { console.error("❌ Missing Template: " + templateKey); return; }

  // Replace Placeholders
  for (const [key, value] of Object.entries(data)) {
    const regex = new RegExp(`{{${key}}}`, "g");
    subject = subject.replace(regex, value);
    body = body.replace(regex, value);
  }
  
  // Send Email (Now supports Attachments for Certificates/Receipts)
  MailApp.sendEmail({
    to: data.Email,
    subject: subject,
    htmlBody: body, 
    name: "Project Allottees Association",
    attachments: attachments
  });
}

// ==========================================
// 🕒 ENGINE 8: TIME MACHINE (BACKUP ROTATION)
// ==========================================

function performSystemBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd_HHmm");
  const MAX_BACKUPS = 5; // Keep last 5 backups

  if (!BACKUP_FOLDER_ID || BACKUP_FOLDER_ID === 'PASTE_YOUR_BACKUP_FOLDER_ID_HERE') {
    console.error("Backup Folder ID not set."); return;
  }

  try {
    const backupFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
    const backupName = `BACKUP_Project_System_${dateStr}`;
    
    // Create Copy
    DriveApp.getFileById(ss.getId()).makeCopy(backupName, backupFolder);
    console.log(`✅ Backup created: ${backupName}`);

    // ROTATION LOGIC (Delete Oldest)
    const fileList = [];
    const allFiles = backupFolder.getFiles();
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      if (f.getName().includes("BACKUP_Project_System_")) {
        fileList.push({ id: f.getId(), date: f.getDateCreated(), file: f });
      }
    }
    fileList.sort((a, b) => b.date - a.date); // Sort Newest First

    // Delete excess files
    if (fileList.length > MAX_BACKUPS) {
      for (let i = MAX_BACKUPS; i < fileList.length; i++) {
        fileList[i].file.setTrashed(true);
        console.log("🗑️ Deleted old backup: " + fileList[i].file.getName());
      }
    }
  } catch (e) {
    MailApp.sendEmail(MNGR_ADMIN_EMAIL, "🚨 BACKUP FAILED", "Error: " + e.message);
  }
}

// ==========================================
// 🛠️ ADMIN UTILITIES & MENUS
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⭐ Project Admin')
    .addItem('📧 Open Email Sidebar', 'openEmailSidebar')
    .addSeparator()
    .addItem('💳 Send Payment Request', 'sendPaymentRequest')
    .addItem('🧾 Send Payment Receipt', 'sendPaymentReceipts')
    .addItem('🤝 Send Welcome Pack', 'sendWelcomePack')
    .addSeparator()
    .addItem('🔄 Force Sync Selected Row', 'manualSync') 
    .addItem('🛡️ Run Security Audit (Batch)', 'provisionAndAuditDrive')
    .addItem('⚖️ Refresh NCLT Watchlist', 'updateNCLTWatchlist')
    .addItem('🕒 Schedule Weekly Backup', 'setupWeeklyBackupTrigger')
    .addItem('🛠️ Initialize Ledger', 'initializeLedgerFromResponded')
    .addToUi();
}

function updateCommitteeSheet(ss, timestamp, flat, email, name, phone, link) {
  const commSheet = ss.getSheetByName("MNGR_Worksheet");
  if (!commSheet) return;

  const data = commSheet.getDataRange().getValues();
  const headers = data[0];
  const col = (name) => headers.indexOf(name);
  
  const normalize = (s) => String(s).toLowerCase().replace(/\s+/g, '').replace(/[–—]/g, '-');
  const key = normalize(flat);
  const flatColIdx = col("Flat");
  
  let targetRow = -1;
  // Dynamic Search (Robust against moved columns)
  if (flatColIdx > -1) {
    for (let i = 1; i < data.length; i++) { 
      if (normalize(data[i][flatColIdx]) === key) { targetRow = i + 1; break; } 
    }
  }

  const write = (header, val) => {
    const c = col(header);
    if (c > -1) commSheet.getRange(targetRow, c + 1).setValue(val);
  };

  if (targetRow === -1) {
    commSheet.appendRow([""]); 
    targetRow = commSheet.getLastRow();
    write("Status", "New");
    write("Remarks", "Unassigned");
  }

  write("Timestamp", timestamp);
  write("Flat", flat);
  write("Email", email);
  write("Owner Name", name);
  write("Phone Number", phone);
  write("Folder_Link", link);
}

function logToAudit(ss, category, flat, object, status, reason, actionItem) {
  let auditSheet = ss.getSheetByName("Sync_Audit_Log");
  // Auto-Create Log Sheet if missing
  if (!auditSheet) {
    auditSheet = ss.insertSheet("Sync_Audit_Log");
    auditSheet.appendRow(["Timestamp", "Category", "Flat/Email", "Object", "Status", "Reason", "Action Item"]);
    auditSheet.getRange("A1:G1").setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
    auditSheet.setFrozenRows(1);
  }
  auditSheet.appendRow([new Date(), category, flat, object, status, reason, actionItem]);
}

function initializeLedgerFromResponded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const respSheet = ss.getSheetByName("Responded");
  const structSheet = ss.getSheetByName("Structure");
  
  const structData = structSheet.getDataRange().getValues();
  const respData = respSheet.getDataRange().getValues();
  
  const colUploaded = structData[0].indexOf("Uploaded IDs");
  if (colUploaded === -1) { SpreadsheetApp.getUi().alert("❌ Column 'Uploaded IDs' missing."); return; }

  const flatMap = {};
  for (let i = 1; i < structData.length; i++) {
    const f = String(structData[i][1] || "").toLowerCase().replace(/\s/g, ''); 
    if(f) flatMap[f] = i + 1;
  }

  // Uses fallback indices if headers not found (2=Flat, 6=File)
  const respHeaders = respData[0];
  const safeFlatIdx = respHeaders.indexOf("Flat") > -1 ? respHeaders.indexOf("Flat") : 2; 
  const safeFileIdx = respHeaders.indexOf("Upload Ownership Documents") > -1 ? respHeaders.indexOf("Upload Ownership Documents") : 6;

  const pending = {};
  for (let r = 1; r < respData.length; r++) {
    const flat = String(respData[r][safeFlatIdx]).toLowerCase().replace(/\s/g, '');
    const ids = String(respData[r][safeFileIdx]).match(/[-\w]{25,}/g) || [];
    if (flatMap[flat] && ids.length) {
      if (!pending[flatMap[flat]]) pending[flatMap[flat]] = new Set();
      ids.forEach(id => pending[flatMap[flat]].add(id));
    }
  }

  for (const [row, set] of Object.entries(pending)) {
    structSheet.getRange(row, colUploaded + 1).setValue(Array.from(set).join(","));
  }
  SpreadsheetApp.getUi().alert("✅ Ledger Initialized!");
}

function setupWeeklyBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) { if (t.getHandlerFunction() === 'performSystemBackup') ScriptApp.deleteTrigger(t); }
  ScriptApp.newTrigger('performSystemBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(14).create();
  SpreadsheetApp.getUi().alert("✅ Backup Scheduled: Sundays @ 2 PM");
}

function manualSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== "Responded") { SpreadsheetApp.getUi().alert("⚠️ Go to 'Responded' sheet."); return; }
  processRow(sheet.getActiveCell().getRow(), ss);
  SpreadsheetApp.getUi().alert("✅ Sync Complete");
}


function getFullPath(folder, rootId) {
  let path = folder.getName();
  let parent = folder.getParents();
  while (parent.hasNext()) {
    const p = parent.next();
    path = p.getName() + "/" + path;
    if (p.getId() === rootId) break;
    parent = p.getParents();
  }
  return path;
}

function batchWriteUpdates(sheet, updates) {
  if (updates.length === 0) return;
  updates.forEach(u => sheet.getRange(u.row, u.col).setValue(u.val));
}

function openEmailSidebar() {
  const html = HtmlService.createHtmlOutput(HTML_SIDEBAR_CODE).setTitle('📧 Project Email Assistant').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getEmailTemplates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email_Templates");
  return sheet ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues() : [];
}

function createDraftFromSidebar(templateIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MNGR_Worksheet");
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return "⚠️ Select a row.";
  
  const templates = getEmailTemplates();
  const t = templates[templateIndex];
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = (n) => headers.indexOf(n);
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const flat = data[col("Flat")] || "";
  const name = data[col("Owner Name")] || "";
  const link = data[col("Folder_Link")] || "";

  let body = t[2]
    .replace(/{{Flat}}/g, flat)
    .replace(/{{Name}}/g, name)
    .replace(/{{Folder_Link}}/g, link)
    .replace(/{{Link}}/g, link);
    
  let subject = t[1].replace(/{{Flat}}/g, flat).replace(/{{Name}}/g, name);

  GmailApp.createDraft(data[col("Email")], subject, body, { htmlBody: body });
  return "✅ Draft created.";
}

const HTML_SIDEBAR_CODE = `<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:sans-serif;padding:10px}select,button{width:100%;margin-bottom:10px}</style></head><body><h3>📧 Email Assistant</h3><select id="tSelect"></select><button onclick="createDraft()">📝 Create Draft</button><script>function loadTemplates(){ google.script.run.withSuccessHandler(data => { const s = document.getElementById('tSelect'); data.forEach((t, i) => { let o = document.createElement('option'); o.value=i; o.text=t[0]; s.add(o); }); }).getEmailTemplates(); } function createDraft(){ const i = document.getElementById('tSelect').value; google.script.run.withSuccessHandler(m=>alert(m)).createDraftFromSidebar(i); } window.onload = loadTemplates;</script></body></html>`;