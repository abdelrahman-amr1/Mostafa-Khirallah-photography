/**
 * Photographer Client Delivery Portal - Backend Apps Script API Router
 * Photographer: Mostafa Khirallah
 * 
 * This script runs in Google Apps Script, serving as a JSON API for
 * the frontend hosted on GitHub Pages, connecting to Sheets and Drive.
 */

// Configuration
const ADMIN_PASSCODE = "2026"; // Default passcode for admin panel
const CLIENTS_SHEET_NAME = "Clients";
const FAVORITES_SHEET_NAME = "Favorites";

/**
 * Serves the JSON API and handles CORS.
 */
function doGet(e) {
  const action = e.parameter.action;
  let response = { success: false, message: "الطلب غير صالح." };
  
  try {
    initSpreadsheet();
    
    // Web GUI interface if loaded directly in browser
    if (!action) {
      let activeUrl = "";
      try {
        activeUrl = ScriptApp.getService().getUrl();
      } catch (err) {
        activeUrl = "رابط التطبيق بعد النشر (Web App URL)";
      }
      
      return HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; text-align: center; margin-top: 80px; direction: rtl; padding: 20px; background-color: #0b0b0c; color: #f1f5f9; min-height: 100vh;">
          <div style="max-width: 600px; margin: 0 auto; background: rgba(255,255,255,0.03); border: 1px solid rgba(212,175,55,0.2); padding: 40px; border-radius: 16px; box-shadow: 0 10px 30px rgba(0,0,0,0.5);">
            <h2 style="color: #d4af37; font-family: serif; font-size: 28px; margin-bottom: 10px;">📸 خادم تسليم الصور يعمل بنجاح!</h2>
            <p style="color: #9ca3af; font-size: 14px; line-height: 1.6;">بوابة الـ API الخاصة بك نشطة وجاهزة للتوصيل بموقع GitHub الخاص بك.</p>
            <hr style="border: 0; border-top: 1px solid rgba(255,255,255,0.1); margin: 25px 0;">
            <p style="font-size: 13px; color: #d4af37; margin-bottom: 8px; font-weight: bold;">قم بنسخ رابط الـ API أدناه وضعه في ملف index.html على GitHub:</p>
            <div style="background-color: #121216; border: 1px solid #1f1f26; padding: 15px; border-radius: 8px; word-break: break-all; margin-top: 10px;">
              <code style="color: #e5c158; font-size: 13px; font-family: monospace;">${activeUrl}</code>
            </div>
            <p style="color: #64748b; font-size: 11px; margin-top: 25px;">MOSTAFA KHIRALLAH PHOTOGRAPHY © 2026</p>
          </div>
        </div>
      `).setTitle("Mostafa Khirallah API Server");
    }
    
    // API Routing
    if (action === 'validateClient') {
      response = validateClient(e.parameter.mobile, e.parameter.invoice);
    } 
    else if (action === 'submitFavorites') {
      let favoriteList = [];
      try {
        favoriteList = JSON.parse(e.parameter.favorites);
      } catch (err) {
        favoriteList = [];
      }
      response = submitFavorites(e.parameter.invoice, e.parameter.name, e.parameter.mobile, favoriteList, e.parameter.notes);
    } 
    else if (action === 'adminGetDashboardData') {
      response = adminGetDashboardData(e.parameter.passcode);
    } 
    else if (action === 'adminCreateSession') {
      response = adminCreateSession(e.parameter.passcode, e.parameter.name, e.parameter.mobile, e.parameter.invoice, e.parameter.folderUrl, e.parameter.status);
    } 
    else if (action === 'adminDeleteSession') {
      response = adminDeleteSession(e.parameter.passcode, e.parameter.invoice);
    }
  } catch (err) {
    response = { success: false, message: "حدث خطأ في خادم جوجل: " + err.toString() };
  }
  
  // Return JSON and set CORS header via native ContentService behavior
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Initializes the Spreadsheet with required sheets and headers if they do not exist.
 */
function initSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Clients Sheet
  let clientSheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
  if (!clientSheet) {
    clientSheet = ss.insertSheet(CLIENTS_SHEET_NAME);
    clientSheet.appendRow([
      "Client Name", 
      "Mobile Number", 
      "Invoice Number", 
      "Drive Folder URL", 
      "Status", 
      "Created At"
    ]);
    clientSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    clientSheet.setFrozenRows(1);
  }
  
  // 2. Favorites Sheet
  let favSheet = ss.getSheetByName(FAVORITES_SHEET_NAME);
  if (!favSheet) {
    favSheet = ss.insertSheet(FAVORITES_SHEET_NAME);
    favSheet.appendRow([
      "Timestamp", 
      "Invoice Number", 
      "Client Name", 
      "Mobile Number",
      "Selected Favorites (IDs/Names)", 
      "Notes / Instructions",
      "Status"
    ]);
    favSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    favSheet.setFrozenRows(1);
  }
}

/**
 * Helper to extract Google Drive Folder ID from a URL.
 */
function extractFolderId(url) {
  if (!url) return null;
  const match = url.match(/folders\/([a-zA-Z0-9-_]+)/) || url.match(/id=([a-zA-Z0-9-_]+)/);
  return match ? match[1] : url.trim();
}

/**
 * Validates the client login credentials (Mobile and Invoice).
 * Returns the client session details and list of images from the Drive folder.
 */
function validateClient(mobileNumber, invoiceNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // Clean inputs
    const cleanMobile = String(mobileNumber).replace(/\s+/g, '').trim();
    const cleanInvoice = String(invoiceNumber).replace(/\s+/g, '').toUpperCase().trim();
    
    let matchedRow = null;
    
    // Search for client (skip headers)
    for (let i = 1; i < data.length; i++) {
      const rowMobile = String(data[i][1]).replace(/\s+/g, '').trim();
      const rowInvoice = String(data[i][2]).replace(/\s+/g, '').toUpperCase().trim();
      
      if (rowMobile === cleanMobile && rowInvoice === cleanInvoice) {
        matchedRow = {
          name: data[i][0],
          mobile: data[i][1],
          invoice: data[i][2],
          folderUrl: data[i][3],
          status: data[i][4],
          rowNum: i + 1
        };
        break;
      }
    }
    
    if (!matchedRow) {
      return { success: false, message: "البيانات غير صحيحة، يرجى التحقق من رقم الموبايل ورقم الفاتورة." };
    }
    
    // Parse Google Drive Folder
    const folderId = extractFolderId(matchedRow.folderUrl);
    if (!folderId) {
      return { success: false, message: "رابط المجلد الخاص بك غير صالح. يرجى التواصل مع المصور." };
    }
    
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      return { 
        success: false, 
        message: "تعذر الوصول إلى صورك. يرجى التأكد من أن مجلد جوجل درايف تم مشاركته كـ (Anyone with link can view)." 
      };
    }
    
    const files = folder.getFiles();
    const images = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      
      if (mime.indexOf('image/') === 0) {
        let thumbUrl = "";
        try {
          const thumb = file.getThumbnailLink();
          if (thumb) {
            thumbUrl = thumb.replace(/=s\d+$/, '=s1600');
          } else {
            thumbUrl = file.getUrl();
          }
        } catch (err) {
          thumbUrl = file.getUrl();
        }
        
        images.push({
          id: file.getId(),
          name: file.getName(),
          downloadUrl: file.getDownloadUrl() || file.getUrl(),
          viewUrl: file.getUrl(),
          thumbnail: thumbUrl,
          size: formatBytes(file.getSize())
        });
      }
    }
    
    images.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
    
    return {
      success: true,
      clientName: matchedRow.name,
      invoice: matchedRow.invoice,
      mobile: matchedRow.mobile,
      status: matchedRow.status,
      folderName: folder.getName(),
      images: images
    };
    
  } catch (err) {
    return { success: false, message: "حدث خطأ غير متوقع: " + err.toString() };
  }
}

/**
 * Formats bytes to readable file size.
 */
function formatBytes(bytes, decimals = 1) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

/**
 * Client submits their selected favorite images.
 */
function submitFavorites(invoiceNumber, clientName, mobileNumber, favoriteList, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FAVORITES_SHEET_NAME);
    
    const favCount = Array.isArray(favoriteList) ? favoriteList.length : 0;
    const favString = JSON.stringify(favoriteList);
    const timestamp = new Date();
    
    const data = sheet.getDataRange().getValues();
    const cleanInvoice = String(invoiceNumber).trim().toUpperCase();
    let rowUpdated = false;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === cleanInvoice) {
        sheet.getRange(i + 1, 1).setValue(timestamp);
        sheet.getRange(i + 1, 5).setValue(favString);
        sheet.getRange(i + 1, 6).setValue(notes);
        sheet.getRange(i + 1, 7).setValue("Updated");
        rowUpdated = true;
        break;
      }
    }
    
    if (!rowUpdated) {
      sheet.appendRow([
        timestamp,
        invoiceNumber,
        clientName,
        mobileNumber,
        favString,
        notes,
        "Pending Review"
      ]);
    }
    
    sendPhotographerEmail(clientName, invoiceNumber, favCount, notes, favoriteList);
    
    return { success: true, message: "تم حفظ الصور المختارة بنجاح وإرسال إشعار للمصور مصطفى خير الله!" };
  } catch (err) {
    return { success: false, message: "حدث خطأ أثناء حفظ الاختيارات: " + err.toString() };
  }
}

/**
 * Sends email alert to photographer.
 */
function sendPhotographerEmail(clientName, invoiceNumber, favCount, notes, favoriteList) {
  try {
    const ownerEmail = Session.getActiveUser().getEmail();
    if (!ownerEmail) return;
    
    const subject = `📸 صور مختارة جديدة: ${clientName} (فاتورة ${invoiceNumber})`;
    
    let htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px; direction: rtl; text-align: right;">
        <h2 style="color: #d4af37; border-bottom: 2px solid #f1f5f9; padding-bottom: 10px; margin-top: 0;">إشعار اختيار الصور 📸</h2>
        <p>مرحباً مصطفى، قام العميل <strong>${clientName}</strong> باختيار الصور المفضلة لديه.</p>
        
        <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #f1f5f9; font-weight: bold; width: 35%;">اسم العميل:</td>
            <td style="padding: 8px; border-bottom: 1px solid #f1f5f9;">${clientName}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #f1f5f9; font-weight: bold;">رقم الفاتورة:</td>
            <td style="padding: 8px; border-bottom: 1px solid #f1f5f9; color: #d4af37;">${invoiceNumber}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border-bottom: 1px solid #f1f5f9; font-weight: bold;">عدد الصور المختارة:</td>
            <td style="padding: 8px; border-bottom: 1px solid #f1f5f9; font-weight: bold; color: #10b981;">${favCount} صورة</td>
          </tr>
        </table>
        
        <div style="background-color: #f8fafc; border-right: 4px solid #d4af37; padding: 12px; margin: 15px 0; border-radius: 4px;">
          <h4 style="margin: 0 0 5px 0; color: #475569;">ملاحظات العميل / طلبات التعديل:</h4>
          <p style="margin: 0; color: #334155; font-style: italic;">${notes ? notes : "لا توجد ملاحظات إضافية"}</p>
        </div>
        
        <h4 style="margin-top: 20px; color: #475569;">قائمة الصور المختارة:</h4>
        <ul style="padding-right: 20px; color: #334155; line-height: 1.5;">
    `;
    
    favoriteList.forEach(img => {
      htmlBody += `<li><strong>${img.name}</strong></li>`;
    });
    
    htmlBody += `
        </ul>
        <p style="margin-top: 25px; font-size: 13px; color: #64748b; border-top: 1px solid #f1f5f9; padding-top: 15px;">
          تم إرسال هذا البريد تلقائياً من تطبيق نظام تسليم الصور الخاص بك.
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: ownerEmail,
      subject: subject,
      htmlBody: htmlBody
    });
  } catch (err) {
    Logger.log("Email notify error: " + err.toString());
  }
}

/* ==========================================================================
   ADMIN PANEL ACTIONS
   ========================================================================== */

/**
 * Validates admin passcode.
 */
function validateAdminPasscode(passcode) {
  return passcode === ADMIN_PASSCODE;
}

/**
 * Fetches dashboard details: all client rows, counts, and favorites data.
 */
function adminGetDashboardData(passcode) {
  if (!validateAdminPasscode(passcode)) {
    return { success: false, message: "رمز المرور غير صحيح!" };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Clients
    const clientSheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
    const clientsData = clientSheet.getDataRange().getValues();
    const clients = [];
    
    for (let i = 1; i < clientsData.length; i++) {
      clients.push({
        name: clientsData[i][0],
        mobile: clientsData[i][1],
        invoice: clientsData[i][2],
        folderUrl: clientsData[i][3],
        status: clientsData[i][4],
        createdAt: clientsData[i][5] ? new Date(clientsData[i][5]).toLocaleDateString() : ""
      });
    }
    
    // Favorites
    const favSheet = ss.getSheetByName(FAVORITES_SHEET_NAME);
    const favsData = favSheet.getDataRange().getValues();
    const favorites = [];
    
    for (let i = 1; i < favsData.length; i++) {
      let photoList = [];
      try {
        photoList = JSON.parse(favsData[i][4]);
      } catch (e) {
        photoList = [];
      }
      
      favorites.push({
        timestamp: favsData[i][0] ? new Date(favsData[i][0]).toLocaleString() : "",
        invoice: favsData[i][1],
        clientName: favsData[i][2],
        mobile: favsData[i][3],
        photosCount: photoList.length,
        photos: photoList,
        notes: favsData[i][5],
        status: favsData[i][6]
      });
    }
    
    return {
      success: true,
      clients: clients.reverse(),
      favorites: favorites.reverse(),
      stats: {
        totalClients: clients.length,
        totalFavorites: favorites.length,
        pendingDeliveries: clients.filter(c => c.status === "تحت التجهيز" || c.status === "Processing").length
      }
    };
    
  } catch (err) {
    return { success: false, message: "خطأ في جلب البيانات: " + err.toString() };
  }
}

/**
 * Creates a new client photoshoot delivery session.
 */
function adminCreateSession(passcode, clientName, mobileNumber, invoiceNumber, folderUrl, status) {
  if (!validateAdminPasscode(passcode)) {
    return { success: false, message: "رمز المرور غير صحيح!" };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    const cleanInvoice = String(invoiceNumber).trim().toUpperCase();
    const cleanMobile = String(mobileNumber).replace(/\s+/g, '').trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toUpperCase() === cleanInvoice) {
        return { success: false, message: "عذراً، رقم الفاتورة هذا مسجل بالفعل لعميل آخر." };
      }
    }
    
    sheet.appendRow([
      clientName.trim(),
      cleanMobile,
      cleanInvoice,
      folderUrl.trim(),
      status,
      new Date()
    ]);
    
    return { success: true, message: "تم إضافة العميل الجديد بنجاح وربطه بمجلد درايف!" };
    
  } catch (err) {
    return { success: false, message: "حدث خطأ أثناء إضافة العميل: " + err.toString() };
  }
}

/**
 * Deletes a client session.
 */
function adminDeleteSession(passcode, invoiceNumber) {
  if (!validateAdminPasscode(passcode)) {
    return { success: false, message: "رمز المرور غير صحيح!" };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CLIENTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const cleanInvoice = String(invoiceNumber).trim().toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toUpperCase() === cleanInvoice) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "تم حذف جلسة العميل بنجاح!" };
      }
    }
    
    return { success: false, message: "لم يتم العثور على العميل بهذا الرقم من الفاتورة." };
  } catch (err) {
    return { success: false, message: "حدث خطأ: " + err.toString() };
  }
}
