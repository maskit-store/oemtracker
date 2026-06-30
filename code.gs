/**
 * OEM TRACKER SYSTEM - BACKEND SYSTEM (Code.gs)
 * Deskripsi: Penanganan database Google Sheets, manajemen folder Google Drive,
 * autentikasi pengguna, dan proses unggah dokumen base64.
 * * Penggunaan: Salin seluruh isi berkas ini ke editor Google Apps Script Anda.
 */

const SPREADSHEET_ID = "1rKy-0w7etx1bPemgEIEJTDU849LPBtxcIevdXbv8JAM";
const DRIVE_FOLDER_ID = "1fcp5BtFHOvcabjG4xFTbZewwmRhzqESq";

/**
 * Menginisialisasi Web App dan merender file index.html
 */
function doGet() {
  try {
    // Inisialisasi awal database saat Web App diakses pertama kali
    initializeDatabase();
    
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle("OEM Tracker - PT SAINSGO KARYA INDONESIA")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(
      "<h3>Terjadi Error Sistem:</h3><p>" + error.toString() + "</p>" +
      "<p>Pastikan Anda telah memberikan izin akses Spreadsheet dan Google Drive saat deployment.</p>"
    );
  }
}

/**
 * Membuka Spreadsheet secara aman dengan proteksi try-catch khusus.
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error("Gagal mengakses Google Sheet. Pastikan ID Spreadsheet benar dan akun Apps Script Anda memiliki hak akses edit.");
  }
}

/**
 * Helper untuk mendapatkan Sheet secara aman. Jika Sheet tidak ditemukan,
 * sistem akan melakukan inisialisasi ulang otomatis.
 */
function getSheetSafely(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    initializeDatabase();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

/**
 * Menginisialisasi Tab Sheets & struktur kolom database secara dinamis.
 */
function initializeDatabase() {
  const ss = getSpreadsheet();
  
  // 1. Users Sheet
  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["ID", "Name", "Division", "Username", "Password", "Role"]);
    // Kredensial Admin Utama Default
    usersSheet.appendRow(["USR101", "Administrator", "Management", "admin", "admin123", "admin"]);
  }
  
  // 2. Customers Sheet
  let custSheet = ss.getSheetByName("Customers");
  if (!custSheet) {
    custSheet = ss.insertSheet("Customers");
    custSheet.appendRow([
      "ID", "Customer Name", "Phone", "Email", "Company Name", "Brand Name", "Address", 
      "Social Media", "KTP Link", "NPWP Link", "NIB Link", "HAKI Link", "MOU Link", 
      "BPOM Link", "Other Link", "Notes"
    ]);
  }
  
  // 3. Prospects Sheet
  let prosSheet = ss.getSheetByName("Prospects");
  if (!prosSheet) {
    prosSheet = ss.insertSheet("Prospects");
    prosSheet.appendRow(["ID", "Date", "Customer ID", "Customer Name", "Status", "Document Link"]);
  }
  
  // 4. Projects Sheet
  let projSheet = ss.getSheetByName("Projects");
  if (!projSheet) {
    projSheet = ss.insertSheet("Projects");
    const projHeaders = [
      "ID", "Customer ID", "Customer Name", "Product Category", "Product Name", "Qty", "Price", "Packaging Description",
      "Step1_Status", "Step1_Url", "Step2_Status", "Step2_Url", "Step3_Status", "Step3_Url", "Step4_Status", "Step4_Url",
      "Step5_Status", "Step5_Url", "Step6_Status", "Step6_Url", "Step7_Status", "Step7_Url", "Step8_Status", "Step8_Url",
      "Step9_Status", "Step9_Url", "Step10_Status", "Step10_Url", "Step11_Status", "Step11_Url", "Step12_Status", "Step12_Url",
      "OverallProgress", "DP_Payment_Date"
    ];
    projSheet.appendRow(projHeaders);
  } else {
    // Memastikan kolom ke-34 (DP_Payment_Date) ada jika sheet sudah terbentuk sebelumnya
    const lastCol = projSheet.getLastColumn();
    const headers = projSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf("DP_Payment_Date") === -1) {
      projSheet.getRange(1, lastCol + 1).setValue("DP_Payment_Date");
    }
  }
}

/**
 * Memverifikasi kredensial login pengguna.
 */
function loginUser(username, password) {
  try {
    const sheet = getSheetSafely("Users");
    const data = sheet.getDataRange().getValues();
    
    const cleanUsername = username.toString().trim().toLowerCase();
    const cleanPassword = password.toString().trim();
    
    for (let i = 1; i < data.length; i++) {
      const dbUser = data[i][3].toString().trim().toLowerCase();
      const dbPass = data[i][4].toString().trim();
      
      if (dbUser === cleanUsername && dbPass === cleanPassword) {
        return {
          success: true,
          user: {
            id: data[i][0],
            name: data[i][1],
            division: data[i][2],
            username: data[i][3],
            role: data[i][5]
          }
        };
      }
    }
    return { success: false, message: "Kombinasi Username atau Password salah!" };
  } catch (error) {
    return { success: false, message: "Kesalahan internal: " + error.toString() };
  }
}

/**
 * Mengambil daftar seluruh pengguna terdaftar.
 */
function getUsers() {
  try {
    const sheet = getSheetSafely("Users");
    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        id: data[i][0],
        name: data[i][1],
        division: data[i][2],
        username: data[i][3],
        role: data[i][5]
      });
    }
    return users;
  } catch (error) {
    return [];
  }
}

/**
 * Menyimpan atau memperbarui data pengguna baru/lama.
 */
function saveUser(user) {
  try {
    const sheet = getSheetSafely("Users");
    const data = sheet.getDataRange().getValues();
    
    // Periksa keunikan Username
    for (let i = 1; i < data.length; i++) {
      if (data[i][3].toString().toLowerCase() === user.username.toString().toLowerCase() && data[i][0] !== user.id) {
        return { success: false, message: "Username sudah digunakan oleh pengguna lain!" };
      }
    }
    
    if (user.id) {
      // Mode Edit
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === user.id) {
          sheet.getRange(i + 1, 2, 1, 5).setValues([[user.name, user.division, user.username, user.password, user.role]]);
          return { success: true };
        }
      }
    } else {
      // Mode Tambah Baru
      const newId = "USR" + (data.length + 100);
      sheet.appendRow([newId, user.name, user.division, user.username, user.password, user.role]);
      return { success: true };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Menghapus pengguna secara permanen.
 */
function deleteUser(id) {
  try {
    const sheet = getSheetSafely("Users");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        if (data[i][3].toString().toLowerCase() === "admin") {
          return { success: false, message: "Akun admin utama tidak dapat dihapus!" };
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ID Pengguna tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mengambil list Customer
 */
function getCustomers() {
  try {
    const sheet = getSheetSafely("Customers");
    const data = sheet.getDataRange().getValues();
    const customers = [];
    for (let i = 1; i < data.length; i++) {
      customers.push({
        id: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        email: data[i][3],
        company: data[i][4],
        brand: data[i][5],
        address: data[i][6],
        socialMedia: data[i][7],
        ktpLink: data[i][8],
        npwpLink: data[i][9],
        nibLink: data[i][10],
        hakiLink: data[i][11],
        mouLink: data[i][12],
        bpomLink: data[i][13],
        otherLink: data[i][14],
        notes: data[i][15]
      });
    }
    return customers;
  } catch (error) {
    return [];
  }
}

/**
 * Menyimpan data Customer baru/lama
 */
function saveCustomer(customer) {
  try {
    const sheet = getSheetSafely("Customers");
    const data = sheet.getDataRange().getValues();
    
    const rowValues = [
      customer.name, customer.phone, customer.email, customer.company, customer.brand, customer.address,
      customer.socialMedia, customer.ktpLink, customer.npwpLink, customer.nibLink, customer.hakiLink, 
      customer.mouLink, customer.bpomLink, customer.otherLink, customer.notes
    ];
    
    if (customer.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === customer.id) {
          sheet.getRange(i + 1, 2, 1, 15).setValues([rowValues]);
          return { success: true };
        }
      }
    } else {
      const newId = "CUST" + (data.length + 100);
      sheet.appendRow([newId, ...rowValues]);
      return { success: true };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Menghapus data Customer
 */
function deleteCustomer(id) {
  try {
    const sheet = getSheetSafely("Customers");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Data customer tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mengambil daftar Prospects
 */
function getProspects() {
  try {
    const sheet = getSheetSafely("Prospects");
    const data = sheet.getDataRange().getValues();
    const prospects = [];
    for (let i = 1; i < data.length; i++) {
      let dateVal = data[i][1];
      let formattedDate = "";
      if (dateVal) {
        try {
          formattedDate = Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch (e) {
          formattedDate = dateVal.toString();
        }
      }
      prospects.push({
        id: data[i][0],
        date: formattedDate,
        customerId: data[i][2],
        customerName: data[i][3],
        status: data[i][4],
        documentLink: data[i][5]
      });
    }
    return prospects;
  } catch (error) {
    return [];
  }
}

/**
 * Menyimpan data Prospect
 */
function saveProspect(prospect) {
  try {
    const sheet = getSheetSafely("Prospects");
    const data = sheet.getDataRange().getValues();
    
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    if (prospect.id) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === prospect.id) {
          sheet.getRange(i + 1, 2, 1, 5).setValues([[todayStr, prospect.customerId, prospect.customerName, prospect.status, prospect.documentLink]]);
          return { success: true };
        }
      }
    } else {
      const newId = "PRSP" + (data.length + 100);
      sheet.appendRow([newId, todayStr, prospect.customerId, prospect.customerName, prospect.status, prospect.documentLink]);
      return { success: true };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Menghapus data Prospect
 */
function deleteProspect(id) {
  try {
    const sheet = getSheetSafely("Prospects");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Prospect tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mengambil daftar Project OEM
 */
function getProjects() {
  try {
    const sheet = getSheetSafely("Projects");
    const data = sheet.getDataRange().getValues();
    const projects = [];
    
    for (let i = 1; i < data.length; i++) {
      let dpDateStr = "";
      if (data[i][33]) {
        try {
          dpDateStr = Utilities.formatDate(new Date(data[i][33]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch (e) {
          dpDateStr = data[i][33].toString();
        }
      }

      const proj = {
        id: data[i][0],
        customerId: data[i][1],
        customerName: data[i][2],
        productCategory: data[i][3],
        productName: data[i][4],
        qty: data[i][5],
        price: data[i][6],
        packagingDescription: data[i][7],
        steps: [],
        overallProgress: data[i][32],
        dpPaymentDate: dpDateStr
      };
      
      // Mengkonstruksi status dan url dari 12 tahapan OEM
      for (let s = 1; s <= 12; s++) {
        const statusIdx = 8 + (s - 1) * 2;
        const urlIdx = statusIdx + 1;
        proj.steps.push({
          status: data[i][statusIdx] || "Pending",
          url: data[i][urlIdx] || ""
        });
      }
      projects.push(proj);
    }
    return projects;
  } catch (error) {
    return [];
  }
}

/**
 * Menyimpan data Project OEM Baru
 */
function saveProject(project) {
  try {
    const sheet = getSheetSafely("Projects");
    const data = sheet.getDataRange().getValues();
    
    const newRow = [
      "PROJ" + (data.length + 100),
      project.customerId,
      project.customerName,
      project.productCategory,
      project.productName,
      project.qty,
      project.price,
      project.packagingDescription
    ];
    
    // Memberikan nilai default "Pending" dan "" untuk 12 langkah tahapan
    for (let i = 0; i < 12; i++) {
      newRow.push("Pending", "");
    }
    newRow.push(0); // Overall progress default (0%)
    newRow.push(""); // DP_Payment_Date default (kosong)
    
    sheet.appendRow(newRow);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mengubah tahapan/step spesifik pada Project OEM beserta kalkulasi % kemajuan otomatis.
 */
function updateProjectStep(projectId, stepIndex, status, docUrl) {
  try {
    const sheet = getSheetSafely("Projects");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        const statusCol = 9 + (stepIndex * 2); // Indeks kolom (1-based) untuk status
        const urlCol = statusCol + 1;
        
        sheet.getRange(i + 1, statusCol).setValue(status);
        sheet.getRange(i + 1, urlCol).setValue(docUrl);
        
        // Kalkulasi ulang presentase kemajuan secara real-time
        let completedSteps = 0;
        const rowData = sheet.getRange(i + 1, 1, 1, 34).getValues()[0];
        for (let s = 0; s < 12; s++) {
          const checkStatus = (s === stepIndex) ? status : rowData[8 + (s * 2)];
          if (checkStatus === "Completed") {
            completedSteps++;
          }
        }
        const progressPercent = Math.round((completedSteps / 12) * 100);
        sheet.getRange(i + 1, 33).setValue(progressPercent);
        
        return { success: true, progressPercent: progressPercent };
      }
    }
    return { success: false, message: "ID Project tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Memperbarui Tanggal Pembayaran DP secara spesifik di Google Sheets
 */
function updateProjectDpDate(projectId, dpDate) {
  try {
    const sheet = getSheetSafely("Projects");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        sheet.getRange(i + 1, 34).setValue(dpDate); // Mengisi kolom DP_Payment_Date (Kolom ke-34)
        return { success: true };
      }
    }
    return { success: false, message: "ID Project tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Menghapus Project OEM
 */
function deleteProject(id) {
  try {
    const sheet = getSheetSafely("Projects");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Project tidak ditemukan." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mengunggah berkas Base64 langsung ke Target Google Drive Folder secara aman.
 */
function uploadFileToDrive(base64Data, fileName) {
  try {
    if (!base64Data || !base64Data.includes(",")) {
      return { success: false, message: "Format file base64 tidak valid." };
    }
    
    let folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch (errFolder) {
      // Fallback jika Folder ID bermasalah / tidak bisa diakses, buat di Root
      folder = DriveApp.getRootFolder();
    }
    
    const splitData = base64Data.split(',');
    const contentTypeMatch = splitData[0].match(/:(.*?);/);
    const contentType = contentTypeMatch ? contentTypeMatch[1] : "application/octet-stream";
    const rawData = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(rawData, contentType, fileName);
    
    // 1. Pembuatan berkas utama
    const file = folder.createFile(blob);
    
    // 2. Pembungkusan terisolasi untuk pengaturan izin berbagi (Sharing)
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      Logger.log("Pemberitahuan: setSharing dilewati karena pembatasan organisasi: " + sharingError.toString());
    }
    
    // 3. Pembungkusan terisolasi untuk penarikan URL berkas
    let fileUrl = "";
    try {
      fileUrl = file.getUrl();
    } catch (urlError) {
      // Fallback url menggunakan ID jika getUrl() ditolak sistem organisasi
      fileUrl = "https://drive.google.com/open?id=" + file.getId();
    }
    
    return { success: true, url: fileUrl, name: fileName };
  } catch (error) {
    return { success: false, message: "Gagal mengunggah berkas ke Google Drive: " + error.toString() };
  }
}
