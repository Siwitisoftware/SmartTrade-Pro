/* ==============================================
SMARTTRADE PRO – SERVER API (VERSION 9.1)
ADDED: DEVELOPER DASHBOARD FOR USER MANAGEMENT
DEVELOPER EMAIL: daxsystems.tz.co@gmail.com
============================================== */

// Global constants
var LOGO_FOLDER_ID = "1CTMJcyrE5wiz4Kev0SZZfCQ36WGMrJy2";
var MASTER_DB_NAME = "SMARTTRADE_MASTER_DATABASE";
var LOGO_BASE_URL = "https://drive.google.com/thumbnail?id=";
var LOGO_ALT_URL = "https://drive.google.com/uc?export=view&id=";
var LOGO_DIRECT_URL = "https://lh3.googleusercontent.com/d/";
var DEFAULT_LOGO_URL = "https://i.imgur.com/XrJpD6d.png";
// Debt status constants
var DEBT_STATUS_PENDING = "pending";
var DEBT_STATUS_PAID = "paid";
var DEBT_STATUS_PARTIAL = "partial";
var DEBT_STATUS_DENIED = "denied";
var DEBT_PHOTOS_FOLDER_ID = "1CTMJcyrE5wiz4Kev0SZZfCQ36WGMrJy2";
var REPORTS_FOLDER_ID = "1O5qWUmfUNhjcrNc_OcN-KsrCgUX89fzx";

// DEVELOPER EMAIL - MUHIMU SANA
var DEVELOPER_EMAIL = "daxsystems.tz.co@gmail.com";

// DEFAULT UNITS - seeded into every business database's own UNITS sheet so
// that ALL units (including these defaults) live in the database, not just
// hardcoded in the app. Businesses can add their own on top of this list.
var DEFAULT_UNITS = [
  // Count / packaging
  "Pcs", "Piece", "Dozen", "Pair", "Set", "Each", "Unit",
  "Box", "Carton", "Crate", "Bag", "Sack", "Bundle", "Pack", "Packet",
  "Pallet", "Roll", "Reel", "Ream", "Sheet", "Bale", "Tray",
  "Tin", "Can", "Jar", "Bottle", "Tube", "Drum", "Barrel", "Jerrycan",
  "Bar", "Block", "Basket", "Sachet", "Container", "Bucket", "Basin",
  "Bunch", "Head", "Bin", "Hamper", "Punnet", "Bushel", "Stack",
  // Weight
  "Kg", "Gram", "Milligram", "Ton", "Quintal", "Pound (lb)", "Ounce (oz)",
  // Volume
  "Litre", "Millilitre", "Gallon", "Cubic Meter",
  // Length / area
  "Meter", "Centimeter", "Millimeter", "Yard", "Inch", "Foot", "Square Meter",
  // Services / miscellaneous business units
  "Trip", "Lumpsum", "Load", "Truckload", "Job", "Service", "Session",
  "Hour", "Day", "Batch", "Lot", "Consultation", "Order"
];

// DEFAULT CATEGORIES - seeded into every business database's own
// CATEGORIES sheet, same pattern as UNITS: all categories (including these
// defaults) live in the database and businesses can add their own on top.
var DEFAULT_CATEGORIES = [
  "Groceries", "Beverages", "Alcoholic Beverages", "Snacks", "Confectionery",
  "Bakery", "Dairy", "Meat & Poultry", "Fish & Seafood", "Fruits & Vegetables",
  "Frozen Foods", "Cooking Oil & Fats", "Grains & Cereals", "Spices & Seasoning",
  "Canned & Packaged Foods", "Baby Products", "Personal Care", "Cosmetics & Beauty",
  "Health & Wellness", "Pharmaceuticals", "Household Cleaning", "Detergents & Soaps",
  "Kitchenware", "Home Appliances", "Electronics", "Mobile Phones & Accessories",
  "Computers & Accessories", "Electrical Supplies", "Hardware & Tools",
  "Building Materials", "Plumbing Supplies", "Furniture", "Home Decor",
  "Bedding & Linen", "Clothing - Men", "Clothing - Women", "Clothing - Children",
  "Footwear", "Bags & Luggage", "Jewelry & Watches", "Textiles & Fabric",
  "Stationery", "Books & Educational", "Office Supplies", "Printing Supplies",
  "Automotive Parts", "Motor Oil & Lubricants", "Tyres", "Motorcycle Parts",
  "Agriculture Inputs", "Fertilizer", "Seeds", "Animal Feed", "Veterinary Supplies",
  "Construction Equipment", "Sporting Goods", "Toys & Games", "Gifts & Novelty",
  "Pet Supplies", "Garden & Outdoor", "Lighting", "Paint & Coatings",
  "Safety Equipment", "Packaging Materials", "Fuel & Gas", "Water & Ice",
  "Tobacco Products", "Salon & Barbershop Supplies", "Restaurant Supplies",
  "Services", "Other"
];


// email the client sends" model). TTL is long on purpose: business users may
// not reopen the app with internet for a while, and we don't want to force a
// re-login just because they were offline. The offline UI itself never calls
// the server, so none of this affects offline usage - it only guards the
// moment a request actually reaches doPost.
var BUSINESS_TIME_ZONE = "UTC";
var CURRENT_BUSINESS_TIME_ZONE = null;

function isValidBusinessTimeZone_(timeZone) {
  if (!timeZone || typeof timeZone !== "string") return false;
  try { Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss"); return true; } catch (e) { return false; }
}

function getCurrentBusinessTimeZone_() {
  return isValidBusinessTimeZone_(CURRENT_BUSINESS_TIME_ZONE) ? CURRENT_BUSINESS_TIME_ZONE : BUSINESS_TIME_ZONE;
}

function getBusinessTimeZoneFromSettings_(databaseId) {
  try {
    if (!databaseId) return null;
    const settings = getUserSettingsFromDB(databaseId);
    const saved = settings && settings.PROFILE && settings.PROFILE.time_zone ? settings.PROFILE.time_zone.value : "";
    return isValidBusinessTimeZone_(saved) ? saved : null;
  } catch (e) { return null; }
}

function ensureBusinessTimeZone_(databaseId, candidateTimeZone) {
  // SmartTrade is global. The client sends the device's real IANA timezone
  // on every authenticated request. Prefer that current timezone so an old
  // saved value (for example UTC) cannot make the local calendar date wrong.
  const tz = isValidBusinessTimeZone_(candidateTimeZone)
    ? candidateTimeZone
    : (isValidBusinessTimeZone_(CURRENT_BUSINESS_TIME_ZONE)
        ? CURRENT_BUSINESS_TIME_ZONE
        : BUSINESS_TIME_ZONE);
  try {
    const db = SpreadsheetApp.openById(databaseId);
    const settingsSheet = db.getSheetByName("SETTINGS");
    if (settingsSheet) updateSettingsKey(settingsSheet, "PROFILE", "time_zone", tz);
  } catch (e) {}
  CURRENT_BUSINESS_TIME_ZONE = tz;
  return tz;
}

function formatBusinessDate_(dateValue, pattern) {
  return Utilities.formatDate(new Date(dateValue), getCurrentBusinessTimeZone_(), pattern || "yyyy-MM-dd");
}

function formatBusinessDateTime_(dateValue) {
  return Utilities.formatDate(new Date(dateValue), getCurrentBusinessTimeZone_(), "dd/MM/yyyy HH:mm:ss");
}
var SESSION_SHEET_NAME = "SESSIONS";
var SESSION_TTL_DAYS = 60;

/* ============================================
   UNITS MANAGEMENT (per-business, stored in DB)
   ============================================ */

function ensureUnitsSheet(userDB) {
  var sheet = userDB.getSheetByName("UNITS");
  if (!sheet) {
    sheet = userDB.insertSheet("UNITS");
    sheet.appendRow(["Unit"]);
    sheet.getRange("A1:A1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    // Batch-write every default unit in ONE call instead of looping
    // appendRow() (each loop iteration was its own slow round-trip to
    // Sheets) - this is a big part of why creating a new business used
    // to take so long.
    if (DEFAULT_UNITS.length > 0) {
      var unitRows = DEFAULT_UNITS.map(function(u) { return [u]; });
      sheet.getRange(2, 1, unitRows.length, 1).setValues(unitRows);
    }
  }
  return sheet;
}

function getUnits(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    var userDB = SpreadsheetApp.openById(d.databaseId);
    var sheet = ensureUnitsSheet(userDB);
    var data = sheet.getDataRange().getValues();
    var units = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        units.push(String(data[i][0]).trim());
      }
    }
    return output({ success: true, units: units });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ============================================
   NON-STOCK PRODUCTS (services sold without stock tracking)
   Kept in their own sheet so search only ever shows names that
   genuinely belong to this business - not arbitrary text mined out
   of the sales log.
   ============================================ */

function ensureNonStockProductsSheet(userDB) {
  var sheet = userDB.getSheetByName("NON_STOCK_PRODUCTS");
  if (!sheet) {
    sheet = userDB.insertSheet("NON_STOCK_PRODUCTS");
    sheet.appendRow(["Name", "Unit", "LastUpdated"]);
    sheet.getRange("A1:C1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

function getNonStockProducts(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    var userDB = SpreadsheetApp.openById(d.databaseId);
    var sheet = ensureNonStockProductsSheet(userDB);
    var data = sheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        items.push({
          name: String(data[i][0]).trim(),
          unit: data[i][1] ? String(data[i][1]).trim() : ""
        });
      }
    }
    return output({ success: true, items: items });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Called whenever a non-stock sale is recorded (online sync or offline
// batch sync) - keeps the sheet current without needing a separate request.
function upsertNonStockProduct(userDB, name, unit) {
  try {
    name = (name || "").toString().trim();
    if (!name) return;
    var sheet = ensureNonStockProductsSheet(userDB);
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === name.toLowerCase()) {
        if (unit) {
          sheet.getRange(i + 1, 2).setValue(unit);
        }
        sheet.getRange(i + 1, 3).setValue(now);
        return;
      }
    }
    sheet.appendRow([name, unit || "", now]);
  } catch (error) {
    console.error("upsertNonStockProduct error:", error);
  }
}

function addNonStockProduct(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    var name = (d.name || "").trim();
    if (!name) {
      return output({ success: false, error: "Service/product name required" });
    }
    var userDB = SpreadsheetApp.openById(d.databaseId);
    upsertNonStockProduct(userDB, name, d.unit || "");
    return getNonStockProducts(d);
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function addUnit(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    var newUnit = (d.unit || "").trim();
    if (!newUnit) {
      return output({ success: false, error: "Unit name required" });
    }
    var userDB = SpreadsheetApp.openById(d.databaseId);
    var sheet = ensureUnitsSheet(userDB);
    var data = sheet.getDataRange().getValues();
    var exists = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === newUnit.toLowerCase()) {
        exists = true;
        break;
      }
    }
    if (!exists) {
      sheet.appendRow([newUnit]);
    }
    var units = [];
    var updated = sheet.getDataRange().getValues();
    for (var j = 1; j < updated.length; j++) {
      if (updated[j][0] && String(updated[j][0]).trim()) {
        units.push(String(updated[j][0]).trim());
      }
    }
    return output({ success: true, units: units, added: !exists });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ============================================
   CATEGORIES MANAGEMENT (per-business, stored in DB)
   Same pattern as Units: a large default list is seeded once, and any new
   category typed anywhere (Stock, Purchase, Purchase Request) gets saved
   automatically so it's searchable everywhere afterward.
   ============================================ */

function ensureCategoriesSheet(userDB) {
  var sheet = userDB.getSheetByName("CATEGORIES");
  if (!sheet) {
    sheet = userDB.insertSheet("CATEGORIES");
    sheet.appendRow(["Category"]);
    sheet.getRange("A1:A1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    // Batch-write every default category in ONE call instead of looping
    // appendRow() - same fix as ensureUnitsSheet(), same reason.
    if (DEFAULT_CATEGORIES.length > 0) {
      var categoryRows = DEFAULT_CATEGORIES.map(function(c) { return [c]; });
      sheet.getRange(2, 1, categoryRows.length, 1).setValues(categoryRows);
    }
  }
  return sheet;
}

function getCategories(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    var userDB = SpreadsheetApp.openById(d.databaseId);
    var sheet = ensureCategoriesSheet(userDB);
    var data = sheet.getDataRange().getValues();
    var categories = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        categories.push(String(data[i][0]).trim());
      }
    }
    return output({ success: true, categories: categories });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function addCategory(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    var newCategory = (d.category || "").trim();
    if (!newCategory) {
      return output({ success: false, error: "Category name required" });
    }
    var userDB = SpreadsheetApp.openById(d.databaseId);
    var sheet = ensureCategoriesSheet(userDB);
    var data = sheet.getDataRange().getValues();
    var exists = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === newCategory.toLowerCase()) {
        exists = true;
        break;
      }
    }
    if (!exists) {
      sheet.appendRow([newCategory]);
    }
    var categories = [];
    var updated = sheet.getDataRange().getValues();
    for (var j = 1; j < updated.length; j++) {
      if (updated[j][0] && String(updated[j][0]).trim()) {
        categories.push(String(updated[j][0]).trim());
      }
    }
    return output({ success: true, categories: categories, added: !exists });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


/* ===== doGet - Serves HTML Frontend ===== */
function doGet(e) {
  // Check if request is from API (has format=json parameter)
  const format = e?.parameter?.format;
  
  if (format === 'json') {
    // Return JSON for API calls
    return output({
      success: true,
      message: "SmartTrade Pro API is running",
      version: "9.0",
      features: ["Individual DB per user", "Logo management", "Offline support", "Stock validation", "Customers", "Suppliers", "Purchases", "Sellers", "Branches", "Multi-user", "Drive API only", "Developer Dashboard"]
    });
  }
  
  // Check if HTML file exists
  try {
    // Return HTML page for browser
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('SmartTrade Pro - Business Management System')
      .setFaviconUrl('https://i.imgur.com/XrJpD6d.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=yes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (htmlError) {
    // If Index.html doesn't exist, return JSON response
    return output({
      success: true,
      message: "SmartTrade Pro API is running. HTML frontend not found. Please add Index.html file.",
      version: "9.0",
      note: "To use the full application, create an Index.html file in this project."
    });
  }
}

/* ===== IDEMPOTENT API ENTRYPOINT =====
   Every mutating request may carry operationId. A committed operation is
   remembered in the business database, so a timeout/connection drop followed
   by Retry cannot execute the same business operation twice.
*/
function ensureOperationLog_(db) {
  let sh = db.getSheetByName("OPERATION_LOG");
  if (!sh) {
    sh = db.insertSheet("OPERATION_LOG");
    sh.appendRow(["Operation ID", "Action", "Status", "Response JSON", "Created At", "Completed At"]);
    sh.getRange("A1:F1").setFontWeight("bold");
  }
  return sh;
}

function findOperation_(sh, operationId) {
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(operationId)) {
      return { rowIndex: i + 1, status: values[i][2], response: values[i][3] || "" };
    }
  }
  return null;
}

function doPost(e) {
  let parsed = null;
  try {
    if (!e || !e.postData || !e.postData.contents) return output({success:false,error:"No request data"});
    parsed = JSON.parse(e.postData.contents);
  } catch (err) {
    return output({success:false,error:"Invalid request data"});
  }

  const action = parsed.action;
  const mutating = [
    "addStock","bulkAddStock","updateStock","deleteStockRow",
    "updateSale","deleteSaleRow",
    "addPurchase","updatePurchase","deletePurchaseRow",
    "addDebt","updateDebt","deleteDebt","returnDebt","denyDebt",
    "addExpense","addUnit","addCategory","addNonStockProduct",
    "saveDraftPurchaseRequest","submitPurchaseRequest","submitAdminPurchaseList","approvePurchaseRequest",
    "rejectPurchaseRequest","saveReviewDraft","deletePurchaseRequest","completePurchaseRequestItem",
    "updateSettings","updateProfile","changePassword","syncOfflineData"
  ];
  const opId = parsed.operationId;
  // addSale already has its own atomic Operation ID + Receipt No. protection.
  if (!mutating.includes(action) || !opId || action === "addSale") return doPostCore(e);

  let db = null;
  try {
    if (parsed.databaseId) db = SpreadsheetApp.openById(parsed.databaseId);
  } catch (_) {}
  if (!db) return doPostCore(e); // doPostCore will perform the authoritative session validation.

  const sh = ensureOperationLog_(db);
  const lock = LockService.getScriptLock();
  let operationRow = -1;
  try {
    lock.waitLock(15000);
    const existing = findOperation_(sh, opId);
    if (existing) {
      if (existing.status === "DONE" && existing.response) {
        return ContentService.createTextOutput(existing.response).setMimeType(ContentService.MimeType.JSON);
      }
      if (existing.status === "PROCESSING") {
        return output({success:false, retryable:true, operationInProgress:true, error:"This operation is already being processed. Please retry shortly."});
      }
    }
    sh.appendRow([opId, action, "PROCESSING", "", new Date(), ""]);
    operationRow = sh.getLastRow();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }

  const response = doPostCore(e);
  let body = null;
  try { body = JSON.parse(response.getContent()); } catch (_) {}
  try {
    if (operationRow > 1) {
      if (body && body.success) {
        sh.getRange(operationRow, 3).setValue("DONE");
        const storedResponse = action === "syncOfflineData"
          ? JSON.stringify({success:true, alreadyProcessed:true, operationId:opId})
          : response.getContent();
        sh.getRange(operationRow, 4).setValue(storedResponse);
        sh.getRange(operationRow, 6).setValue(new Date());
      } else {
        sh.deleteRow(operationRow);
      }
    }
  } catch (_) {}
  return response;
}

/* ===== doPost - Handles API Calls (Preserved) ===== */
function doPostCore(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return output({ success: false, error: "No request data" });
    }
    
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    // Only these actions can run without an existing session - they either
    // create one (login) or don't touch any account data (register, forgot).
    // linkDeviceToSession/pushScannedItemToSession are also here: a
    // phone/POS/tablet in "Remote Scanner" mode never logs in at all - it
    // only ever knows a short-lived pairing code (never the database ID or
    // account email directly), so there's nothing sensitive it could
    // impersonate even without a session token.
    const noAuthActions = ["register", "login", "forgot", "linkDeviceToSession", "pushScannedItemToSession", "removeScannedItemFromSession", "getConfirmedScanIds"];
    
    if (!noAuthActions.includes(action)) {
      const session = validateSession(data.token);
      if (!session.valid) {
        return output({
          success: false,
          error: "Your session has expired or is invalid. Please login again.",
          sessionExpired: true
        });
      }
      // Trust boundary: from here on, identity/database come from the
      // verified session - never from fields the client sent in the payload.
      // This is what stops someone from just typing another user's email (or
      // the developer's email) into a request to impersonate them.
      data.email = session.email;
      if (session.databaseId) {
        data.databaseId = session.databaseId;
        ensureBusinessTimeZone_(data.databaseId, data.clientTimeZone);
      }
    }
    
    // Handle syncOfflineData separately (now runs AFTER the session check
    // above, so offline data can only ever sync into the caller's own database)
    if (action === "syncOfflineData") {
      return processOfflineData(data);
    }
    
    // Public actions list (no SUBSCRIPTION check needed - these are exempt
    // from the expiry gate below, but they still had to pass the session
    // check above, except for register/login/forgot which need neither).
    // linkDeviceToSession/pushScannedItemToSession are here too: a phone/
    // POS/tablet in Remote Scanner mode never has a data.email at all (it
    // never logged in), so running the access/subscription check against
    // an undefined email would just break the feature outright.
    const publicActions = [
      "register", "login", "forgot", "checkUserAccess",
      "getAllUsers", "getAllUsersWithSubscription", "deleteUser",
      "updateUserBusiness", "resetUserPassword", "getUserDatabaseAccess",
      "updateUserSubscription", "getUserSubscription", "saveSubscriptionToLocal",
      "checkSubscriptionOffline", "linkDeviceToSession", "pushScannedItemToSession",
      "removeScannedItemFromSession", "getConfirmedScanIds"
    ];
    
    // Developer bypass - safe now because data.email is verified, not client-supplied
    const isDeveloper = (data.email === DEVELOPER_EMAIL);
    
    // Check access for non-public, non-developer actions
    if (!isDeveloper && !publicActions.includes(action)) {
      const accessCheck = checkUserAccessOnServer(data.email);
      if (!accessCheck.allowed) {
        return output({ 
          success: false, 
          error: accessCheck.message,
          accessDenied: true,
          isExpired: accessCheck.isExpired || false
        });
      }
    }
    
    // Execute action based on type
    switch(action){
      // AUTHENTICATION ACTIONS
      case "register": return registerUser(data);
      case "login": return loginUser(data);
      case "getMyBusinesses": return getMyBusinesses(data);
      case "addBusiness": return addBusiness(data);
      case "switchBusiness": return switchBusiness(data);
      case "forgot": return forgotPassword(data);
      case "logout": return logoutUser(data);
      
      // ACCOUNTING PERIOD ACTIONS
      case "previewAccountingPeriod": return previewAccountingPeriod(data);
      case "closeAccountingPeriod": return closeAccountingPeriod(data);
      case "getClosingReports": return getClosingReports(data);
      case "getExpensesData": return getExpensesData(data);
      case "addExpense": return addExpense(data);
      
      // STOCK ACTIONS
      case "addStock": return addStock(data);
      case "bulkAddStock": return bulkAddStock(data);
      case "createDeviceLinkSession": return createDeviceLinkSession(data);
      case "linkDeviceToSession": return linkDeviceToSession(data);
      case "pushScannedItemToSession": return pushScannedItemToSession(data);
      case "pollScannedItems": return pollScannedItems(data);
      case "acknowledgeScannedItems": return acknowledgeScannedItems(data);
      case "removeScannedItemFromSession": return removeScannedItemFromSession(data);
      case "acknowledgeRemovals": return acknowledgeRemovals(data);
      case "getConfirmedScanIds": return getConfirmedScanIds(data);
      case "endDeviceLinkSession": return endDeviceLinkSession(data);
      case "getStock": return getStock(data);
      case "updateStock": return updateStock(data);
      case "deleteStockRow": return deleteStockRow(data);
      
      // SALES ACTIONS
      case "addSale": return addSale(data);
      case "getSales": return getSales(data);
      case "updateSale": return updateSale(data);
      case "deleteSaleRow": return deleteSaleRow(data);
      
      // UNITS ACTIONS
      case "getUnits": return getUnits(data);
      case "addUnit": return addUnit(data);
      
      // CATEGORIES ACTIONS
      case "getCategories": return getCategories(data);
      case "addCategory": return addCategory(data);
      
      // NON-STOCK PRODUCTS ACTIONS
      case "getNonStockProducts": return getNonStockProducts(data);
      case "addNonStockProduct": return addNonStockProduct(data);
      
      // PURCHASES ACTIONS
      case "addPurchase": return addPurchase(data);
      case "getPurchases": return getPurchases(data);
      case "updatePurchase": return updatePurchase(data);
      case "deletePurchaseRow": return deletePurchaseRow(data);
      
      // PURCHASE REQUEST ACTIONS (Seller -> Admin approval workflow)
      case "saveDraftPurchaseRequest": return saveDraftPurchaseRequest(data);
      case "submitPurchaseRequest": return submitPurchaseRequest(data);
      case "submitAdminPurchaseList": return submitAdminPurchaseList(data);
      case "getPurchaseRequests": return getPurchaseRequests(data);
      case "approvePurchaseRequest": return approvePurchaseRequest(data);
      case "rejectPurchaseRequest": return rejectPurchaseRequest(data);
      case "saveReviewDraft": return saveReviewDraft(data);
      case "deletePurchaseRequest": return deletePurchaseRequest(data);
      case "completePurchaseRequestItem": return completePurchaseRequestItem(data);
      
      // REPORT ACTIONS
      case "getReport": return getReport(data);
      case "getDashboard": return getDashboard(data);
      case "generateReport": return generateReport(data);
      case "generateFilteredReport": return generateFilteredReport(data);
      
      // SETTINGS ACTIONS
      case "getSettings": return getUserSettings(data);
      case "updateSettings": return updateUserSettings(data);
      case "updateProfile": return updateUserProfile(data);
      case "changePassword": return changePassword(data);
      case "deleteAccount": return deleteAccount(data);
      
      // LOGO ACTIONS
      case "uploadLogo": return uploadUserLogo(data);
      case "getLogo": return getUserLogo(data);
      case "deleteLogo": return deleteUserLogo(data);
      case "uploadLogoBase64": return uploadUserLogoBase64(data);
      case "deleteLogoBase64": return deleteUserLogoBase64(data);
      
      // SYNC ACTIONS
      case "syncAllData": return syncAllUserData(data);
      case "syncOfflineData": return processOfflineData(data);
      
      // PRODUCT ACTIONS
      case "getProducts": return getProductsList(data);
      case "checkStock": return checkProductStock(data);
      
      // DEBT ACTIONS
      case "addDebt": return addDebt(data);
      case "getDebts": return getDebts(data);
      case "updateDebt": return updateDebt(data);
      case "deleteDebt": return deleteDebt(data);
      case "returnDebt": return returnDebt(data);
      case "denyDebt": return denyDebt(data);
      case "checkAndDenyExpiredDebts": return checkAndDenyExpiredDebts(data);
      case "uploadBorrowerPhoto": return uploadBorrowerPhoto(data);
      
      // CUSTOMER ACTIONS
      case "addCustomer": return addCustomer(data);
      case "getCustomers": return getCustomers(data);
      case "updateCustomer": return updateCustomer(data);
      case "deleteCustomer": return deleteCustomer(data);
      
      // SUPPLIER ACTIONS
      case "addSupplier": return addSupplier(data);
      case "getSuppliers": return getSuppliers(data);
      case "updateSupplier": return updateSupplier(data);
      case "deleteSupplier": return deleteSupplier(data);
      
      // SELLER ACTIONS
      case "addSeller": return addSeller(data);
      case "getSellers": return getSellers(data);
      case "setSellerPin": return setSellerPin(data);
      case "resetSellerPin": return resetSellerPin(data);
      case "verifySellerPin": return verifySellerPin(data);
      case "getMySellerPinStatus": return getMySellerPinStatus(data);
      case "setAdminPin": return setAdminPin(data);
      case "getMyAdminPinStatus": return getMyAdminPinStatus(data);
      case "setModalRestrictions": return setModalRestrictions(data);
      case "getModalRestrictions": return getModalRestrictions(data);
      case "updateSeller": return updateSeller(data);
      case "deleteSeller": return deleteSeller(data);
      case "getSellerSales": return getSellerSales(data);
      case "getSellerPurchases": return getSellerPurchases(data);
      
      // BRANCH ACTIONS
      case "addBranch": return addBranch(data);
      case "getBranches": return getBranches(data);
      case "updateBranch": return updateBranch(data);
      case "deleteBranch": return deleteBranch(data);
      
      // MIGRATION ACTIONS
      case "ensureSheetsExist": return ensureSheetsExist(data);
      
      // DEVELOPER DASHBOARD ACTIONS
      case "getAllUsers": return getAllUsers(data);
      case "getAllUsersWithSubscription": return getAllUsersWithSubscription(data);
      case "deleteUser": return deleteUser(data);
      case "updateUserBusiness": return updateUserBusiness(data);
      case "resetUserPassword": return resetUserPassword(data);
      case "getUserDatabaseAccess": return getUserDatabaseAccess(data);
      
      // SUBSCRIPTION ACTIONS
      case "updateUserSubscription": return updateUserSubscription(data);
      case "getUserSubscription": return getUserSubscription(data);
      case "checkUserAccess": return checkUserAccess(data);
      case "saveSubscriptionToLocal": return saveSubscriptionToLocal(data);
      case "checkSubscriptionOffline": return checkSubscriptionOffline(data);
      // Katika doPost, ongeza:
case "getDebtHistory": return getDebtHistory(data);
case "getBorrowerDebtHistory": return getBorrowerDebtHistory(data);
      
      default:
        return output({ success: false, error: "Invalid action: " + action });
    }
    
  } catch(err){
    return output({ 
      success: false, 
      error: err.toString(),
      stack: err.stack 
    });
  }
}

/* ===== JSON OUTPUT Helper ===== */
function output(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// DEVELOPER DASHBOARD FUNCTIONS
// ============================================

function getAllUsers(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    const users = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) { // Email exists
        users.push({
          business: data[i][0] || "",
          email: data[i][1],
          password: "********",
          databaseId: data[i][3] || "",
          created: data[i][4] ? data[i][4].toString() : null,
          lastLogin: data[i][5] ? data[i][5].toString() : null,
          logoId: data[i][6] || null,
          userType: data[i][7] || "admin",
          rowIndex: i + 1
        });
      }
    }
    
    // Sort by created date (newest first)
    users.sort((a, b) => {
      const dateA = a.created ? new Date(a.created) : new Date(0);
      const dateB = b.created ? new Date(b.created) : new Date(0);
      return dateB - dateA;
    });
    
    return output({
      success: true,
      users: users,
      count: users.length,
      developerEmail: DEVELOPER_EMAIL
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function deleteUser(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    if (!d.targetEmail) {
      return output({ success: false, error: "Target email required" });
    }
    
    // Cannot delete developer account
    if (d.targetEmail === DEVELOPER_EMAIL) {
      return output({ success: false, error: "Cannot delete developer account" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    let targetRow = -1;
    let databaseId = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.targetEmail) {
        targetRow = i + 1;
        databaseId = data[i][3];
        break;
      }
    }
    
    if (targetRow === -1) {
      return output({ success: false, error: "User not found" });
    }
    
    // Delete user's spreadsheet if exists
    if (databaseId) {
      try {
        Drive.Files.remove(databaseId);
      } catch (e) {
        console.log("Could not delete user spreadsheet:", e);
      }
    }
    
    // Delete user from master sheet
    masterSheet.deleteRow(targetRow);
    
    return output({
      success: true,
      message: "User deleted successfully"
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function updateUserBusiness(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    if (!d.targetEmail || !d.businessName) {
      return output({ success: false, error: "Target email and business name required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.targetEmail) {
        masterSheet.getRange(i + 1, 1).setValue(d.businessName);
        
        // Also update in user's own settings
        const databaseId = data[i][3];
        if (databaseId) {
          try {
            const userDB = SpreadsheetApp.openById(databaseId);
            const settingsSheet = userDB.getSheetByName("SETTINGS");
            if (settingsSheet) {
              const settingsData = settingsSheet.getDataRange().getValues();
              let found = false;
              for (let j = 1; j < settingsData.length; j++) {
                if (settingsData[j][0] === "PROFILE" && settingsData[j][1] === "business_name") {
                  settingsSheet.getRange(j + 1, 3).setValue(d.businessName);
                  settingsSheet.getRange(j + 1, 4).setValue(new Date());
                  found = true;
                  break;
                }
              }
              if (!found) {
                settingsSheet.appendRow(["PROFILE", "business_name", d.businessName, new Date()]);
              }
            }
          } catch (e) {
            console.log("Could not update user settings:", e);
          }
        }
        
        return output({
          success: true,
          message: "Business name updated successfully"
        });
      }
    }
    
    return output({ success: false, error: "User not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function resetUserPassword(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    if (!d.targetEmail || !d.newPassword) {
      return output({ success: false, error: "Target email and new password required" });
    }
    
    if (d.newPassword.length < 6) {
      return output({ success: false, error: "Password must be at least 6 characters" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.targetEmail) {
        masterSheet.getRange(i + 1, 3).setValue(d.newPassword);
        
        // Also update password for sellers if this is a seller account
        const databaseId = data[i][3];
        const userType = data[i][7] || "admin";
        
        if (databaseId && userType !== "admin") {
          try {
            const userDB = SpreadsheetApp.openById(databaseId);
            const sellersSheet = userDB.getSheetByName("SELLERS");
            if (sellersSheet) {
              const sellersData = sellersSheet.getDataRange().getValues();
              for (let j = 1; j < sellersData.length; j++) {
                if (sellersData[j][3] === d.targetEmail) {
                  sellersSheet.getRange(j + 1, 5).setValue(d.newPassword);
                  break;
                }
              }
            }
          } catch (e) {
            console.log("Could not update seller password:", e);
          }
        }
        
        return output({
          success: true,
          message: "Password reset successfully"
        });
      }
    }
    
    return output({ success: false, error: "User not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== GET USER LOGO ID FROM MASTER DATABASE ===== */
function getUserLogoIdFromMaster(email, databaseId) {
  try {
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    // Match by email AND databaseId when a databaseId is provided - the
    // Logo is per-BUSINESS (each row in MASTER_USERS has its own Logo ID
    // column), so matching by email alone would always return whichever
    // business was created FIRST instead of the one currently active.
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email && (!databaseId || data[i][3] === databaseId)) {
        // Column 7 (index 6) stores the Logo ID
        const logoId = data[i][6];
        if (logoId && logoId !== "") {
          console.log("Found logo ID for user:", email, "->", logoId);
          return logoId;
        } else {
          console.log("No logo ID found for user:", email);
          return null;
        }
      }
    }
    
    console.log("User not found in master database:", email);
    return null;
  } catch (error) {
    console.error("Error getting logo ID from master:", error.toString());
    return null;
  }
}

function getUserDatabaseAccess(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    if (!d.targetEmail) {
      return output({ success: false, error: "Target email required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.targetEmail) {
        const databaseId = data[i][3];
        if (databaseId) {
          try {
            const file = Drive.Files.get(databaseId);
            return output({
              success: true,
              databaseId: databaseId,
              databaseName: file.title,
              webViewLink: "https://docs.google.com/spreadsheets/d/" + databaseId + "/edit",
              canAccess: true
            });
          } catch (e) {
            return output({
              success: true,
              databaseId: databaseId,
              canAccess: false,
              error: "Cannot access file"
            });
          }
        }
        return output({
          success: false,
          error: "No database found for this user"
        });
      }
    }
    
    return output({ success: false, error: "User not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


function ensureSheetsExist(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const results = {
      customers: false,
      suppliers: false,
      sellers: false,
      branches: false,
      debts: false,
      units: false,
      categories: false,
      nonStockProducts: false,
      purchaseRequests: false
    };
    
    // Check and create CUSTOMERS sheet
    let customersSheet = userDB.getSheetByName("CUSTOMERS");
    if (!customersSheet) {
      customersSheet = userDB.insertSheet("CUSTOMERS");
      customersSheet.appendRow(["ID", "Name", "Phone", "Email", "Address", "Total Purchases", "Last Purchase", "Notes", "Created"]);
      customersSheet.getRange("A1:I1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
      results.customers = true;
    }
    
    // Check and create SUPPLIERS sheet
    let suppliersSheet = userDB.getSheetByName("SUPPLIERS");
    if (!suppliersSheet) {
      suppliersSheet = userDB.insertSheet("SUPPLIERS");
      suppliersSheet.appendRow(["ID", "Name", "Contact Person", "Phone", "Email", "Address", "Products Supplied", "Total Supplied", "Last Supply", "Notes", "Created"]);
      suppliersSheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
      results.suppliers = true;
    }
    
    // Check and create SELLERS sheet
    let sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      sellersSheet = userDB.insertSheet("SELLERS");
      sellersSheet.appendRow(["ID", "Name", "Phone", "Email", "Password", "Total Sales", "Total Purchases", "Last Active", "Created", "Status", "Branch ID"]);
      sellersSheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
      results.sellers = true;
    }
    
    // Check and create BRANCHES sheet
    let branchesSheet = userDB.getSheetByName("BRANCHES");
    if (!branchesSheet) {
      branchesSheet = userDB.insertSheet("BRANCHES");
      branchesSheet.appendRow(["ID", "Name", "Location", "Manager", "Phone", "Created", "Status"]);
      branchesSheet.getRange("A1:G1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
      branchesSheet.appendRow(["MAIN", "Main Branch", "Head Office", "", "", new Date(), "active"]);
      results.branches = true;
    }
    
    // Check and create DEBTS sheet - NO DUE DATE
    let debtsSheet = userDB.getSheetByName("DEBTS");
    if (!debtsSheet) {
      debtsSheet = userDB.insertSheet("DEBTS");
      debtsSheet.appendRow(["ID", "Date", "Borrower Name", "Borrower Phone", "Borrower NIDA", "Photo ID", "Product", "Quantity", "Unit Price", "Total Amount", "Paid Amount", "Status", "Branch ID", "Seller ID", "Notes", "Last Updated", "Unit", "Photo Base64", "Photo MIME Type"]);
      debtsSheet.getRange("A1:S1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
      results.debts = true;
    }
    
    // Check and create UNITS sheet (older businesses created before this
    // feature existed won't have one yet - seed it with the defaults now)
    if (!userDB.getSheetByName("UNITS")) {
      ensureUnitsSheet(userDB);
      results.units = true;
    }
    
    // Check and create CATEGORIES sheet (older businesses too)
    if (!userDB.getSheetByName("CATEGORIES")) {
      ensureCategoriesSheet(userDB);
      results.categories = true;
    }
    
    // Check and create NON_STOCK_PRODUCTS sheet (older businesses too)
    if (!userDB.getSheetByName("NON_STOCK_PRODUCTS")) {
      ensureNonStockProductsSheet(userDB);
      results.nonStockProducts = true;
    }
    
    // Check and create PURCHASE_REQUESTS sheet (older businesses too)
    if (!userDB.getSheetByName("PURCHASE_REQUESTS")) {
      ensurePurchaseRequestsSheet(userDB);
      results.purchaseRequests = true;
    }
    
    return output({
      success: true,
      message: "Sheets checked and created if missing",
      created: results
    });
    
  } catch (error) {
    return output({ 
      success: false, 
      error: "Failed to ensure sheets: " + error.toString() 
    });
  }
}

/* ===== GET PRODUCTS LIST (BY BRANCH) ===== */
function getProductsList(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const branchId = d.branchId || "MAIN";
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    const data = stockSheet.getDataRange().getValues();
    const products = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        const productName = data[i][1];
        const itemBranchId = data[i][8] || "MAIN"; // Branch ID column
        
        // Only include products from the selected branch
        if (itemBranchId !== branchId) {
          continue;
        }
        
        const quantity = data[i][2] || 0;
        const sellPrice = data[i][4] || 0;
        const unit = data[i][9] || "";
        
        if (!products[productName]) {
          products[productName] = {
            quantity: quantity,
            sellPrice: sellPrice,
            buyPrice: data[i][3] || 0,
            category: data[i][6] || "",
            unit: unit
          };
        } else {
          // Aggregate quantities for the same product in this branch
          products[productName].quantity += quantity;
          products[productName].sellPrice = sellPrice || products[productName].sellPrice;
          if (unit) products[productName].unit = unit;
        }
      }
    }
    
    return output({
      success: true,
      products: products,
      branch: branchId
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== CHECK PRODUCT STOCK (BY BRANCH) ===== */
function checkProductStock(d) {
  try {
    if (!d.databaseId || !d.product) {
      return output({ success: false, error: "Database ID and product required" });
    }
    
    const branchId = d.branchId || "MAIN"; // Default to MAIN branch
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    const data = stockSheet.getDataRange().getValues();
    let totalQuantity = 0;
    let sellPrice = 0;
    let unit = "";
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1] === d.product) {
        const itemBranchId = data[i][8] || "MAIN"; // Branch ID column
        
        // Only count stock from the specified branch
        if (itemBranchId === branchId) {
          totalQuantity += (data[i][2] || 0);
          sellPrice = data[i][4] || 0; // Get latest sell price
          if (data[i][9]) unit = data[i][9];
        }
      }
    }
    
    return output({
      success: true,
      product: d.product,
      branch: branchId,
      available: totalQuantity,
      sellPrice: sellPrice,
      unit: unit
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== SYNC ALL USER DATA ===== */
function syncAllUserData(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    ensureSheetsExist(d);
    
    const settings = getUserSettingsFromDB(d.databaseId);
    const businessTimeZone = ensureBusinessTimeZone_(d.databaseId, d.clientTimeZone);
    if (!settings.PROFILE) settings.PROFILE = {};
    settings.PROFILE.time_zone = { value: businessTimeZone, updated: new Date() };
    let logoUrl = DEFAULT_LOGO_URL;
    let userType = d.userType || "admin";
    let branchId = d.branchId || null;
    let sellerId = d.sellerId || null;
    
    // Pata business name kutoka master database
    let businessName = "";
    try {
      const masterDB = getMasterDatabase();
      const masterSheet = masterDB.getSheetByName("MASTER_USERS");
      const masterData = masterSheet.getDataRange().getValues();
      
      for (let i = 1; i < masterData.length; i++) {
        if (masterData[i][1] === d.email && masterData[i][3] === d.databaseId) {
          businessName = masterData[i][0];
          break;
        }
      }
    } catch (e) {
      console.log("Error reading master database:", e);
    }
    
    // Logo: prefer the base64 copy stored right in this business's own
    // Settings (no external Drive fetch needed, so it can never silently
    // fail to load) - fall back to the older Drive-based logo for
    // accounts that haven't re-uploaded since this feature was added.
    const logoInfo = getLogoBase64FromSettings(settings);
    if (logoInfo) {
      logoUrl = "data:" + logoInfo.mimeType + ";base64," + logoInfo.base64;
    } else {
      const logoId = getUserLogoIdFromMaster(d.email, d.databaseId);
      if (logoId) {
        logoUrl = LOGO_BASE_URL + logoId;
      }
    }
    
    // Get all data with branch filtering for sellers
    const stockData = getStockDataForSync(d.databaseId, branchId);
    const salesData = getSalesDataForSync(d.databaseId, branchId);
    const purchasesData = getPurchasesDataForSync(d.databaseId, branchId);
    const productsData = getProductsForSync(d.databaseId, branchId);
    const debtsData = getDebtsDataForSync(d.databaseId, sellerId, branchId);
    const unitsData = getUnitsListForSync(d.databaseId);
    const categoriesData = getCategoriesListForSync(d.databaseId);
    const nonStockProductsData = getNonStockProductsListForSync(d.databaseId);
    const purchaseRequestsData = getPurchaseRequestsForSync(d.databaseId);
    
    // Get customers, suppliers, sellers for admin only (private business
    // contact data) - but branch NAMES aren't sensitive and Sellers need
    // them too, to show "at <Branch Name>" instead of a raw branch ID
    // anywhere in their own UI (Recent Activity, purchase records, etc.).
    let customersData = [];
    let suppliersData = [];
    let sellersData = [];
    let branchesData = getBranchesDataForSync(d.databaseId);
    
    if (userType === "admin") {
      customersData = getCustomersDataForSync(d.databaseId);
      suppliersData = getSuppliersDataForSync(d.databaseId);
      sellersData = getSellersDataForSync(d.databaseId);
    }
    
    return output({
      success: true,
      lastSync: new Date().toISOString(),
      settings: settings,
      logoUrl: logoUrl,
      userType: userType,
      branchId: branchId,
      branchName: userType === "seller" ? getBranchNameById(d.databaseId, branchId) : null,
      businessName: businessName, // MUHIMU: Tuma business name
      data: {
        stock: stockData,
        sales: salesData,
        purchases: purchasesData,
        products: productsData,
        customers: customersData,
        suppliers: suppliersData,
        sellers: sellersData,
        branches: branchesData,
        debts: debtsData,
        units: unitsData,
        categories: categoriesData,
        nonStockProducts: nonStockProductsData,
        purchaseRequests: purchaseRequestsData
      }
    });
    
  } catch (error) {
    return output({ 
      success: false, 
      error: "Failed to sync data: " + error.toString() 
    });
  }
}

/* ===== GET STOCK DATA FOR SYNC (by branch) ===== */
function getUnitsListForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sheet = ensureUnitsSheet(userDB);
    const data = sheet.getDataRange().getValues();
    const units = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        units.push(String(data[i][0]).trim());
      }
    }
    return units;
  } catch (error) {
    return [];
  }
}

function getCategoriesListForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sheet = ensureCategoriesSheet(userDB);
    const data = sheet.getDataRange().getValues();
    const categories = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        categories.push(String(data[i][0]).trim());
      }
    }
    return categories;
  } catch (error) {
    return [];
  }
}

function getNonStockProductsListForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sheet = ensureNonStockProductsSheet(userDB);
    const data = sheet.getDataRange().getValues();
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        items.push({
          name: String(data[i][0]).trim(),
          unit: data[i][1] ? String(data[i][1]).trim() : ""
        });
      }
    }
    return items;
  } catch (error) {
    return [];
  }
}

// All purchase requests get bundled into the regular sync so the Purchase
// tabs load instantly from local cache (like Stock/Sales) and keep working
// offline - the small per-business row count makes sending everything
// simpler and safer than trying to filter server-side per role/tab.
function getPurchaseRequestsForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const data = sheet.getDataRange().getValues();
    const requests = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        requests.push(parseRequestRow(data[i]));
      }
    }
    requests.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
    return requests;
  } catch (error) {
    return [];
  }
}

function getStockDataForSync(databaseId, branchId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) return [];
    
    const data = stockSheet.getDataRange().getValues();
    const stockData = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const itemBranchId = data[i][8] || "MAIN"; // Branch ID column
        
        // Filter by branch if seller
        if (branchId && itemBranchId !== branchId) {
          continue;
        }
        
        stockData.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          buyPrice: data[i][3],
          sellPrice: data[i][4],
          value: data[i][5],
          category: data[i][6],
          sellerId: data[i][7] || null,
          branchId: itemBranchId, // Include branch ID
          unit: data[i][9] || "",
          id: `stock_${i}_${Date.now()}`
        });
      }
    }
    
    return stockData;
  } catch (error) {
    return [];
  }
}

/* ===== OPERATIONAL PERIOD BOUNDARIES (READ ONCE PER REQUEST) ===== */
function getOperationalOpenPeriodStarts_(userDB) {
  const result = { ALL: null, branches: {} };
  try {
    const sheet = userDB.getSheetByName("ACCOUNTING_PERIODS");
    if (!sheet) return result;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] !== "CLOSED") continue;
      const branch = data[i][3] || "ALL";
      const closedAt = new Date(data[i][5]).getTime();
      if (isNaN(closedAt)) continue;
      if (branch === "ALL") {
        if (result.ALL === null || closedAt > result.ALL) result.ALL = closedAt;
      } else if (result.branches[branch] === undefined || result.branches[branch] === null || closedAt > result.branches[branch]) {
        result.branches[branch] = closedAt;
      }
    }
    return result;
  } catch (error) {
    throw new Error("Unable to read accounting close boundaries: " + error.toString());
  }
}

function getOperationalOpenPeriodStartFromMap_(boundaries, branchId) {
  const branch = branchId || "MAIN";
  const branchClose = boundaries.branches[branch] !== undefined ? boundaries.branches[branch] : null;
  const allClose = boundaries.ALL !== undefined ? boundaries.ALL : null;
  if (branchClose === null && allClose === null) return null;
  if (branchClose === null) return allClose;
  if (allClose === null) return branchClose;
  return Math.max(branchClose, allClose);
}

/* ===== GET SALES DATA FOR SYNC (by branch) ===== */
function getSalesDataForSync(databaseId, branchId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const salesSheet = userDB.getSheetByName("SALES");
    const sellersSheet = userDB.getSheetByName("SELLERS");
    
    if (!salesSheet) return [];
    
    const sellerMap = {};
    if (sellersSheet) {
      const sellersData = sellersSheet.getDataRange().getValues();
      for (let i = 1; i < sellersData.length; i++) {
        if (sellersData[i][0]) {
          sellerMap[sellersData[i][0]] = sellersData[i][1];
        }
      }
    }
    
    ensureSalesReceiptNumbers_(salesSheet);
    const data = salesSheet.getDataRange().getValues();
    const operationalBoundaries = getOperationalOpenPeriodStarts_(userDB);

    const salesData = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const recordBranchId = data[i][8] || "MAIN"; // Branch ID column
        
        // Filter by branch if seller
        if (branchId && recordBranchId !== branchId) {
          continue;
        }

        // Sync must return ONLY the current open-period sales. This is
        // important because syncAllData can repopulate the local cache
        // after Close Accounts; returning historical rows here would
        // immediately make closed receipts reappear in All Sales.
        const saleTime = new Date(data[i][0]).getTime();
        if (isNaN(saleTime)) continue;
        const openPeriodStart = getOperationalOpenPeriodStartFromMap_(operationalBoundaries, recordBranchId);
        if (openPeriodStart !== null && saleTime <= openPeriodStart) continue;
        
        const recordSellerId = data[i][7];
        const sellerName = recordSellerId ? (sellerMap[recordSellerId] || "Unknown Seller") : "Admin";
        
        salesData.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          unitPrice: data[i][3],
          total: data[i][4],
          customer: data[i][5],
          paymentMethod: data[i][6],
          sellerId: recordSellerId,
          sellerName: sellerName,
          branchId: recordBranchId, // Include branch ID
          unit: data[i][9] || "",
          receiptNo: data[i][10] || "",
          rowIndex: i + 1,
          id: `sale_${i}_${Date.now()}`
        });
      }
    }
    
    return salesData;
  } catch (error) {
    return [];
  }
}

/* ===== GET PURCHASES DATA FOR SYNC (by branch) ===== */
function getPurchasesDataForSync(databaseId, branchId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    
    if (!purchasesSheet) return [];
    
    const data = purchasesSheet.getDataRange().getValues();
    const operationalBoundaries = getOperationalOpenPeriodStarts_(userDB);
    const purchasesData = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const recordBranchId = data[i][8] || "MAIN"; // Branch ID column
        
        // Filter by branch if seller
        if (branchId && recordBranchId !== branchId) {
          continue;
        }

        // Sync must return ONLY purchases from the current open period.
        // Otherwise a background sync can repopulate closed purchases
        // into the local cache after the Purchase table was already reset.
        const purchaseTime = new Date(data[i][0]).getTime();
        if (isNaN(purchaseTime)) continue;
        const openPeriodStart = getOperationalOpenPeriodStartFromMap_(operationalBoundaries, recordBranchId);
        if (openPeriodStart !== null && purchaseTime <= openPeriodStart) continue;
        
        purchasesData.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          unitCost: data[i][3],
          total: data[i][4],
          supplier: data[i][5],
          paymentMethod: data[i][6],
          sellerId: data[i][7] || null,
          branchId: recordBranchId, // Include branch ID
          category: data[i][9] || "",
          unit: data[i][10] || "",
          isStockProduct: data[i][11] !== "nonstock",
          rowIndex: i + 1,
          id: `purchase_${i}_${Date.now()}`
        });
      }
    }
    
    return purchasesData;
  } catch (error) {
    return [];
  }
}

/* ===== GET PRODUCTS FOR SYNC (BY BRANCH) ===== */
function getProductsForSync(databaseId, branchId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) return {};
    
    const data = stockSheet.getDataRange().getValues();
    const products = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        const productName = data[i][1];
        const itemBranchId = data[i][8] || "MAIN";
        
        // For sellers, only include products from their branch
        if (branchId && itemBranchId !== branchId) {
          continue;
        }
        
        const quantity = data[i][2] || 0;
        const sellPrice = data[i][4] || 0;
        
        if (!products[productName]) {
          products[productName] = {
            quantity: quantity,
            sellPrice: sellPrice,
            buyPrice: data[i][3] || 0,
            category: data[i][6] || ""
          };
        } else {
          products[productName].quantity += quantity;
          products[productName].sellPrice = sellPrice || products[productName].sellPrice;
        }
      }
    }
    
    return products;
  } catch (error) {
    return {};
  }
}

/* ===== GET BRANCHES DATA FOR SYNC ===== */
function getBranchesDataForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const branchesSheet = userDB.getSheetByName("BRANCHES");
    
    if (!branchesSheet) return [];
    
    const data = branchesSheet.getDataRange().getValues();
    const branches = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        branches.push({
          id: data[i][0],
          name: data[i][1],
          location: data[i][2] || "",
          manager: data[i][3] || "",
          phone: data[i][4] || "",
          created: data[i][5],
          status: data[i][6] || "active"
        });
      }
    }
    
    return branches;
  } catch (error) {
    return [];
  }
}

/* ===== GET SELLERS DATA FOR SYNC (with branch ID) ===== */
function getSellersDataForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    
    if (!sellersSheet) return [];
    
    const data = sellersSheet.getDataRange().getValues();
    const sellers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        sellers.push({
          id: data[i][0],
          name: data[i][1],
          phone: data[i][2],
          email: data[i][3],
          password: data[i][4],
          totalSales: data[i][5] || 0,
          totalPurchases: data[i][6] || 0,
          lastActive: data[i][7],
          created: data[i][8],
          status: data[i][9] || "active",
          branchId: data[i][10] || "MAIN" // NEW: Branch ID
        });
      }
    }
    
    return sellers;
  } catch (error) {
    return [];
  }
}

/* ===== GET CUSTOMERS DATA FOR SYNC ===== */
function getCustomersDataForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const customersSheet = userDB.getSheetByName("CUSTOMERS");
    if (!customersSheet) return [];
    
    const data = customersSheet.getDataRange().getValues();
    const customers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        customers.push({
          id: data[i][0],
          name: data[i][1],
          phone: data[i][2],
          email: data[i][3],
          address: data[i][4],
          totalPurchases: data[i][5] || 0,
          lastPurchase: data[i][6],
          notes: data[i][7],
          created: data[i][8]
        });
      }
    }
    
    return customers;
  } catch (error) {
    return [];
  }
}

/* ===== GET SUPPLIERS DATA FOR SYNC ===== */
function getSuppliersDataForSync(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const suppliersSheet = userDB.getSheetByName("SUPPLIERS");
    if (!suppliersSheet) return [];
    
    const data = suppliersSheet.getDataRange().getValues();
    const suppliers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        suppliers.push({
          id: data[i][0],
          name: data[i][1],
          contact: data[i][2],
          phone: data[i][3],
          email: data[i][4],
          address: data[i][5],
          products: data[i][6],
          totalSupplied: data[i][7] || 0,
          lastSupply: data[i][8],
          notes: data[i][9],
          created: data[i][10]
        });
      }
    }
    
    return suppliers;
  } catch (error) {
    return [];
  }
}

/* ===== GET MASTER DATABASE USING DRIVE API ONLY ===== */
function getMasterDatabase() {
  try {
    // Search for existing master database using Drive API
    const query = "title='" + MASTER_DB_NAME + "' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
    const files = Drive.Files.list({
      q: query,
      fields: "items(id,title)"
    });
    
    if (files.items && files.items.length > 0) {
      return SpreadsheetApp.openById(files.items[0].id);
    } else {
      // Create new master database
      const ss = SpreadsheetApp.create(MASTER_DB_NAME);
      const masterSheet = ss.getActiveSheet();
      masterSheet.setName("MASTER_USERS");
      masterSheet.appendRow(["Business", "Email", "Password", "DatabaseID", "Created", "LastLogin", "LogoID", "UserType"]);
      masterSheet.getRange("A1:H1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
      return ss;
    }
  } catch (error) {
    throw new Error("Master database error: " + error);
  }
}

/* ============================================
   SESSION MANAGEMENT (server-verified tokens)
   ============================================ */

/* ============================================
   MASTER SELLERS REGISTRY (in the shared master database)
   Every Seller across every business is recorded here too, with enough
   info (password, name, status, branch) to fully authenticate a login
   WITHOUT ever opening the business's own database - that extra open+scan
   was the real source of slow Seller logins. The business database is
   only opened afterward, once, to fetch settings/products for a
   successful login.
   ============================================ */

function ensureMasterSellersSheet(masterDB) {
  var sheet = masterDB.getSheetByName("MASTER_SELLERS");
  if (!sheet) {
    sheet = masterDB.insertSheet("MASTER_SELLERS");
    sheet.appendRow(["SellerEmail", "AdminEmail", "SellerID", "DatabaseID", "Created", "Password", "SellerName", "Status", "BranchID"]);
    sheet.getRange("A1:I1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  } else {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.length < 6) sheet.getRange(1, 6).setValue("Password");
    if (headers.length < 7) sheet.getRange(1, 7).setValue("SellerName");
    if (headers.length < 8) sheet.getRange(1, 8).setValue("Status");
    if (headers.length < 9) sheet.getRange(1, 9).setValue("BranchID");
  }
  return sheet;
}

// Finds a Seller's row in the registry by their sellerId - used to keep
// the registry's cached copy (password/name/status/branch) in sync
// whenever updateSeller() changes any of those fields.
function findMasterSellerRow(masterSellersSheet, sellerId) {
  var data = masterSellersSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === sellerId) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function ensureSessionsSheet(masterDB) {
  var sheet = masterDB.getSheetByName(SESSION_SHEET_NAME);
  if (!sheet) {
    sheet = masterDB.insertSheet(SESSION_SHEET_NAME);
    sheet.appendRow(["Token", "Email", "DatabaseID", "UserType", "SellerID", "BranchID", "CreatedAt", "ExpiresAt", "LastSeenAt"]);
    sheet.getRange("A1:I1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

// Cheap best-effort cleanup so the sheet doesn't grow forever. Only runs when
// the sheet has actually gotten large, so normal logins stay fast.
function purgeExpiredSessions(sheet) {
  try {
    if (sheet.getLastRow() < 300) return;
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    for (var i = data.length - 1; i >= 1; i--) {
      var exp = data[i][7] ? new Date(data[i][7]) : null;
      if (exp && exp.getTime() < now.getTime()) {
        sheet.deleteRow(i + 1);
      }
    }
  } catch (e) {
    console.error("purgeExpiredSessions error:", e);
  }
}

// Creates a real, server-tracked session and returns the opaque token for it.
function createSession(email, databaseId, userType, sellerId, branchId) {
  var masterDB = getMasterDatabase();
  var sheet = ensureSessionsSheet(masterDB);
  purgeExpiredSessions(sheet);
  var token = Utilities.getUuid() + "-" + Utilities.getUuid();
  var now = new Date();
  var expires = new Date(now.getTime() + SESSION_TTL_DAYS * 24 * 60 * 60 * 1000);
  sheet.appendRow([token, email, databaseId, userType || "admin", sellerId || "", branchId || "", now, expires, now]);
  return token;
}

// Validates a token and returns the VERIFIED identity behind it. Callers must
// use the returned email/databaseId instead of anything the client claimed.
function validateSession(token) {
  if (!token) return { valid: false, reason: "No token" };
  try {
    var masterDB = getMasterDatabase();
    var sheet = ensureSessionsSheet(masterDB);
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === token) {
        var expiresAt = data[i][7] ? new Date(data[i][7]) : null;
        if (!expiresAt || expiresAt.getTime() < now.getTime()) {
          return { valid: false, reason: "Session expired" };
        }
        // Sliding expiry: an active session keeps renewing itself.
        var newExpiry = new Date(now.getTime() + SESSION_TTL_DAYS * 24 * 60 * 60 * 1000);
        sheet.getRange(i + 1, 8).setValue(newExpiry);
        sheet.getRange(i + 1, 9).setValue(now);
        return {
          valid: true,
          email: data[i][1],
          databaseId: data[i][2],
          userType: data[i][3] || "admin",
          sellerId: data[i][4] || null,
          branchId: data[i][5] || null
        };
      }
    }
    return { valid: false, reason: "Invalid session" };
  } catch (error) {
    return { valid: false, reason: "Session check failed: " + error.toString() };
  }
}

function deleteSession(token) {
  if (!token) return;
  try {
    var masterDB = getMasterDatabase();
    var sheet = ensureSessionsSheet(masterDB);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === token) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
  } catch (error) {
    console.error("Error deleting session:", error);
  }
}

function logoutUser(d) {
  try {
    if (d && d.token) deleteSession(d.token);
    return output({ success: true, message: "Logged out" });
  } catch (error) {
    // Logout should never really "fail" from the client's point of view.
    return output({ success: true, message: "Logged out" });
  }
}

/* ===== CREATE USER DATABASE ===== */
function createUserDatabase(email, business, timeZone) {
  try {
    const cleanEmail = email.replace(/[@.]/g, "_").toLowerCase();
    const cleanBusiness = (business || "business").replace(/\s+/g, "_").replace(/[^\w_]/g, "").toLowerCase();
    
    const dbName = `SMARTTRADE_${cleanEmail}_${cleanBusiness}_${Date.now()}`.substring(0, 100);
    const userDB = SpreadsheetApp.create(dbName);
    const dbId = userDB.getId();
    
    // Pass the selected/detected business timezone into the sheet initializer.
    // This fixes signup failures caused by createDefaultSheets() referencing
    // timeZone without having it in scope.
    const safeTimeZone = isValidBusinessTimeZone_(timeZone) ? timeZone : BUSINESS_TIME_ZONE;
    createDefaultSheets(userDB, safeTimeZone);
    
    return {
      success: true,
      databaseId: dbId,
      databaseName: dbName
    };
    
  } catch (error) {
    return {
      success: false,
      error: "Failed to create user database: " + error
    };
  }
}

function createDefaultSheets(spreadsheet, timeZone) {
  // Always define a safe timezone for newly created businesses.
  // Registration may come from any country, and an invalid/missing client
  // timezone must never prevent account creation.
  const safeTimeZone = isValidBusinessTimeZone_(timeZone) ? timeZone : BUSINESS_TIME_ZONE;
  // Stock sheet
  const stockSheet = spreadsheet.insertSheet("STOCK");
  stockSheet.appendRow(["Date", "Product", "Quantity", "Buy Price", "Sell Price", "Stock Value", "Category", "Seller ID", "Branch ID", "Unit"]);
  stockSheet.getRange("A1:J1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Sales sheet
  const salesSheet = spreadsheet.insertSheet("SALES");
  salesSheet.appendRow(["Date", "Product", "Quantity", "Unit Price", "Total Amount", "Customer", "Payment Method", "Seller ID", "Branch ID", "Unit"]);
  salesSheet.getRange("A1:J1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Purchases sheet
  const purchasesSheet = spreadsheet.insertSheet("PURCHASES");
  purchasesSheet.appendRow(["Date", "Product", "Quantity", "Unit Cost", "Total Cost", "Supplier", "Payment Method", "Seller ID", "Branch ID", "Category", "Unit", "Item Type"]);
  purchasesSheet.getRange("A1:L1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // DEBTS sheet - NO DUE DATE
  const debtsSheet = spreadsheet.insertSheet("DEBTS");
  debtsSheet.appendRow([
    "ID",                    // A
    "Date",                  // B
    "Borrower Name",         // C
    "Borrower Phone",        // D
    "Borrower NIDA",         // E
    "Photo ID",              // F
    "Product",               // G
    "Quantity",              // H
    "Unit Price",            // I
    "Total Amount",          // J
    "Paid Amount",           // K
    "Status",                // L
    "Branch ID",             // M
    "Seller ID",             // N
    "Notes",                 // O
    "Last Updated"           // P
  ]);
  debtsSheet.getRange("A1:P1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Customers sheet
  const customersSheet = spreadsheet.insertSheet("CUSTOMERS");
  customersSheet.appendRow(["ID", "Name", "Phone", "Email", "Address", "Total Purchases", "Last Purchase", "Notes", "Created"]);
  customersSheet.getRange("A1:I1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Suppliers sheet
  const suppliersSheet = spreadsheet.insertSheet("SUPPLIERS");
  suppliersSheet.appendRow(["ID", "Name", "Contact Person", "Phone", "Email", "Address", "Products Supplied", "Total Supplied", "Last Supply", "Notes", "Created"]);
  suppliersSheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Sellers sheet
  const sellersSheet = spreadsheet.insertSheet("SELLERS");
  sellersSheet.appendRow(["ID", "Name", "Phone", "Email", "Password", "Total Sales", "Total Purchases", "Last Active", "Created", "Status", "Branch ID"]);
  sellersSheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // BRANCHES sheet
  const branchesSheet = spreadsheet.insertSheet("BRANCHES");
  branchesSheet.appendRow(["ID", "Name", "Location", "Manager", "Phone", "Created", "Status"]);
  branchesSheet.getRange("A1:G1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Add default MAIN BRANCH
  branchesSheet.appendRow(["MAIN", "Main Branch", "Head Office", "", "", new Date(), "active"]);
  
  // Settings sheet
  const settingsSheet = spreadsheet.insertSheet("SETTINGS");
  settingsSheet.appendRow(["Section", "Key", "Value", "Updated"]);
  settingsSheet.getRange("A1:D1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  
  // Add default settings
  const defaultSettings = [
    ["PROFILE", "business_name", "", new Date()],
    ["PROFILE", "email", "", new Date()],
    ["PROFILE", "phone", "", new Date()],
    ["PROFILE", "address", "", new Date()],
    ["DISPLAY", "theme", "dark", new Date()],
    ["DISPLAY", "language", "sw", new Date()],
    ["DISPLAY", "currency", "TZS", new Date()],
    ["DISPLAY", "date_format", "dd/mm/yyyy", new Date()],
    ["PROFILE", "time_zone", safeTimeZone, new Date()]
  ];
  
  // Add default settings - one batched write instead of 8 separate
  // appendRow() calls.
  if (defaultSettings.length > 0) {
    settingsSheet.getRange(2, 1, defaultSettings.length, 4).setValues(defaultSettings);
  }
  
  // UNITS sheet - seeded with the default unit list, plus anything the
  // business adds later via "Add Unit"
  ensureUnitsSheet(spreadsheet);
  
  // CATEGORIES sheet - seeded with a large default category list
  ensureCategoriesSheet(spreadsheet);
  
  // NON_STOCK_PRODUCTS sheet - fills up automatically as non-stock
  // services get sold
  ensureNonStockProductsSheet(spreadsheet);
  
  // PURCHASE_REQUESTS sheet - Seller purchase request workflow
  ensurePurchaseRequestsSheet(spreadsheet);
  
  spreadsheet.deleteSheet(spreadsheet.getSheetByName("Sheet1"));
}



/* ================== AUTH FUNCTIONS ================== */

/* ===== REGISTER USER ===== */
function registerUser(d) {
  try {
    if (!d.business || !d.email || !d.password) {
      return output({ success: false, error: "Business name, email, and password are required" });
    }
    
    if (!isPasswordStrongEnough(d.password)) {
      return output({ success: false, error: "Password must be at least 8 characters and include a mix of uppercase, lowercase, numbers or symbols" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    
    const data = masterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email) {
        return output({ success: false, error: "Email already registered. Please login or use another email." });
      }
    }
    
    const dbResult = createUserDatabase(d.email, d.business, d.clientTimeZone);
    if (!dbResult.success) {
      return output(dbResult);
    }
    
    const timestamp = new Date();
    masterSheet.appendRow([
      d.business || "My Business",
      d.email,
      d.password,
      dbResult.databaseId,
      timestamp,
      timestamp,
      "",
      "admin"
    ]);
    
    return output({ 
      success: true, 
      message: "Account created successfully!",
      business: d.business,
      databaseId: dbResult.databaseId,
      timestamp: timestamp.toISOString()
    });
    
  } catch (error) {
    return output({ 
      success: false, 
      error: "Registration failed: " + error.toString() 
    });
  }
}

/* ============================================
   MULTI-BUSINESS SUPPORT
   One Admin account (one email+password) can own several completely
   separate businesses (each its own database - Stock, Sales, Settings,
   everything independent). These three actions let the Admin list their
   businesses, add a new one without re-entering a password (they're
   already authenticated), and switch the active session to a different
   one instantly.
   ============================================ */

// Lists every business owned by the currently logged-in Admin's email -
// used to refresh the business-switcher menu (e.g. right after adding one).
function getMyBusinesses(d) {
  try {
    if (!d.email) {
      return output({ success: false, error: "Email required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    const businesses = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && (data[i][7] === "admin" || !data[i][7])) {
        businesses.push({ business: data[i][0], databaseId: data[i][3] });
      }
    }
    
    return output({ success: true, businesses: businesses });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Adds a brand-new, fully independent business under the SAME account -
// no password needed since d.email is already the verified session
// identity; the existing password on file is simply reused for the new
// MASTER_USERS row (all businesses under one account share one login).
function addBusiness(d) {
  try {
    if (!d.email || !d.business || !d.business.trim()) {
      return output({ success: false, error: "Business name required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    let existingPassword = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email) {
        existingPassword = data[i][2];
        break;
      }
    }
    
    if (!existingPassword) {
      return output({ success: false, error: "Account not found" });
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && data[i][0] === d.business.trim()) {
        return output({ success: false, error: "You already have a business with this name" });
      }
    }
    
    const dbResult = createUserDatabase(d.email, d.business.trim(), d.clientTimeZone);
    if (!dbResult.success) {
      return output(dbResult);
    }
    
    const timestamp = new Date();
    masterSheet.appendRow([
      d.business.trim(),
      d.email,
      existingPassword,
      dbResult.databaseId,
      timestamp,
      timestamp,
      "",
      "admin"
    ]);
    
    return output({
      success: true,
      message: "Business added successfully!",
      business: d.business.trim(),
      databaseId: dbResult.databaseId
    });
  } catch (error) {
    return output({ success: false, error: "Failed to add business: " + error.toString() });
  }
}

// Switches the Admin's active session to a different one of their own
// businesses - issues a fresh session token scoped to that business's
// database, since every action in the app trusts the session's databaseId
// exclusively. Verifies the target databaseId genuinely belongs to this
// email before switching, so one business owner can never switch into
// (or guess their way into) another account's business.
function switchBusiness(d) {
  try {
    // IMPORTANT: the target business must travel as d.targetDatabaseId,
    // never d.databaseId - doPost's session check always overwrites
    // d.databaseId with the CURRENT session's database (the one being
    // switched AWAY from) for every authenticated action, so reading
    // d.databaseId here would always just confirm the old business again
    // and the switch would silently do nothing.
    if (!d.email || !d.targetDatabaseId) {
      return output({ success: false, error: "Target business required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    let matchedRow = null;
    let matchedRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && data[i][3] === d.targetDatabaseId) {
        matchedRow = data[i];
        matchedRowIndex = i + 1;
        break;
      }
    }
    
    if (!matchedRow) {
      return output({ success: false, error: "You don't have access to this business" });
    }
    
    // Mark this business as the most-recently-active one for this account,
    // so logging out and back in lands here again automatically instead
    // of always returning to whichever business was created first.
    masterSheet.getRange(matchedRowIndex, 6).setValue(new Date());
    
    const business = matchedRow[0];
    const logoId = matchedRow[6];
    const databaseId = d.targetDatabaseId;
    
    const settings = getUserSettingsFromDB(databaseId);
    
    let logoUrl = DEFAULT_LOGO_URL;
    let altLogoUrl = DEFAULT_LOGO_URL;
    let directLogoUrl = DEFAULT_LOGO_URL;
    
    // Prefer the base64 logo stored in THIS business's own Settings - no
    // external Drive fetch needed, so it can never silently fail to show.
    const logoInfo = getLogoBase64FromSettings(settings);
    if (logoInfo) {
      const dataUri = "data:" + logoInfo.mimeType + ";base64," + logoInfo.base64;
      logoUrl = dataUri;
      altLogoUrl = dataUri;
      directLogoUrl = dataUri;
    } else if (logoId) {
      try {
        DriveApp.getFileById(logoId);
        logoUrl = LOGO_BASE_URL + logoId;
        altLogoUrl = LOGO_ALT_URL + logoId;
        directLogoUrl = LOGO_DIRECT_URL + logoId;
      } catch (e) {
        // Logo missing/inaccessible - fall back to defaults, harmless.
      }
    }
    
    const token = createSession(d.email, databaseId, "admin", null, null);
    
    return output({
      success: true,
      business: business,
      email: d.email,
      token: token,
      databaseId: databaseId,
      logoUrl: logoUrl,
      altLogoUrl: altLogoUrl,
      directLogoUrl: directLogoUrl,
      settings: settings,
      defaultLogoUrl: DEFAULT_LOGO_URL,
      userType: "admin",
      products: getProductsForSync(databaseId)
    });
  } catch (error) {
    return output({ success: false, error: "Failed to switch business: " + error.toString() });
  }
}


// ============================================
// ACCOUNTING PERIODS - lets an Admin "close" a date range as a permanent,
// unchangeable financial snapshot (like a monthly bookkeeping close),
// without ever touching the underlying Sales/Purchases/Stock/Debts
// records themselves. This is purely additive on top of everything that
// already exists: it reads those sheets to compute totals for the chosen
// range, and writes one new row to its own ACCOUNTING_PERIODS sheet. It
// never deletes, resets, or modifies a single existing record - stock,
// sales, purchases and debts all keep behaving exactly as they always
// have, for every other feature in the app.
// ============================================

function ensureAccountingPeriodsSheet(userDB) {
  let sheet = userDB.getSheetByName("ACCOUNTING_PERIODS");
  if (!sheet) {
    sheet = userDB.insertSheet("ACCOUNTING_PERIODS");
    sheet.appendRow([
      "Close ID", "Period Start", "Period End", "Branch ID", "Status",
      "Closed At", "Closed By", "Total Sales", "Total Purchases",
      "Estimated COGS", "Gross Profit", "Operating Expenses", "Net Profit",
      "Closing Stock Value", "Stock Entry Count", "Outstanding Debts", "Notes"
    ]);
    sheet.getRange("A1:Q1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

// ============================================
// EXPENSES - genuine business OPERATING costs (rent, electricity,
// transport, salaries, internet, etc), completely separate from
// Purchases (buying goods to resell, whether Stock-tracked or not, is
// never an Expense - it's Cost of Goods Sold once actually sold, and
// Purchases already has its own sheet/flow for that). Its own dedicated
// sheet - not reusing Purchases for this, since they represent
// fundamentally different things in real accounting.
// ============================================
function ensureExpensesSheet(userDB) {
  let sheet = userDB.getSheetByName("EXPENSES");
  if (!sheet) {
    sheet = userDB.insertSheet("EXPENSES");
    sheet.appendRow([
      "Date", "Category", "Description", "Amount", "Payment Method",
      "Branch ID", "Seller ID", "Recorded At", "Operation ID"
    ]);
    sheet.getRange("A1:I1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

function addExpense(d) {
  try {
    const session = validateSession(d.token);
    if (!session.valid) {
      return output({ success: false, error: "Your session has expired or is invalid. Please login again.", sessionExpired: true });
    }
    
    if (!d.description || !d.amount) {
      return output({ success: false, error: "Description and amount are required" });
    }
    
    const amount = Number(d.amount);
    if (isNaN(amount) || amount <= 0) {
      return output({ success: false, error: "Amount must be a valid positive number" });
    }
    
    const userDB = SpreadsheetApp.openById(session.databaseId);
    const sheet = ensureExpensesSheet(userDB);
    
    // The person can pick a real Date (e.g. logging a rent payment from a
    // few days ago) - defaults to right now if none is given.
    const expenseDate = d.date ? new Date(d.date) : new Date();
    const recordedAt = d.recordedAt ? new Date(d.recordedAt) : new Date();
    const safeRecordedAt = isNaN(recordedAt.getTime()) ? new Date() : recordedAt;
    const branchId = d.branchId || "MAIN";
    const category = d.category || "Uncategorized";
    const paymentMethod = d.paymentMethod || "cash";
    
    // IMPORTANT: Date (column A) is the user-selected business calendar date,
    // while Recorded At (column H) is the exact real transaction instant.
    // Accounting-period membership MUST use Recorded At so a transaction
    // created at 09:30 is closed when accounts are closed at 09:37, even
    // though both records share the same calendar date.
    // Each Expense carries its own operationId. This protects offline sync
    // from duplicates if the network drops after the server has already
    // committed the row and the device retries later. Existing 8-column
    // Expense sheets are upgraded non-destructively by ensureExpensesSheet.
    const operationId = d.operationId ? String(d.operationId) : "";
    if (operationId && sheet.getLastColumn() < 9) {
      sheet.getRange(1, 9).setValue("Operation ID");
    }
    const expenseLock = LockService.getScriptLock();
    expenseLock.waitLock(10000);
    try {
      if (operationId) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const opValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
          for (let oi = 0; oi < opValues.length; oi++) {
            if (String(opValues[oi][0] || "") === operationId) {
              return output({ success: true, alreadyProcessed: true, operationId: operationId });
            }
          }
        }
      }

      if (sheet.getLastColumn() < 9) sheet.getRange(1, 9).setValue("Operation ID");
      sheet.appendRow([
        expenseDate, category, d.description, amount, paymentMethod,
        branchId, d.sellerId || "", safeRecordedAt, operationId
      ]);
      
      return output({ success: true, operationId: operationId });
    } finally {
      try { expenseLock.releaseLock(); } catch (_) {}
    }
  } catch (error) {
    return output({ success: false, error: "Failed to save expense: " + error.toString() });
  }
}

// Computes every figure the Accounting page needs for one period+branch,
// live, straight from the existing Sales/Purchases/Stock/Debts sheets -
// used both for "Preview" (nothing saved) and for the final numbers
// written when a period is actually closed.
//
// Estimated COGS: individual sales don't record which exact purchase
// batch/cost they came from, so this uses each product's CURRENT
// weighted-average buy cost (quantity-weighted across every matching row
// in Stock right now, per product+branch+unit) as the best available
// stand-in for what those units genuinely cost. This is always labeled
// "Estimated" everywhere it's shown, since it is an estimate rather than
// an exact historical cost trace - Sell Price is never used for this
// (that would just be circular).
function computeAccountingPeriodStats(userDB, periodStart, periodEnd, branchId) {
  // Business-timezone-aware boundaries (see getBusinessDayBoundary above)
  // - critical for a business ahead of UTC: without this, "close today"
  // would silently cut off partway through the actual local day.
  const startTime = getBusinessDayBoundary(periodStart, false);
  const endTime = getBusinessDayBoundary(periodEnd, true);
  const allBranches = !branchId || branchId === "ALL";
  
  // --- Weighted-average buy cost per product+branch+unit, from STOCK -
  // for genuinely stock-tracked products (the STOCK sheet only ever
  // holds Stock Purchases; Non-Stock Purchases never touch it) ---
  const stockSheet = userDB.getSheetByName("STOCK");
  const avgBuyCostStock = {};
  let closingStockValue = 0;
  let stockEntryCount = 0;
  
  if (stockSheet) {
    const stockData = stockSheet.getDataRange().getValues();
    const totals = {};
    for (let i = 1; i < stockData.length; i++) {
      if (!stockData[i][1]) continue;
      const rowBranch = stockData[i][8] || "MAIN";
      if (!allBranches && rowBranch !== branchId) continue;
      
      stockEntryCount++;
      closingStockValue += Number(stockData[i][5]) || 0;
      
      const key = stockData[i][1] + "|||" + rowBranch + "|||" + (stockData[i][9] || "");
      const qty = Number(stockData[i][2]) || 0;
      const buyPrice = Number(stockData[i][3]) || 0;
      if (!totals[key]) totals[key] = { qty: 0, cost: 0 };
      totals[key].qty += qty;
      totals[key].cost += qty * buyPrice;
    }
    Object.keys(totals).forEach(key => {
      avgBuyCostStock[key] = totals[key].qty > 0 ? totals[key].cost / totals[key].qty : 0;
    });
  }
  
  // --- Weighted-average buy cost per product+branch+unit, from
  // Non-Stock Purchases - products bought to resell but NOT tracked as
  // inventory quantity. A sale of one of these has no matching STOCK
  // entry at all, so THIS is the only available cost estimate for it -
  // built from all-time purchase history (not period-limited), matching
  // how Stock's own average cost above works the same way, since an
  // item sold this period could have been bought in an earlier one.
  const purchasesSheet = userDB.getSheetByName("PURCHASES");
  const avgBuyCostNonStock = {};
  let totalPurchases = 0;
  
  if (purchasesSheet) {
    const purchasesData = purchasesSheet.getDataRange().getValues();
    const nonStockTotals = {};
    for (let i = 1; i < purchasesData.length; i++) {
      if (!purchasesData[i][0]) continue;
      const rowBranch = purchasesData[i][8] || "MAIN";
      if (!allBranches && rowBranch !== branchId) continue;
      
      const purchaseTime = new Date(purchasesData[i][0]).getTime();
      const inPeriod = !isNaN(purchaseTime) && purchaseTime >= startTime && purchaseTime <= endTime;
      const total = Number(purchasesData[i][4]) || 0;
      
      // Total Purchases (Stock + Non-Stock combined) - this period only.
      if (inPeriod) totalPurchases += total;
      
      if (purchasesData[i][11] === "nonstock") {
        const key = purchasesData[i][1] + "|||" + rowBranch + "|||" + (purchasesData[i][10] || "");
        const qty = Number(purchasesData[i][2]) || 0;
        const unitCost = Number(purchasesData[i][3]) || 0;
        if (!nonStockTotals[key]) nonStockTotals[key] = { qty: 0, cost: 0 };
        nonStockTotals[key].qty += qty;
        nonStockTotals[key].cost += qty * unitCost;
      }
    }
    Object.keys(nonStockTotals).forEach(key => {
      avgBuyCostNonStock[key] = nonStockTotals[key].qty > 0 ? nonStockTotals[key].cost / nonStockTotals[key].qty : 0;
    });
  }
  
  // --- Sales + Estimated COGS - a Stock product uses Stock's own
  // average cost; anything with NO matching Stock entry (a Non-Stock
  // product) falls back to the Non-Stock Purchases average cost for
  // that same product+branch+unit - this is what correctly treats
  // Non-Stock Purchases as Cost of Goods Sold once actually sold,
  // rather than as an immediate Operating Expense at time of purchase. ---
  const salesSheet = userDB.getSheetByName("SALES");
  let totalSales = 0;
  let cogs = 0;
  if (salesSheet) {
    const salesData = salesSheet.getDataRange().getValues();
    for (let i = 1; i < salesData.length; i++) {
      if (!salesData[i][0]) continue;
      const rowBranch = salesData[i][8] || "MAIN";
      if (!allBranches && rowBranch !== branchId) continue;
      
      const saleTime = new Date(salesData[i][0]).getTime();
      if (isNaN(saleTime) || saleTime < startTime || saleTime > endTime) continue;
      
      totalSales += Number(salesData[i][4]) || 0;
      
      const key = salesData[i][1] + "|||" + rowBranch + "|||" + (salesData[i][9] || "");
      const qty = Number(salesData[i][2]) || 0;
      const unitCost = avgBuyCostStock[key] !== undefined ? avgBuyCostStock[key]
                     : (avgBuyCostNonStock[key] !== undefined ? avgBuyCostNonStock[key] : 0);
      cogs += qty * unitCost;
    }
  }
  
  // --- Operating Expenses - genuine business running costs (rent,
  // electricity, transport, salaries, internet, etc), from the dedicated
  // EXPENSES sheet. Completely separate from Purchases now: buying goods
  // to resell (Stock or Non-Stock) is Cost of Goods Sold territory
  // (handled above), never an Operating Expense. ---
  const expensesSheet = userDB.getSheetByName("EXPENSES");
  let operatingExpenses = 0;
  if (expensesSheet) {
    const expensesData = expensesSheet.getDataRange().getValues();
    for (let i = 1; i < expensesData.length; i++) {
      if (!expensesData[i][0]) continue;
      const rowBranch = expensesData[i][5] || "MAIN";
      if (!allBranches && rowBranch !== branchId) continue;
      
      const expenseTime = new Date(expensesData[i][0]).getTime();
      if (isNaN(expenseTime) || expenseTime < startTime || expenseTime > endTime) continue;
      
      operatingExpenses += Number(expensesData[i][3]) || 0;
    }
  }
  
  // Outstanding debts are a live, point-in-time balance - never filtered
  // by period, since a debt stays owed regardless of when it was created
  // (this matches the explicit requirement that Close Accounts must never
  // reset or affect outstanding debts).
  const debtsSheet = userDB.getSheetByName("DEBTS");
  let outstandingDebts = 0;
  if (debtsSheet) {
    const debtsData = debtsSheet.getDataRange().getValues();
    for (let i = 1; i < debtsData.length; i++) {
      if (!debtsData[i][0]) continue;
      const rowBranch = debtsData[i][12] || "MAIN";
      if (!allBranches && rowBranch !== branchId) continue;
      
      const status = debtsData[i][11] || "pending";
      if (status === "denied" || status === "paid") continue;
      const totalAmount = Number(debtsData[i][9]) || 0;
      const paidAmount = Number(debtsData[i][10]) || 0;
      outstandingDebts += (totalAmount - paidAmount);
    }
  }
  
  // Net Profit = Sales − COGS − Operating Expenses
  const grossProfit = totalSales - cogs;
  const netProfit = grossProfit - operatingExpenses;
  
  return {
    totalSales: totalSales, totalPurchases: totalPurchases, cogs: cogs,
    grossProfit: grossProfit, operatingExpenses: operatingExpenses, netProfit: netProfit,
    closingStockValue: closingStockValue, stockEntryCount: stockEntryCount,
    outstandingDebts: outstandingDebts
  };
}

// Finds an existing CLOSED record matching this exact period+branch, if
// one exists - this is what lets the page show "Period already closed"
// with the real saved figures (never recomputed) instead of silently
// allowing the same period to be closed twice.
function findExistingClosedPeriod(userDB, periodStart, periodEnd, branchId) {
  const sheet = userDB.getSheetByName("ACCOUNTING_PERIODS");
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const normalizedBranch = branchId || "ALL";
  const startStr = Utilities.formatDate(new Date(periodStart), getCurrentBusinessTimeZone_(), "yyyy-MM-dd");
  const endStr = Utilities.formatDate(new Date(periodEnd), getCurrentBusinessTimeZone_(), "yyyy-MM-dd");
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] !== "CLOSED") continue;
    if (Utilities.formatDate(new Date(data[i][1]), getCurrentBusinessTimeZone_(), "yyyy-MM-dd") !== startStr) continue;
    if (Utilities.formatDate(new Date(data[i][2]), getCurrentBusinessTimeZone_(), "yyyy-MM-dd") !== endStr) continue;
    if ((data[i][3] || "ALL") !== normalizedBranch) continue;
    
    return {
      closeId: data[i][0], periodStart: data[i][1], periodEnd: data[i][2],
      branchId: data[i][3], status: data[i][4], closedAt: data[i][5], closedBy: data[i][6],
      totalSales: data[i][7], totalPurchases: data[i][8], cogs: data[i][9],
      grossProfit: data[i][10], operatingExpenses: data[i][11], netProfit: data[i][12],
      closingStockValue: data[i][13], stockEntryCount: data[i][14],
      outstandingDebts: data[i][15], notes: data[i][16]
    };
  }
  return null;
}

// Finds any CLOSED period whose date range overlaps the requested one BY
// EVEN A SINGLE DAY, for a branch this new period would actually affect -
// that means the exact branch requested, AND "ALL" (closing "ALL
// Branches" already finalized every branch's figures for that range, so
// a specific branch can never be re-closed inside it either), AND, if
// the request itself is for "ALL", any single branch already closed
// somewhere inside that range too (closing "ALL" must not silently
// re-finalize a branch that was already individually closed). This is
// what makes "you cannot close into an already-closed period, even by
// one overlapping day" actually true, rather than only catching an
// exact duplicate date-range match.
function findOverlappingClosedPeriod(userDB, periodStart, periodEnd, branchId) {
  const sheet = userDB.getSheetByName("ACCOUNTING_PERIODS");
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const normalizedBranch = branchId || "ALL";
  // Same business-timezone-aware boundaries as everywhere else - a new
  // period's start is the beginning of its day, its end is the end of
  // its day, exactly matching how an existing closed period's own range
  // was computed when it was closed.
  const newStart = getBusinessDayBoundary(periodStart, false);
  const newEnd = getBusinessDayBoundary(periodEnd, true);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] !== "CLOSED") continue;
    
    const rowBranch = data[i][3] || "ALL";
    const branchRelevant = (rowBranch === normalizedBranch) || (rowBranch === "ALL") || (normalizedBranch === "ALL");
    if (!branchRelevant) continue;
    
    const existingStart = getBusinessDayBoundary(data[i][1], false);
    const existingEnd = getBusinessDayBoundary(data[i][2], true);
    
    // Standard range-overlap test - true the moment the two ranges share
    // even a single day.
    if (existingStart <= newEnd && existingEnd >= newStart) {
      return {
        closeId: data[i][0], periodStart: data[i][1], periodEnd: data[i][2],
        branchId: data[i][3]
      };
    }
  }
  return null;
}

// Shared validation for both Preview and Close - a period can never
// genuinely make sense with its start after its end, or with an end date
// that hasn't happened yet (there is no real, final data for days in the
// future). Returns an error string, or null if the dates are valid.
function validateAccountingPeriodDates(periodStart, periodEnd) {
  // Period Start/End are calendar dates (YYYY-MM-DD), not instants.
  // Validate the strings directly so Apps Script cannot reinterpret a
  // date-only value as UTC midnight and shift it into another calendar day.
  const start = String(periodStart || '').trim();
  const end = String(periodEnd || '').trim();
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;

  if (!datePattern.test(start) || !datePattern.test(end)) {
    return "Please provide valid Period Start and Period End dates.";
  }
  if (start > end) {
    return "Period Start must be before Period End.";
  }

  // Today is the actual calendar date in the current device/business
  // timezone. No hardcoded UTC+3 and no server-local timezone is used.
  const businessToday = Utilities.formatDate(new Date(), getCurrentBusinessTimeZone_(), "yyyy-MM-dd");
  if (end > businessToday) {
    return "Period End cannot be a future date - a period can only be closed once it has actually happened.";
  }
  
  return null;
}

// Preview (or re-check) a period - never writes anything. If this exact
// period+branch was already closed, returns the SAVED figures from that
// closing (so the page shows history, not a live recompute that could
// drift from what was actually locked in); if it overlaps a different
// already-closed period by even one day, that's flagged too (Close is
// never actually reachable for it - see closeAccountingPeriod, which
// enforces this for real); otherwise computes live.
function previewAccountingPeriod(d) {
  try {
    if (!d.databaseId || !d.periodStart || !d.periodEnd) {
      return output({ success: false, error: "Period start and end are required" });
    }
    
    const dateError = validateAccountingPeriodDates(d.periodStart, d.periodEnd);
    if (dateError) {
      return output({ success: false, error: dateError });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const branchId = d.branchId || "ALL";
    
    const existing = findExistingClosedPeriod(userDB, d.periodStart, d.periodEnd, branchId);
    if (existing) {
      return output({ success: true, alreadyClosed: true, period: existing });
    }
    
    const overlap = findOverlappingClosedPeriod(userDB, d.periodStart, d.periodEnd, branchId);
    if (overlap) {
      return output({ success: true, overlapsClosedPeriod: true, overlap: overlap });
    }
    
    const stats = computeAccountingPeriodStats(userDB, d.periodStart, d.periodEnd, branchId);
    return output({ success: true, alreadyClosed: false, period: stats });
  } catch (error) {
    return output({ success: false, error: "Failed to preview period: " + error.toString() });
  }
}

// Permanently closes a period - Admin only (re-verified here from the
// session token itself, never trusting a userType field the client might
// send, since doPost's session check does not itself gate individual
// actions by role). Computes the final figures and appends one permanent
// row - it never touches Sales, Purchases, Stock, or Debts records
// themselves, so every one of those keeps working completely normally,
// with full history, both before and after a close.
function closeAccountingPeriod(d) {
  try {
    const session = validateSession(d.token);
    if (!session.valid) {
      return output({ success: false, error: "Your session has expired or is invalid. Please login again.", sessionExpired: true });
    }
    if (session.userType !== "admin") {
      return output({ success: false, error: "Only an admin can close an accounting period." });
    }
    
    if (!d.periodStart || !d.periodEnd) {
      return output({ success: false, error: "Period start and end are required" });
    }
    
    const dateError = validateAccountingPeriodDates(d.periodStart, d.periodEnd);
    if (dateError) {
      return output({ success: false, error: dateError });
    }
    
    const userDB = SpreadsheetApp.openById(session.databaseId);
    const branchId = d.branchId || "ALL";
    
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      // Re-checked here, inside the lock, for the same reason as the
      // duplicate-close check below: closes the small window where two
      // requests could otherwise both pass validation and both get
      // appended. This is the authoritative enforcement of "never close
      // into an already-closed period, even by one overlapping day" -
      // Preview can only ever warn about it in advance.
      const overlap = findOverlappingClosedPeriod(userDB, d.periodStart, d.periodEnd, branchId);
      if (overlap) {
        return output({
          success: false,
          error: "This period overlaps an already-closed period (" + overlap.closeId + "). A period can never be closed inside one that's already been finalized, even by a single day.",
          overlapsClosedPeriod: true,
          overlap: overlap
        });
      }
      
      // Duplicate-close check - an EXACT re-match of an existing closed
      // period is reported as a success (idempotent - the person just
      // gets shown the same figures again), rather than an error, since
      // it's harmless and likely just a repeated tap.
      const existing = findExistingClosedPeriod(userDB, d.periodStart, d.periodEnd, branchId);
      if (existing) {
        return output({ success: true, alreadyClosed: true, period: existing });
      }
      
      const stats = computeAccountingPeriodStats(userDB, d.periodStart, d.periodEnd, branchId);
      const now = new Date();
      const closeId = "CLOSE-" + Utilities.formatDate(now, getCurrentBusinessTimeZone_(), "yyyyMMdd-HHmmss") + "-" + Utilities.getUuid().slice(0, 8);
      
      const sheet = ensureAccountingPeriodsSheet(userDB);
      sheet.appendRow([
        closeId, new Date(d.periodStart), new Date(d.periodEnd), branchId, "CLOSED",
        now, session.email, stats.totalSales, stats.totalPurchases, stats.cogs,
        stats.grossProfit, stats.operatingExpenses, stats.netProfit,
        stats.closingStockValue, stats.stockEntryCount, stats.outstandingDebts,
        d.notes || ""
      ]);
      
      return output({
        success: true, alreadyClosed: false,
        period: Object.assign({ closeId: closeId, periodStart: d.periodStart, periodEnd: d.periodEnd, branchId: branchId, status: "CLOSED", closedAt: now, closedBy: session.email }, stats)
      });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return output({ success: false, error: "Failed to close accounting period: " + error.toString() });
  }
}

// Lists every closed period (most recent first), optionally filtered by
// branch - powers the "Closing Reports" tab. Read-only.
function getClosingReports(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = userDB.getSheetByName("ACCOUNTING_PERIODS");
    if (!sheet) {
      return output({ success: true, reports: [] });
    }
    
    const data = sheet.getDataRange().getValues();
    const branchFilter = d.branchId || null;
    const reports = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] !== "CLOSED") continue;
      if (branchFilter && branchFilter !== "ALL" && (data[i][3] || "ALL") !== branchFilter) continue;
      
      reports.push({
        closeId: data[i][0], periodStart: data[i][1], periodEnd: data[i][2],
        branchId: data[i][3], status: data[i][4], closedAt: data[i][5], closedBy: data[i][6],
        totalSales: data[i][7], totalPurchases: data[i][8], cogs: data[i][9],
        grossProfit: data[i][10], operatingExpenses: data[i][11], netProfit: data[i][12],
        closingStockValue: data[i][13], stockEntryCount: data[i][14],
        outstandingDebts: data[i][15], notes: data[i][16]
      });
    }
    
    reports.sort((a, b) => new Date(b.closedAt).getTime() - new Date(a.closedAt).getTime());
    
    return output({ success: true, reports: reports });
  } catch (error) {
    return output({ success: false, error: "Failed to load closing reports: " + error.toString() });
  }
}

// ============================================
// EXPENSES (Accounting page, "Expenses" tab) - built entirely from
// existing PURCHASES records that are marked "nonstock" (a purchase of
// something that ISN'T inventory to resell - rent, utilities, transport,
// etc). Stock Purchases (buying products to sell) are deliberately
// excluded here - they're a different thing entirely (an asset/COGS
// input, not an operating expense) and must never be double-counted as
// both a Purchase and an Expense. No new sheet, no new data model -
// entirely additive read-only aggregation over what already exists.
// ============================================
function getExpensesData(d) {
  try {
    const session = validateSession(d.token);
    if (!session.valid) {
      return output({ success: false, error: "Your session has expired or is invalid. Please login again.", sessionExpired: true });
    }
    
    const userDB = SpreadsheetApp.openById(session.databaseId);
    const expensesSheet = userDB.getSheetByName("EXPENSES");
    
    const branchId = d.branchId || "ALL";
    const allBranches = branchId === "ALL";
    const categoryFilter = d.category || null;
    
    const timeZone = getCurrentBusinessTimeZone_();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    
    // Read all applicable Close Accounts boundaries once. Every expense
    // row is then tested against the exact Closed At timestamp for its
    // own branch (or an ALL-Branches close). This is the same operational
    // boundary rule used by live Sales and Purchases.
    const operationalBoundaries = getOperationalOpenPeriodStarts_(userDB);
    
    const dateFromTime = d.dateFrom ? getBusinessDayBoundary(d.dateFrom, false) : null;
    const dateToTime = d.dateTo ? getBusinessDayBoundary(d.dateTo, true) : null;
    
    let todayExpenses = 0;
    let periodExpenses = 0;
    let totalExpenses = 0;
    const categoryTotals = {};
    const filteredList = [];
    
    if (expensesSheet) {
      const data = expensesSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        
        const rowBranch = data[i][5] || "MAIN";
        if (!allBranches && rowBranch !== branchId) continue;
        
        const expenseDate = new Date(data[i][0]);
        if (isNaN(expenseDate.getTime())) continue;
        const recordedAt = data[i][7] ? new Date(data[i][7]) : expenseDate;
        if (isNaN(recordedAt.getTime())) continue;
        const category = data[i][1] || "Uncategorized";
        const description = data[i][2];
        const total = Number(data[i][3]) || 0;
        const openPeriodStart = getOperationalOpenPeriodStartFromMap_(operationalBoundaries, rowBranch);

        // Accounting periods are defined by the exact moment the expense
        // was recorded, NOT by the date-only value in column A. This keeps
        // the boundary accurate to date + hour + minute + second + ms.
        // Example: expense recorded 05/09/2026 09:30 and accounts closed
        // 05/09/2026 09:37 => excluded; a new expense at 09:38 => included.
        const expenseTime = recordedAt.getTime();
        // Defensive rule: once Close Accounts has a Closed At boundary for
        // this branch (or ALL branches), an expense at/before that exact
        // moment belongs only to Closing Reports, never to live Expenses.
        const isInCurrentPeriod = (openPeriodStart === null || expenseTime > openPeriodStart);

        // A live Expenses table must NEVER show an expense belonging to a
        // closed accounting period. The exact Closed At instant is the
        // boundary, so rows at/before that timestamp stay in Closing
        // Reports and are excluded from the new live period.
        if (!isInCurrentPeriod) continue;

        // Current-period statistics are calculated only from the open
        // period. Closed-period expenses therefore cannot leak into the
        // new period after Close Accounts.
        periodExpenses += total;
        const dateStr = Utilities.formatDate(expenseDate, timeZone, "yyyy-MM-dd");
        if (dateStr === todayStr) {
          todayExpenses += total;
        }

        // Category breakdown and Recent Expenses use the selected filters,
        // but ONLY within the currently open accounting period.
        if (categoryFilter && category !== categoryFilter) continue;
        if (dateFromTime !== null && expenseTime < dateFromTime) continue;
        if (dateToTime !== null && expenseTime > dateToTime) continue;
        
        categoryTotals[category] = (categoryTotals[category] || 0) + total;
        
        filteredList.push({
          date: data[i][0], recordedAt: data[i][7] || data[i][0], description: description, total: total,
          paymentMethod: data[i][4], category: category, branchId: rowBranch
        });
      }
    }
    
    filteredList.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
    
    const categories = Object.keys(categoryTotals)
      .map(name => ({ name: name, total: categoryTotals[name] }))
      .sort((a, b) => b.total - a.total);
    
    return output({
      success: true,
      todayExpenses: todayExpenses,
      periodExpenses: periodExpenses,
      totalExpenses: periodExpenses,
      categories: categories,
      recent: filteredList.slice(0, 50)
    });
  } catch (error) {
    return output({ success: false, error: "Failed to load expenses: " + error.toString() });
  }
}

function getBranchNameFromDB(userDB, branchId) {


  try {
    if (!branchId || branchId === "MAIN") return "Main Branch";
    const branchesSheet = userDB.getSheetByName("BRANCHES");
    if (!branchesSheet) return "Main Branch";
    const data = branchesSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === branchId) {
        return data[i][1] || branchId;
      }
    }
    return branchId;
  } catch (e) {
    return branchId || "Main Branch";
  }
}

function getBranchNameById(databaseId, branchId) {
  try {
    if (!branchId || branchId === "MAIN") return "Main Branch";
    const userDB = SpreadsheetApp.openById(databaseId);
    return getBranchNameFromDB(userDB, branchId);
  } catch (e) {
    return branchId || "Main Branch";
  }
}

/* ===== LOGIN USER ===== */
function loginUser(d) {
  try {
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    
    const data = masterSheet.getDataRange().getValues();
    
    // Check for admin - gather every business row that matches this
    // email+password (they all share the same login credentials), then
    // land on whichever one was MOST RECENTLY active - i.e. the last
    // business the Admin was actually using (whether that's from a
    // regular visit or an explicit "switch business") - not necessarily
    // the very first business ever created under this account.
    let matchingRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && data[i][2] === d.password) {
        matchingRows.push({ rowIndex: i + 1, row: data[i] });
      }
    }
    
    if (matchingRows.length > 0) {
      matchingRows.sort((a, b) => new Date(b.row[5] || 0) - new Date(a.row[5] || 0));
      const active = matchingRows[0];
      const i = active.rowIndex - 1; // keep the rest of this block unchanged below
      {
        const userType = active.row[7] || "admin";
        masterSheet.getRange(i + 1, 6).setValue(new Date());
        
        const business = active.row[0]; // Jina kutoka master database
        const databaseId = active.row[3];
        const logoId = active.row[6];
        
        const settings = getUserSettingsFromDB(databaseId);
        const businessTimeZone = ensureBusinessTimeZone_(databaseId, d.clientTimeZone);
        if (!settings.PROFILE) settings.PROFILE = {};
        settings.PROFILE.time_zone = { value: businessTimeZone, updated: new Date() };
        const token = createSession(d.email, databaseId, "admin", null, null);
        
        let logoUrl = DEFAULT_LOGO_URL;
        let altLogoUrl = DEFAULT_LOGO_URL;
        let directLogoUrl = DEFAULT_LOGO_URL;
        
        // Prefer the base64 logo stored in THIS business's own Settings -
        // no external Drive fetch needed, so it can never silently fail
        // to show up.
        const logoInfo = getLogoBase64FromSettings(settings);
        if (logoInfo) {
          const dataUri = "data:" + logoInfo.mimeType + ";base64," + logoInfo.base64;
          logoUrl = dataUri;
          altLogoUrl = dataUri;
          directLogoUrl = dataUri;
        } else if (logoId) {
          try {
            const file = DriveApp.getFileById(logoId);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            
            logoUrl = LOGO_BASE_URL + logoId;
            altLogoUrl = LOGO_ALT_URL + logoId;
            directLogoUrl = LOGO_DIRECT_URL + logoId;
          } catch (e) {
            masterSheet.getRange(i + 1, 7).clear();
          }
        }
        
        // One account (same email+password) can own several businesses,
        // each with its own completely separate database - gather all of
        // them here so the client can offer the business-switcher menu
        // right after login, without a second round trip.
        const businesses = [];
        for (let k = 1; k < data.length; k++) {
          if (data[k][1] === d.email && (data[k][7] === "admin" || !data[k][7])) {
            businesses.push({ business: data[k][0], databaseId: data[k][3] });
          }
        }
        
        return output({
          success: true,
          business: business, // Jina kutoka master database
          email: d.email,
          token: token,
          databaseId: databaseId,
          logoUrl: logoUrl,
          altLogoUrl: altLogoUrl,
          directLogoUrl: directLogoUrl,
          settings: settings,
          defaultLogoUrl: DEFAULT_LOGO_URL,
          userType: "admin",
          products: getProductsForSync(databaseId),
          businesses: businesses
        });
      }
    }
    
    // Check for seller - authenticate ENTIRELY against the master registry
    // (email, password, status all live there now) so the business's own
    // database is never opened just to check credentials. It's opened
    // exactly once, only after a successful match, to fetch settings and
    // products for the login response.
    try {
      const masterSellersSheet = ensureMasterSellersSheet(masterDB);
      const masterSellersData = masterSellersSheet.getDataRange().getValues();
      for (let m = 1; m < masterSellersData.length; m++) {
        if (masterSellersData[m][0] === d.email) {
          const databaseId = masterSellersData[m][3];
          const sellerId = masterSellersData[m][2];
          const registryPassword = masterSellersData[m][5];
          const sellerName = masterSellersData[m][6] || "";
          const sellerStatus = masterSellersData[m][7] || "active";
          const sellerBranchId = masterSellersData[m][8] || "MAIN";
          
          if (registryPassword !== d.password) {
            return output({ success: false, error: "Invalid email or password. Please try again." });
          }
          if (sellerStatus !== "active") {
            return output({ success: false, error: "Your account is inactive. Please contact administrator." });
          }
          
          const userDB = SpreadsheetApp.openById(databaseId);
          
          // Best-effort "last active" stamp - never blocks or fails login.
          try {
            const sellersSheet = userDB.getSheetByName("SELLERS");
            if (sellersSheet) {
              const sellersData = sellersSheet.getDataRange().getValues();
              for (let j = 1; j < sellersData.length; j++) {
                if (sellersData[j][0] === sellerId) {
                  sellersSheet.getRange(j + 1, 8).setValue(new Date());
                  break;
                }
              }
            }
          } catch (e) {
            console.log("Could not update seller's last active timestamp:", e);
          }
          
          let business = "";
          for (let bi = 1; bi < data.length; bi++) {
            if (data[bi][3] === databaseId) {
              business = data[bi][0];
              break;
            }
          }
          
          const settings = getUserSettingsFromDB(databaseId);
          const businessTimeZone = ensureBusinessTimeZone_(databaseId, d.clientTimeZone);
          if (!settings.PROFILE) settings.PROFILE = {};
          settings.PROFILE.time_zone = { value: businessTimeZone, updated: new Date() };
          const token = createSession(d.email, databaseId, "seller", sellerId, sellerBranchId);
          
          return output({
            success: true,
            business: business,
            email: d.email,
            token: token,
            databaseId: databaseId,
            logoUrl: DEFAULT_LOGO_URL,
            altLogoUrl: DEFAULT_LOGO_URL,
            directLogoUrl: DEFAULT_LOGO_URL,
            settings: settings,
            defaultLogoUrl: DEFAULT_LOGO_URL,
            userType: "seller",
            sellerId: sellerId,
            sellerName: sellerName,
            branchId: sellerBranchId,
            branchName: getBranchNameFromDB(userDB, sellerBranchId),
            products: getProductsForSync(databaseId, sellerBranchId)
          });
        }
      }
    } catch (e) {
      console.log("Master seller registry lookup failed:", e);
    }
    
    return output({ success: false, error: "Invalid email or password. Please try again." });
    
  } catch (error) {
    return output({ success: false, error: "Login failed: " + error.toString() });
  }
}


/* ===== FORGOT PASSWORD ===== */
function forgotPassword(d) {
  try {
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    
    const data = masterSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && data[i][7] === "admin") {
        const password = data[i][2];
        const business = data[i][0];
        
        try {
          MailApp.sendEmail({
            to: d.email,
            subject: "SmartTrade Pro - Password Recovery",
            htmlBody: createPasswordEmailTemplate(business, password)
          });
          
          return output({ success: true, message: "Password has been sent to your email address." });
        } catch (emailError) {
          return output({ success: false, error: "Could not send email. Please contact support." });
        }
      }
    }
    
    return output({ success: false, error: "Email not found in our system." });
    
  } catch (error) {
    return output({ success: false, error: "Password recovery failed: " + error.toString() });
  }
}

/* ===== GET USER SETTINGS ===== */
function getUserSettings(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const settings = getUserSettingsFromDB(d.databaseId);
    let logoUrl = DEFAULT_LOGO_URL;
    
    const logoInfo = getLogoBase64FromSettings(settings);
    if (logoInfo) {
      logoUrl = "data:" + logoInfo.mimeType + ";base64," + logoInfo.base64;
    } else {
      const logoId = getUserLogoIdFromMaster(d.email, d.databaseId);
      if (logoId) {
        logoUrl = LOGO_BASE_URL + logoId;
      }
    }
    
    return output({
      success: true,
      settings: settings,
      logoUrl: logoUrl,
      defaultLogoUrl: DEFAULT_LOGO_URL
    });
    
  } catch (error) {
    return output({ success: false, error: "Failed to get settings: " + error.toString() });
  }
}

/* ===== GET USER SETTINGS FROM DB ===== */
function getUserSettingsFromDB(databaseId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    
    if (!settingsSheet) {
      return {};
    }
    
    const data = settingsSheet.getDataRange().getValues();
    const settings = {};
    
    for (let i = 1; i < data.length; i++) {
      const [section, key, value] = data[i];
      
      if (!settings[section]) {
        settings[section] = {};
      }
      
      settings[section][key] = { value: value, updated: data[i][3] };
    }
    
    // The system default language changed from Swahili to English - a
    // DISPLAY.language value saved before that change reflects the OLD
    // default the system silently picked for everyone, not a genuine
    // choice this person ever made themselves. Treating it as unset here
    // lets them see English (the new default) until they explicitly pick
    // a language on the Settings page - which saves a fresh, later
    // timestamp and is then always honored, exactly like any other
    // deliberate choice.
    if (settings.DISPLAY && settings.DISPLAY.language) {
      const languageChangedAt = new Date(settings.DISPLAY.language.updated).getTime();
      const languageDefaultCutoff = new Date("2026-08-30T00:00:00Z").getTime();
      if (isNaN(languageChangedAt) || languageChangedAt < languageDefaultCutoff) {
        delete settings.DISPLAY.language;
      }
    }
    
    return settings;
    
  } catch (error) {
    console.error("Error getting settings:", error);
    return {};
  }
}

/* ===== UPDATE USER SETTINGS ===== */
function updateUserSettings(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    
    if (!settingsSheet) {
      return output({ success: false, error: "Settings sheet not found" });
    }
    
    const data = settingsSheet.getDataRange().getValues();
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.section && data[i][1] === d.key) {
        settingsSheet.getRange(i + 1, 3).setValue(d.value);
        settingsSheet.getRange(i + 1, 4).setValue(new Date());
        found = true;
        break;
      }
    }
    
    if (!found) {
      settingsSheet.appendRow([d.section, d.key, d.value, new Date()]);
    }
    
    if (d.section === "PROFILE" && d.key === "business_name") {
      updateMasterBusinessName(d.email, d.databaseId, d.value);
    }
    
    return output({ success: true, message: "Settings updated successfully" });
    
  } catch (error) {
    return output({ success: false, error: "Failed to update settings: " + error.toString() });
  }
}

function updateMasterBusinessName(email, databaseId, businessName) {
  try {
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    // Match by email AND databaseId - one account can own several
    // businesses (multi-business support) that all share the same email,
    // so matching by email alone would always rename whichever business
    // was created FIRST instead of the one currently switched into.
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email && data[i][3] === databaseId) {
        masterSheet.getRange(i + 1, 1).setValue(businessName);
        break;
      }
    }
  } catch (error) {
    console.error("Error updating master business name:", error);
  }
}

/* ===== UPDATE USER PROFILE ===== */
function updateUserProfile(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && data[i][3] === d.databaseId) {
        if (d.business) {
          masterSheet.getRange(i + 1, 1).setValue(d.business);
        }
        break;
      }
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    
    if (settingsSheet) {
      if (d.business) {
        updateSettingsKey(settingsSheet, "PROFILE", "business_name", d.business);
      }
      if (d.phone) {
        updateSettingsKey(settingsSheet, "PROFILE", "phone", d.phone);
      }
      if (d.address) {
        updateSettingsKey(settingsSheet, "PROFILE", "address", d.address);
      }
    }
    
    return output({ success: true, message: "Profile updated successfully" });
    
  } catch (error) {
    return output({ success: false, error: "Failed to update profile: " + error.toString() });
  }
}

function updateSettingsKey(settingsSheet, section, key, value) {
  const data = settingsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === section && data[i][1] === key) {
      settingsSheet.getRange(i + 1, 3).setValue(value);
      settingsSheet.getRange(i + 1, 4).setValue(new Date());
      return true;
    }
  }
  
  settingsSheet.appendRow([section, key, value, new Date()]);
  return true;
}

/* ============================================
   LOGO STORAGE - BASE64, PER BUSINESS (in that business's own SETTINGS
   sheet, not the shared master database). This replaces depending on
   Google Drive at PDF/report-generation time (fetching a Drive file
   there was the main reason the logo sometimes silently failed to show
   up at all) - the image bytes are simply already sitting in Settings,
   ready to embed directly wherever needed.
   
   A single Sheets cell holds at most ~50,000 characters, so a base64
   string is split across several rows ("logo_base64_0", "logo_base64_1",
   ...) and reassembled on read - this means an image is NEVER truncated
   or corrupted regardless of size, it just uses more rows.
   ============================================ */

const LOGO_CELL_CHUNK_SIZE = 49000;

function setLogoBase64InSettings(settingsSheet, base64Data, mimeType) {
  clearLogoBase64FromSettings(settingsSheet);
  
  const chunks = [];
  for (let i = 0; i < base64Data.length; i += LOGO_CELL_CHUNK_SIZE) {
    chunks.push(base64Data.substring(i, i + LOGO_CELL_CHUNK_SIZE));
  }
  
  updateSettingsKey(settingsSheet, "PROFILE", "logo_mimetype", mimeType);
  updateSettingsKey(settingsSheet, "PROFILE", "logo_chunks", String(chunks.length));
  chunks.forEach((chunk, idx) => {
    updateSettingsKey(settingsSheet, "PROFILE", "logo_base64_" + idx, chunk);
  });
}

function clearLogoBase64FromSettings(settingsSheet) {
  const data = settingsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const key = String(data[i][1]);
    if (data[i][0] === "PROFILE" && (key === "logo_mimetype" || key === "logo_chunks" || key.indexOf("logo_base64_") === 0)) {
      settingsSheet.deleteRow(i + 1);
    }
  }
}

// Reads the (possibly chunked) base64 logo out of an already-fetched
// settings object (as returned by getUserSettingsFromDB) - no extra
// sheet reads needed since the caller almost always has this already.
function getLogoBase64FromSettings(settings) {
  if (!settings || !settings.PROFILE) return null;
  
  const chunkCountStr = settings.PROFILE.logo_chunks?.value;
  const mimeType = settings.PROFILE.logo_mimetype?.value || "image/jpeg";
  
  if (!chunkCountStr) return null;
  
  const chunkCount = parseInt(chunkCountStr, 10) || 0;
  let combined = "";
  for (let i = 0; i < chunkCount; i++) {
    combined += settings.PROFILE["logo_base64_" + i]?.value || "";
  }
  
  return combined ? { base64: combined, mimeType: mimeType } : null;
}

function uploadUserLogoBase64(d) {
  try {
    if (!d.databaseId || !d.logoData) {
      return output({ success: false, error: "Database ID and logo data required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    if (!settingsSheet) {
      return output({ success: false, error: "Settings sheet not found" });
    }
    
    const mimeType = d.mimeType || "image/jpeg";
    setLogoBase64InSettings(settingsSheet, d.logoData, mimeType);
    
    return output({
      success: true,
      message: "Logo uploaded successfully",
      logoUrl: "data:" + mimeType + ";base64," + d.logoData
    });
  } catch (error) {
    return output({ success: false, error: "Failed to upload logo: " + error.toString() });
  }
}

function deleteUserLogoBase64(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    if (settingsSheet) {
      clearLogoBase64FromSettings(settingsSheet);
    }
    
    return output({ success: true, message: "Logo deleted successfully", logoUrl: DEFAULT_LOGO_URL });
  } catch (error) {
    return output({ success: false, error: "Failed to delete logo: " + error.toString() });
  }
}

/* ===== UPLOAD USER LOGO USING DRIVE API ONLY (REPLACE EXISTING) ===== */
function uploadUserLogo(d) {
  try {
    if (!d.email || !d.logoData) {
      return output({ success: false, error: "Email and logo data required" });
    }
    
    // Get existing logo ID
    const existingLogoId = getUserLogoIdFromMaster(d.email, d.databaseId);
    
    let fileId = null;
    let fileName = `logo_${d.email.replace(/[@.]/g, "_")}.${d.fileExtension || 'png'}`;
    const blob = Utilities.newBlob(Utilities.base64Decode(d.logoData), d.mimeType, fileName);
    
    if (existingLogoId) {
      // REPLACE existing file - update content
      try {
        // First, check if file exists
        const existingFile = Drive.Files.get(existingLogoId);
        if (existingFile && existingFile.id) {
          // Update file content (replace)
          Drive.Files.update({
            'title': fileName,
            'mimeType': d.mimeType
          }, existingLogoId, blob, {
            'convert': false
          });
          fileId = existingLogoId;
          console.log("Existing logo replaced:", existingLogoId);
        } else {
          // File doesn't exist, create new
          const fileMetadata = {
            'title': fileName,
            'mimeType': d.mimeType,
            'parents': [{'id': LOGO_FOLDER_ID}]
          };
          const newFile = Drive.Files.insert(fileMetadata, blob, { 'convert': false });
          fileId = newFile.id;
          console.log("New logo created (existing not found):", fileId);
        }
      } catch (e) {
        // Error accessing existing file, create new
        console.log("Error accessing existing logo, creating new:", e.toString());
        const fileMetadata = {
          'title': fileName,
          'mimeType': d.mimeType,
          'parents': [{'id': LOGO_FOLDER_ID}]
        };
        const newFile = Drive.Files.insert(fileMetadata, blob, { 'convert': false });
        fileId = newFile.id;
        console.log("New logo created after error:", fileId);
      }
    } else {
      // No existing logo, create new file
      const fileMetadata = {
        'title': fileName,
        'mimeType': d.mimeType,
        'parents': [{'id': LOGO_FOLDER_ID}]
      };
      const newFile = Drive.Files.insert(fileMetadata, blob, { 'convert': false });
      fileId = newFile.id;
      console.log("New logo created (first time):", fileId);
    }
    
    // Set sharing permissions (always ensure public access)
    try {
      Drive.Permissions.insert({
        'role': 'reader',
        'type': 'anyone'
      }, fileId, {
        'sendNotificationEmails': 'false'
      });
    } catch (permError) {
      console.log("Permission error (might already exist):", permError.toString());
    }
    
    // Wait a moment for permissions to apply
    Utilities.sleep(1000);
    
    // Update master database with logo ID (same or new)
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    let userFound = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && (!d.databaseId || data[i][3] === d.databaseId)) {
        masterSheet.getRange(i + 1, 7).setValue(fileId);
        userFound = true;
        console.log("Master database updated with logo ID:", fileId);
        break;
      }
    }
    
    if (!userFound) {
      console.log("User not found in master database:", d.email);
    }
    
    // Generate URLs
    const logoUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=s200";
    const altLogoUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
    const directLogoUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    
    return output({
      success: true,
      message: existingLogoId ? "Logo replaced successfully" : "Logo uploaded successfully",
      logoUrl: logoUrl,
      altLogoUrl: altLogoUrl,
      directLogoUrl: directLogoUrl,
      logoId: fileId,
      replaced: existingLogoId ? true : false
    });
    
  } catch (error) {
    console.error("Upload error:", error.toString());
    return output({ success: false, error: "Failed to upload logo: " + error.toString() });
  }
}

/* ===== GET USER LOGO - prefers base64 stored in this business's own
   Settings, falls back to Drive for accounts that haven't re-uploaded
   since this feature was added. ===== */
function getUserLogo(d) {
  try {
    if (d.databaseId) {
      const settings = getUserSettingsFromDB(d.databaseId);
      const logoInfo = getLogoBase64FromSettings(settings);
      if (logoInfo) {
        const dataUri = "data:" + logoInfo.mimeType + ";base64," + logoInfo.base64;
        return output({
          success: true,
          logoUrl: dataUri,
          altLogoUrl: dataUri,
          directLogoUrl: dataUri,
          defaultLogoUrl: DEFAULT_LOGO_URL
        });
      }
    }
    
    const logoId = getUserLogoIdFromMaster(d.email, d.databaseId);
    let logoUrl = DEFAULT_LOGO_URL;
    let altLogoUrl = DEFAULT_LOGO_URL;
    let directLogoUrl = DEFAULT_LOGO_URL;
    
    if (logoId) {
      try {
        const file = Drive.Files.get(logoId);
        if (file && file.id) {
          // Use the most reliable URL format
          logoUrl = "https://drive.google.com/uc?export=view&id=" + logoId;
          altLogoUrl = "https://drive.google.com/thumbnail?id=" + logoId + "&sz=s200";
          directLogoUrl = "https://lh3.googleusercontent.com/d/" + logoId;
          
          console.log("Logo found:", file.title);
          
          return output({
            success: true,
            logoUrl: logoUrl,
            altLogoUrl: altLogoUrl,
            directLogoUrl: directLogoUrl,
            logoId: logoId,
            fileName: file.title,
            defaultLogoUrl: DEFAULT_LOGO_URL
          });
        }
      } catch (e) {
        console.log("Logo file not found: " + e.toString());
        // Clear invalid logo ID
        const masterDB = getMasterDatabase();
        const masterSheet = masterDB.getSheetByName("MASTER_USERS");
        const data = masterSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][1] === d.email) {
            masterSheet.getRange(i + 1, 7).clear();
            break;
          }
        }
      }
    }
    
    return output({
      success: true,
      logoUrl: DEFAULT_LOGO_URL,
      altLogoUrl: DEFAULT_LOGO_URL,
      directLogoUrl: DEFAULT_LOGO_URL,
      defaultLogoUrl: DEFAULT_LOGO_URL
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== DELETE USER LOGO USING DRIVE API ONLY ===== */
function deleteUserLogo(d) {
  try {
    const logoId = getUserLogoIdFromMaster(d.email, d.databaseId);
    
    if (!logoId) {
      return output({ success: true, message: "No logo to delete" });
    }
    
    try {
      // Delete file using Drive API
      Drive.Files.remove(logoId);
    } catch (e) {
      console.log("File already deleted or not found: " + e.toString());
    }
    
    // Update master database
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && (!d.databaseId || data[i][3] === d.databaseId)) {
        masterSheet.getRange(i + 1, 7).clear();
        break;
      }
    }
    
    return output({
      success: true,
      message: "Logo deleted successfully",
      logoUrl: DEFAULT_LOGO_URL
    });
    
  } catch (error) {
    return output({ success: false, error: "Failed to delete logo: " + error.toString() });
  }
}

/* ===== ADD STOCK (with consolidation) ===== */
// Bulk-imports many products at once from the Stock page's Excel
// template feature - one batched write for all rows (much faster than
// looping addStock()/appendRow() one at a time). Merge rules:
//  - Same Product+Unit+Branch already exists -> ADD the template quantity
//    onto the existing quantity (0 in the template = leave quantity
//    unchanged), and ALWAYS keep the existing Buy/Sell price (the
//    template's price is ignored for an existing entry, whether it
//    differs or is blank).
//  - No match -> create a brand-new entry using the template's own
//    quantity (even 0) and price.
// Also registers any brand-new Units/Categories encountered.
function bulkAddStock(d) {
  try {
    if (!d.databaseId || !Array.isArray(d.items) || d.items.length === 0) {
      return output({ success: false, error: "No items to import" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    const timestamp = new Date();
    const data = stockSheet.getDataRange().getValues();
    
    // Rows already on the sheet, keyed by Product+Unit+Branch.
    const existingIndex = {};
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const key = data[i][1] + "|||" + (data[i][9] || "") + "|||" + (data[i][8] || "MAIN");
      existingIndex[key] = {
        rowNumber: i + 1,
        quantity: data[i][2] || 0,
        buyPrice: data[i][3] || 0,
        category: data[i][6] || "",
        dirty: false
      };
    }
    
    // Brand-new rows to append this run, keyed the same way, so several
    // template rows for the SAME new Product+Unit combine into one entry
    // instead of creating duplicates.
    const pendingNew = {};
    const newUnits = {};
    const newCategories = {};
    
    d.items.forEach(function(item) {
      const branchId = item.branchId || d.branchId || "MAIN";
      const unit = String(item.unit || "").trim();
      const category = String(item.category || "").trim();
      const templateQty = Number(item.quantity) || 0;
      const templateBuyPrice = Number(item.buyPrice) || 0;
      const templateSellPrice = Number(item.sellPrice) || 0;
      const key = item.product + "|||" + unit + "|||" + branchId;
      const addedQty = templateQty > 0 ? templateQty : 0;
      
      if (existingIndex[key]) {
        const existing = existingIndex[key];
        existing.quantity += addedQty;
        if (!existing.category && category) existing.category = category;
        existing.dirty = true;
      } else if (pendingNew[key]) {
        pendingNew[key].quantity += addedQty;
        if (!pendingNew[key].category && category) pendingNew[key].category = category;
      } else {
        pendingNew[key] = {
          product: item.product,
          unit: unit,
          branchId: branchId,
          quantity: templateQty,
          buyPrice: templateBuyPrice,
          sellPrice: templateSellPrice,
          category: category
        };
      }
      
      if (unit) newUnits[unit] = true;
      if (category) newCategories[category] = true;
    });
    
    // Write merges into their existing rows.
    let mergedCount = 0;
    Object.keys(existingIndex).forEach(function(key) {
      const existing = existingIndex[key];
      if (existing.dirty) {
        stockSheet.getRange(existing.rowNumber, 3).setValue(existing.quantity);
        stockSheet.getRange(existing.rowNumber, 6).setValue(existing.quantity * existing.buyPrice);
        if (existing.category) stockSheet.getRange(existing.rowNumber, 7).setValue(existing.category);
        mergedCount++;
      }
    });
    
    // Append every brand-new row in ONE batched write.
    const newRows = Object.values(pendingNew).map(function(p) {
      return [timestamp, p.product, p.quantity, p.buyPrice, p.sellPrice, p.quantity * p.buyPrice, p.category, "", p.branchId, p.unit];
    });
    if (newRows.length > 0) {
      const startRow = stockSheet.getLastRow() + 1;
      stockSheet.getRange(startRow, 1, newRows.length, 10).setValues(newRows);
    }
    
    // Register any brand-new units/categories encountered, skipping ones
    // that already exist.
    const unitsSheet = ensureUnitsSheet(userDB);
    const existingUnitsData = unitsSheet.getDataRange().getValues();
    const existingUnits = existingUnitsData.slice(1).map(function(r) { return r[0]; });
    const unitsToAdd = Object.keys(newUnits).filter(function(u) { return existingUnits.indexOf(u) === -1; });
    if (unitsToAdd.length > 0) {
      unitsSheet.getRange(unitsSheet.getLastRow() + 1, 1, unitsToAdd.length, 1).setValues(unitsToAdd.map(function(u) { return [u]; }));
    }
    
    const categoriesSheet = ensureCategoriesSheet(userDB);
    const existingCategoriesData = categoriesSheet.getDataRange().getValues();
    const existingCategories = existingCategoriesData.slice(1).map(function(r) { return r[0]; });
    const categoriesToAdd = Object.keys(newCategories).filter(function(c) { return existingCategories.indexOf(c) === -1; });
    if (categoriesToAdd.length > 0) {
      categoriesSheet.getRange(categoriesSheet.getLastRow() + 1, 1, categoriesToAdd.length, 1).setValues(categoriesToAdd.map(function(c) { return [c]; }));
    }
    
    return output({
      success: true,
      added: newRows.length,
      merged: mergedCount,
      message: newRows.length + " new, " + mergedCount + " updated"
    });
  } catch (error) {
    return output({ success: false, error: "Bulk import failed: " + error.toString() });
  }
}

function addStock(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    const timestamp = new Date();
    const branchId = d.branchId || "MAIN";
    const product = d.product;
    const newQuantity = d.qty;
    const newBuyPrice = d.buy;
    const newSellPrice = d.sell;
    const category = d.category || "";
    const unit = d.unit || "";
    const sellerId = d.sellerId || null;
    
    // Check if this product already exists in this branch WITH THE SAME
    // UNIT - "Rice (Kg)" and "Rice (Bags)" must never be merged together,
    // since their quantities aren't comparable. A legacy row that has no
    // unit recorded yet is still allowed to match/backfill once, for
    // backward compatibility with stock added before the Unit feature.
    const data = stockSheet.getDataRange().getValues();
    let existingRow = -1;
    let existingQuantity = 0;
    let existingBuyPrice = 0;
    let existingTotalValue = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1] === product) {
        const itemBranchId = data[i][8] || "MAIN";
        const itemUnit = data[i][9] || "";
        if (itemBranchId === branchId && (itemUnit === unit || !itemUnit)) {
          existingRow = i + 1; // Row number (1-based)
          existingQuantity = data[i][2] || 0;
          existingBuyPrice = data[i][3] || 0;
          existingTotalValue = data[i][5] || 0;
          break;
        }
      }
    }
    
    if (existingRow !== -1) {
      // Product exists in this branch - UPDATE existing row
      
      // Calculate new weighted average buy price
      const oldTotalValue = existingQuantity * existingBuyPrice;
      const newTotalValue = newQuantity * newBuyPrice;
      const totalQuantity = existingQuantity + newQuantity;
      const avgBuyPrice = (oldTotalValue + newTotalValue) / totalQuantity;
      
      // Update the existing row
      stockSheet.getRange(existingRow, 3).setValue(totalQuantity);           // New quantity
      stockSheet.getRange(existingRow, 4).setValue(avgBuyPrice);            // Weighted average buy price
      stockSheet.getRange(existingRow, 5).setValue(newSellPrice);           // Latest sell price
      stockSheet.getRange(existingRow, 6).setValue(totalQuantity * avgBuyPrice); // New total value
      
      // Update category if provided
      if (category) {
        stockSheet.getRange(existingRow, 7).setValue(category);
      }
      
      // Update/fill in the Unit (also backfills old rows that never had one)
      if (unit) {
        stockSheet.getRange(existingRow, 10).setValue(unit);
      }
      
      // Update timestamp (Last Updated changes)
      stockSheet.getRange(existingRow, 1).setValue(timestamp);
      
      return output({ 
        success: true, 
        message: `Added ${newQuantity} units to existing ${product} stock at ${branchId} branch. Total now: ${totalQuantity}`,
        consolidated: true,
        oldQuantity: existingQuantity,
        newQuantity: totalQuantity,
        timestamp: timestamp.toISOString()
      });
      
    } else {
      // New product in this branch - CREATE new row
      const stockValue = newQuantity * newBuyPrice;
      
      // Ensure sheet has all columns
      const headers = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
      if (headers.length < 9) {
        stockSheet.getRange(1, 8).setValue("Seller ID");
        stockSheet.getRange(1, 9).setValue("Branch ID");
      }
      if (headers.length < 10) {
        stockSheet.getRange(1, 10).setValue("Unit");
      }
      
      stockSheet.appendRow([
        timestamp,
        product,
        newQuantity,
        newBuyPrice,
        newSellPrice,
        stockValue,
        category,
        sellerId,
        branchId,
        unit
      ]);
      
      return output({ 
        success: true, 
        message: `Added ${newQuantity} units of ${product} to stock at ${branchId} branch`,
        consolidated: false,
        stockValue: stockValue,
        timestamp: timestamp.toISOString()
      });
    }
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}
/* ===== GET STOCK (by branch) ===== */
function getStock(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    const branchId = d.branchId || null; // For sellers
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    const data = stockSheet.getDataRange().getValues();
    const stockData = [];
    let totalValue = 0;
    let totalItems = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && i > 0) {
        const itemBranchId = data[i][8] || "MAIN";
        
        // Filter by branch if seller
        if (branchId && itemBranchId !== branchId) {
          continue;
        }
        
        stockData.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          buyPrice: data[i][3],
          sellPrice: data[i][4],
          value: data[i][5],
          category: data[i][6],
          sellerId: data[i][7] || null,
          branchId: itemBranchId,
          unit: data[i][9] || ""
        });
        
        totalValue += (data[i][5] || 0);
        totalItems += (data[i][2] || 0);
      }
    }
    
    return output({
      success: true,
      data: stockData,
      summary: {
        totalItems: totalItems,
        totalValue: totalValue,
        count: stockData.length
      }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== SALES RECEIPT NUMBER HELPERS ===== */
function generateSalesReceiptNo_() {
  var now = new Date();
  var tz = getCurrentBusinessTimeZone_();
  var stamp = Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');
  return 'ST-' + stamp + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
}

function ensureSalesReceiptNumbers_(salesSheet) {
  if (!salesSheet) return;
  var lastCol = salesSheet.getLastColumn();
  if (lastCol < 11) {
    salesSheet.getRange(1, 11).setValue('Receipt No');
  }
  if (salesSheet.getLastColumn() < 12) {
    salesSheet.getRange(1, 12).setValue('Operation ID');
  }
  var lastRow = salesSheet.getLastRow();
  if (lastRow < 2) return;

  var values = salesSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var missing = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] && !values[i][10]) {
      missing.push({ row: i + 2, receiptNo: generateSalesReceiptNo_() });
    }
  }
  if (missing.length) {
    missing.forEach(function(item) {
      salesSheet.getRange(item.row, 11).setValue(item.receiptNo);
    });
  }
}

/* ===== GET SALES (BY BRANCH / SELLER) ===== */
function getSales(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const salesSheet = userDB.getSheetByName("SALES");
    
    if (!salesSheet) {
      return output({ success: false, error: "Sales sheet not found" });
    }
    
    ensureSalesReceiptNumbers_(salesSheet);
    const data = salesSheet.getDataRange().getValues();
    const operationalBoundaries = getOperationalOpenPeriodStarts_(userDB);

    // Include the actual Seller Name in direct getSales responses.
    const sellerMap = {};
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (sellersSheet) {
      const sellersData = sellersSheet.getDataRange().getValues();
      for (let i = 1; i < sellersData.length; i++) {
        if (sellersData[i][0]) sellerMap[String(sellersData[i][0])] = sellersData[i][1] || "";
      }
    }

    const salesData = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const rowSellerId = data[i][7] || null;
        const rowBranchId = data[i][8] || "MAIN";
        
        // Optional filters (sellers viewing their own sales, or a specific branch)
        if (d.sellerId && rowSellerId !== d.sellerId) continue;
        if (d.branchId && rowBranchId !== d.branchId) continue;

        // All Sales is an operational view. Once Accounts are closed,
        // anything at or before the exact Closed At timestamp belongs to
        // the historical period and must not appear here.
        const saleTime = new Date(data[i][0]).getTime();
        if (isNaN(saleTime)) continue;
        const openPeriodStart = getOperationalOpenPeriodStartFromMap_(operationalBoundaries, rowBranchId);
        if (openPeriodStart !== null && saleTime <= openPeriodStart) continue;
        
        salesData.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          unitPrice: data[i][3],
          total: data[i][4],
          customer: data[i][5],
          paymentMethod: data[i][6],
          sellerId: rowSellerId,
          sellerName: rowSellerId ? (sellerMap[String(rowSellerId)] || "Unknown Seller") : "Admin",
          branchId: rowBranchId,
          unit: data[i][9] || "",
          receiptNo: data[i][10] || "",
          operationId: data[i][11] || "",
          rowIndex: i + 1
        });
      }
    }
    
    return output({
      success: true,
      data: salesData
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== SALE IDEMPOTENCY / ATOMIC STOCK HELPERS ===== */
function findExistingSaleByOperation_(salesSheet, operationId, receiptNo, product, branchId, unit) {
  if (!salesSheet || salesSheet.getLastRow() < 2) return null;
  var lastCol = Math.max(12, salesSheet.getLastColumn());
  var data = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowReceipt = String(data[i][10] || '');
    var rowOperation = String(data[i][11] || '');
    var rowProduct = String(data[i][1] || '');
    var rowBranch = String(data[i][8] || 'MAIN');
    var rowUnit = String(data[i][9] || '');
    if (operationId && rowOperation && rowOperation === String(operationId)) {
      return { rowIndex: i + 2, receiptNo: rowReceipt || receiptNo || '', alreadyProcessed: true };
    }
    // Backward-compatible idempotency for offline records created before
    // Operation ID was introduced.  One receipt may contain several products,
    // so the fallback key includes product + branch + unit.
    if (!operationId && receiptNo && rowReceipt === String(receiptNo) &&
        rowProduct === String(product || '') && rowBranch === String(branchId || 'MAIN') &&
        rowUnit === String(unit || '')) {
      return { rowIndex: i + 2, receiptNo: rowReceipt, alreadyProcessed: true };
    }
  }
  return null;
}

function ensureSalesColumns_(salesSheet) {
  if (!salesSheet) throw new Error('Sales sheet not found');
  if (salesSheet.getLastColumn() < 9) {
    salesSheet.getRange(1, 8).setValue('Seller ID');
    salesSheet.getRange(1, 9).setValue('Branch ID');
  }
  if (salesSheet.getLastColumn() < 10) salesSheet.getRange(1, 10).setValue('Unit');
  if (salesSheet.getLastColumn() < 11) salesSheet.getRange(1, 11).setValue('Receipt No');
  if (salesSheet.getLastColumn() < 12) salesSheet.getRange(1, 12).setValue('Operation ID');
}

/* ===== ADD SALE (with branch validation) ===== */
function addSale(d) {
  var lock = LockService.getScriptLock();
  try {
    if (!d.databaseId) return output({ success: false, error: "Database ID required" });
    lock.waitLock(30000);

    const branchId = d.branchId || "MAIN";
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const salesSheet = userDB.getSheetByName("SALES");
    if (!salesSheet) return output({ success: false, error: "Sales sheet not found" });
    ensureSalesColumns_(salesSheet);

    const receiptNo = d.receiptNo || generateSalesReceiptNo_();
    const operationId = d.operationId || null;

    // CRITICAL: check for an already-committed request BEFORE touching stock.
    // If the browser retries because the previous response was lost, this
    // returns success without appending another row or reducing stock again.
    const existing = findExistingSaleByOperation_(salesSheet, operationId, receiptNo, d.product, branchId, d.unit);
    if (existing) {
      return output({
        success: true,
        alreadyProcessed: true,
        receiptNo: existing.receiptNo || receiptNo,
        rowIndex: existing.rowIndex,
        message: "Sale was already recorded; duplicate request ignored."
      });
    }

    const qty = Number(d.qty);
    if (!isFinite(qty) || qty <= 0) {
      return output({ success: false, error: "Invalid sale quantity" });
    }

    // Stock validation and the actual stock update happen under the SAME
    // script lock as the duplicate check and sale insert.
    if (d.isStockProduct) {
      const stockCheck = checkBranchStockAvailability(d.databaseId, d.product, qty, branchId, d.unit);
      if (!stockCheck.available) {
        return output({
          success: false,
          error: `Insufficient stock at ${branchId} branch. Available: ${stockCheck.currentStock}, Requested: ${qty}`,
          currentStock: stockCheck.currentStock
        });
      }
    }

    const timestamp = new Date();
    const total = Number(d.total) || 0;
    const unitPrice = Number(d.unitPrice) || (total / qty);
    const newRow = salesSheet.getLastRow() + 1;

    salesSheet.getRange(newRow, 1, 1, 12).setValues([[
      timestamp, d.product, qty, unitPrice, total,
      d.customer || "Walk-in", d.paymentMethod || "cash",
      d.sellerId || null, branchId, d.unit || "", receiptNo, operationId || ""
    ]]);

    try {
      if (d.isStockProduct) {
        updateStockAfterSale(d.databaseId, d.product, qty, branchId, d.unit);
      } else {
        upsertNonStockProduct(userDB, d.product, d.unit);
      }
    } catch (stockError) {
      // Do not leave a sale without its corresponding stock movement.
      try { salesSheet.deleteRow(newRow); } catch (rollbackError) {}
      throw stockError;
    }

    if (d.sellerId) updateSellerTotals(d.databaseId, d.sellerId, total, 'sale');

    return output({
      success: true,
      message: `Sale recorded: ${qty} x ${d.product} at ${branchId} branch`,
      total: total,
      timestamp: timestamp.toISOString(),
      receiptNo: receiptNo,
      operationId: operationId,
      rowIndex: newRow,
      alreadyProcessed: false
    });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ===== ADD PURCHASE (WITH AUTO STOCK UPDATE) ===== */
/* ============================================
   PURCHASE REQUESTS (Seller -> Admin approval workflow)
   Sellers request items they need; Admin reviews, can adjust
   quantities/prices, and approves; Seller then completes items
   (one at a time or all together) which adds them straight to Stock.
   ============================================ */

function ensurePurchaseRequestsSheet(userDB) {
  var sheet = userDB.getSheetByName("PURCHASE_REQUESTS");
  if (!sheet) {
    sheet = userDB.insertSheet("PURCHASE_REQUESTS");
    sheet.appendRow(["RequestID", "Date", "SellerID", "SellerName", "BranchID", "Status", "ItemsJSON", "Notes", "ApprovedBy", "ApprovedDate", "CompletedDate", "Title", "RejectedReason", "RejectedDate", "ApprovalNotes", "ReviewDraftJSON"]);
    sheet.getRange("A1:P1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  } else {
    // Backfill for sheets created before these columns existed
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.length < 12) sheet.getRange(1, 12).setValue("Title");
    if (headers.length < 13) sheet.getRange(1, 13).setValue("RejectedReason");
    if (headers.length < 14) sheet.getRange(1, 14).setValue("RejectedDate");
    if (headers.length < 15) sheet.getRange(1, 15).setValue("ApprovalNotes");
    if (headers.length < 16) sheet.getRange(1, 16).setValue("ReviewDraftJSON");
  }
  return sheet;
}

function parseRequestRow(row) {
  var items = [];
  try {
    items = JSON.parse(row[6] || "[]");
  } catch (e) {
    items = [];
  }
  var reviewDraft = null;
  try {
    reviewDraft = row[15] ? JSON.parse(row[15]) : null;
  } catch (e) {
    reviewDraft = null;
  }
  return {
    requestId: row[0],
    date: row[1],
    sellerId: row[2],
    sellerName: row[3],
    branchId: row[4],
    status: row[5],
    items: items,
    notes: row[7] || "",
    approvedBy: row[8] || "",
    approvedDate: row[9] || "",
    completedDate: row[10] || "",
    title: row[11] || "",
    rejectedReason: row[12] || "",
    rejectedDate: row[13] || "",
    approvalNotes: row[14] || "",
    reviewDraft: reviewDraft
  };
}

function normalizeRequestItems(rawItems) {
  return (rawItems || []).map(function(item) {
    return {
      product: item.product,
      qty: Number(item.qty) || 0,
      unit: item.unit || "",
      category: item.category || "",
      supplier: item.supplier || "",
      paymentMethod: item.paymentMethod || "",
      isStockProduct: item.isStockProduct !== false, // defaults true for backward compatibility
      suggestedBuyPrice: item.suggestedBuyPrice ? Number(item.suggestedBuyPrice) : "",
      suggestedSellPrice: item.suggestedSellPrice ? Number(item.suggestedSellPrice) : "",
      price: item.price ? Number(item.price) : "",
      sellPrice: item.sellPrice ? Number(item.sellPrice) : "",
      completed: !!item.completed,
      completedQty: Number(item.completedQty) || 0
    };
  });
}

// Finds a request row by ID. Returns { rowIndex (1-based), req } or null.
function findRequestRow(sheet, requestId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      return { rowIndex: i + 1, req: parseRequestRow(data[i]) };
    }
  }
  return null;
}

// Creates or updates a DRAFT request. The client always generates and sends
// its own requestId (so the same ID works whether this was created offline
// or online) - if that ID already exists as a draft or rejected request, it
// gets updated in place; otherwise a new draft row is created.
function saveDraftPurchaseRequest(d) {
  try {
    if (!d.databaseId || !d.requestId) {
      return output({ success: false, error: "Database ID and Request ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const items = normalizeRequestItems(d.items);
    const found = findRequestRow(sheet, d.requestId);
    
    if (found) {
      // Only drafts/rejected requests can be edited in place
      if (found.req.status !== "draft" && found.req.status !== "rejected") {
        return output({ success: false, error: "This request can no longer be edited as a draft" });
      }
      sheet.getRange(found.rowIndex, 6).setValue("draft");
      sheet.getRange(found.rowIndex, 7).setValue(JSON.stringify(items));
      sheet.getRange(found.rowIndex, 8).setValue(d.notes || "");
      sheet.getRange(found.rowIndex, 12).setValue(d.title || "");
      sheet.getRange(found.rowIndex, 13).setValue("");
      sheet.getRange(found.rowIndex, 14).setValue("");
    } else {
      sheet.appendRow([
        d.requestId,
        new Date(),
        d.sellerId || "",
        d.sellerName || "",
        d.branchId || "MAIN",
        "draft",
        JSON.stringify(items),
        d.notes || "",
        "", "", "",
        d.title || "",
        "", ""
      ]);
    }
    
    return output({ success: true, requestId: d.requestId });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Seller submits a wishlist of items they need restocked - either turning
// an existing draft into a pending request, or creating one fresh.
function submitPurchaseRequest(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    if (!d.items || !Array.isArray(d.items) || d.items.length === 0) {
      return output({ success: false, error: "At least one item is required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const items = normalizeRequestItems(d.items);
    const requestId = d.requestId || ("PR-" + Utilities.getUuid().substring(0, 8).toUpperCase());
    
    const found = findRequestRow(sheet, requestId);
    
    if (found) {
      sheet.getRange(found.rowIndex, 2).setValue(new Date());
      sheet.getRange(found.rowIndex, 6).setValue("pending");
      sheet.getRange(found.rowIndex, 7).setValue(JSON.stringify(items));
      sheet.getRange(found.rowIndex, 8).setValue(d.notes || "");
      sheet.getRange(found.rowIndex, 12).setValue(d.title || "");
      sheet.getRange(found.rowIndex, 13).setValue("");
      sheet.getRange(found.rowIndex, 14).setValue("");
    } else {
      sheet.appendRow([
        requestId,
        new Date(),
        d.sellerId || "",
        d.sellerName || "",
        d.branchId || "MAIN",
        "pending",
        JSON.stringify(items),
        d.notes || "",
        "", "", "",
        d.title || "",
        "", ""
      ]);
    }
    
    return output({ success: true, requestId: requestId, message: "Purchase request submitted" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Admin's own shopping list - reuses the exact same PURCHASE_REQUESTS
// sheet/workflow as Seller requests, but skips the approval step entirely:
// Admin sets prices directly and the request lands straight in "approved"
// status (sellerId="ADMIN"), ready for immediate item-by-item completion.
function submitAdminPurchaseList(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    if (!d.items || !Array.isArray(d.items) || d.items.length === 0) {
      return output({ success: false, error: "At least one item is required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const items = normalizeRequestItems(d.items).map(function(item) {
      // Admin sets the real price directly (no separate approval), so
      // fold whatever price was given into the final "price" field.
      var finalPrice = item.price || item.suggestedBuyPrice || 0;
      var finalSellPrice = item.sellPrice || item.suggestedSellPrice || (finalPrice * 1.2);
      return Object.assign({}, item, { price: finalPrice, sellPrice: finalSellPrice });
    });
    const requestId = d.requestId || ("PL-" + Utilities.getUuid().substring(0, 8).toUpperCase());
    const now = new Date();
    
    const found = findRequestRow(sheet, requestId);
    
    if (found) {
      sheet.getRange(found.rowIndex, 2).setValue(now);
      sheet.getRange(found.rowIndex, 6).setValue("approved");
      sheet.getRange(found.rowIndex, 7).setValue(JSON.stringify(items));
      sheet.getRange(found.rowIndex, 8).setValue(d.notes || "");
      sheet.getRange(found.rowIndex, 9).setValue(d.approvedBy || "");
      sheet.getRange(found.rowIndex, 10).setValue(now);
      sheet.getRange(found.rowIndex, 12).setValue(d.title || "");
    } else {
      sheet.appendRow([
        requestId,
        now,
        "ADMIN",
        d.sellerName || "Admin",
        d.branchId || "MAIN",
        "approved",
        JSON.stringify(items),
        d.notes || "",
        d.approvedBy || "",
        now,
        "",
        d.title || "",
        "", "", "", ""
      ]);
    }
    
    return output({ success: true, requestId: requestId, message: "Added to your list" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Returns purchase requests, optionally filtered by seller and/or status.
// d.status can be a single string or an array of strings.
function getPurchaseRequests(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const data = sheet.getDataRange().getValues();
    
    var statusFilter = null;
    if (d.status) {
      statusFilter = Array.isArray(d.status) ? d.status : [d.status];
    }
    
    const requests = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const req = parseRequestRow(data[i]);
      
      if (d.sellerId && req.sellerId !== d.sellerId) continue;
      if (statusFilter && statusFilter.indexOf(req.status) === -1) continue;
      
      requests.push(req);
    }
    
    requests.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
    
    return output({ success: true, requests: requests });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Admin reviews a pending request: can edit quantities, set/override prices
// (the Seller's prices are only suggestions), and drop items they don't
// want to approve (send a smaller items array).
function approvePurchaseRequest(d) {
  try {
    if (!d.databaseId || !d.requestId) {
      return output({ success: false, error: "Database ID and Request ID required" });
    }
    if (!d.items || !Array.isArray(d.items) || d.items.length === 0) {
      return output({ success: false, error: "At least one approved item is required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const found = findRequestRow(sheet, d.requestId);
    
    if (!found) {
      return output({ success: false, error: "Request not found" });
    }
    
    const approvedItems = d.items.map(function(item) {
      return {
        product: item.product,
        qty: Number(item.approvedQty || item.qty) || 0,
        unit: item.unit || "",
        category: item.category || "",
        supplier: item.supplier || "",
        paymentMethod: item.paymentMethod || "",
        isStockProduct: item.isStockProduct !== false,
        price: Number(item.approvedPrice) || 0,
        sellPrice: Number(item.sellPrice) || (Number(item.approvedPrice) || 0) * 1.2,
        completed: false,
        completedQty: 0,
        adminNote: item.adminNote || ""
      };
    });
    
    sheet.getRange(found.rowIndex, 6).setValue("approved");
    sheet.getRange(found.rowIndex, 7).setValue(JSON.stringify(approvedItems));
    sheet.getRange(found.rowIndex, 9).setValue(d.approvedBy || "");
    sheet.getRange(found.rowIndex, 10).setValue(new Date());
    if (d.approvalNotes) {
      sheet.getRange(found.rowIndex, 15).setValue(d.approvalNotes);
    }
    sheet.getRange(found.rowIndex, 16).setValue(""); // clear the in-progress review draft, now finalized
    
    return output({ success: true, message: "Request approved" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Admin rejects a pending request with a reason - Seller can then edit and
// resubmit it (like a draft) or delete it entirely.
function rejectPurchaseRequest(d) {
  try {
    if (!d.databaseId || !d.requestId) {
      return output({ success: false, error: "Database ID and Request ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const found = findRequestRow(sheet, d.requestId);
    
    if (!found) {
      return output({ success: false, error: "Request not found" });
    }
    
    sheet.getRange(found.rowIndex, 6).setValue("rejected");
    sheet.getRange(found.rowIndex, 13).setValue(d.reason || "No reason given");
    sheet.getRange(found.rowIndex, 14).setValue(new Date());
    sheet.getRange(found.rowIndex, 16).setValue(""); // clear the in-progress review draft
    
    return output({ success: true, message: "Request rejected" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Auto-saves the Admin's in-progress review edits (quantity/price changes,
// added items, removed items) WITHOUT changing the request's status - so
// nothing is lost if the Admin navigates away before tapping "Approve".
function saveReviewDraft(d) {
  try {
    if (!d.databaseId || !d.requestId) {
      return output({ success: false, error: "Database ID and Request ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const found = findRequestRow(sheet, d.requestId);
    
    if (!found) {
      return output({ success: false, error: "Request not found" });
    }
    if (found.req.status !== "pending") {
      return output({ success: false, error: "This request is no longer pending" });
    }
    
    const draft = {
      items: d.items || [],
      removedItems: d.removedItems || []
    };
    sheet.getRange(found.rowIndex, 16).setValue(JSON.stringify(draft));
    
    return output({ success: true });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Deletes a draft or rejected request entirely (Seller's own cleanup).
function deletePurchaseRequest(d) {
  try {
    if (!d.databaseId || !d.requestId) {
      return output({ success: false, error: "Database ID and Request ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const found = findRequestRow(sheet, d.requestId);
    
    if (!found) {
      return output({ success: true, message: "Already removed" });
    }
    if (found.req.status !== "draft" && found.req.status !== "rejected") {
      return output({ success: false, error: "Only draft or rejected requests can be deleted" });
    }
    
    sheet.deleteRow(found.rowIndex);
    return output({ success: true, message: "Request deleted" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function completePurchaseRequestItem(d) {
  try {
    if (!d.databaseId || !d.requestId || d.itemIndex === undefined || d.itemIndex === null) {
      return output({ success: false, error: "Database ID, Request ID, and item index required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sheet = ensurePurchaseRequestsSheet(userDB);
    const found = findRequestRow(sheet, d.requestId);
    
    if (!found) {
      return output({ success: false, error: "Request not found" });
    }
    
    const req = found.req;
    const idx = Number(d.itemIndex);
    
    if (!req.items[idx]) {
      return output({ success: false, error: "Item not found in this request" });
    }
    if (req.items[idx].completed) {
      return output({ success: false, error: "This item was already completed" });
    }
    
    const item = req.items[idx];
    const alreadyCompleted = Number(item.completedQty) || 0;
    const remaining = Number(item.qty) - alreadyCompleted;
    
    // How much to complete THIS time - the Seller can do partial amounts
    // across multiple visits (e.g. 2 of 6 now, the rest later).
    var completeQty = Number(d.completeQty);
    if (!completeQty || completeQty <= 0 || completeQty > remaining) {
      completeQty = remaining;
    }
    
    // Add straight to Stock (unless it's a Non-Stock item) + log as a
    // Purchase (mirrors addPurchase()) for only the quantity being
    // completed right now.
    addItemToStockAndPurchases(
      userDB,
      item.product,
      completeQty,
      item.unit,
      item.price,
      item.sellPrice || (item.price * 1.2),
      req.branchId,
      req.sellerId,
      item.supplier || "",
      item.category || "",
      item.isStockProduct !== false,
      item.paymentMethod || ""
    );
    
    const newCompletedQty = alreadyCompleted + completeQty;
    req.items[idx].completedQty = newCompletedQty;
    req.items[idx].completed = newCompletedQty >= Number(item.qty);
    const allDone = req.items.every(function(it) { return it.completed; });
    
    sheet.getRange(found.rowIndex, 7).setValue(JSON.stringify(req.items));
    if (allDone) {
      sheet.getRange(found.rowIndex, 6).setValue("completed");
      sheet.getRange(found.rowIndex, 11).setValue(new Date());
    }
    
    if (req.sellerId && req.sellerId !== "ADMIN") {
      updateSellerTotals(d.databaseId, req.sellerId, completeQty * item.price, 'purchase');
    }
    
    return output({ success: true, message: "Item added to stock", allCompleted: allDone, itemCompleted: req.items[idx].completed, completedQty: newCompletedQty });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Shared by completePurchaseRequestItem() - same stock-consolidation logic
// as addPurchase(), but also stores the Unit (which addPurchase's stock
// write predates and doesn't set). Pass isStockProduct=false for Non-Stock
// items (services, one-off expenses) - the Purchase record is still kept,
// but nothing is added to Stock.
function addItemToStockAndPurchases(userDB, product, qty, unit, unitCost, sellPrice, branchId, sellerId, supplierLabel, category, isStockProduct, paymentMethod) {
  const purchasesSheet = userDB.getSheetByName("PURCHASES");
  const stockSheet = userDB.getSheetByName("STOCK");
  const timestamp = new Date();
  
  if (purchasesSheet) {
    const headers = purchasesSheet.getRange(1, 1, 1, purchasesSheet.getLastColumn()).getValues()[0];
    if (headers.length < 9) {
      purchasesSheet.getRange(1, 8).setValue("Seller ID");
      purchasesSheet.getRange(1, 9).setValue("Branch ID");
    }
    if (headers.length < 10) {
      purchasesSheet.getRange(1, 10).setValue("Category");
    }
    if (headers.length < 11) {
      purchasesSheet.getRange(1, 11).setValue("Unit");
    }
    if (headers.length < 12) {
      purchasesSheet.getRange(1, 12).setValue("Item Type");
    }
    purchasesSheet.appendRow([timestamp, product, qty, unitCost, qty * unitCost, supplierLabel || "", paymentMethod || "", sellerId || "", branchId, category || "", unit || "", isStockProduct === false ? "nonstock" : "stock"]);
  }
  
  // Non-Stock items (services, one-off expenses) are recorded as a
  // Purchase above but never touch the Stock sheet.
  if (isStockProduct === false) return;
  
  if (!stockSheet) return;
  
  const stockData = stockSheet.getDataRange().getValues();
  let existingRow = -1, existingQuantity = 0, existingBuyPrice = 0;
  
  for (let i = 1; i < stockData.length; i++) {
    if (stockData[i][0] && stockData[i][1] === product) {
      const itemBranchId = stockData[i][8] || "MAIN";
      const itemUnit = stockData[i][9] || "";
      if (itemBranchId === branchId && (itemUnit === (unit || "") || !itemUnit)) {
        existingRow = i + 1;
        existingQuantity = stockData[i][2] || 0;
        existingBuyPrice = stockData[i][3] || 0;
        break;
      }
    }
  }
  
  if (existingRow !== -1) {
    const oldTotalValue = existingQuantity * existingBuyPrice;
    const newTotalValue = qty * unitCost;
    const totalQuantity = existingQuantity + qty;
    const avgBuyPrice = (oldTotalValue + newTotalValue) / totalQuantity;
    
    stockSheet.getRange(existingRow, 3).setValue(totalQuantity);
    stockSheet.getRange(existingRow, 4).setValue(avgBuyPrice);
    stockSheet.getRange(existingRow, 5).setValue(sellPrice);
    stockSheet.getRange(existingRow, 6).setValue(totalQuantity * avgBuyPrice);
    stockSheet.getRange(existingRow, 1).setValue(timestamp);
    if (unit) stockSheet.getRange(existingRow, 10).setValue(unit);
    if (category) stockSheet.getRange(existingRow, 7).setValue(category);
  } else {
    const stockValue = qty * unitCost;
    const stockHeaders = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
    if (stockHeaders.length < 9) {
      stockSheet.getRange(1, 8).setValue("Seller ID");
      stockSheet.getRange(1, 9).setValue("Branch ID");
    }
    if (stockHeaders.length < 10) {
      stockSheet.getRange(1, 10).setValue("Unit");
    }
    stockSheet.appendRow([timestamp, product, qty, unitCost, sellPrice, stockValue, category || "", sellerId || "", branchId, unit || ""]);
  }
}

function addPurchase(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    const stockSheet = userDB.getSheetByName("STOCK");
    const branchId = d.branchId || "MAIN";
    
    if (!purchasesSheet) {
      return output({ success: false, error: "Purchases sheet not found" });
    }
    
    const isStockProduct = d.isStockProduct !== false;
    
    if (isStockProduct && !stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    // Ensure sheet has all columns
    const headers = purchasesSheet.getRange(1, 1, 1, purchasesSheet.getLastColumn()).getValues()[0];
    if (headers.length < 9) {
      purchasesSheet.getRange(1, 8).setValue("Seller ID");
      purchasesSheet.getRange(1, 9).setValue("Branch ID");
    }
    if (headers.length < 10) {
      purchasesSheet.getRange(1, 10).setValue("Category");
    }
    if (headers.length < 11) {
      purchasesSheet.getRange(1, 11).setValue("Unit");
    }
    if (headers.length < 12) {
      purchasesSheet.getRange(1, 12).setValue("Item Type");
    }
    
    // Purchases always record the real moment they're entered - "now" -
    // exactly as before. The one narrow exception: Add Expense explicitly
    // lets someone pick a Date (e.g. logging a rent payment from a few
    // days ago), so it alone passes customTimestamp - a field the normal
    // Add Purchase flow never sends, so ordinary purchases are completely
    // unaffected by this and keep behaving exactly as they always have.
    const timestamp = d.customTimestamp ? new Date(d.customTimestamp) : new Date();
    const product = d.product;
    const quantity = d.qty;
    const unitCost = d.unitCost || (d.total / d.qty);
    const total = d.total;
    const supplier = d.supplier || "Local Supplier";
    const paymentMethod = d.paymentMethod || "cash";
    const sellerId = d.sellerId || null;
    const notes = d.notes || "";
    const category = d.category || "";
    const unit = d.unit || "";
    
    // Record purchase
    purchasesSheet.appendRow([
      timestamp,
      product,
      quantity,
      unitCost,
      total,
      supplier,
      paymentMethod,
      sellerId,
      branchId,
      category,
      unit,
      isStockProduct ? "stock" : "nonstock"
    ]);
    
    // ============================================
    // ADD TO STOCK INVENTORY AUTOMATICALLY (skipped for Non-Stock items -
    // services / one-off expenses that shouldn't affect inventory)
    // ============================================
    if (isStockProduct) {
    // Calculate sell price (20% markup or use provided)
    let sellPrice = d.sellPrice || (unitCost * 1.2);
    
    // Check if this product already exists in this branch
    const stockData = stockSheet.getDataRange().getValues();
    let existingRow = -1;
    let existingQuantity = 0;
    let existingBuyPrice = 0;
    let existingTotalValue = 0;
    
    // Check if this product already exists in this branch WITH THE SAME
    // UNIT - two different units of the same product name must never be
    // merged into one row (their quantities aren't comparable).
    for (let i = 1; i < stockData.length; i++) {
      if (stockData[i][0] && stockData[i][1] === product) {
        const itemBranchId = stockData[i][8] || "MAIN";
        const itemUnit = stockData[i][9] || "";
        if (itemBranchId === branchId && (itemUnit === (d.unit || "") || !itemUnit)) {
          existingRow = i + 1;
          existingQuantity = stockData[i][2] || 0;
          existingBuyPrice = stockData[i][3] || 0;
          existingTotalValue = stockData[i][5] || 0;
          break;
        }
      }
    }
    
    if (existingRow !== -1) {
      // Product exists - UPDATE with weighted average
      const oldTotalValue = existingQuantity * existingBuyPrice;
      const newTotalValue = quantity * unitCost;
      const totalQuantity = existingQuantity + quantity;
      const avgBuyPrice = (oldTotalValue + newTotalValue) / totalQuantity;
      
      stockSheet.getRange(existingRow, 3).setValue(totalQuantity);           // New quantity
      stockSheet.getRange(existingRow, 4).setValue(avgBuyPrice);            // Weighted average buy price
      stockSheet.getRange(existingRow, 5).setValue(sellPrice);              // Suggest sell price
      stockSheet.getRange(existingRow, 6).setValue(totalQuantity * avgBuyPrice); // New total value
      stockSheet.getRange(existingRow, 1).setValue(timestamp);              // Update timestamp
      if (d.unit) stockSheet.getRange(existingRow, 10).setValue(d.unit);    // Backfill unit if provided
      
      console.log("Stock updated for existing product:", product, "New quantity:", totalQuantity);
    } else {
      // New product - CREATE new stock entry
      const stockValue = quantity * unitCost;
      
      // Ensure stock sheet has all columns
      const stockHeaders = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
      if (stockHeaders.length < 9) {
        stockSheet.getRange(1, 8).setValue("Seller ID");
        stockSheet.getRange(1, 9).setValue("Branch ID");
      }
      if (stockHeaders.length < 10) {
        stockSheet.getRange(1, 10).setValue("Unit");
      }
      
      stockSheet.appendRow([
        timestamp,
        product,
        quantity,
        unitCost,
        sellPrice,
        stockValue,
        d.category || "",
        sellerId,
        branchId,
        d.unit || ""
      ]);
      
      console.log("New stock entry created for product:", product);
    }
    } // end isStockProduct block
    
    // Update seller's total purchases if sellerId exists
    if (sellerId) {
      updateSellerTotals(d.databaseId, sellerId, total, 'purchase');
    }
    
    return output({ 
      success: true, 
      message: isStockProduct
        ? `Purchase recorded: ${quantity} x ${product} at ${branchId} branch. Stock updated automatically.`
        : `Purchase recorded: ${quantity} x ${product} at ${branchId} branch (Non-Stock - not added to inventory).`,
      total: total,
      stockUpdated: isStockProduct,
      timestamp: timestamp.toISOString()
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


/* ===== UPDATE STOCK ===== */
function updateStock(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    if (!d.timestamp || !d.product || !d.branchId) {
      return output({ success: false, error: "Timestamp, product and branch ID required to identify the row" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    // Tafuta row sahihi kwa kutumia timestamp, product, na branch
    const data = stockSheet.getDataRange().getValues();
    let targetRow = -1;
    let existingDate = null;
    let existingSellerId = null;
    let existingBranchId = null;
    
    // Tafuta kwa kutumia timestamp (column 1), product (column 2), na branch (column 9)
    for (let i = 1; i < data.length; i++) {
      const rowTimestamp = data[i][0];
      const rowProduct = data[i][1];
      const rowBranchId = data[i][8] || "MAIN";
      
      // Convert timestamp to comparable format
      const rowTime = new Date(rowTimestamp).getTime();
      const targetTime = new Date(d.timestamp).getTime();
      
      // Compare timestamps (within 1 second tolerance)
      const timeDiff = Math.abs(rowTime - targetTime);
      
      if (timeDiff < 1000 && rowProduct === d.product && rowBranchId === d.branchId) {
        targetRow = i + 1; // +1 because sheet rows are 1-based
        existingDate = rowTimestamp;
        existingSellerId = data[i][7] || null;
        existingBranchId = rowBranchId;
        break;
      }
    }
    
    if (targetRow === -1) {
      return output({ 
        success: false, 
        error: `Could not find the exact stock entry to update. Try refreshing the page.` 
      });
    }
    
    // Hesabu stock value mpya
    const stockValue = d.quantity * d.buyPrice;
    
    // Set values kwenye row iliyopatikana
    stockSheet.getRange(targetRow, 2).setValue(d.newProduct || d.product); // Product (allow rename)
    stockSheet.getRange(targetRow, 3).setValue(d.quantity);     // Quantity
    stockSheet.getRange(targetRow, 4).setValue(d.buyPrice);     // Buy Price
    stockSheet.getRange(targetRow, 5).setValue(d.sellPrice);    // Sell Price
    stockSheet.getRange(targetRow, 6).setValue(stockValue);     // Stock Value
    stockSheet.getRange(targetRow, 7).setValue(d.category || ""); // Category
    if (d.unit !== undefined) {
      stockSheet.getRange(targetRow, 10).setValue(d.unit || ""); // Unit
    }
    if (d.newBranchId && d.newBranchId !== existingBranchId) {
      stockSheet.getRange(targetRow, 9).setValue(d.newBranchId); // Branch (moved)
    }
    
    // HAZIBADILISHI (isipokuwa Unit na Branch endapo zimetumwa hapo juu):
    // Column 1: Date - inabaki vile
    // Column 8: Seller ID - inabaki vile
    
    return output({
      success: true,
      message: "Stock updated successfully",
      timestamp: existingDate.toISOString(),
      branchId: (d.newBranchId || existingBranchId),
      sellerId: existingSellerId,
      product: d.product
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


function processOfflineData(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    ensureSheetsExist(d);
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    let results = { success: true, operations: [] };
    
    // ============================================
    // PROCESS STOCK DATA (with consolidation)
    // ============================================
    if (d.stock && d.stock.length > 0) {
      const stockSheet = userDB.getSheetByName("STOCK");
      if (stockSheet) {
        for (var s = 0; s < d.stock.length; s++) {
          var item = d.stock[s];
          try {
            var timestamp = new Date(item.timestamp || Date.now());
            var branchId = item.branchId || "MAIN";
            var product = item.product;
            var newQuantity = item.qty;
            var newBuyPrice = item.buy;
            var newSellPrice = item.sell;
            var category = item.category || "";
            var unit = item.unit || "";
            var sellerId = item.sellerId || null;
            
            var data = stockSheet.getDataRange().getValues();
            var existingRow = -1;
            var existingQuantity = 0;
            var existingBuyPrice = 0;
            
            for (var i = 1; i < data.length; i++) {
              if (data[i][0] && data[i][1] === product) {
                var itemBranchId = data[i][8] || "MAIN";
                var itemUnit = data[i][9] || "";
                if (itemBranchId === branchId && (itemUnit === unit || !itemUnit)) {
                  existingRow = i + 1;
                  existingQuantity = data[i][2] || 0;
                  existingBuyPrice = data[i][3] || 0;
                  break;
                }
              }
            }
            
            if (existingRow !== -1) {
              var oldTotalValue = existingQuantity * existingBuyPrice;
              var newTotalValue = newQuantity * newBuyPrice;
              var totalQuantity = existingQuantity + newQuantity;
              var avgBuyPrice = (oldTotalValue + newTotalValue) / totalQuantity;
              
              stockSheet.getRange(existingRow, 3).setValue(totalQuantity);
              stockSheet.getRange(existingRow, 4).setValue(avgBuyPrice);
              stockSheet.getRange(existingRow, 5).setValue(newSellPrice);
              stockSheet.getRange(existingRow, 6).setValue(totalQuantity * avgBuyPrice);
              if (category) stockSheet.getRange(existingRow, 7).setValue(category);
              if (unit) stockSheet.getRange(existingRow, 10).setValue(unit);
              stockSheet.getRange(existingRow, 1).setValue(timestamp);
              
              results.operations.push({ type: "stock_add", product: product, branch: branchId, status: "success", consolidated: true });
            } else {
              var stockValue = newQuantity * newBuyPrice;
              var headers = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
              if (headers.length < 9) {
                stockSheet.getRange(1, 8).setValue("Seller ID");
                stockSheet.getRange(1, 9).setValue("Branch ID");
              }
              if (headers.length < 10) {
                stockSheet.getRange(1, 10).setValue("Unit");
              }
              stockSheet.appendRow([timestamp, product, newQuantity, newBuyPrice, newSellPrice, stockValue, category, sellerId, branchId, unit]);
              results.operations.push({ type: "stock_add", product: product, branch: branchId, status: "success", consolidated: false });
            }
          } catch (e) {
            results.operations.push({ type: "stock_add", product: item.product, status: "failed", error: e.toString() });
          }
        }
      }
    }
    
    // ============================================
    // PROCESS SALES DATA (IDEMPOTENT)
    // ============================================
    if (d.sales && d.sales.length > 0) {
      const salesSheet = userDB.getSheetByName("SALES");
      if (salesSheet) {
        for (var sl = 0; sl < d.sales.length; sl++) {
          var item = d.sales[sl];
          try {
            // Reuse the exact same atomic/idempotent sale path as online sales.
            // This is essential: a sync request may reach the server, commit,
            // then lose its response before the phone receives it and retries.
            var saleResponse = addSale({
              databaseId: d.databaseId,
              product: item.product,
              qty: item.qty,
              unitPrice: item.unitPrice,
              total: item.total,
              customer: item.customer,
              paymentMethod: item.paymentMethod,
              isStockProduct: item.isStockProduct,
              branchId: item.branchId || "MAIN",
              timestamp: item.timestamp,
              sellerId: item.sellerId || null,
              unit: item.unit || "",
              receiptNo: item.receiptNo || null,
              operationId: item.operationId || null
            });
            var saleResult = {};
            try { saleResult = JSON.parse(saleResponse.getContent()); } catch (parseErr) {}
            if (saleResult.success) {
              results.operations.push({
                type: "sale",
                operationId: item.operationId || saleResult.operationId || null,
                product: item.product,
                branch: item.branchId || "MAIN",
                status: "success",
                alreadyProcessed: !!saleResult.alreadyProcessed,
                rowIndex: saleResult.rowIndex,
                receiptNo: saleResult.receiptNo || item.receiptNo || ""
              });
            } else {
              results.operations.push({ type: "sale", product: item.product, status: "failed", error: saleResult.error || "Sale failed" });
            }
          } catch (e) {
            results.operations.push({ type: "sale", product: item.product, status: "failed", error: e.toString() });
          }
        }
      }
    }
    
    // ============================================
    // PROCESS PURCHASES DATA (WITH AUTO STOCK UPDATE)
    // ============================================
    if (d.purchases && d.purchases.length > 0) {
      const purchasesSheet = userDB.getSheetByName("PURCHASES");
      const stockSheet = userDB.getSheetByName("STOCK");
      
      if (purchasesSheet && stockSheet) {
        for (var p = 0; p < d.purchases.length; p++) {
          var item = d.purchases[p];
          try {
            var branchId = item.branchId || "MAIN";
            var timestamp = new Date(item.timestamp || Date.now());
            var product = item.product;
            var quantity = item.qty;
            var unitCost = item.unitCost || (item.total / item.qty);
            var total = item.total;
            var supplier = item.supplier || "Local Supplier";
            var paymentMethod = item.paymentMethod || "cash";
            var sellerId = item.sellerId || null;
            
            var headers = purchasesSheet.getRange(1, 1, 1, purchasesSheet.getLastColumn()).getValues()[0];
            if (headers.length < 9) {
              purchasesSheet.getRange(1, 8).setValue("Seller ID");
              purchasesSheet.getRange(1, 9).setValue("Branch ID");
            }
            if (headers.length < 10) {
              purchasesSheet.getRange(1, 10).setValue("Category");
            }
            if (headers.length < 11) {
              purchasesSheet.getRange(1, 11).setValue("Unit");
            }
            if (headers.length < 12) {
              purchasesSheet.getRange(1, 12).setValue("Item Type");
            }
            purchasesSheet.appendRow([timestamp, product, quantity, unitCost, total, supplier, paymentMethod, sellerId, branchId, item.category || "", item.unit || "", item.isStockProduct === false ? "nonstock" : "stock"]);
            
            // ============================================
            // ADD TO STOCK INVENTORY AUTOMATICALLY (skipped for Non-Stock items)
            // ============================================
            if (item.isStockProduct === false) {
              if (sellerId) updateSellerTotals(d.databaseId, sellerId, total, 'purchase');
              results.operations.push({ type: "purchase", product: product, branch: branchId, status: "success", stockUpdated: false });
              continue;
            }
            var sellPrice = item.sellPrice || (unitCost * 1.2);
            
            var stockData = stockSheet.getDataRange().getValues();
            var existingRow = -1;
            var existingQuantity = 0;
            var existingBuyPrice = 0;
            
            for (var j = 1; j < stockData.length; j++) {
              if (stockData[j][0] && stockData[j][1] === product) {
                var itemBranchId = stockData[j][8] || "MAIN";
                var itemUnit = stockData[j][9] || "";
                if (itemBranchId === branchId && (itemUnit === (item.unit || "") || !itemUnit)) {
                  existingRow = j + 1;
                  existingQuantity = stockData[j][2] || 0;
                  existingBuyPrice = stockData[j][3] || 0;
                  break;
                }
              }
            }
            
            if (existingRow !== -1) {
              var oldTotalValue = existingQuantity * existingBuyPrice;
              var newTotalValue = quantity * unitCost;
              var totalQuantity = existingQuantity + quantity;
              var avgBuyPrice = (oldTotalValue + newTotalValue) / totalQuantity;
              
              stockSheet.getRange(existingRow, 3).setValue(totalQuantity);
              stockSheet.getRange(existingRow, 4).setValue(avgBuyPrice);
              stockSheet.getRange(existingRow, 5).setValue(sellPrice);
              stockSheet.getRange(existingRow, 6).setValue(totalQuantity * avgBuyPrice);
              if (item.category) stockSheet.getRange(existingRow, 7).setValue(item.category);
              if (item.unit) stockSheet.getRange(existingRow, 10).setValue(item.unit);
              stockSheet.getRange(existingRow, 1).setValue(timestamp);
            } else {
              var stockValue = quantity * unitCost;
              var stockHeaders = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
              if (stockHeaders.length < 9) {
                stockSheet.getRange(1, 8).setValue("Seller ID");
                stockSheet.getRange(1, 9).setValue("Branch ID");
              }
              if (stockHeaders.length < 10) {
                stockSheet.getRange(1, 10).setValue("Unit");
              }
              stockSheet.appendRow([timestamp, product, quantity, unitCost, sellPrice, stockValue, item.category || "", sellerId, branchId, item.unit || ""]);
            }
            
            if (sellerId) updateSellerTotals(d.databaseId, sellerId, total, 'purchase');
            results.operations.push({ type: "purchase", product: product, branch: branchId, status: "success", stockUpdated: true });
          } catch (e) {
            results.operations.push({ type: "purchase", product: item.product, status: "failed", error: e.toString() });
          }
        }
      }
    }
    
    // ============================================
    // PROCESS EXPENSES DATA (genuine operating costs - rent, electricity,
    // transport, salaries, etc - completely separate from Purchases;
    // never touches Stock)
    // ============================================
    if (d.expenses && d.expenses.length > 0) {
      const expensesSheet = ensureExpensesSheet(userDB);
      for (var e2 = 0; e2 < d.expenses.length; e2++) {
        var expItem = d.expenses[e2];
        try {
          // Reuse the same validated/idempotent Expense path as online saves.
          // This preserves recordedAt and the per-expense operationId instead
          // of generating a new server time for an offline transaction.
          var expenseResponse = addExpense({
            token: d.token,
            databaseId: d.databaseId,
            operationId: expItem.operationId || null,
            date: expItem.date || expItem.timestamp || null,
            recordedAt: expItem.recordedAt || expItem.timestamp || null,
            category: expItem.category || "Uncategorized",
            description: expItem.description || "",
            amount: Number(expItem.amount || expItem.total) || 0,
            paymentMethod: expItem.paymentMethod || "cash",
            branchId: expItem.branchId || "MAIN",
            sellerId: expItem.sellerId || ""
          });
          var expenseResult = {};
          try { expenseResult = JSON.parse(expenseResponse.getContent()); } catch (parseErr) {}
          results.operations.push({
            type: "expense",
            description: expItem.description,
            branch: expItem.branchId || "MAIN",
            status: expenseResult.success ? "success" : "failed",
            alreadyProcessed: !!expenseResult.alreadyProcessed,
            operationId: expItem.operationId || "",
            error: expenseResult.success ? "" : (expenseResult.error || "Expense failed")
          });
        } catch (e) {
          results.operations.push({ type: "expense", description: expItem.description, status: "failed", error: e.toString() });
        }
      }
    }
    
    // ============================================
    // PROCESS DEBTS DATA
    // ============================================
    if (d.debts && d.debts.length > 0) {
      const debtsSheet = userDB.getSheetByName("DEBTS");
      if (debtsSheet) {
        for (var db = 0; db < d.debts.length; db++) {
          var item = d.debts[db];
          try {
            var branchId = item.branchId || "MAIN";
            var timestamp = new Date(item.timestamp || Date.now());
            var debtId = item.id || "DEBT_" + Date.now() + "_" + Math.random().toString(36).substr(2, 8).toUpperCase();
            
            var photoId = item.photoId || '';
            var existingDebtRow = -1;
            var existingDebtData = debtsSheet.getDataRange().getValues();
            for (var edi = 1; edi < existingDebtData.length; edi++) {
              if (String(existingDebtData[edi][0] || '') === String(debtId)) { existingDebtRow = edi + 1; break; }
            }
            if (existingDebtRow !== -1) {
              results.operations.push({ type: "debt", debtId: debtId, product: item.product, branch: branchId, status: "success", alreadyProcessed: true, borrower: item.borrowerName });
              continue;
            }
            var photoBase64 = String(item.photoBase64 || '');
            var photoMimeType = String(item.photoMimeType || 'image/jpeg');
            if (photoBase64.indexOf('data:') === 0) {
              var commaPos = photoBase64.indexOf(',');
              if (commaPos !== -1) photoBase64 = photoBase64.substring(commaPos + 1);
            }
            if (photoBase64.length > 48000) {
              results.operations.push({ type: "debt", product: item.product, status: "failed", error: "Borrower photo exceeds 50,000 character cell limit" });
              continue;
            }
            
            if (item.isStockProduct) {
              var stockCheck = checkBranchStockAvailability(d.databaseId, item.product, item.quantity, branchId, item.unit);
              if (!stockCheck.available) {
                results.operations.push({ type: "debt", product: item.product, status: "failed", error: "Insufficient stock" });
                continue;
              }
            }
            
            var totalAmount = item.quantity * item.unitPrice;
            var headers = debtsSheet.getRange(1, 1, 1, debtsSheet.getLastColumn()).getValues()[0];
            if (headers.length < 16) {
              debtsSheet.getRange(1, 1, 1, 16).setValues([["ID", "Date", "Borrower Name", "Borrower Phone", "Borrower NIDA", "Photo ID", "Product", "Quantity", "Unit Price", "Total Amount", "Paid Amount", "Status", "Branch ID", "Seller ID", "Notes", "Last Updated"]]);
            }
            ensureDebtPhotoColumns_(debtsSheet);
            debtsSheet.appendRow([debtId, timestamp, item.borrowerName, item.borrowerPhone || "", item.borrowerNIDA || "", photoId, item.product, item.quantity, item.unitPrice, totalAmount, item.paidAmount || 0, item.status || "pending", branchId, item.sellerId || null, item.notes || "", timestamp, item.unit || "", photoBase64, photoMimeType]);
            if (item.isStockProduct) updateStockAfterSale(d.databaseId, item.product, item.quantity, branchId, item.unit);
            results.operations.push({ type: "debt", debtId: debtId, product: item.product, branch: branchId, status: "success", borrower: item.borrowerName });
          } catch (e) {
            results.operations.push({ type: "debt", product: item.product, status: "failed", error: e.toString() });
          }
        }
      }
    }
    
    // ============================================
    // PROCESS CUSTOMERS DATA
    // ============================================
    if (d.customers && d.customers.length > 0) {
      const customersSheet = userDB.getSheetByName("CUSTOMERS");
      if (customersSheet) {
        for (var c = 0; c < d.customers.length; c++) {
          var item = d.customers[c];
          try {
            var data = customersSheet.getDataRange().getValues();
            var exists = false;
            for (var k = 1; k < data.length; k++) {
              if (data[k][1] === item.name) { exists = true; break; }
            }
            if (!exists) {
              customersSheet.appendRow([item.id || ("CUST_" + Date.now()), item.name, item.phone || "", item.email || "", item.address || "", item.totalPurchases || 0, item.lastPurchase || "", item.notes || "", new Date(item.created || Date.now())]);
              results.operations.push({ type: "customer_add", name: item.name, status: "success" });
            }
          } catch (e) {
            results.operations.push({ type: "customer_add", name: item.name, status: "failed", error: e.toString() });
          }
        }
      }
    }
    
    // ============================================
    // PROCESS SUPPLIERS DATA
    // ============================================
    if (d.suppliers && d.suppliers.length > 0) {
      const suppliersSheet = userDB.getSheetByName("SUPPLIERS");
      if (suppliersSheet) {
        for (var sp = 0; sp < d.suppliers.length; sp++) {
          var item = d.suppliers[sp];
          try {
            var data = suppliersSheet.getDataRange().getValues();
            var exists = false;
            for (var k = 1; k < data.length; k++) {
              if (data[k][1] === item.name) { exists = true; break; }
            }
            if (!exists) {
              suppliersSheet.appendRow([item.id || ("SUPP_" + Date.now()), item.name, item.contact || "", item.phone || "", item.email || "", item.address || "", item.products || "", item.totalSupplied || 0, item.lastSupply || "", item.notes || "", new Date(item.created || Date.now())]);
              results.operations.push({ type: "supplier_add", name: item.name, status: "success" });
            }
          } catch (e) {
            results.operations.push({ type: "supplier_add", name: item.name, status: "failed", error: e.toString() });
          }
        }
      }
    }
    
    // ============================================
    // PROCESS QUEUED PURCHASE REQUEST ACTIONS (Seller <-> Admin workflow
    // performed while offline - drafts, submits, approvals, rejections,
    // and item completions all get replayed here in the order they
    // happened on the device)
    // ============================================
    if (d.purchaseRequestActions && d.purchaseRequestActions.length > 0) {
      for (var pa = 0; pa < d.purchaseRequestActions.length; pa++) {
        var action = d.purchaseRequestActions[pa];
        try {
          var actionPayload = Object.assign({}, action.payload, { databaseId: d.databaseId });
          var actionResult = null;
          
          switch (action.type) {
            case "saveDraft":
              actionResult = saveDraftPurchaseRequest(actionPayload);
              break;
            case "submit":
              actionResult = submitPurchaseRequest(actionPayload);
              break;
            case "approve":
              actionResult = approvePurchaseRequest(actionPayload);
              break;
            case "reject":
              actionResult = rejectPurchaseRequest(actionPayload);
              break;
            case "delete":
              actionResult = deletePurchaseRequest(actionPayload);
              break;
            case "completeItem":
              actionResult = completePurchaseRequestItem(actionPayload);
              break;
            case "saveReviewDraft":
              actionResult = saveReviewDraft(actionPayload);
              break;
            case "submitAdminList":
              actionResult = submitAdminPurchaseList(actionPayload);
              break;
          }
          
          results.operations.push({ type: "purchase_request_" + action.type, requestId: action.payload && action.payload.requestId, status: "success" });
        } catch (e) {
          results.operations.push({ type: "purchase_request_" + action.type, status: "failed", error: e.toString() });
        }
      }
    }
    
    // ============================================
    // PROCESS QUEUED EDITS (Stock/Sale/Purchase records edited while
    // offline - these previously had no offline path at all and would
    // just fail outright with a network error).
    // ============================================
    if (d.stockEdits && d.stockEdits.length > 0) {
      for (var se = 0; se < d.stockEdits.length; se++) {
        var edit = d.stockEdits[se];
        try {
          updateStock(Object.assign({}, edit, { databaseId: d.databaseId }));
          results.operations.push({ type: "stock_edit", product: edit.newProduct || edit.product, status: "success" });
        } catch (e) {
          results.operations.push({ type: "stock_edit", product: edit.product, status: "failed", error: e.toString() });
        }
      }
    }
    
    if (d.saleEdits && d.saleEdits.length > 0) {
      for (var sve = 0; sve < d.saleEdits.length; sve++) {
        var edit = d.saleEdits[sve];
        try {
          updateSale(Object.assign({}, edit, { databaseId: d.databaseId }));
          results.operations.push({ type: "sale_edit", product: edit.product, status: "success" });
        } catch (e) {
          results.operations.push({ type: "sale_edit", product: edit.product, status: "failed", error: e.toString() });
        }
      }
    }
    
    if (d.purchaseEdits && d.purchaseEdits.length > 0) {
      for (var pe = 0; pe < d.purchaseEdits.length; pe++) {
        var edit = d.purchaseEdits[pe];
        try {
          updatePurchase(Object.assign({}, edit, { databaseId: d.databaseId }));
          results.operations.push({ type: "purchase_edit", product: edit.product, status: "success" });
        } catch (e) {
          results.operations.push({ type: "purchase_edit", product: edit.product, status: "failed", error: e.toString() });
        }
      }
    }
    
    if (d.debtEdits && d.debtEdits.length > 0) {
      for (var de = 0; de < d.debtEdits.length; de++) {
        var edit = d.debtEdits[de];
        try {
          updateDebt(Object.assign({}, edit, { databaseId: d.databaseId }));
          results.operations.push({ type: "debt_edit", debtId: edit.debtId, status: "success" });
        } catch (e) {
          results.operations.push({ type: "debt_edit", debtId: edit.debtId, status: "failed", error: e.toString() });
        }
      }
    }
    
    // ============================================
    // PROCESS QUEUED DELETES (Stock/Sale/Purchase/Debt deleted while
    // offline - these previously had no offline path and would just fail
    // outright with a network error, silently losing the delete request).
    // Sale/Purchase are identified by row index, so within each type they
    // must be processed in DESCENDING row-index order - deleting a row
    // shifts every row after it up by one, which would otherwise point a
    // later delete in the same batch at the wrong row.
    // ============================================
    if (d.deletes && d.deletes.length > 0) {
      var saleDeletes = d.deletes.filter(function(x) { return x.type === "sale"; }).sort(function(a, b) { return b.rowIndex - a.rowIndex; });
      var purchaseDeletes = d.deletes.filter(function(x) { return x.type === "purchase"; }).sort(function(a, b) { return b.rowIndex - a.rowIndex; });
      var stockDeletes = d.deletes.filter(function(x) { return x.type === "stock"; });
      var debtDeletes = d.deletes.filter(function(x) { return x.type === "debt"; });
      
      saleDeletes.forEach(function(del) {
        try {
          // IMPORTANT: deleteSaleRow returns an output() response instead of
          // throwing on a normal lookup failure. Parse the actual response so
          // a failed delete is NOT falsely reported as successful.
          var deleteResponse = deleteSaleRow(Object.assign({}, del, { databaseId: d.databaseId }));
          var deleteResult = {};
          try { deleteResult = JSON.parse(deleteResponse.getContent()); } catch (parseErr) {}
          if (deleteResult && deleteResult.success) {
            results.operations.push({ type: "sale_delete", rowIndex: del.rowIndex, operationId: del.operationId || null, status: "success" });
          } else {
            results.operations.push({ type: "sale_delete", rowIndex: del.rowIndex, operationId: del.operationId || null, status: "failed", error: deleteResult.error || "Sale delete failed" });
          }
        } catch (e) {
          results.operations.push({ type: "sale_delete", rowIndex: del.rowIndex, operationId: del.operationId || null, status: "failed", error: e.toString() });
        }
      });
      
      purchaseDeletes.forEach(function(del) {
        try {
          deletePurchaseRow(Object.assign({}, del, { databaseId: d.databaseId }));
          results.operations.push({ type: "purchase_delete", rowIndex: del.rowIndex, status: "success" });
        } catch (e) {
          results.operations.push({ type: "purchase_delete", rowIndex: del.rowIndex, status: "failed", error: e.toString() });
        }
      });
      
      stockDeletes.forEach(function(del) {
        try {
          deleteStockRow(Object.assign({}, del, { databaseId: d.databaseId }));
          results.operations.push({ type: "stock_delete", product: del.product, status: "success" });
        } catch (e) {
          results.operations.push({ type: "stock_delete", product: del.product, status: "failed", error: e.toString() });
        }
      });
      
      debtDeletes.forEach(function(del) {
        try {
          deleteDebt(Object.assign({}, del, { databaseId: d.databaseId, email: d.email }));
          results.operations.push({ type: "debt_delete", debtId: del.debtId, status: "success" });
        } catch (e) {
          results.operations.push({ type: "debt_delete", debtId: del.debtId, status: "failed", error: e.toString() });
        }
      });
    }
    
    var failedCount = results.operations.filter(function(op) { return op.status === "failed"; }).length;
    results.success = failedCount === 0;
    results.failedCount = failedCount;
    return output(results);
    
  } catch (error) {
    return output({ success: false, error: "Failed to sync offline data: " + error.toString() });
  }
}

function uploadBorrowerPhotoDirect(d) {
    try {
        if (!d.email || !d.photoData) {
            return { success: false, error: "Email and photo data required" };
        }
        var fileName = "borrower_" + (d.debtId || Date.now()) + "_" + d.email.replace(/[@.]/g, "_") + "." + (d.fileExtension || 'jpg');
        var blob = Utilities.newBlob(Utilities.base64Decode(d.photoData), d.mimeType, fileName);
        var fileMetadata = { 'title': fileName, 'mimeType': d.mimeType, 'parents': [{'id': DEBT_PHOTOS_FOLDER_ID}] };
        var file = Drive.Files.insert(fileMetadata, blob, { 'convert': false });
        Drive.Permissions.insert({ 'role': 'reader', 'type': 'anyone' }, file.id, { 'sendNotificationEmails': 'false' });
        var lh3Url = "https://lh3.googleusercontent.com/d/" + file.id + "=s200";
        var thumbnailUrl = "https://drive.google.com/thumbnail?id=" + file.id + "&sz=s200";
        var directUrl = "https://drive.google.com/uc?export=view&id=" + file.id;
        return { success: true, photoId: file.id, photoUrl: lh3Url, thumbnailUrl: thumbnailUrl, directUrl: directUrl, lh3Url: lh3Url };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}


/* ===== HELPER FUNCTION: Check Branch Stock Availability ===== */
function checkBranchStockAvailability(databaseId, product, requestedQty, branchId, unit) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return { available: false, currentStock: 0 };
    }
    
    const data = stockSheet.getDataRange().getValues();
    let totalQuantity = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1] === product) {
        const itemBranchId = data[i][8] || "MAIN";
        const itemUnit = data[i][9] || "";
        // Match the SAME unit variant only - "Rice" sold by the Kg and by
        // the Bag are different items and must never be added together.
        if (itemBranchId === branchId && (!unit || itemUnit === unit)) {
          totalQuantity += (data[i][2] || 0);
        }
      }
    }
    
    return {
      available: totalQuantity >= requestedQty,
      currentStock: totalQuantity
    };
  } catch (error) {
    return { available: false, currentStock: 0 };
  }
}

/* ===== HELPER FUNCTION: Update Stock After Sale ===== */
function updateStockAfterSale(databaseId, product, soldQty, branchId, unit) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) return;
    
    const data = stockSheet.getDataRange().getValues();
    let remainingToReduce = soldQty;
    
    // Get all rows for this EXACT product + unit variant in this branch -
    // never reduce a different unit's stock just because the name matches.
    const productRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1] === product) {
        const itemBranchId = data[i][8] || "MAIN";
        const itemUnit = data[i][9] || "";
        if (itemBranchId === branchId && (!unit || itemUnit === unit)) {
          productRows.push({
            row: i + 1,
            quantity: data[i][2] || 0,
            buyPrice: data[i][3] || 0
          });
        }
      }
    }
    
    // Sort by date (oldest first) for FIFO
    productRows.sort((a, b) => a.row - b.row);
    
    // Reduce from oldest entries first (FIFO)
    for (let i = 0; i < productRows.length && remainingToReduce > 0; i++) {
      const row = productRows[i];
      
      if (row.quantity > 0) {
        const reduceBy = Math.min(row.quantity, remainingToReduce);
        const newQty = row.quantity - reduceBy;
        
        stockSheet.getRange(row.row, 3).setValue(newQty);
        const newValue = newQty * row.buyPrice;
        stockSheet.getRange(row.row, 6).setValue(newValue);
        
        remainingToReduce -= reduceBy;
      }
    }
    if (remainingToReduce > 0) {
      throw new Error("Stock update incomplete: unable to reduce the full sale quantity.");
    }
    
  } catch (error) {
    console.error("Error updating stock after sale:", error);
    throw error;
  }
}

/* ===== HELPER FUNCTION: Update Seller Totals ===== */
function updateSellerTotals(databaseId, sellerId, amount, type) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    
    if (!sellersSheet) return;
    
    const data = sellersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sellerId) {
        const row = i + 1;
        const currentTotalSales = data[i][5] || 0;
        const currentTotalPurchases = data[i][6] || 0;
        
        if (type === 'sale') {
          sellersSheet.getRange(row, 6).setValue(currentTotalSales + amount);
        } else if (type === 'purchase') {
          sellersSheet.getRange(row, 7).setValue(currentTotalPurchases + amount);
        }
        
        sellersSheet.getRange(row, 8).setValue(new Date());
        break;
      }
    }
  } catch (error) {
    console.error("Error updating seller totals:", error);
  }
}


/* ===== CUSTOMER MANAGEMENT ===== */

function addCustomer(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    let customersSheet = userDB.getSheetByName("CUSTOMERS");
    if (!customersSheet) {
      customersSheet = userDB.insertSheet("CUSTOMERS");
      customersSheet.appendRow(["ID", "Name", "Phone", "Email", "Address", "Total Purchases", "Last Purchase", "Notes", "Created"]);
      customersSheet.getRange("A1:I1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    }
    
    const timestamp = new Date();
    const customerId = "CUST_" + Date.now() + "_" + Math.random().toString(36).substr(2, 5);
    
    customersSheet.appendRow([
      customerId,
      d.name,
      d.phone || "",
      d.email || "",
      d.address || "",
      0,
      "",
      d.notes || "",
      timestamp
    ]);
    
    return output({ success: true, message: `Customer ${d.name} added successfully` });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function getCustomers(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const customersSheet = userDB.getSheetByName("CUSTOMERS");
    if (!customersSheet) {
      return output({ success: true, data: [], count: 0 });
    }
    
    const data = customersSheet.getDataRange().getValues();
    const customers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        customers.push({
          id: data[i][0],
          name: data[i][1],
          phone: data[i][2],
          email: data[i][3],
          address: data[i][4],
          totalPurchases: data[i][5] || 0,
          lastPurchase: data[i][6],
          notes: data[i][7],
          created: data[i][8]
        });
      }
    }
    
    return output({ success: true, data: customers, count: customers.length });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function updateCustomer(d) {
  try {
    if (!d.databaseId || !d.customerId) {
      return output({ success: false, error: "Database ID and Customer ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const customersSheet = userDB.getSheetByName("CUSTOMERS");
    if (!customersSheet) {
      return output({ success: false, error: "Customers sheet not found" });
    }
    
    const data = customersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.customerId) {
        const row = i + 1;
        
        if (d.name) customersSheet.getRange(row, 2).setValue(d.name);
        if (d.phone !== undefined) customersSheet.getRange(row, 3).setValue(d.phone);
        if (d.email !== undefined) customersSheet.getRange(row, 4).setValue(d.email);
        if (d.address !== undefined) customersSheet.getRange(row, 5).setValue(d.address);
        if (d.notes !== undefined) customersSheet.getRange(row, 8).setValue(d.notes);
        
        return output({ success: true, message: "Customer updated successfully" });
      }
    }
    
    return output({ success: false, error: "Customer not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function deleteCustomer(d) {
  try {
    if (!d.databaseId || !d.customerId) {
      return output({ success: false, error: "Database ID and Customer ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const customersSheet = userDB.getSheetByName("CUSTOMERS");
    if (!customersSheet) {
      return output({ success: false, error: "Customers sheet not found" });
    }
    
    const data = customersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.customerId) {
        customersSheet.deleteRow(i + 1);
        return output({ success: true, message: "Customer deleted successfully" });
      }
    }
    
    return output({ success: false, error: "Customer not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== SUPPLIER MANAGEMENT ===== */

function addSupplier(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    let suppliersSheet = userDB.getSheetByName("SUPPLIERS");
    if (!suppliersSheet) {
      suppliersSheet = userDB.insertSheet("SUPPLIERS");
      suppliersSheet.appendRow(["ID", "Name", "Contact Person", "Phone", "Email", "Address", "Products Supplied", "Total Supplied", "Last Supply", "Notes", "Created"]);
      suppliersSheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    }
    
    const timestamp = new Date();
    const supplierId = "SUPP_" + Date.now() + "_" + Math.random().toString(36).substr(2, 5);
    
    suppliersSheet.appendRow([
      supplierId,
      d.name,
      d.contact || "",
      d.phone || "",
      d.email || "",
      d.address || "",
      d.products || "",
      0,
      "",
      d.notes || "",
      timestamp
    ]);
    
    return output({ success: true, message: `Supplier ${d.name} added successfully` });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function getSuppliers(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const suppliersSheet = userDB.getSheetByName("SUPPLIERS");
    if (!suppliersSheet) {
      return output({ success: true, data: [], count: 0 });
    }
    
    const data = suppliersSheet.getDataRange().getValues();
    const suppliers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        suppliers.push({
          id: data[i][0],
          name: data[i][1],
          contact: data[i][2],
          phone: data[i][3],
          email: data[i][4],
          address: data[i][5],
          products: data[i][6],
          totalSupplied: data[i][7] || 0,
          lastSupply: data[i][8],
          notes: data[i][9],
          created: data[i][10]
        });
      }
    }
    
    return output({ success: true, data: suppliers, count: suppliers.length });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function updateSupplier(d) {
  try {
    if (!d.databaseId || !d.supplierId) {
      return output({ success: false, error: "Database ID and Supplier ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const suppliersSheet = userDB.getSheetByName("SUPPLIERS");
    if (!suppliersSheet) {
      return output({ success: false, error: "Suppliers sheet not found" });
    }
    
    const data = suppliersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.supplierId) {
        const row = i + 1;
        
        if (d.name) suppliersSheet.getRange(row, 2).setValue(d.name);
        if (d.contact !== undefined) suppliersSheet.getRange(row, 3).setValue(d.contact);
        if (d.phone !== undefined) suppliersSheet.getRange(row, 4).setValue(d.phone);
        if (d.email !== undefined) suppliersSheet.getRange(row, 5).setValue(d.email);
        if (d.address !== undefined) suppliersSheet.getRange(row, 6).setValue(d.address);
        if (d.products !== undefined) suppliersSheet.getRange(row, 7).setValue(d.products);
        if (d.notes !== undefined) suppliersSheet.getRange(row, 10).setValue(d.notes);
        
        return output({ success: true, message: "Supplier updated successfully" });
      }
    }
    
    return output({ success: false, error: "Supplier not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function deleteSupplier(d) {
  try {
    if (!d.databaseId || !d.supplierId) {
      return output({ success: false, error: "Database ID and Supplier ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const suppliersSheet = userDB.getSheetByName("SUPPLIERS");
    if (!suppliersSheet) {
      return output({ success: false, error: "Suppliers sheet not found" });
    }
    
    const data = suppliersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.supplierId) {
        suppliersSheet.deleteRow(i + 1);
        return output({ success: true, message: "Supplier deleted successfully" });
      }
    }
    
    return output({ success: false, error: "Supplier not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== SELLER MANAGEMENT ===== */

function addSeller(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    let sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      sellersSheet = userDB.insertSheet("SELLERS");
      sellersSheet.appendRow(["ID", "Name", "Phone", "Email", "Password", "Total Sales", "Total Purchases", "Last Active", "Created", "Status", "Branch ID"]);
      sellersSheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    }
    
    // IMPORTANT: the new Seller's email must travel as d.sellerEmail, not
    // d.email - doPost's session check always overwrites d.email with the
    // logged-in Admin's OWN verified email for every authenticated action,
    // so using d.email here would silently save the Admin's email instead
    // of the Seller's.
    if (!d.sellerEmail) {
      return output({ success: false, error: "Seller email required" });
    }
    
    // Global uniqueness check across EVERY business, not just this one -
    // prevents two different Admin accounts from creating a Seller with
    // the same email and colliding at login.
    const masterDB = getMasterDatabase();
    const masterSellersSheet = ensureMasterSellersSheet(masterDB);
    const masterSellersData = masterSellersSheet.getDataRange().getValues();
    for (let m = 1; m < masterSellersData.length; m++) {
      if (masterSellersData[m][0] === d.sellerEmail) {
        return output({ success: false, error: "This email is already registered to a Seller (possibly under a different business). Please use another email." });
      }
    }
    
    const data = sellersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === d.sellerEmail) {
        return output({ success: false, error: "Email already used by another seller" });
      }
    }
    
    const timestamp = new Date();
    const sellerId = "SELL_" + Date.now() + "_" + Math.random().toString(36).substr(2, 5);
    const branchId = d.branchId || "MAIN"; // Assign seller to branch
    
    sellersSheet.appendRow([
      sellerId,
      d.name,
      d.phone || "",
      d.sellerEmail,
      d.password,
      0,
      0,
      timestamp,
      timestamp,
      "active",
      branchId
    ]);
    
    masterSellersSheet.appendRow([d.sellerEmail, d.email, sellerId, d.databaseId, timestamp, d.password, d.name, "active", branchId]);
    
    return output({ success: true, message: `Seller ${d.name} added successfully` });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function getSellers(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: true, data: [], count: 0 });
    }
    
    const data = sellersSheet.getDataRange().getValues();
    const sellers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        sellers.push({
          id: data[i][0],
          name: data[i][1],
          phone: data[i][2],
          email: data[i][3],
          password: data[i][4],
          totalSales: data[i][5] || 0,
          totalPurchases: data[i][6] || 0,
          lastActive: data[i][7],
          created: data[i][8],
          status: data[i][9] || "active",
          branchId: data[i][10] || "MAIN" // Include branch ID
        });
      }
    }
    
    return output({ success: true, data: sellers, count: sellers.length });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ============================================
   SELLER PURCHASE-COMPLETION PIN
   A Seller sets their own PIN to protect "Complete Purchase" actions -
   Admin can force a reset (in case it's forgotten) but can never see the
   PIN itself, only whether one is set. The hash is also returned to the
   Seller's own device so it can be verified even while offline.
   ============================================ */

// Matches the frontend's rule exactly: at least 8 characters AND at least
// 3 of the 4 character categories (lower, upper, digit, symbol).
function isPasswordStrongEnough(password) {
  if (!password || String(password).length < 8) return false;
  var categories = 0;
  if (/[a-z]/.test(password)) categories++;
  if (/[A-Z]/.test(password)) categories++;
  if (/[0-9]/.test(password)) categories++;
  if (/[^A-Za-z0-9]/.test(password)) categories++;
  return categories >= 3;
}

function hashPin(pin) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin));
  return digest.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function findSellerRow(sellersSheet, sellerId) {
  var data = sellersSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === sellerId) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

// Seller sets or changes their own PIN. Only the Seller who is logged in
// as d.sellerId should ever call this (enforced by the app's own session,
// same trust model as the rest of this Seller-scoped workflow).
function setSellerPin(d) {
  try {
    if (!d.databaseId || !d.sellerId || !d.pin) {
      return output({ success: false, error: "Database ID, Seller ID, and PIN required" });
    }
    if (!d.password) {
      return output({ success: false, error: "Please confirm your account password" });
    }
    if (String(d.pin).length < 4) {
      return output({ success: false, error: "PIN must be at least 4 digits" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: false, error: "Sellers sheet not found" });
    }
    
    const headers = sellersSheet.getRange(1, 1, 1, sellersSheet.getLastColumn()).getValues()[0];
    if (headers.length < 12) {
      sellersSheet.getRange(1, 12).setValue("PinHash");
    }
    
    const found = findSellerRow(sellersSheet, d.sellerId);
    if (!found) {
      return output({ success: false, error: "Seller not found" });
    }
    
    // Must confirm the account password before a PIN can be set/changed.
    const storedPassword = found.row[4];
    if (storedPassword !== d.password) {
      return output({ success: false, error: "Incorrect password" });
    }
    
    const pinHash = hashPin(d.pin);
    sellersSheet.getRange(found.rowIndex, 12).setValue(pinHash);
    
    // Returned only to the Seller who just set their own PIN, so their
    // device can verify it offline too - never exposed via getSellers().
    return output({ success: true, pinHash: pinHash, message: "PIN saved" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Admin forces a reset (clears the PIN, forcing the Seller to set a new
// one) - Admin never learns the actual PIN or its hash.
function resetSellerPin(d) {
  try {
    if (!d.databaseId || !d.sellerId) {
      return output({ success: false, error: "Database ID and Seller ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: false, error: "Sellers sheet not found" });
    }
    
    const found = findSellerRow(sellersSheet, d.sellerId);
    if (!found) {
      return output({ success: false, error: "Seller not found" });
    }
    
    sellersSheet.getRange(found.rowIndex, 12).setValue("");
    
    return output({ success: true, message: "PIN reset - the seller will need to set a new one" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Verifies a PIN attempt against the stored hash (used online; the Seller's
// device also keeps its own cached copy of the hash for offline checks).
function verifySellerPin(d) {
  try {
    if (!d.databaseId || !d.sellerId || !d.pin) {
      return output({ success: false, error: "Database ID, Seller ID, and PIN required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: false, error: "Sellers sheet not found" });
    }
    
    const found = findSellerRow(sellersSheet, d.sellerId);
    if (!found) {
      return output({ success: false, error: "Seller not found" });
    }
    
    const storedHash = found.row[11] || "";
    const valid = storedHash && storedHash === hashPin(d.pin);
    
    return output({ success: true, valid: !!valid, hasPinSet: !!storedHash });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Tells the Seller's own device whether a PIN is already set (and its
// hash, so it can be cached for offline verification) right after login -
// scoped to their own sellerId only.
function getMySellerPinStatus(d) {
  try {
    if (!d.databaseId || !d.sellerId) {
      return output({ success: false, error: "Database ID and Seller ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: true, hasPinSet: false, pinHash: "" });
    }
    
    const found = findSellerRow(sellersSheet, d.sellerId);
    if (!found) {
      return output({ success: true, hasPinSet: false, pinHash: "" });
    }
    
    const storedHash = found.row[11] || "";
    return output({ success: true, hasPinSet: !!storedHash, pinHash: storedHash });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ============================================
   ADMIN'S OWN PIN (used to unlock the Record Sale form)
   Stored in the generic SETTINGS sheet (section "SECURITY") that already
   exists per business - same hash-based, offline-verifiable approach as
   the Seller's Purchase-completion PIN.
   ============================================ */

function setAdminPin(d) {
  try {
    if (!d.databaseId || !d.pin) {
      return output({ success: false, error: "Database ID and PIN required" });
    }
    if (!d.password) {
      return output({ success: false, error: "Please confirm your account password" });
    }
    if (String(d.pin).length < 4) {
      return output({ success: false, error: "PIN must be at least 4 digits" });
    }
    
    // Must confirm the account (login) password before a PIN can be set/changed.
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const masterData = masterSheet.getDataRange().getValues();
    let passwordOk = false;
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][1] === d.email && masterData[i][2] === d.password) {
        passwordOk = true;
        break;
      }
    }
    if (!passwordOk) {
      return output({ success: false, error: "Incorrect password" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    if (!settingsSheet) {
      return output({ success: false, error: "Settings sheet not found" });
    }
    
    const pinHash = hashPin(d.pin);
    updateSettingsKey(settingsSheet, "SECURITY", "admin_pin_hash", pinHash);
    
    return output({ success: true, pinHash: pinHash, message: "PIN saved" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function getMyAdminPinStatus(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const settings = getUserSettingsFromDB(d.databaseId);
    const storedHash = (settings.SECURITY && settings.SECURITY.admin_pin_hash && settings.SECURITY.admin_pin_hash.value) || "";
    
    return output({ success: true, hasPinSet: !!storedHash, pinHash: storedHash });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ============================================
   MODAL PIN RESTRICTIONS (Admin-only setting)
   Admin chooses which "Add" modals require a PIN to open. The list is
   stored in the generic SETTINGS sheet and already rides along with every
   syncAllUserData response, so both Admin and Sellers have it cached for
   offline use the moment they log in / sync.
   ============================================ */

function setModalRestrictions(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    if (!Array.isArray(d.restrictedModals)) {
      return output({ success: false, error: "restrictedModals must be an array" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    if (!settingsSheet) {
      return output({ success: false, error: "Settings sheet not found" });
    }
    
    updateSettingsKey(settingsSheet, "MODAL_RESTRICTIONS", "restricted_list", JSON.stringify(d.restrictedModals));
    
    return output({ success: true, message: "Restrictions saved" });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Lightweight dedicated fetch (no other business data) so Sellers/Admin can
// poll this frequently (every few seconds) without the cost of a full sync,
// letting an Admin's restriction change reach other devices almost instantly.
function getModalRestrictions(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    const settings = getUserSettingsFromDB(d.databaseId);
    const raw = (settings.MODAL_RESTRICTIONS && settings.MODAL_RESTRICTIONS.restricted_list && settings.MODAL_RESTRICTIONS.restricted_list.value) || "[]";
    let restrictedModals = [];
    try {
      restrictedModals = JSON.parse(raw);
    } catch (e) {
      restrictedModals = [];
    }
    return output({ success: true, restrictedModals: restrictedModals });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function updateSeller(d) {
  try {
    if (!d.databaseId || !d.sellerId) {
      return output({ success: false, error: "Database ID and Seller ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: false, error: "Sellers sheet not found" });
    }
    
    const data = sellersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.sellerId) {
        const row = i + 1;
        
        if (d.name) sellersSheet.getRange(row, 2).setValue(d.name);
        if (d.phone !== undefined) sellersSheet.getRange(row, 3).setValue(d.phone);
        // Same collision as addSeller() - the new email must arrive as
        // d.sellerEmail, since d.email is always the Admin's own verified
        // identity, overwritten by doPost's session check.
        if (d.sellerEmail) {
          const oldEmail = data[i][3];
          if (d.sellerEmail !== oldEmail) {
            // Keep the master registry's global uniqueness check honest.
            const masterDBCheck = getMasterDatabase();
            const masterSellersSheetCheck = ensureMasterSellersSheet(masterDBCheck);
            const masterSellersDataCheck = masterSellersSheetCheck.getDataRange().getValues();
            for (let m = 1; m < masterSellersDataCheck.length; m++) {
              if (masterSellersDataCheck[m][0] === d.sellerEmail) {
                return output({ success: false, error: "This email is already registered to a Seller (possibly under a different business). Please use another email." });
              }
            }
          }
          sellersSheet.getRange(row, 4).setValue(d.sellerEmail);
        }
        if (d.password) sellersSheet.getRange(row, 5).setValue(d.password);
        if (d.status) sellersSheet.getRange(row, 10).setValue(d.status);
        if (d.branchId) sellersSheet.getRange(row, 11).setValue(d.branchId); // Update branch
        
        // Keep the master registry's cached copy (email/password/name/
        // status/branch) in sync, so login never needs to open this
        // business's own database just to check credentials.
        try {
          const masterDB = getMasterDatabase();
          const masterSellersSheet = ensureMasterSellersSheet(masterDB);
          const found = findMasterSellerRow(masterSellersSheet, d.sellerId);
          if (found) {
            if (d.sellerEmail) masterSellersSheet.getRange(found.rowIndex, 1).setValue(d.sellerEmail);
            if (d.password) masterSellersSheet.getRange(found.rowIndex, 6).setValue(d.password);
            if (d.name) masterSellersSheet.getRange(found.rowIndex, 7).setValue(d.name);
            if (d.status) masterSellersSheet.getRange(found.rowIndex, 8).setValue(d.status);
            if (d.branchId) masterSellersSheet.getRange(found.rowIndex, 9).setValue(d.branchId);
          }
        } catch (e) {
          console.log("Could not sync seller update to master registry:", e);
        }
        
        return output({ success: true, message: "Seller updated successfully" });
      }
    }
    
    return output({ success: false, error: "Seller not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function deleteSeller(d) {
  try {
    if (!d.databaseId || !d.sellerId) {
      return output({ success: false, error: "Database ID and Seller ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    if (!sellersSheet) {
      return output({ success: false, error: "Sellers sheet not found" });
    }
    
    const data = sellersSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.sellerId) {
        sellersSheet.deleteRow(i + 1);
        
        // Also remove this Seller from the master registry.
        try {
          const masterDB = getMasterDatabase();
          const masterSellersSheet = ensureMasterSellersSheet(masterDB);
          const masterSellersData = masterSellersSheet.getDataRange().getValues();
          for (let m = 1; m < masterSellersData.length; m++) {
            if (masterSellersData[m][2] === d.sellerId) {
              masterSellersSheet.deleteRow(m + 1);
              break;
            }
          }
        } catch (e) {
          console.log("Could not remove seller from master registry:", e);
        }
        
        return output({ success: true, message: "Seller deleted successfully" });
      }
    }
    
    return output({ success: false, error: "Seller not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== BRANCHES MANAGEMENT (NEW) ===== */

function addBranch(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    let branchesSheet = userDB.getSheetByName("BRANCHES");
    if (!branchesSheet) {
      branchesSheet = userDB.insertSheet("BRANCHES");
      branchesSheet.appendRow(["ID", "Name", "Location", "Manager", "Phone", "Created", "Status"]);
      branchesSheet.getRange("A1:G1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    }
    
    // Generate branch ID
    const branchId = "BR_" + Date.now() + "_" + Math.random().toString(36).substr(2, 5).toUpperCase();
    
    const timestamp = new Date();
    
    branchesSheet.appendRow([
      branchId,
      d.name,
      d.location || "",
      d.manager || "",
      d.phone || "",
      timestamp,
      "active"
    ]);
    
    return output({ 
      success: true, 
      message: `Branch ${d.name} added successfully`,
      branchId: branchId
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function getBranches(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const branchesSheet = userDB.getSheetByName("BRANCHES");
    
    if (!branchesSheet) {
      return output({ success: true, data: [], count: 0 });
    }
    
    const data = branchesSheet.getDataRange().getValues();
    const branches = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        branches.push({
          id: data[i][0],
          name: data[i][1],
          location: data[i][2] || "",
          manager: data[i][3] || "",
          phone: data[i][4] || "",
          created: data[i][5],
          status: data[i][6] || "active"
        });
      }
    }
    
    return output({
      success: true,
      data: branches,
      count: branches.length
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function updateBranch(d) {
  try {
    if (!d.databaseId || !d.branchId) {
      return output({ success: false, error: "Database ID and Branch ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const branchesSheet = userDB.getSheetByName("BRANCHES");
    
    if (!branchesSheet) {
      return output({ success: false, error: "Branches sheet not found" });
    }
    
    const data = branchesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.branchId) {
        const row = i + 1;
        
        if (d.name) branchesSheet.getRange(row, 2).setValue(d.name);
        if (d.location !== undefined) branchesSheet.getRange(row, 3).setValue(d.location);
        if (d.manager !== undefined) branchesSheet.getRange(row, 4).setValue(d.manager);
        if (d.phone !== undefined) branchesSheet.getRange(row, 5).setValue(d.phone);
        if (d.status) branchesSheet.getRange(row, 7).setValue(d.status);
        
        return output({ success: true, message: "Branch updated successfully" });
      }
    }
    
    return output({ success: false, error: "Branch not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function deleteBranch(d) {
  try {
    if (!d.databaseId || !d.branchId) {
      return output({ success: false, error: "Database ID and Branch ID required" });
    }
    
    if (d.branchId === "MAIN") {
      return output({ success: false, error: "Cannot delete Main Branch" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const branchesSheet = userDB.getSheetByName("BRANCHES");
    
    if (!branchesSheet) {
      return output({ success: false, error: "Branches sheet not found" });
    }
    
    const data = branchesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.branchId) {
        branchesSheet.deleteRow(i + 1);
        return output({ success: true, message: "Branch deleted successfully" });
      }
    }
    
    return output({ success: false, error: "Branch not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== GET SELLER SALES ===== */
function getSellerSales(d) {
  try {
    if (!d.databaseId || !d.sellerId) {
      return output({ success: false, error: "Database ID and Seller ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const salesSheet = userDB.getSheetByName("SALES");
    
    if (!salesSheet) {
      return output({ success: false, error: "Sales sheet not found" });
    }
    
    const data = salesSheet.getDataRange().getValues();
    const sales = [];
    let totalAmount = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][7] === d.sellerId) {
        sales.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          unitPrice: data[i][3],
          total: data[i][4],
          customer: data[i][5],
          paymentMethod: data[i][6],
          branchId: data[i][8] || "MAIN"
        });
        totalAmount += (data[i][4] || 0);
      }
    }
    
    return output({
      success: true,
      data: sales,
      summary: { count: sales.length, totalAmount: totalAmount }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== GET SELLER PURCHASES ===== */
function getSellerPurchases(d) {
  try {
    if (!d.databaseId || !d.sellerId) {
      return output({ success: false, error: "Database ID and Seller ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    
    if (!purchasesSheet) {
      return output({ success: false, error: "Purchases sheet not found" });
    }
    
    const data = purchasesSheet.getDataRange().getValues();
    const purchases = [];
    let totalAmount = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][7] === d.sellerId) {
        purchases.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          unitCost: data[i][3],
          total: data[i][4],
          supplier: data[i][5],
          paymentMethod: data[i][6],
          branchId: data[i][8] || "MAIN"
        });
        totalAmount += (data[i][4] || 0);
      }
    }
    
    return output({
      success: true,
      data: purchases,
      summary: { count: purchases.length, totalAmount: totalAmount }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== GET PURCHASES ===== */
function getPurchases(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    const sellersSheet = userDB.getSheetByName("SELLERS");
    
    if (!purchasesSheet) {
      return output({ success: false, error: "Purchases sheet not found" });
    }
    
    const sellerMap = {};
    if (sellersSheet) {
      const sellersData = sellersSheet.getDataRange().getValues();
      for (let i = 1; i < sellersData.length; i++) {
        if (sellersData[i][0]) {
          sellerMap[sellersData[i][0]] = sellersData[i][1];
        }
      }
    }
    
    const data = purchasesSheet.getDataRange().getValues();
    const operationalBoundaries = getOperationalOpenPeriodStarts_(userDB);
    const purchasesData = [];
    let totalAmount = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const recordSellerId = data[i][7];
        
        if (d.sellerId && recordSellerId !== d.sellerId) {
          continue;
        }
        if (d.branchId && (data[i][8] || "MAIN") !== d.branchId) {
          continue;
        }

        // Purchases shown in the live Purchase table also reset at the
        // exact Close Accounts moment.
        const purchaseTime = new Date(data[i][0]).getTime();
        if (isNaN(purchaseTime)) continue;
        const purchaseBranch = data[i][8] || "MAIN";
        const openPeriodStart = getOperationalOpenPeriodStartFromMap_(operationalBoundaries, purchaseBranch);
        if (openPeriodStart !== null && purchaseTime <= openPeriodStart) continue;
        
        const sellerName = recordSellerId ? (sellerMap[recordSellerId] || "Unknown Seller") : "Admin";
        
        purchasesData.push({
          date: data[i][0],
          product: data[i][1],
          quantity: data[i][2],
          unitCost: data[i][3],
          total: data[i][4],
          supplier: data[i][5],
          paymentMethod: data[i][6],
          sellerId: recordSellerId,
          sellerName: sellerName,
          branchId: data[i][8] || "MAIN",
          category: data[i][9] || "",
          unit: data[i][10] || "",
          isStockProduct: data[i][11] !== "nonstock",
          rowIndex: i + 1
        });
        
        totalAmount += (data[i][4] || 0);
      }
    }
    
    return output({
      success: true,
      data: purchasesData,
      summary: {
        totalAmount: totalAmount,
        count: purchasesData.length
      }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== GET REPORT ===== */
function getReport(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    const salesSheet = userDB.getSheetByName("SALES");
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    const customersSheet = userDB.getSheetByName("CUSTOMERS");
    
    // Get stock data
    let stockCount = 0;
    let totalStockValue = 0;
    let stockData = [];
    
    if (stockSheet) {
      const data = stockSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          stockCount++;
          totalStockValue += (data[i][5] || 0);
          stockData.push({
            date: data[i][0],
            product: data[i][1],
            quantity: data[i][2],
            buyPrice: data[i][3],
            sellPrice: data[i][4],
            value: data[i][5],
            category: data[i][6] || "",
            branchId: data[i][8] || "MAIN"
          });
        }
      }
    }
    
    // Get sales data
    let salesCount = 0;
    let totalSales = 0;
    let salesData = [];
    
    if (salesSheet) {
      const data = salesSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          if (d.sellerId && data[i][7] !== d.sellerId) {
            continue;
          }
          salesCount++;
          totalSales += (data[i][4] || 0);
          salesData.push({
            date: data[i][0],
            product: data[i][1],
            quantity: data[i][2],
            unitPrice: data[i][3],
            total: data[i][4],
            customer: data[i][5] || "Walk-in",
            paymentMethod: data[i][6],
            sellerId: data[i][7] || null,
            branchId: data[i][8] || "MAIN"
          });
        }
      }
    }
    
    // Get purchases data
    let purchasesCount = 0;
    let totalPurchases = 0;
    let purchasesData = [];
    
    if (purchasesSheet) {
      const data = purchasesSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          if (d.sellerId && data[i][7] !== d.sellerId) {
            continue;
          }
          purchasesCount++;
          totalPurchases += (data[i][4] || 0);
          purchasesData.push({
            date: data[i][0],
            product: data[i][1],
            quantity: data[i][2],
            unitCost: data[i][3],
            total: data[i][4],
            supplier: data[i][5] || "Local Supplier",
            paymentMethod: data[i][6],
            sellerId: data[i][7] || null,
            branchId: data[i][8] || "MAIN"
          });
        }
      }
    }
    
    // Get customers data
    let customersCount = 0;
    if (customersSheet) {
      const data = customersSheet.getDataRange().getValues();
      customersCount = data.length - 1; // Subtract header
    }
    
    const profit = totalSales - totalPurchases;
    
    return output({
      success: true,
      report: {
        stock: {
          items: stockCount,
          totalValue: totalStockValue,
          data: stockData
        },
        sales: {
          transactions: salesCount,
          totalAmount: totalSales,
          data: salesData
        },
        purchases: {
          transactions: purchasesCount,
          totalAmount: totalPurchases,
          data: purchasesData
        },
        customers: {
          count: customersCount
        },
        summary: {
          profit: profit,
          profitMargin: totalSales > 0 ? ((profit / totalSales) * 100).toFixed(2) + "%" : "0%",
          totalTransactions: salesCount + purchasesCount
        }
      }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


/* ===== GET DASHBOARD DATA ===== */
// Converts a "YYYY-MM-DD" calendar date (as picked by someone in this
// business's own local time - e.g. via an HTML date picker) into the
// real UTC timestamp for either the START (00:00:00.000) or END
// (23:59:59.999) of that date AS ACTUALLY EXPERIENCED in the business's
// own timezone (this Apps Script project's configured Time Zone, under
// Project Settings) - never assumed to be UTC. Without this, a business
// several hours away from UTC (e.g. Tanzania, UTC+3) would have its "day
// boundary" quietly shifted by that many hours: a period closed to cover
// "all of today" would in practice only genuinely cover until about 3am
// today, silently pushing everything sold after breakfast into what the
// system considers "the next period" - which is exactly the bug this
// fixes (Sales/All Sales not resetting to zero after a close).
function getBusinessDayBoundary(dateValue, endOfDay) {
  const timeZone = getCurrentBusinessTimeZone_();
  const dateOnly = Utilities.formatDate(new Date(dateValue), "UTC", "yyyy-MM-dd");
  const naiveUtc = new Date(dateOnly + "T00:00:00Z").getTime();
  
  // What local wall-clock time does that UTC instant show in this
  // business's timezone? The gap between that and "00:00:00 of the same
  // calendar date" reveals the timezone's real offset at this moment
  // (handles both fixed offsets and any daylight-saving shift correctly,
  // since it's derived live rather than hardcoded).
  const localWallClock = Utilities.formatDate(new Date(naiveUtc), timeZone, "yyyy-MM-dd'T'HH:mm:ss");
  const localWallClockAsUtc = new Date(localWallClock + "Z").getTime();
  const offsetMs = localWallClockAsUtc - naiveUtc;
  const localMidnightUtc = naiveUtc - offsetMs;
  
  return endOfDay ? (localMidnightUtc + 86399999) : localMidnightUtc;
}

// Finds the most recent CLOSED period's ACTUAL closing moment relevant
// to a given branch (that exact branch, or an "ALL Branches" close,
// which finalizes every branch together) - returns its real timestamp,
// or null if nothing has ever been closed for it.
//
// Uses "Closed At" (the exact real moment Close Accounts was clicked) -
// NOT "end of Period End's calendar day". If accounts are closed
// mid-day (e.g. 23:04), a sale made at 23:03 is rightly part of the
// closed period, while a sale made at 23:10 the SAME calendar day must
// immediately count toward the new open period - using "end of day" as
// the boundary would have silently stranded that 23:10 sale in limbo
// until midnight (excluded from the closed period's locked-in figures,
// but ALSO excluded from the new period's live figures). Anchoring to
// the real closing instant means the new period genuinely starts
// exactly where the old one ended, down to the second.
function getLatestClosedPeriodEnd(userDB, branchId) {
  const sheet = userDB.getSheetByName("ACCOUNTING_PERIODS");
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const normalizedBranch = branchId || "ALL";
  let latestEnd = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] !== "CLOSED") continue;
    const rowBranch = data[i][3] || "ALL";
    if (rowBranch !== normalizedBranch && rowBranch !== "ALL") continue;
    
    const endTime = new Date(data[i][5]).getTime(); // "Closed At" column
    if (isNaN(endTime)) continue;
    if (latestEnd === null || endTime > latestEnd) latestEnd = endTime;
  }
  return latestEnd;
}

// Exact closing boundary that applies to one branch record. Both a
// branch-specific close and an ALL-Branches close may apply; the latest
// applicable timestamp is the start of the new live period.
function getOperationalOpenPeriodStart_(userDB, recordBranchId) {
  const branch = recordBranchId || "MAIN";
  const branchClose = getLatestClosedPeriodEnd(userDB, branch);
  const allClose = getLatestClosedPeriodEnd(userDB, "ALL");
  if (branchClose === null && allClose === null) return null;
  if (branchClose === null) return allClose;
  if (allClose === null) return branchClose;
  return Math.max(branchClose, allClose);
}

function getDashboard(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    const salesSheet = userDB.getSheetByName("SALES");
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    
    const branchId = d.branchId || null; // For sellers
    const sellerId = d.sellerId || null;
    
    // The current OPEN accounting period's start - "All Sales" on the
    // Dashboard means sales since this period opened, NOT all-time sales,
    // so the figure genuinely goes back to zero right after a Close
    // Accounts (matching the explicit requirement that operational
    // stats reset on close while Stock/Debts carry forward). If nothing
    // has ever been closed for this branch, there's no boundary - "All
    // Sales" is simply every sale ever recorded, exactly as it always
    // used to behave.
    const currentPeriodStartTime = getLatestClosedPeriodEnd(userDB, branchId || "ALL");
    
    // Get stock data
    let stockCount = 0;
    let totalStockValue = 0;
    
    if (stockSheet) {
      const stockData = stockSheet.getDataRange().getValues();
      const products = new Set();
      
      for (let i = 1; i < stockData.length; i++) {
        if (stockData[i][0]) {
          const itemBranchId = stockData[i][8] || "MAIN";
          
          // Filter by branch if seller
          if (branchId && itemBranchId !== branchId) {
            continue;
          }
          
          products.add(stockData[i][1]);
          totalStockValue += (stockData[i][5] || 0);
        }
      }
      stockCount = products.size;
    }
    
    // Get sales data
    let salesCount = 0;
    let totalSales = 0;
    let todaySales = 0;
    const today = new Date();
    // "Today" as a calendar date in the BUSINESS's own timezone, not
    // UTC - otherwise, for a business ahead of UTC (like Tanzania,
    // UTC+3), the roughly 3 hours after local midnight would still be
    // counted as "yesterday" (since it's still yesterday in UTC),
    // silently excluding this morning's sales from "Today's Sales" and
    // showing yesterday's late-night UTC sales there instead.
    const dashboardTimeZone = getCurrentBusinessTimeZone_();
    const todayString = Utilities.formatDate(today, dashboardTimeZone, "yyyy-MM-dd");
    
    if (salesSheet) {
      const salesData = salesSheet.getDataRange().getValues();
      for (let i = 1; i < salesData.length; i++) {
        if (salesData[i][0]) {
          const recordSellerId = salesData[i][7];
          const recordBranchId = salesData[i][8] || "MAIN";
          
          // Filter by seller if user is seller
          if (sellerId && recordSellerId !== sellerId) {
            continue;
          }
          
          // Filter by branch if user is seller
          if (branchId && recordBranchId !== branchId) {
            continue;
          }
          
          const saleDate = new Date(salesData[i][0]);
          const saleTotal = salesData[i][4] || 0;
          
          // A sale genuinely belongs to the current open period only if
          // it happened AFTER the last Close Accounts moment - this same
          // condition gates both figures below, since "Today's Sales" is
          // really just "All Sales" narrowed to today's calendar date,
          // not a separate rule. Without this same check here too, a
          // sale made just before closing (e.g. 23:03, with the close
          // itself happening at 23:04) would keep showing up under
          // Today's Sales forever after being closed - exactly the bug
          // this fixes.
          const isInCurrentPeriod = (currentPeriodStartTime === null || saleDate.getTime() > currentPeriodStartTime);
          
          if (isInCurrentPeriod) {
            salesCount++;
            totalSales += saleTotal;
          }
          
          // Check if today's sale
          const saleDateString = Utilities.formatDate(saleDate, dashboardTimeZone, "yyyy-MM-dd");
          if (saleDateString === todayString && isInCurrentPeriod) {
            todaySales += saleTotal;
          }
        }
      }
    }
    
    // Tumia purchases kwa ajili ya taarifa tu, sio kwa profit
    let purchasesCount = 0;
    let totalPurchases = 0;
    
    if (purchasesSheet) {
      const purchasesData = purchasesSheet.getDataRange().getValues();
      for (let i = 1; i < purchasesData.length; i++) {
        if (purchasesData[i][0]) {
          const recordSellerId = purchasesData[i][7];
          const recordBranchId = purchasesData[i][8] || "MAIN";
          
          if (sellerId && recordSellerId !== sellerId) {
            continue;
          }
          
          if (branchId && recordBranchId !== branchId) {
            continue;
          }
          
          purchasesCount++;
          totalPurchases += (purchasesData[i][4] || 0);
        }
      }
    }
    
    // Profit itahesabiwa kwenye frontend kutoka stock
    // Hapa tunatuma data ghafi
    const lowStockItems = getLowStockItemsByBranch(d.databaseId, branchId);
    const recentActivity = getRecentActivityByBranch(d.databaseId, sellerId, branchId);
    
    console.log('Dashboard calculated - Today Sales:', todaySales, 'Total Sales:', totalSales);
    
    return output({
      success: true,
      dashboard: {
        quickStats: {
          stockItems: stockCount,
          stockValue: totalStockValue,
          todaySales: todaySales,
          totalSales: totalSales,
          totalPurchases: totalPurchases,
          lowStock: lowStockItems.length,
          salesCount: salesCount,
          purchasesCount: purchasesCount,
          currentPeriodStartTime: currentPeriodStartTime
        },
        recentActivity: recentActivity,
        lowStockItems: lowStockItems
      }
    });
    
  } catch (error) {
    return output({ success: false, error: "Failed to load dashboard: " + error.toString() });
  }
}


function getTotalStockValue(userDB) {
  try {
    const sheet = userDB.getSheetByName("STOCK");
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    let total = 0;
    
    for (let i = 1; i < data.length; i++) {
      total += (data[i][5] || 0);
    }
    
    return total;
  } catch (e) {
    return 0;
  }
}

function getSalesSummary(userDB) {
  try {
    const sheet = userDB.getSheetByName("SALES");
    if (!sheet) return { today: 0, month: 0, total: 0, count: 0 };
    
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    const todayStr = Utilities.formatDate(today, getCurrentBusinessTimeZone_(), "yyyy-MM-dd");
    const currentMonth = Number(Utilities.formatDate(today, getCurrentBusinessTimeZone_(), "MM")) - 1;
    const currentYear = Number(Utilities.formatDate(today, getCurrentBusinessTimeZone_(), "yyyy"));
    
    let todayTotal = 0;
    let monthTotal = 0;
    let total = 0;
    let count = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      const amount = data[i][4] || 0;
      total += amount;
      count++;
      
      const saleDate = new Date(data[i][0]);
      const saleDateStr = Utilities.formatDate(saleDate, getCurrentBusinessTimeZone_(), "yyyy-MM-dd");
      
      if (saleDateStr === todayStr) {
        todayTotal += amount;
      }
      
      const saleMonth = Number(Utilities.formatDate(saleDate, getCurrentBusinessTimeZone_(), "MM")) - 1;
      const saleYear = Number(Utilities.formatDate(saleDate, getCurrentBusinessTimeZone_(), "yyyy"));
      if (saleMonth === currentMonth && saleYear === currentYear) {
        monthTotal += amount;
      }
    }
    
    return {
      today: todayTotal,
      month: monthTotal,
      total: total,
      count: count
    };
  } catch (e) {
    return { today: 0, month: 0, total: 0, count: 0 };
  }
}

function getPurchasesSummary(userDB) {
  try {
    const sheet = userDB.getSheetByName("PURCHASES");
    if (!sheet) return { month: 0, total: 0, count: 0 };
    
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const currentMonth = Number(Utilities.formatDate(now, getCurrentBusinessTimeZone_(), "MM")) - 1;
    const currentYear = Number(Utilities.formatDate(now, getCurrentBusinessTimeZone_(), "yyyy"));
    
    let monthTotal = 0;
    let total = 0;
    let count = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      const amount = data[i][4] || 0;
      total += amount;
      count++;
      
      const purchaseDate = new Date(data[i][0]);
      const purchaseMonth = Number(Utilities.formatDate(purchaseDate, getCurrentBusinessTimeZone_(), "MM")) - 1;
      const purchaseYear = Number(Utilities.formatDate(purchaseDate, getCurrentBusinessTimeZone_(), "yyyy"));
      if (purchaseMonth === currentMonth && purchaseYear === currentYear) {
        monthTotal += amount;
      }
    }
    
    return {
      month: monthTotal,
      total: total,
      count: count
    };
  } catch (e) {
    return { month: 0, total: 0, count: 0 };
  }
}

function getCustomersCount(userDB) {
  try {
    const sheet = userDB.getSheetByName("CUSTOMERS");
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    return Math.max(0, data.length - 1); // Subtract header
  } catch (e) {
    return 0;
  }
}

/* ===== GET LOW STOCK ITEMS BY BRANCH ===== */
function getLowStockItemsByBranch(databaseId, branchId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) return [];
    
    const data = stockSheet.getDataRange().getValues();
    const products = {};
    const lowStock = [];
    
    let threshold = 10;
    const settingsSheet = userDB.getSheetByName("SETTINGS");
    if (settingsSheet) {
      const settingsData = settingsSheet.getDataRange().getValues();
      for (let i = 1; i < settingsData.length; i++) {
        if (settingsData[i][0] === "BUSINESS" && settingsData[i][1] === "low_stock_threshold") {
          threshold = parseInt(settingsData[i][2]) || 10;
          break;
        }
      }
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        const product = data[i][1];
        const itemBranchId = data[i][8] || "MAIN";
        const quantity = data[i][2] || 0;
        
        // Filter by branch if specified
        if (branchId && itemBranchId !== branchId) {
          continue;
        }
        
        if (!products[product]) {
          products[product] = 0;
        }
        products[product] += quantity;
      }
    }
    
    for (const [product, quantity] of Object.entries(products)) {
      if (quantity < threshold) {
        lowStock.push({
          product: product,
          quantity: quantity,
          threshold: threshold,
          branch: branchId || "All Branches"
        });
      }
    }
    
    return lowStock;
  } catch (error) {
    return [];
  }
}

/* ===== GET RECENT ACTIVITY BY BRANCH ===== */
function getRecentActivityByBranch(databaseId, sellerId, branchId) {
  const activities = [];
  const isAdmin = !sellerId;
  const limit = 10;
  
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const sellersSheet = userDB.getSheetByName("SELLERS");
    const branchesSheet = userDB.getSheetByName("BRANCHES");
    
    const sellerMap = {};
    if (sellersSheet) {
      const sellersData = sellersSheet.getDataRange().getValues();
      for (let i = 1; i < sellersData.length; i++) {
        if (sellersData[i][0]) {
          sellerMap[sellersData[i][0]] = sellersData[i][1];
        }
      }
    }
    
    const branchMap = {};
    if (branchesSheet) {
      const branchesData = branchesSheet.getDataRange().getValues();
      for (let i = 1; i < branchesData.length; i++) {
        if (branchesData[i][0]) {
          branchMap[branchesData[i][0]] = branchesData[i][1];
        }
      }
    }
    
    // Add sales activities
    const salesSheet = userDB.getSheetByName("SALES");
    if (salesSheet) {
      const salesData = salesSheet.getDataRange().getValues();
      for (let i = 1; i < salesData.length; i++) {
        if (salesData[i][0]) {
          const recordSellerId = salesData[i][7];
          const recordBranchId = salesData[i][8] || "MAIN";
          
          // Filter by seller
          if (sellerId && recordSellerId !== sellerId) {
            continue;
          }
          
          // Filter by branch
          if (branchId && recordBranchId !== branchId) {
            continue;
          }
          
          const sellerName = recordSellerId ? (sellerMap[recordSellerId] || "Unknown Seller") : "Admin";
          const branchName = branchMap[recordBranchId] || recordBranchId;
          const customer = salesData[i][5] || "Walk-in";
          const product = salesData[i][1];
          const quantity = salesData[i][2];
          
          let details;
          if (isAdmin) {
            if (recordSellerId) {
              details = `${sellerName} sold ${quantity} of ${product} to ${customer} at ${branchName}`;
            } else {
              details = `You sold ${quantity} of ${product} to ${customer} at ${branchName}`;
            }
          } else {
            details = `You sold ${quantity} of ${product} to ${customer} at ${branchName}`;
          }
          
          activities.push({
            timestamp: salesData[i][0],
            action: "SALE",
            details: details,
            sellerId: recordSellerId,
            branchId: recordBranchId
          });
        }
      }
    }
    
    // Add purchases activities
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    if (purchasesSheet) {
      const purchasesData = purchasesSheet.getDataRange().getValues();
      for (let i = 1; i < purchasesData.length; i++) {
        if (purchasesData[i][0]) {
          const recordSellerId = purchasesData[i][7];
          const recordBranchId = purchasesData[i][8] || "MAIN";
          
          if (sellerId && recordSellerId !== sellerId) {
            continue;
          }
          
          if (branchId && recordBranchId !== branchId) {
            continue;
          }
          
          const sellerName = recordSellerId ? (sellerMap[recordSellerId] || "Unknown Seller") : "Admin";
          const branchName = branchMap[recordBranchId] || recordBranchId;
          const supplier = purchasesData[i][5] || "Local Supplier";
          const product = purchasesData[i][1];
          const quantity = purchasesData[i][2];
          
          let details;
          if (isAdmin) {
            if (recordSellerId) {
              details = `${sellerName} purchased ${quantity} of ${product} from ${supplier} at ${branchName}`;
            } else {
              details = `You purchased ${quantity} of ${product} from ${supplier} at ${branchName}`;
            }
          } else {
            details = `You purchased ${quantity} of ${product} from ${supplier} at ${branchName}`;
          }
          
          activities.push({
            timestamp: purchasesData[i][0],
            action: "PURCHASE",
            details: details,
            sellerId: recordSellerId,
            branchId: recordBranchId
          });
        }
      }
    }
    
    // Add stock activities
    const stockSheet = userDB.getSheetByName("STOCK");
    if (stockSheet) {
      const stockData = stockSheet.getDataRange().getValues();
      for (let i = 1; i < stockData.length; i++) {
        if (stockData[i][0]) {
          const recordSellerId = stockData[i][7];
          const recordBranchId = stockData[i][8] || "MAIN";
          
          if (sellerId && recordSellerId !== sellerId) {
            continue;
          }
          
          if (branchId && recordBranchId !== branchId) {
            continue;
          }
          
          const sellerName = recordSellerId ? (sellerMap[recordSellerId] || "Unknown Seller") : "Admin";
          const branchName = branchMap[recordBranchId] || recordBranchId;
          const product = stockData[i][1];
          const quantity = stockData[i][2];
          
          let details;
          if (isAdmin) {
            if (recordSellerId) {
              details = `${sellerName} added ${quantity} of ${product} at ${branchName}`;
            } else {
              details = `You added ${quantity} of ${product} at ${branchName}`;
            }
          } else {
            details = `You added ${quantity} of ${product} at ${branchName}`;
          }
          
          activities.push({
            timestamp: stockData[i][0],
            action: "STOCK ADD",
            details: details,
            sellerId: recordSellerId,
            branchId: recordBranchId
          });
        }
      }
    }
    
    activities.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return activities.slice(0, limit);
    
  } catch (error) {
    return [];
  }
}

/* ===== CREATE PASSWORD EMAIL TEMPLATE ===== */
function createPasswordEmailTemplate(business, password) {
  return `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
      <h2 style="color: #0f172a; text-align: center;">SmartTrade Pro</h2>
      <p>Hello,</p>
      <p>Your password for <strong>${business}</strong> account is:</p>
      <div style="background: #f1f5f9; padding: 20px; border-radius: 8px; margin: 20px 0; text-align: center;">
        <h3 style="margin: 0; color: #0f172a;">${password}</h3>
      </div>
      <p>For security reasons, we recommend changing your password after logging in.</p>
      <p>If you didn't request this, please ignore this email.</p>
      <br>
      <p>Best regards,<br><strong>SmartTrade Pro Team</strong></p>
      <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;">
      <p style="font-size: 12px; color: #666; text-align: center;">
        This is an automated message. Please do not reply to this email.
      </p>
    </div>
  `;
}

/* ===== GET CONSOLIDATED STOCK ===== */
function getConsolidatedStock(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    const data = stockSheet.getDataRange().getValues();
    const productMap = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        const product = data[i][1];
        const quantity = data[i][2] || 0;
        const buyPrice = data[i][3] || 0;
        const sellPrice = data[i][4] || 0;
        const category = data[i][6] || "";
        const branchId = data[i][8] || "MAIN";
        
        const key = product + "|" + branchId;
        
        if (!productMap[key]) {
          productMap[key] = {
            product: product,
            branch: branchId,
            totalQuantity: 0,
            totalBuyValue: 0,
            avgBuyPrice: 0,
            latestSellPrice: sellPrice,
            categories: new Set(),
            entries: []
          };
        }
        
        productMap[key].totalQuantity += quantity;
        productMap[key].totalBuyValue += (quantity * buyPrice);
        if (category) productMap[key].categories.add(category);
        
        productMap[key].entries.push({
          date: data[i][0],
          sellPrice: sellPrice,
          buyPrice: buyPrice
        });
      }
    }
    
    const consolidated = [];
    
    for (const [key, info] of Object.entries(productMap)) {
      if (info.totalQuantity > 0) {
        info.avgBuyPrice = info.totalBuyValue / info.totalQuantity;
      }
      
      if (info.entries.length > 0) {
        info.entries.sort((a, b) => new Date(b.date) - new Date(a.date));
        info.latestSellPrice = info.entries[0].sellPrice;
      }
      
      consolidated.push({
        product: info.product,
        branch: info.branch,
        quantity: info.totalQuantity,
        avgBuyPrice: info.avgBuyPrice,
        latestSellPrice: info.latestSellPrice,
        totalValue: info.totalBuyValue,
        categories: Array.from(info.categories),
        lastUpdated: info.entries.length > 0 ? info.entries[0].date : null
      });
    }
    
    consolidated.sort((a, b) => a.product.localeCompare(b.product));
    
    return output({
      success: true,
      data: consolidated,
      summary: {
        totalProducts: consolidated.length,
        totalItems: consolidated.reduce((sum, item) => sum + item.quantity, 0),
        totalValue: consolidated.reduce((sum, item) => sum + item.totalValue, 0)
      }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== CHANGE PASSWORD ===== */
function changePassword(d) {
  try {
    if (!d.email || !d.currentPassword || !d.newPassword) {
      return output({ success: false, error: "All fields are required" });
    }
    
    if (d.newPassword.length < 6) {
      return output({ success: false, error: "New password must be at least 6 characters" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    
    const data = masterSheet.getDataRange().getValues();
    
    let foundAny = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email) {
        if (!foundAny && data[i][2] !== d.currentPassword) {
          return output({ success: false, error: "Current password is incorrect" });
        }
        // One account can own several businesses that all share the same
        // login password - update EVERY row for this email, not just the
        // first, so they never fall out of sync with each other.
        masterSheet.getRange(i + 1, 3).setValue(d.newPassword);
        foundAny = true;
      }
    }
    
    if (!foundAny) {
      return output({ success: false, error: "User not found" });
    }
    
    return output({ success: true, message: "Password changed successfully" });
    
  } catch (error) {
    return output({ success: false, error: "Failed to change password: " + error.toString() });
  }
}


/* ===== DELETE ACCOUNT USING DRIVE API ONLY ===== */
function deleteAccount(d) {
  try {
    if (!d.email || !d.confirmBusiness) {
      return output({ success: false, error: "All fields are required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    
    const data = masterSheet.getDataRange().getValues();
    let rowToDelete = -1;
    let databaseId = null;
    let logoId = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.email && data[i][0] === d.confirmBusiness) {
        rowToDelete = i + 1;
        databaseId = data[i][3];
        logoId = data[i][6];
        break;
      }
    }
    
    if (rowToDelete === -1) {
      // Email exists but no business under it matches the confirmation
      // name typed - or the email itself was never found. Either way,
      // scanning every row (instead of stopping at the first email match)
      // means this correctly finds a SPECIFIC business even when the
      // account owns several.
      let emailExists = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === d.email) { emailExists = true; break; }
      }
      return output({ success: false, error: emailExists ? "Business name does not match" : "User not found" });
    }
    
    // Delete user's spreadsheet using Drive API
    if (databaseId) {
      try {
        Drive.Files.remove(databaseId);
      } catch (e) {
        console.log("Could not delete user spreadsheet:", e);
      }
    }
    
    // Delete user's logo using Drive API
    if (logoId) {
      try {
        Drive.Files.remove(logoId);
      } catch (e) {
        console.log("Could not delete user logo:", e);
      }
    }
    
    // Delete user from master sheet
    masterSheet.deleteRow(rowToDelete);
    
    return output({ success: true, message: "Account deleted successfully" });
    
  } catch (error) {
    return output({ success: false, error: "Failed to delete account: " + error.toString() });
  }
}


/* ===== UPDATE SALE ===== */
function updateSale(d) {
  // Stock reconciliation for sale edits is deliberately performed here on
  // the server, under one script lock. The client is never trusted for the
  // old quantity or old product because a stale/offline client can otherwise
  // create negative stock or silently accept an impossible edit.
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!d.databaseId || !d.rowIndex) {
      return output({ success: false, error: "Database ID and row index required" });
    }

    const userDB = SpreadsheetApp.openById(d.databaseId);
    const salesSheet = userDB.getSheetByName("SALES");
    if (!salesSheet) return output({ success: false, error: "Sales sheet not found" });

    const row = Number(d.rowIndex);
    const lastRow = salesSheet.getLastRow();
    if (!Number.isInteger(row) || row < 2 || row > lastRow) {
      return output({ success: false, error: "Invalid row index" });
    }

    const existing = salesSheet.getRange(row, 1, 1, 10).getValues()[0];
    const existingDate = existing[0];
    const oldProduct = String(existing[1] || "").trim();
    const oldQuantity = Number(existing[2] || 0);
    const oldBranchId = String(existing[8] || "MAIN");
    const oldUnit = String(existing[9] || "");
    const existingSellerId = existing[7] || null;

    const newProduct = String(d.product || "").trim();
    const newQuantity = Number(d.quantity);
    const newUnit = String(d.unit || "");
    const newBranchId = String(d.newBranchId || oldBranchId);
    const newUnitPrice = Number(d.unitPrice);
    const newTotal = Number(d.total);

    if (!newProduct) return output({ success: false, error: "Product is required" });
    if (!Number.isFinite(newQuantity) || newQuantity <= 0) {
      return output({ success: false, error: "Quantity must be greater than zero" });
    }
    if (!Number.isFinite(newUnitPrice) || newUnitPrice < 0) {
      return output({ success: false, error: "Unit price is invalid" });
    }
    if (!Number.isFinite(newTotal) || newTotal < 0) {
      return output({ success: false, error: "Total amount is invalid" });
    }
    if (!Number.isFinite(oldQuantity) || oldQuantity < 0) {
      return output({ success: false, error: "Existing sale quantity is invalid" });
    }

    const stockSheet = userDB.getSheetByName("STOCK");
    const oldKey = makeStockKey_(oldProduct, oldBranchId, oldUnit);
    const newKey = makeStockKey_(newProduct, newBranchId, newUnit);

    // Stock validation uses the CURRENT available stock in STOCK.
    // For the same product, only an increase over the original sold quantity
    // consumes additional stock. A decrease returns stock.
    const stockRows = stockSheet ? getStockRowsForEdit_(stockSheet) : [];
    const oldRows = stockRows.filter(r => r.key === oldKey);
    const newRows = stockRows.filter(r => r.key === newKey);
    const oldIsStock = oldRows.length > 0;
    const newIsStock = newRows.length > 0;

    const currentNewStock = newRows.reduce((sum, r) => sum + Math.max(0, r.quantity), 0);
    const sameStockItem = oldIsStock && oldKey === newKey;
    const additionalNeeded = newIsStock ? Math.max(0, newQuantity - (sameStockItem ? oldQuantity : 0)) : 0;

    if (newIsStock && additionalNeeded > currentNewStock + 1e-9) {
      return output({
        success: false,
        error: 'Insufficient stock. Available Stock: ' + formatQuantityForError_(currentNewStock) +
               '. Additional required: ' + formatQuantityForError_(additionalNeeded) + '.'
      });
    }

    // Snapshot the exact cells that will change so a later failure can be
    // rolled back. This is important because Sheets has no SQL-style
    // transaction spanning multiple ranges.
    const stockSnapshot = {};
    const rowsToSnapshot = {};
    oldRows.concat(newRows).forEach(r => { rowsToSnapshot[r.row] = true; });
    Object.keys(rowsToSnapshot).forEach(rowNo => {
      const rr = Number(rowNo);
      stockSnapshot[rr] = stockSheet.getRange(rr, 3, 1, 4).getValues()[0];
    });

    const oldSaleValues = salesSheet.getRange(row, 2, 1, 9).getValues()[0];

    try {
      // Apply the stock delta exactly once. If old/new are the same stock
      // key, this becomes +oldQty - newQty. If the keys differ, the old
      // stock is fully returned and the new stock is fully consumed.
      if (oldIsStock && newIsStock && oldKey === newKey) {
        const delta = oldQuantity - newQuantity;
        if (Math.abs(delta) > 1e-12) {
          applyStockDelta_(stockSheet, newRows, delta);
        }
      } else {
        if (oldIsStock && oldQuantity > 0) {
          applyStockDelta_(stockSheet, oldRows, oldQuantity);
        }
        if (newIsStock && newQuantity > 0) {
          // Re-read the new stock after returning the old quantity when the
          // keys differ, then deduct the requested quantity.
          const fresh = getStockRowsForEdit_(stockSheet).filter(r => r.key === newKey);
          const freshAvailable = fresh.reduce((sum, r) => sum + r.quantity, 0);
          if (newQuantity > freshAvailable + 1e-9) {
            throw new Error("Insufficient stock after stock reconciliation. Available: " + formatQuantityForError_(freshAvailable));
          }
          applyStockDelta_(stockSheet, fresh, -newQuantity);
        }
      }

      // Only after stock validation/reconciliation succeeds do we write the
      // edited sale row.
      salesSheet.getRange(row, 2, 1, 5).setValues([[
        newProduct,
        newQuantity,
        newUnitPrice,
        newTotal,
        d.customer || "Walk-in"
      ]]);
      salesSheet.getRange(row, 7).setValue(d.paymentMethod || "cash");
      if (d.unit !== undefined) salesSheet.getRange(row, 10).setValue(newUnit);
      if (d.newBranchId && d.newBranchId !== oldBranchId) salesSheet.getRange(row, 9).setValue(newBranchId);

    } catch (writeError) {
      // Roll back every stock cell touched by this edit, then restore the
      // original sale fields. This prevents a failed edit from leaving the
      // sale and stock out of sync.
      Object.keys(stockSnapshot).forEach(rowNo => {
        stockSheet.getRange(Number(rowNo), 3, 1, 4).setValues([stockSnapshot[rowNo]]);
      });
      salesSheet.getRange(row, 2, 1, 9).setValues([oldSaleValues]);
      throw writeError;
    }

    return output({
      success: true,
      message: "Sale updated successfully",
      timestamp: existingDate instanceof Date ? existingDate.toISOString() : String(existingDate),
      branchId: newBranchId,
      sellerId: existingSellerId
    });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

function makeStockKey_(product, branchId, unit) {
  return String(product || "").trim().toLowerCase() + "||" + String(branchId || "MAIN") + "||" + String(unit || "").trim().toLowerCase();
}

function getStockRowsForEdit_(stockSheet) {
  const lastRow = stockSheet.getLastRow();
  if (lastRow < 2) return [];
  const data = stockSheet.getRange(1, 1, lastRow, Math.max(10, stockSheet.getLastColumn())).getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    rows.push({
      row: i + 1,
      product: String(data[i][1] || "").trim(),
      quantity: Number(data[i][2] || 0),
      buyPrice: Number(data[i][3] || 0),
      branchId: String(data[i][8] || "MAIN"),
      unit: String(data[i][9] || ""),
      key: makeStockKey_(data[i][1], data[i][8] || "MAIN", data[i][9] || "")
    });
  }
  return rows;
}

function applyStockDelta_(stockSheet, rows, delta) {
  if (!rows || !rows.length || !Number.isFinite(delta)) {
    if (Math.abs(delta) < 1e-12) return;
    throw new Error("Stock record not found");
  }

  let remaining = Math.abs(delta);
  const direction = delta > 0 ? 1 : -1;

  // Return stock to the first matching stock row. For a deduction, use FIFO
  // (oldest rows first) so the existing stock valuation convention remains.
  const ordered = rows.slice().sort((a, b) => a.row - b.row);

  if (direction > 0) {
    const target = ordered[0];
    const newQty = target.quantity + remaining;
    stockSheet.getRange(target.row, 3).setValue(newQty);
    stockSheet.getRange(target.row, 6).setValue(newQty * target.buyPrice);
    return;
  }

  const totalAvailable = ordered.reduce((sum, r) => sum + Math.max(0, r.quantity), 0);
  if (remaining > totalAvailable + 1e-9) {
    throw new Error("Insufficient stock. Available: " + formatQuantityForError_(totalAvailable) + ". Requested: " + formatQuantityForError_(remaining) + ".");
  }

  for (let i = 0; i < ordered.length && remaining > 1e-12; i++) {
    const r = ordered[i];
    const take = Math.min(Math.max(0, r.quantity), remaining);
    if (take <= 0) continue;
    const newQty = r.quantity - take;
    stockSheet.getRange(r.row, 3).setValue(newQty);
    stockSheet.getRange(r.row, 6).setValue(newQty * r.buyPrice);
    remaining -= take;
  }
}

function formatQuantityForError_(value) {
  const n = Number(value || 0);
  return Number.isInteger(n) ? String(n) : String(Number(n.toFixed(6)));
}

/* ===== UPDATE PURCHASE ===== */
function updatePurchase(d) {
  try {
    if (!d.databaseId || !d.rowIndex) {
      return output({ success: false, error: "Database ID and row index required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    
    if (!purchasesSheet) {
      return output({ success: false, error: "Purchases sheet not found" });
    }
    
    const row = d.rowIndex;
    const lastRow = purchasesSheet.getLastRow();
    
    if (row < 2 || row > lastRow) {
      return output({ success: false, error: "Invalid row index" });
    }
    
    // Get existing data to preserve timestamp, branch and seller
    const existingData = purchasesSheet.getRange(row, 1, 1, 9).getValues()[0];
    const existingDate = existingData[0];           // Keep original timestamp
    const existingBranchId = existingData[8] || "MAIN";
    const existingSellerId = existingData[7] || null;
    
    // Update ONLY the editable fields (keep original date in column 1)
    purchasesSheet.getRange(row, 2).setValue(d.product);        // Product
    purchasesSheet.getRange(row, 3).setValue(d.quantity);       // Quantity
    purchasesSheet.getRange(row, 4).setValue(d.unitCost);       // Unit Cost
    purchasesSheet.getRange(row, 5).setValue(d.total);          // Total
    purchasesSheet.getRange(row, 6).setValue(d.supplier || "Local Supplier"); // Supplier
    purchasesSheet.getRange(row, 7).setValue(d.paymentMethod);  // Payment Method
    if (d.category !== undefined) {
      if (purchasesSheet.getLastColumn() < 10) {
        purchasesSheet.getRange(1, 10).setValue("Category");
      }
      purchasesSheet.getRange(row, 10).setValue(d.category || ""); // Category
    }
    if (d.unit !== undefined) {
      if (purchasesSheet.getLastColumn() < 11) {
        purchasesSheet.getRange(1, 11).setValue("Unit");
      }
      purchasesSheet.getRange(row, 11).setValue(d.unit || ""); // Unit
    }
    
    // DO NOT change these:
    // Column 1: Date - stays the same
    // Column 8: Seller ID - stays the same
    // Column 9: Branch ID - stays the same
    
    return output({
      success: true,
      message: "Purchase updated successfully",
      timestamp: existingDate.toISOString(),
      branchId: existingBranchId,
      sellerId: existingSellerId
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== DELETE STOCK ROW ===== */
function deleteStockRow(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    if (!d.timestamp || !d.product || !d.branchId) {
      return output({ success: false, error: "Timestamp, product and branch ID required to identify the row" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!stockSheet) {
      return output({ success: false, error: "Stock sheet not found" });
    }
    
    // Tafuta row sahihi kwa kutumia timestamp, product, na branch
    const data = stockSheet.getDataRange().getValues();
    let targetRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      const rowTimestamp = data[i][0];
      const rowProduct = data[i][1];
      const rowBranchId = data[i][8] || "MAIN";
      
      const rowTime = new Date(rowTimestamp).getTime();
      const targetTime = new Date(d.timestamp).getTime();
      const timeDiff = Math.abs(rowTime - targetTime);
      
      if (timeDiff < 1000 && rowProduct === d.product && rowBranchId === d.branchId) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) {
      return output({ 
        success: false, 
        error: `Could not find the exact stock entry to delete. Try refreshing the page.` 
      });
    }
    
    // Futa row
    stockSheet.deleteRow(targetRow);
    
    return output({
      success: true,
      message: "Stock entry deleted successfully"
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== DELETE SALE ROW ===== */
function deleteSaleRow(d) {
  try {
    if (!d.databaseId || !d.rowIndex) {
      return output({ success: false, error: "Database ID and row index required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const salesSheet = userDB.getSheetByName("SALES");
    
    if (!salesSheet) {
      return output({ success: false, error: "Sales sheet not found" });
    }
    
    const lastRow = salesSheet.getLastRow();
    let row = Number(d.rowIndex) || 0;

    // Prefer the immutable sale snapshot when supplied. This prevents a stale
    // rowIndex from deleting the wrong sale after other rows have shifted.
    // rowIndex remains the compatibility fallback for older clients.
    const target = d.sale && typeof d.sale === 'object' ? d.sale : null;
    if (target) {
      // The sale's own immutable operationId is the strongest identity key.
      // Use it before the descriptive snapshot because date/formatting can
      // differ slightly between browser and Sheets. This also prevents a
      // shifted rowIndex from ever deleting the wrong sale.
      const saleOperationId = String(target.operationId || '').trim();
      if (saleOperationId && salesSheet.getLastColumn() >= 12 && lastRow >= 2) {
        const opValues = salesSheet.getRange(2, 12, lastRow - 1, 1).getValues();
        for (let i = 0; i < opValues.length; i++) {
          if (String(opValues[i][0] || '').trim() === saleOperationId) {
            row = i + 2;
            break;
          }
        }
      }

      // If operationId did not identify the row, fall back to the exact
      // snapshot for compatibility with older sales that predate operationId.
      if (row === Number(d.rowIndex) || !saleOperationId) {
        const rowCount = Math.max(0, lastRow - 1);
        const rows = rowCount ? salesSheet.getRange(2, 1, rowCount, 10).getValues() : [];
        const targetTime = target.date ? new Date(target.date).getTime() : NaN;
        let matchedRow = -1;
        for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const rowTime = r[0] ? new Date(r[0]).getTime() : NaN;
        const sameTime = Number.isNaN(targetTime) || Number.isNaN(rowTime)
          ? String(r[0]) === String(target.date || '')
          : Math.abs(rowTime - targetTime) <= 2000;
        const sameProduct = String(r[1] ?? '') === String(target.product || target.service || '');
        const sameQuantity = Number(r[2] ?? 0) === Number(target.quantity ?? 0);
        const sameUnitPrice = Number(r[3] ?? 0) === Number(target.unitPrice ?? 0);
        const sameTotal = Number(r[4] ?? 0) === Number(target.total ?? 0);
        const sameCustomer = String(r[5] ?? '') === String(target.customer || '');
        const samePayment = String(r[6] ?? '') === String(target.paymentMethod || 'cash');
        const sameSeller = String(r[7] ?? '') === String(target.sellerId || '');
        const sameBranch = String(r[8] || 'MAIN') === String(target.branchId || 'MAIN');
        const sameUnit = String(r[9] ?? '') === String(target.unit || '');
          if (sameTime && sameProduct && sameQuantity && sameUnitPrice && sameTotal && sameCustomer && samePayment && sameSeller && sameBranch && sameUnit) {
            matchedRow = i + 2;
            break;
          }
        }
        if (matchedRow !== -1) row = matchedRow;
        else if (!saleOperationId || row === Number(d.rowIndex)) {
          return output({ success: false, error: "Could not find the exact sale entry to delete." });
        }
      }
    }

    if (row < 2 || row > lastRow) {
      return output({ success: false, error: "Invalid row index" });
    }
    
    // Read the row BEFORE deleting it, so we know exactly what to give
    // back to Stock. A Non-Stock service simply won't be found in the
    // STOCK sheet by returnStockQuantity() below, so this is a harmless
    // no-op for those - no need to check isStockProduct separately here.
    const rowData = salesSheet.getRange(row, 1, 1, 10).getValues()[0];
    const product = rowData[1];
    const quantity = rowData[2] || 0;
    const branchId = rowData[8] || "MAIN";
    const unit = rowData[9] || "";
    
    salesSheet.deleteRow(row);
    
    if (product && quantity) {
      returnStockQuantity(userDB, product, quantity, branchId, unit);
    }
    
    return output({
      success: true,
      message: "Sale entry deleted successfully"
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== DELETE PURCHASE ROW ===== */
function deletePurchaseRow(d) {
  try {
    if (!d.databaseId || !d.rowIndex) {
      return output({ success: false, error: "Database ID and row index required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const purchasesSheet = userDB.getSheetByName("PURCHASES");
    
    if (!purchasesSheet) {
      return output({ success: false, error: "Purchases sheet not found" });
    }
    
    const row = d.rowIndex;
    const lastRow = purchasesSheet.getLastRow();
    
    if (row < 2 || row > lastRow) {
      return output({ success: false, error: "Invalid row index" });
    }
    
    purchasesSheet.deleteRow(row);
    
    return output({
      success: true,
      message: "Purchase entry deleted successfully"
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


/* ===== RETURN DEBT (MARK AS PAID) ===== */
function returnDebt(d) {
  try {
    if (!d.databaseId || !d.debtId) {
      return output({ success: false, error: "Database ID and Debt ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    
    if (!debtsSheet) {
      return output({ success: false, error: "Debts sheet not found" });
    }
    
    const data = debtsSheet.getDataRange().getValues();
    const timestamp = new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.debtId) {
        const row = i + 1;
        const currentStatus = data[i][11];
        const product = data[i][6];
        const quantity = data[i][7];
        const branchId = data[i][12] || "MAIN";
        const oldPaidAmount = data[i][10] || 0;
        const oldStatus = currentStatus;
        
        if (currentStatus === DEBT_STATUS_PAID) {
          return output({ success: false, error: "This debt is already marked as paid" });
        }
        
        const totalAmount = data[i][9];
        const paidAmount = d.paidAmount || totalAmount;
        const newPaidAmount = oldPaidAmount + paidAmount;
        const newStatus = newPaidAmount >= totalAmount ? DEBT_STATUS_PAID : DEBT_STATUS_PARTIAL;
        
        debtsSheet.getRange(row, 11).setValue(newPaidAmount);
        debtsSheet.getRange(row, 12).setValue(newStatus);
        debtsSheet.getRange(row, 16).setValue(timestamp);
        
        // Add history entry
        addDebtHistory(d.databaseId, d.debtId, "PAYMENT", paidAmount, oldStatus, newStatus,
                       oldPaidAmount, newPaidAmount, "Payment received", d.email);
        
        return output({
          success: true,
          message: newStatus === DEBT_STATUS_PAID ? "Debt marked as paid" : `Partial payment of ${paidAmount} recorded`,
          status: newStatus,
          paidAmount: newPaidAmount,
          timestamp: timestamp.toISOString()
        });
      }
    }
    
    return output({ success: false, error: "Debt not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


/* ===== DENY DEBT ===== */
function denyDebt(d) {
  try {
    if (!d.databaseId || !d.debtId) {
      return output({ success: false, error: "Database ID and Debt ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    
    if (!debtsSheet) {
      return output({ success: false, error: "Debts sheet not found" });
    }
    
    const data = debtsSheet.getDataRange().getValues();
    const timestamp = new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.debtId) {
        const row = i + 1;
        const currentStatus = data[i][11];
        const oldPaidAmount = data[i][10] || 0;
        const oldStatus = currentStatus;
        
        if (currentStatus !== DEBT_STATUS_PENDING && currentStatus !== DEBT_STATUS_PARTIAL) {
          return output({ success: false, error: "Only pending or partial debts can be denied" });
        }
        
        debtsSheet.getRange(row, 12).setValue(DEBT_STATUS_DENIED);
        debtsSheet.getRange(row, 16).setValue(timestamp);
        
        let denialNote = d.reason || "Denied by administrator";
        const currentNotes = data[i][14] || "";
        const newNotes = currentNotes + (currentNotes ? "\n" : "") + "[DENIED: " + denialNote + " on " + timestamp.toLocaleString() + "]";
        debtsSheet.getRange(row, 15).setValue(newNotes);
        
        // Add history entry
        addDebtHistory(d.databaseId, d.debtId, "DENY", 0, oldStatus, DEBT_STATUS_DENIED,
                       oldPaidAmount, oldPaidAmount, denialNote, d.email);
        
        return output({
          success: true,
          message: "Debt denied successfully. It will now appear to the seller only.",
          status: DEBT_STATUS_DENIED,
          timestamp: timestamp.toISOString()
        });
      }
    }
    
    return output({ success: false, error: "Debt not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== DELETE DEBT ===== */
function deleteDebt(d) {
  try {
    if (!d.databaseId || !d.debtId) {
      return output({ success: false, error: "Database ID and Debt ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!debtsSheet) {
      return output({ success: false, error: "Debts sheet not found" });
    }
    
    const data = debtsSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.debtId) {
        const row = i + 1;
        const status = data[i][11];
        const product = data[i][6];
        const quantity = data[i][7];
        const branchId = data[i][12] || "MAIN";
        const unit = data[i][16] || "";
        const oldPaidAmount = data[i][10] || 0;
        const oldStatus = status;
        
        // Check if this is a stock product by looking at STOCK sheet
        let isStockProduct = false;
        if (stockSheet) {
          const stockData = stockSheet.getDataRange().getValues();
          for (let j = 1; j < stockData.length; j++) {
            if (stockData[j][1] === product) {
              isStockProduct = true;
              break;
            }
          }
        }
        
        // Add history entry before deletion
        addDebtHistory(d.databaseId, d.debtId, "DELETE", 0, oldStatus, "DELETED",
                       oldPaidAmount, 0, "Debt deleted" + (d.returnStock && isStockProduct ? " with stock return" : ""), d.email);
        
        // Only return stock if:
        // 1. It's a stock product
        // 2. Status is not paid
        // 3. returnStock flag is true
        // Uses returnStockQuantity() - matches by Product+Branch+UNIT (a
        // product can have more than one unit in stock) and only ever
        // ADDS the quantity onto an existing matching row. It deliberately
        // never creates a new Stock row: if no matching row exists (the
        // item was removed from Stock entirely since this debt was
        // recorded), there's nothing safe to merge into, so the quantity
        // is simply not fabricated back into a brand-new row.
        if (isStockProduct && status !== DEBT_STATUS_PAID && d.returnStock === true) {
          returnStockQuantity(userDB, product, quantity, branchId, unit);
        }
        
        // Delete the debt record
        debtsSheet.deleteRow(row);
        
        return output({
          success: true,
          message: "Debt deleted successfully" + (isStockProduct && status !== DEBT_STATUS_PAID && d.returnStock ? " and stock returned" : "")
        });
      }
    }
    
    return output({ success: false, error: "Debt not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

/* ===== GET DEBTS ===== */
function getDebts(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    
    if (!debtsSheet) {
      return output({ success: true, data: [], count: 0 });
    }
    ensureDebtPhotoColumns_(debtsSheet);
    
    const data = debtsSheet.getDataRange().getValues();
    const debts = [];
    const sellerId = d.sellerId || null;
    const branchId = d.branchId || null;
    const isAdmin = !sellerId;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const debtSellerId = data[i][13] || null;
        const debtBranchId = data[i][12] || "MAIN";
        const debtStatus = data[i][11] || "pending";
        const photoId = data[i][5] || "";
        const photoBase64 = data[i][17] || "";
        const photoMimeType = data[i][18] || "image/jpeg";
        
        // Filter by seller if not admin
        if (!isAdmin && sellerId && debtSellerId !== sellerId) {
          continue;
        }
        
        // Filter by branch if specified
        if (branchId && debtBranchId !== branchId) {
          continue;
        }
        
        // For admin: hide denied debts
        if (isAdmin && debtStatus === "denied") {
          continue;
        }
        
        // Generate ALL photo URLs
        let thumbnailUrl = "";
        let directUrl = "";
        let lh3Url = "";
        let viewUrl = "";
        
        if (photoId) {
          thumbnailUrl = "https://drive.google.com/thumbnail?id=" + photoId + "&sz=s200";
          directUrl = "https://drive.google.com/uc?export=view&id=" + photoId;
          lh3Url = "https://lh3.googleusercontent.com/d/" + photoId + "=s200";
          viewUrl = "https://drive.google.com/file/d/" + photoId + "/view";
        }
        
        debts.push({
          id: data[i][0],
          date: data[i][1],
          borrowerName: data[i][2],
          borrowerPhone: data[i][3] || "",
          borrowerNIDA: data[i][4] || "",
          photoId: photoId,
          photoBase64: photoBase64,
          photoMimeType: photoMimeType,
          thumbnailUrl: thumbnailUrl,
          directUrl: directUrl,
          lh3Url: lh3Url,
          viewUrl: viewUrl,
          product: data[i][6],
          quantity: data[i][7],
          unitPrice: data[i][8],
          totalAmount: data[i][9],
          paidAmount: data[i][10] || 0,
          status: debtStatus,
          branchId: debtBranchId,
          sellerId: debtSellerId,
          notes: data[i][14] || "",
          lastUpdated: data[i][15] || data[i][1],
          unit: data[i][16] || ""
        });
      }
    }
    
    return output({
      success: true,
      data: debts,
      count: debts.length
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


/* ===== GET DEBTS DATA FOR SYNC (with all photo URLs) ===== */
function getDebtsDataForSync(databaseId, sellerId, branchId) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    
    if (!debtsSheet) return [];
    ensureDebtPhotoColumns_(debtsSheet);
    
    const data = debtsSheet.getDataRange().getValues();
    const debts = [];
    const isAdmin = !sellerId;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const debtSellerId = data[i][13] || null;
        const debtBranchId = data[i][12] || "MAIN";
        const debtStatus = data[i][11] || "pending";
        const photoId = data[i][5] || "";
        const photoBase64 = data[i][17] || "";
        const photoMimeType = data[i][18] || "image/jpeg";
        
        // For admin: exclude denied debts
        if (isAdmin && debtStatus === "denied") {
          continue;
        }
        
        // For seller: filter by sellerId
        if (!isAdmin && sellerId && debtSellerId !== sellerId) {
          continue;
        }
        
        // Filter by branch if specified
        if (branchId && debtBranchId !== branchId) {
          continue;
        }
        
        // Generate ALL photo URLs - MUHIMU KWA OFFLINE
        let thumbnailUrl = "";
        let directUrl = "";
        let lh3Url = "";
        let viewUrl = "";
        
        if (photoId) {
          thumbnailUrl = "https://drive.google.com/thumbnail?id=" + photoId + "&sz=s200";
          directUrl = "https://drive.google.com/uc?export=view&id=" + photoId;
          lh3Url = "https://lh3.googleusercontent.com/d/" + photoId + "=s200";
          viewUrl = "https://drive.google.com/file/d/" + photoId + "/view";
        }
        
        debts.push({
          id: data[i][0],
          date: data[i][1],
          borrowerName: data[i][2],
          borrowerPhone: data[i][3] || "",
          borrowerNIDA: data[i][4] || "",
          photoId: photoId,
          photoBase64: photoBase64,
          photoMimeType: photoMimeType,
          thumbnailUrl: thumbnailUrl,
          directUrl: directUrl,
          lh3Url: lh3Url,
          viewUrl: viewUrl,
          photoUrl: lh3Url, // Tumia lh3Url kama default kwa offline
          product: data[i][6],
          quantity: data[i][7],
          unitPrice: data[i][8],
          totalAmount: data[i][9],
          paidAmount: data[i][10] || 0,
          status: debtStatus,
          branchId: debtBranchId,
          sellerId: debtSellerId,
          notes: data[i][14] || "",
          lastUpdated: data[i][15] || data[i][1],
          unit: data[i][16] || ""
        });
      }
    }
    
    return debts;
    
  } catch (error) {
    console.error("Error getting debts for sync:", error);
    return [];
  }
}

/* ===== UPDATE DEBT ===== */
function updateDebt(d) {
  try {
    if (!d.databaseId || !d.debtId) {
      return output({ success: false, error: "Database ID and Debt ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    
    if (!debtsSheet) {
      return output({ success: false, error: "Debts sheet not found" });
    }
    
    const data = debtsSheet.getDataRange().getValues();
    const timestamp = new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.debtId) {
        const row = i + 1;
        
        const oldProduct = data[i][6];
        const oldQuantity = data[i][7] || 0;
        const oldUnit = data[i][16] || "";
        const branchId = data[i][12] || "MAIN";
        
        if (d.borrowerName) debtsSheet.getRange(row, 3).setValue(d.borrowerName); // C
        if (d.borrowerPhone !== undefined) debtsSheet.getRange(row, 4).setValue(d.borrowerPhone); // D
        if (d.borrowerNIDA !== undefined) debtsSheet.getRange(row, 5).setValue(d.borrowerNIDA); // E
        if (d.photoId !== undefined) debtsSheet.getRange(row, 6).setValue(d.photoId); // F
        if (d.photoBase64 !== undefined) {
          ensureDebtPhotoColumns_(debtsSheet);
          var editPhotoBase64 = String(d.photoBase64 || '');
          if (editPhotoBase64.indexOf('data:') === 0) { var editComma = editPhotoBase64.indexOf(','); if (editComma !== -1) editPhotoBase64 = editPhotoBase64.substring(editComma + 1); }
          if (editPhotoBase64.length > 48000) return output({ success: false, error: 'Borrower photo exceeds 50,000 character cell limit' });
          debtsSheet.getRange(row, 18).setValue(editPhotoBase64);
          debtsSheet.getRange(row, 19).setValue(d.photoMimeType || 'image/jpeg');
        }
        if (d.notes !== undefined) debtsSheet.getRange(row, 15).setValue(d.notes); // O
        if (d.status) debtsSheet.getRange(row, 12).setValue(d.status); // L
        
        // Editing the item itself (Product/Quantity/Price/Unit) - reconcile
        // Stock so inventory stays accurate: give the OLD quantity back
        // first, then deduct the NEW quantity (matches how a fresh Debt
        // deducts stock in addDebt()).
        if (d.quantity !== undefined && d.unitPrice !== undefined) {
          const newQuantity = Number(d.quantity) || 0;
          const newUnitPrice = Number(d.unitPrice) || 0;
          const newProduct = d.product || oldProduct;
          const newUnit = d.unit !== undefined ? d.unit : oldUnit;
          const newTotal = newQuantity * newUnitPrice;
          
          if (d.isStockProduct !== false) {
            if (oldProduct) {
              returnStockQuantity(userDB, oldProduct, oldQuantity, branchId, oldUnit);
            }
            updateStockAfterSale(d.databaseId, newProduct, newQuantity, branchId, newUnit);
          }
          
          debtsSheet.getRange(row, 7).setValue(newProduct);   // Product
          debtsSheet.getRange(row, 8).setValue(newQuantity);  // Quantity
          debtsSheet.getRange(row, 9).setValue(newUnitPrice); // Unit Price
          debtsSheet.getRange(row, 10).setValue(newTotal);    // Total Amount
          if (debtsSheet.getLastColumn() < 17) {
            debtsSheet.getRange(1, 17).setValue("Unit");
          }
          debtsSheet.getRange(row, 17).setValue(newUnit);     // Unit
        }
        
        // Update last updated
        debtsSheet.getRange(row, 16).setValue(timestamp); // P
        
        return output({
          success: true,
          message: "Debt updated successfully",
          timestamp: timestamp.toISOString()
        });
      }
    }
    
    return output({ success: false, error: "Debt not found" });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Adds a quantity BACK to stock (used when editing a Debt changes the
// item, so the old quantity is returned before the new one is deducted).
// Unit-aware: only merges into a row that matches the SAME unit too.
function returnStockQuantity(userDB, product, quantity, branchId, unit) {
  if (!quantity) return;
  const stockSheet = userDB.getSheetByName("STOCK");
  if (!stockSheet) return;
  
  const stockData = stockSheet.getDataRange().getValues();
  for (let j = 1; j < stockData.length; j++) {
    if (stockData[j][0] && stockData[j][1] === product) {
      const itemBranchId = stockData[j][8] || "MAIN";
      const itemUnit = stockData[j][9] || "";
      if (itemBranchId === branchId && (itemUnit === (unit || "") || !itemUnit)) {
        const newQuantity = (stockData[j][2] || 0) + quantity;
        const buyPrice = stockData[j][3] || 0;
        stockSheet.getRange(j + 1, 3).setValue(newQuantity);
        stockSheet.getRange(j + 1, 6).setValue(newQuantity * buyPrice);
        return;
      }
    }
  }
}




/* ===== UPLOAD BORROWER PHOTO USING DRIVE API ONLY ===== */
function uploadBorrowerPhoto(d) {
  try {
    if (!d.email || !d.photoData) {
      return output({ success: false, error: "Email and photo data required" });
    }
    
    // Prepare file metadata
    const fileName = `borrower_${d.debtId || Date.now()}_${d.email.replace(/[@.]/g, "_")}.${d.fileExtension || 'jpg'}`;
    const blob = Utilities.newBlob(Utilities.base64Decode(d.photoData), d.mimeType, fileName);
    
    // Create file using Drive API
    const fileMetadata = {
      'title': fileName,
      'mimeType': d.mimeType,
      'parents': [{'id': DEBT_PHOTOS_FOLDER_ID}]
    };
    
    const file = Drive.Files.insert(fileMetadata, blob, {
      'convert': false
    });
    
    // Set sharing permissions using Drive API
    Drive.Permissions.insert({
      'role': 'reader',
      'type': 'anyone'
    }, file.id, {
      'sendNotificationEmails': 'false'
    });
    
    // Generate URLs
    const lh3Url = "https://lh3.googleusercontent.com/d/" + file.id + "=s200";
    const thumbnailUrl = "https://drive.google.com/thumbnail?id=" + file.id + "&sz=s200";
    const directUrl = "https://drive.google.com/uc?export=view&id=" + file.id;
    
    return output({
      success: true,
      message: "Photo uploaded successfully",
      photoId: file.id,
      photoUrl: lh3Url,
      thumbnailUrl: thumbnailUrl,
      directUrl: directUrl,
      lh3Url: lh3Url
    });
    
  } catch (error) {
    return output({ success: false, error: "Failed to upload photo: " + error.toString() });
  }
}




/* ===== GENERATE PDF REPORT USING DRIVE API ONLY (LOGO FIX) ===== */
/* ============================================================
   FILTERED REPORT DATA HELPERS (Stock, Sales, Purchases, Debts, Profit)
   Used by both generateReport() (quick "Full Business Report" download)
   and generateFilteredReport() (Report page's type-specific downloads).
   ============================================================ */

function getFilteredStockData(userDB, filters) {
  filters = filters || {};
  const stockSheet = userDB.getSheetByName("STOCK");
  if (!stockSheet) return { items: [], totalValue: 0, totalItems: 0 };
  
  const data = stockSheet.getDataRange().getValues();
  const items = [];
  let totalValue = 0;
  let totalItems = 0;
  
  const branchesSheet = userDB.getSheetByName("BRANCHES");
  const branchMap = {};
  if (branchesSheet) {
    const branchData = branchesSheet.getDataRange().getValues();
    for (let i = 1; i < branchData.length; i++) {
      if (branchData[i][0]) branchMap[branchData[i][0]] = branchData[i][1];
    }
  }
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    const branchId = data[i][8] || "MAIN";
    const category = data[i][6] || "";
    const quantity = data[i][2] || 0;
    
    if (filters.branch && branchId !== filters.branch) continue;
    if (filters.category && category !== filters.category) continue;
    if (filters.status === "low" && !(quantity > 0 && quantity < 10)) continue;
    if (filters.status === "out" && quantity !== 0) continue;
    if (filters.status === "available" && quantity <= 0) continue;
    
    items.push({
      date: formatDateForReport(data[i][0]),
      product: data[i][1],
      branch: branchMap[branchId] || branchId,
      quantity: quantity,
      buyPrice: data[i][3] || 0,
      sellPrice: data[i][4] || 0,
      value: data[i][5] || 0,
      category: category,
      unit: data[i][9] || ""
    });
    
    totalValue += (data[i][5] || 0);
    totalItems += quantity;
  }
  
  if (filters.sort === "product") items.sort((a, b) => (a.product || "").localeCompare(b.product || ""));
  else if (filters.sort === "value") items.sort((a, b) => (b.value || 0) - (a.value || 0));
  else if (filters.sort === "quantity") items.sort((a, b) => (b.quantity || 0) - (a.quantity || 0));
  
  return { items: items, totalValue: totalValue, totalItems: totalItems };
}

function getFilteredSalesData(userDB, filters) {
  filters = filters || {};
  const salesSheet = userDB.getSheetByName("SALES");
  if (!salesSheet) return { items: [], totalAmount: 0 };
  
  const data = salesSheet.getDataRange().getValues();
  const items = [];
  let totalAmount = 0;
  
  const branchesSheet = userDB.getSheetByName("BRANCHES");
  const branchMap = {};
  if (branchesSheet) {
    const branchData = branchesSheet.getDataRange().getValues();
    for (let i = 1; i < branchData.length; i++) {
      if (branchData[i][0]) branchMap[branchData[i][0]] = branchData[i][1];
    }
  }
  
  const sellersSheet = userDB.getSheetByName("SELLERS");
  const sellerMap = {};
  if (sellersSheet) {
    const sellerData = sellersSheet.getDataRange().getValues();
    for (let i = 1; i < sellerData.length; i++) {
      if (sellerData[i][0]) sellerMap[sellerData[i][0]] = sellerData[i][1];
    }
  }
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    const saleDate = new Date(data[i][0]);
    const branchId = data[i][8] || "MAIN";
    const sellerId = data[i][7] || null;
    const product = data[i][1];
    const paymentMethod = data[i][6] || "cash";
    
    if (filters.dateFrom && saleDate < new Date(filters.dateFrom)) continue;
    if (filters.dateTo) {
      const toDate = new Date(filters.dateTo);
      toDate.setHours(23, 59, 59, 999);
      if (saleDate > toDate) continue;
    }
    if (filters.branch && branchId !== filters.branch) continue;
    if (filters.seller && sellerId !== filters.seller) continue;
    if (filters.product && product !== filters.product) continue;
    if (filters.payment && paymentMethod !== filters.payment) continue;
    
    const total = data[i][4] || 0;
    items.push({
      date: formatDateForReport(data[i][0]),
      product: product,
      seller: sellerId ? (sellerMap[sellerId] || "Unknown Seller") : "Admin",
      branch: branchMap[branchId] || branchId,
      quantity: data[i][2] || 0,
      unitPrice: data[i][3] || 0,
      total: total,
      customer: data[i][5] || "Walk-in",
      paymentMethod: paymentMethod
    });
    
    totalAmount += total;
  }
  
  return { items: items, totalAmount: totalAmount, count: items.length };
}

function getFilteredPurchasesData(userDB, filters) {
  filters = filters || {};
  const purchasesSheet = userDB.getSheetByName("PURCHASES");
  if (!purchasesSheet) return { items: [], totalAmount: 0 };
  
  const data = purchasesSheet.getDataRange().getValues();
  const items = [];
  let totalAmount = 0;
  
  const branchesSheet = userDB.getSheetByName("BRANCHES");
  const branchMap = {};
  if (branchesSheet) {
    const branchData = branchesSheet.getDataRange().getValues();
    for (let i = 1; i < branchData.length; i++) {
      if (branchData[i][0]) branchMap[branchData[i][0]] = branchData[i][1];
    }
  }
  
  const sellersSheet = userDB.getSheetByName("SELLERS");
  const sellerMap = {};
  if (sellersSheet) {
    const sellerData = sellersSheet.getDataRange().getValues();
    for (let i = 1; i < sellerData.length; i++) {
      if (sellerData[i][0]) sellerMap[sellerData[i][0]] = sellerData[i][1];
    }
  }
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    const purchaseDate = new Date(data[i][0]);
    const branchId = data[i][8] || "MAIN";
    const sellerId = data[i][7] || null;
    const product = data[i][1];
    const supplier = data[i][5] || "Local Supplier";
    const paymentMethod = data[i][6] || "cash";
    
    if (filters.dateFrom && purchaseDate < new Date(filters.dateFrom)) continue;
    if (filters.dateTo) {
      const toDate = new Date(filters.dateTo);
      toDate.setHours(23, 59, 59, 999);
      if (purchaseDate > toDate) continue;
    }
    if (filters.branch && branchId !== filters.branch) continue;
    if (filters.seller && sellerId !== filters.seller) continue;
    if (filters.supplier && supplier !== filters.supplier) continue;
    if (filters.product && product !== filters.product) continue;
    if (filters.payment && paymentMethod !== filters.payment) continue;
    
    const total = data[i][4] || 0;
    items.push({
      date: formatDateForReport(data[i][0]),
      product: product,
      seller: sellerId ? (sellerMap[sellerId] || "Unknown Seller") : "Admin",
      branch: branchMap[branchId] || branchId,
      quantity: data[i][2] || 0,
      unitCost: data[i][3] || 0,
      total: total,
      supplier: supplier,
      paymentMethod: paymentMethod,
      category: data[i][9] || "",
      unit: data[i][10] || ""
    });
    
    totalAmount += total;
  }
  
  return { items: items, totalAmount: totalAmount };
}

function getFilteredDebtsData(userDB, filters) {
  filters = filters || {};
  const debtsSheet = userDB.getSheetByName("DEBTS");
  if (!debtsSheet) return { items: [], totalAmount: 0, paidAmount: 0, pendingAmount: 0 };
  
  const data = debtsSheet.getDataRange().getValues();
  const items = [];
  let totalAmount = 0, paidAmount = 0, pendingAmount = 0;
  
  const branchesSheet = userDB.getSheetByName("BRANCHES");
  const branchMap = {};
  if (branchesSheet) {
    const branchData = branchesSheet.getDataRange().getValues();
    for (let i = 1; i < branchData.length; i++) {
      if (branchData[i][0]) branchMap[branchData[i][0]] = branchData[i][1];
    }
  }
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    const debtDate = new Date(data[i][1]);
    const branchId = data[i][12] || "MAIN";
    const sellerId = data[i][13] || null;
    const status = data[i][11] || "pending";
    
    if (filters.dateFrom && debtDate < new Date(filters.dateFrom)) continue;
    if (filters.dateTo) {
      const toDate = new Date(filters.dateTo);
      toDate.setHours(23, 59, 59, 999);
      if (debtDate > toDate) continue;
    }
    if (filters.branch && branchId !== filters.branch) continue;
    if (filters.seller && sellerId !== filters.seller) continue;
    if (filters.status && status !== filters.status) continue;
    
    const total = data[i][9] || 0;
    const paid = data[i][10] || 0;
    
    items.push({
      date: formatDateForReport(data[i][1]),
      borrowerName: data[i][2],
      borrowerPhone: data[i][3] || "",
      product: data[i][6],
      quantity: data[i][7] || 0,
      unit: data[i][16] || "",
      totalAmount: total,
      paidAmount: paid,
      status: status,
      branch: branchMap[branchId] || branchId
    });
    
    totalAmount += total;
    paidAmount += paid;
    if (status !== "paid") pendingAmount += (total - paid);
  }
  
  return { items: items, totalAmount: totalAmount, paidAmount: paidAmount, pendingAmount: pendingAmount };
}

function getProfitData(userDB, filters) {
  filters = filters || {};
  const salesResult = getFilteredSalesData(userDB, filters);
  const purchasesResult = getFilteredPurchasesData(userDB, filters);
  
  const totalSales = salesResult.totalAmount || 0;
  const totalPurchases = purchasesResult.totalAmount || 0;
  const profit = totalSales - totalPurchases;
  const profitMargin = totalSales > 0 ? (profit / totalSales) * 100 : 0;
  const salesCount = salesResult.items.length;
  const purchasesCount = purchasesResult.items.length;
  const averageSaleValue = salesCount > 0 ? totalSales / salesCount : 0;
  
  const productTotals = {};
  salesResult.items.forEach(item => {
    if (!productTotals[item.product]) {
      productTotals[item.product] = { product: item.product, quantity: 0, amount: 0 };
    }
    productTotals[item.product].quantity += (item.quantity || 0);
    productTotals[item.product].amount += (item.total || 0);
  });
  
  const topProducts = Object.values(productTotals)
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 5);
  
  return {
    totalSales: totalSales,
    totalPurchases: totalPurchases,
    profit: profit,
    profitMargin: profitMargin,
    salesCount: salesCount,
    purchasesCount: purchasesCount,
    averageSaleValue: averageSaleValue,
    topProducts: topProducts
  };
}

function generateReport(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settings = getUserSettingsFromDB(d.databaseId);
    
    // Get business name
    let businessName = settings.PROFILE?.business_name?.value || "";
    if (!businessName) {
      try {
        const masterDB = getMasterDatabase();
        const masterSheet = masterDB.getSheetByName("MASTER_USERS");
        const masterData = masterSheet.getDataRange().getValues();
        
        for (let i = 1; i < masterData.length; i++) {
          if (masterData[i][1] === d.email && masterData[i][3] === d.databaseId) {
            businessName = masterData[i][0];
            break;
          }
        }
      } catch (e) {
        console.log("Error reading master database:", e);
      }
    }
    if (!businessName) {
      businessName = d.email.split('@')[0] + "'s Business";
    }
    
    // Get data
    const stockData = getFilteredStockData(userDB, {});
    const salesData = getFilteredSalesData(userDB, {});
    const purchasesData = getFilteredPurchasesData(userDB, {});
    const debtsData = getFilteredDebtsData(userDB, {});
    
    // GET LOGO - prefer the base64 copy already stored in this business's
    // own Settings (simple field lookup, no external fetch, so it can
    // never silently fail) - fall back to fetching from Drive only for
    // accounts that haven't re-uploaded since this feature was added.
    let logoBase64 = null;
    let logoContentType = 'image/png';
    
    const storedLogo = getLogoBase64FromSettings(settings);
    if (storedLogo) {
      logoBase64 = storedLogo.base64;
      logoContentType = storedLogo.mimeType;
      console.log("Logo loaded from Settings base64, length:", logoBase64.length);
    } else {
      try {
        // Get logo ID from master database
        const logoId = getUserLogoIdFromMaster(d.email, d.databaseId);
        
        if (logoId) {
          // Get file metadata using Drive API
          const file = Drive.Files.get(logoId);
          
          if (file && file.id) {
            // Download file content directly from Drive
            const response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + logoId + "?alt=media", {
              headers: {
                Authorization: "Bearer " + ScriptApp.getOAuthToken()
              }
            });
            
            const blob = response.getBlob();
            logoBase64 = Utilities.base64Encode(blob.getBytes());
            logoContentType = blob.getContentType();
            console.log("Logo loaded directly from Drive:", file.title);
          }
        }
      } catch (logoError) {
        console.log("Could not fetch logo from Drive:", logoError);
        // Continue without logo
      }
    }
    
    // Create HTML content with logo as base64 - uses the same modern,
    // redesigned Full Report layout as the filtered-report download.
    const reportData = { stock: stockData, sales: salesData, purchases: purchasesData, debts: debtsData };
    const htmlContent = createFullReportHTML(businessName, d.email, 'Full Business Report', reportData, {}, logoBase64, logoContentType, null);
    
    // Create PDF
    const safeBusinessName = businessName.replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = safeBusinessName + '_Full_Report_' + formatBusinessDate_(new Date(), 'yyyy-MM-dd') + '.pdf';
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'report.html');
    const pdfBlob = blob.getAs('application/pdf').setName(fileName);
    
    // Save to Drive using Drive API
    const fileMetadata = {
      'title': fileName,
      'mimeType': 'application/pdf',
      'parents': [{'id': REPORTS_FOLDER_ID}]
    };
    
    const pdfFile = Drive.Files.insert(fileMetadata, pdfBlob);
    
    // Set sharing permissions using Drive API
    Drive.Permissions.insert({
      'role': 'reader',
      'type': 'anyone'
    }, pdfFile.id, {
      'sendNotificationEmails': 'false'
    });
    
    const downloadUrl = "https://drive.google.com/uc?export=download&id=" + pdfFile.id;
    
    // Also return the raw file bytes (base64) so the client can build the
    // download entirely locally (Blob + object URL) instead of depending
    // on navigating to drive.google.com - some restricted WebViews (e.g.
    // app-wrapper tools) block or mishandle that external navigation.
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    return output({
      success: true,
      message: "Report generated successfully",
      fileId: pdfFile.id,
      fileName: pdfFile.title,
      downloadUrl: downloadUrl,
      fileData: pdfBase64,
      mimeType: "application/pdf",
      reportType: "Full Business Report"
    });
    
  } catch (error) {
    return output({ success: false, error: "Failed to generate report: " + error.toString() });
  }
}


/* ===== CREATE REPORT HTML ===== */
function createReportHTML(businessName, email, stockData, salesData, purchasesData, debtsData, logoBlob) {
  const date = new Date();
  const reportDate = formatBusinessDateTime_(date);
  
  // Format currency
  const formatCurrency = (amount) => {
    return Number(amount).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  };
  
  // Encode logo as base64 if exists
  let logoBase64 = '';
  if (logoBlob) {
    logoBase64 = 'data:' + logoBlob.getContentType() + ';base64,' + Utilities.base64Encode(logoBlob.getBytes());
  }
  
  // Calculate totals
  const totalStockValue = stockData.totalValue || 0;
  const totalSales = salesData.totalAmount || 0;
  const totalPurchases = purchasesData.totalAmount || 0;
  const totalDebts = debtsData.totalAmount || 0;
  const totalPaidDebts = debtsData.paidAmount || 0;
  const totalPendingDebts = debtsData.pendingAmount || 0;
  const profit = totalSales - totalPurchases;
  
  // Create HTML
  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <title>${businessName} Report</title>
    <style>
      body {
        font-family: 'Segoe UI', Arial, sans-serif;
        margin: 20px;
        color: #333;
      }
      .header {
        display: flex;
        align-items: center;
        margin-bottom: 30px;
        border-bottom: 2px solid #38bdf8;
        padding-bottom: 15px;
      }
      .logo {
        width: 80px;
        height: 80px;
        border-radius: 50%;
        margin-right: 20px;
        object-fit: cover;
        border: 2px solid #38bdf8;
      }
      .title {
        flex: 1;
      }
      .title h1 {
        margin: 0;
        color: #0f172a;
        font-size: 24px;
      }
      .title p {
        margin: 5px 0 0;
        color: #64748b;
        font-size: 14px;
      }
      .summary-cards {
        display: flex;
        flex-wrap: wrap;
        gap: 15px;
        margin-bottom: 30px;
      }
      .card {
        flex: 1;
        min-width: 150px;
        background: linear-gradient(135deg, #f8fafc, #f1f5f9);
        border-left: 5px solid #38bdf8;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      }
      .card h3 {
        margin: 0 0 10px;
        font-size: 14px;
        color: #64748b;
      }
      .card .value {
        font-size: 20px;
        font-weight: 700;
        color: #0f172a;
      }
      .section {
        margin-bottom: 30px;
        page-break-inside: avoid;
      }
      .section h2 {
        background: #0f172a;
        color: white;
        padding: 10px 15px;
        margin: 0 0 15px;
        border-radius: 5px;
        font-size: 18px;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 15px;
        font-size: 12px;
      }
      th {
        background: #1e293b;
        color: white;
        padding: 8px;
        text-align: left;
        font-weight: 600;
      }
      td {
        padding: 6px 8px;
        border-bottom: 1px solid #e2e8f0;
      }
      tr:nth-child(even) {
        background: #f8fafc;
      }
      .total-row {
        font-weight: 700;
        background: #e2e8f0 !important;
      }
      .status-pending {
        color: #f59e0b;
        font-weight: 600;
      }
      .status-paid {
        color: #22c55e;
        font-weight: 600;
      }
      .status-denied {
        color: #ef4444;
        font-weight: 600;
      }
      .footer {
        margin-top: 40px;
        text-align: center;
        font-size: 12px;
        color: #94a3b8;
        border-top: 1px solid #e2e8f0;
        padding-top: 15px;
      }
    </style>
  </head>
  <body>
    <div class="header">
      ${logoBase64 ? '<img src="' + logoBase64 + '" class="logo" alt="Logo">' : ''}
      <div class="title">
        <h1>${businessName}</h1>
        <p>Email: ${email}</p>
        <p>Report Generated: ${reportDate}</p>
      </div>
    </div>
    
    <div class="summary-cards">
      <div class="card">
        <h3>Stock Value</h3>
        <div class="value">${formatCurrency(totalStockValue)}</div>
      </div>
      <div class="card">
        <h3>Total Sales</h3>
        <div class="value">${formatCurrency(totalSales)}</div>
      </div>
      <div class="card">
        <h3>Total Purchases</h3>
        <div class="value">${formatCurrency(totalPurchases)}</div>
      </div>
      <div class="card">
        <h3>Profit</h3>
        <div class="value">${formatCurrency(profit)}</div>
      </div>
      <div class="card">
        <h3>Total Debts</h3>
        <div class="value">${formatCurrency(totalDebts)}</div>
      </div>
      <div class="card">
        <h3>Pending Debts</h3>
        <div class="value">${formatCurrency(totalPendingDebts)}</div>
      </div>
    </div>
    
    <div class="section">
      <h2>Stock Inventory (${stockData.count} items)</h2>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Branch</th>
            <th>Quantity</th>
            <th>Buy Price</th>
            <th>Sell Price</th>
            <th>Value</th>
            <th>Category</th>
          </tr>
        </thead>
        <tbody>
          ${stockData.items.map(item => `
            <tr>
              <td>${item.date}</td>
              <td>${item.product}</td>
              <td>${item.branch}</td>
              <td>${item.quantity}</td>
              <td>${formatCurrency(item.buyPrice)}</td>
              <td>${formatCurrency(item.sellPrice)}</td>
              <td>${formatCurrency(item.value)}</td>
              <td>${item.category}</td>
            </tr>
          `).join('')}
          <tr class="total-row">
            <td colspan="3"><strong>Total</strong></td>
            <td><strong>${stockData.totalItems}</strong></td>
            <td colspan="2"></td>
            <td><strong>${formatCurrency(stockData.totalValue)}</strong></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <div class="section">
      <h2>Sales (${salesData.count} transactions)</h2>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Unit Price</th>
            <th>Total</th>
            <th>Customer</th>
            <th>Payment</th>
          </tr>
        </thead>
        <tbody>
          ${salesData.items.map(item => `
            <tr>
              <td>${item.date}</td>
              <td>${item.product}</td>
              <td>${item.seller}</td>
              <td>${item.branch}</td>
              <td>${item.quantity}</td>
              <td>${formatCurrency(item.unitPrice)}</td>
              <td>${formatCurrency(item.total)}</td>
              <td>${item.customer}</td>
              <td>${item.paymentMethod}</td>
            </tr>
          `).join('')}
          <tr class="total-row">
            <td colspan="6"><strong>Total</strong></td>
            <td><strong>${formatCurrency(salesData.totalAmount)}</strong></td>
            <td colspan="2"></td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <div class="section">
      <h2>Purchases (${purchasesData.count} transactions)</h2>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Unit Cost</th>
            <th>Total</th>
            <th>Supplier</th>
            <th>Payment</th>
          </tr>
        </thead>
        <tbody>
          ${purchasesData.items.map(item => `
            <tr>
              <td>${item.date}</td>
              <td>${item.product}</td>
              <td>${item.seller}</td>
              <td>${item.branch}</td>
              <td>${item.quantity}</td>
              <td>${formatCurrency(item.unitCost)}</td>
              <td>${formatCurrency(item.total)}</td>
              <td>${item.supplier}</td>
              <td>${item.paymentMethod}</td>
            </tr>
          `).join('')}
          <tr class="total-row">
            <td colspan="6"><strong>Total</strong></td>
            <td><strong>${formatCurrency(purchasesData.totalAmount)}</strong></td>
            <td colspan="2"></td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <div class="section">
      <h2>Debts (${debtsData.count} records)</h2>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Borrower</th>
            <th>Phone</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Total</th>
            <th>Paid</th>
            <th>Balance</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${debtsData.items.map(item => {
            const statusClass = item.status === 'paid' ? 'status-paid' : (item.status === 'denied' ? 'status-denied' : 'status-pending');
            return `
            <tr>
              <td>${item.date}</td>
              <td>${item.borrowerName}</td>
              <td>${item.borrowerPhone}</td>
              <td>${item.product}</td>
              <td>${item.seller}</td>
              <td>${item.branch}</td>
              <td>${item.quantity}</td>
              <td>${formatCurrency(item.totalAmount)}</td>
              <td>${formatCurrency(item.paidAmount)}</td>
              <td>${formatCurrency(item.balance)}</td>
              <td class="${statusClass}">${item.status.toUpperCase()}</td>
            </tr>
          `}).join('')}
          <tr class="total-row">
            <td colspan="7"><strong>Totals</strong></td>
            <td><strong>${formatCurrency(debtsData.totalAmount)}</strong></td>
            <td><strong>${formatCurrency(debtsData.paidAmount)}</strong></td>
            <td><strong>${formatCurrency(debtsData.pendingAmount)}</strong></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <div class="footer">
      <p>Generated by SmartTrade Pro on ${reportDate}</p>
      <p>This is an automated report. All amounts are in local currency.</p>
    </div>
  </body>
  </html>
  `;
  
  return html;
}

/* ===== FORMAT DATE FOR REPORT ===== */
function formatDateForReport(date) {
  if (!date) return '—';
  try {
    const d = new Date(date);
    return formatBusinessDateTime_(d);
  } catch (e) {
    return '—';
  }
}



/* ===== GENERATE FILTERED REPORT - FIXED (NO FRONTEND DEPENDENCY) ===== */
function generateFilteredReport(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const settings = getUserSettingsFromDB(d.databaseId);
    const filters = d.filters || {};
    const reportType = filters.reportType || 'full';
    
    // Get business name
    let businessName = settings.PROFILE?.business_name?.value || "";
    if (!businessName) {
      try {
        const masterDB = getMasterDatabase();
        const masterSheet = masterDB.getSheetByName("MASTER_USERS");
        const masterData = masterSheet.getDataRange().getValues();
        for (let i = 1; i < masterData.length; i++) {
          if (masterData[i][1] === d.email && masterData[i][3] === d.databaseId) {
            businessName = masterData[i][0];
            break;
          }
        }
      } catch (e) {
        console.log("Error reading master database:", e);
      }
    }
    if (!businessName) {
      businessName = d.email.split('@')[0] + "'s Business";
    }
    
    // Get filtered data based on report type
    let reportData = {};
    let reportTitle = "";
    let contentType = "";
    
    switch(reportType) {
      case 'stock':
        reportData = getFilteredStockData(userDB, filters);
        reportTitle = "Stock Report";
        contentType = "stock";
        break;
      case 'sales':
        reportData = getFilteredSalesData(userDB, filters);
        reportTitle = "Sales Report";
        contentType = "sales";
        break;
      case 'purchases':
        reportData = getFilteredPurchasesData(userDB, filters);
        reportTitle = "Purchases Report";
        contentType = "purchases";
        break;
      case 'debts':
        reportData = getFilteredDebtsData(userDB, filters);
        reportTitle = "Debts Report";
        contentType = "debts";
        break;
      case 'profit':
        reportData = getProfitData(userDB, filters);
        reportTitle = "Profit & Loss Report";
        contentType = "profit";
        break;
      default:
        reportData = {
          stock: getFilteredStockData(userDB, {}),
          sales: getFilteredSalesData(userDB, {}),
          purchases: getFilteredPurchasesData(userDB, {}),
          debts: getFilteredDebtsData(userDB, {})
        };
        reportTitle = "Full Business Report";
        contentType = "full";
    }
    
    // ============================================================
    // GET LOGO - prefer the base64 copy already stored in this
    // business's own Settings (simple field lookup, can't silently
    // fail) - fall back to Drive only for accounts that haven't
    // re-uploaded since this feature was added.
    // ============================================================
    let logoBase64 = null;
    let logoContentType = 'image/png';
    let logoFileName = null;
    
    const storedLogo = getLogoBase64FromSettings(settings);
    if (storedLogo) {
      logoBase64 = storedLogo.base64;
      logoContentType = storedLogo.mimeType;
      console.log("Logo loaded from Settings base64, length:", logoBase64.length);
    } else {
      try {
        // Get logo ID from master database (NOT from frontend)
        const logoId = getUserLogoIdFromMaster(d.email, d.databaseId);
        console.log("Logo ID from master database:", logoId);
        
        if (logoId && logoId !== "") {
          try {
            // Verify file exists in Drive
            const file = Drive.Files.get(logoId);
            if (file && file.id) {
              console.log("Logo file found in Drive:", file.title);
              logoFileName = file.title;
              
              // Download file directly from Drive using Drive API
              const response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/files/" + logoId + "?alt=media", {
                headers: {
                  Authorization: "Bearer " + ScriptApp.getOAuthToken()
                },
                muteHttpExceptions: true
              });
              
              if (response.getResponseCode() === 200) {
                const blob = response.getBlob();
                logoBase64 = Utilities.base64Encode(blob.getBytes());
                logoContentType = blob.getContentType() || 'image/png';
                console.log("Logo loaded directly from Drive, base64 length:", logoBase64.length);
              } else {
                console.log("Failed to fetch logo, response code:", response.getResponseCode());
              }
            }
          } catch (e) {
            console.log("Error fetching logo from Drive:", e.toString());
          }
        } else {
          console.log("No logo ID found in master database for user:", d.email);
        }
      } catch (logoError) {
        console.log("Could not fetch logo from Drive:", logoError.toString());
      }
    }
    
    // Debug log
    console.log("Final logo status:", logoBase64 ? "LOGO LOADED (length: " + logoBase64.length + ")" : "NO LOGO");
    
    // Create HTML content based on report type
    let htmlContent = "";
    
    if (contentType === "full") {
      htmlContent = createFullReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    } else if (contentType === "stock") {
      htmlContent = createStockReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    } else if (contentType === "sales") {
      htmlContent = createSalesReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    } else if (contentType === "purchases") {
      htmlContent = createPurchasesReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    } else if (contentType === "debts") {
      htmlContent = createDebtsReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    } else if (contentType === "profit") {
      htmlContent = createProfitReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    } else {
      htmlContent = createFullReportHTML(businessName, d.email, reportTitle, reportData, filters, logoBase64, logoContentType, logoFileName);
    }
    
    // Create PDF
    const safeBusinessName = businessName.replace(/[^a-zA-Z0-9]/g, '_');
    const fileName = safeBusinessName + '_' + reportType + '_report_' + formatBusinessDate_(new Date(), 'yyyy-MM-dd') + '.pdf';
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'report.html');
    const pdfBlob = blob.getAs('application/pdf').setName(fileName);
    
    // Save to Drive
    const fileMetadata = {
      'title': fileName,
      'mimeType': 'application/pdf',
      'parents': [{'id': REPORTS_FOLDER_ID}]
    };
    
    const pdfFile = Drive.Files.insert(fileMetadata, pdfBlob);
    Drive.Permissions.insert({ 'role': 'reader', 'type': 'anyone' }, pdfFile.id, { 'sendNotificationEmails': 'false' });
    
    const downloadUrl = "https://drive.google.com/uc?export=download&id=" + pdfFile.id;
    
    // Also return the raw file bytes (base64) so the client can build the
    // download entirely locally (Blob + object URL) instead of depending
    // on navigating to drive.google.com - some restricted WebViews (e.g.
    // app-wrapper tools) block or mishandle that external navigation.
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    return output({
      success: true,
      message: "Report generated successfully",
      fileId: pdfFile.id,
      fileName: fileName,
      downloadUrl: downloadUrl,
      fileData: pdfBase64,
      mimeType: "application/pdf",
      reportType: reportTitle
    });
    
  } catch (error) {
    return output({ success: false, error: "Failed to generate report: " + error.toString() });
  }
}



// Create Stock Report HTML
/* ============================================================
   MODERN PDF REPORT DESIGN SYSTEM
   Shared styling, header, footer and helpers used by every report type
   below (Stock, Sales, Purchases, Debts, Profit & Loss, Full Report) so
   they all share one clean, consistent, modern look instead of each
   report having its own ad-hoc layout.
   ============================================================ */

function escapeHtmlReport(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatReportCurrency(amount) {
  const n = Number(amount) || 0;
  return n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function getModernReportCSS() {
  return `
    * { box-sizing: border-box; }
    body { font-family: 'Segoe UI', Roboto, Arial, sans-serif; margin: 0; padding: 0; color: #1e293b; background: #ffffff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .page { padding: 30px 34px; }
    
    .report-header { display: flex; align-items: center; gap: 18px; padding-bottom: 18px; border-bottom: 3px solid #0ea5e9; margin-bottom: 22px; }
    .report-logo { width: 60px; height: 60px; border-radius: 12px; object-fit: cover; border: 1px solid #e2e8f0; }
    .report-logo-placeholder { width: 60px; height: 60px; border-radius: 12px; background: linear-gradient(135deg, #0ea5e9, #0369a1); display: flex; align-items: center; justify-content: center; color: #ffffff; font-size: 24px; font-weight: 700; }
    .report-header-text { flex: 1; }
    .report-business-name { font-size: 21px; font-weight: 800; color: #0f172a; margin: 0; }
    .report-title { font-size: 13.5px; font-weight: 700; color: #0ea5e9; margin: 3px 0 0; text-transform: uppercase; letter-spacing: 0.6px; }
    .report-meta { font-size: 10.5px; color: #64748b; margin-top: 7px; }
    .report-meta span { margin-right: 16px; }
    
    .filters-bar { font-size: 10.5px; color: #64748b; background: #f1f5f9; border-radius: 6px; padding: 8px 12px; margin-bottom: 18px; }
    .filters-bar strong { color: #0f172a; }
    
    .kpi-table { width: 100%; border-collapse: separate; border-spacing: 8px 0; margin: 0 0 24px; table-layout: fixed; }
    .kpi-cell { background: #f8fafc; border: 1px solid #e2e8f0; border-top: 3px solid #0ea5e9; border-radius: 10px; padding: 12px 14px; vertical-align: top; }
    .kpi-cell.accent-green { border-top-color: #22c55e; }
    .kpi-cell.accent-amber { border-top-color: #f59e0b; }
    .kpi-cell.accent-red { border-top-color: #ef4444; }
    .kpi-label { font-size: 9.5px; color: #64748b; text-transform: uppercase; letter-spacing: 0.4px; margin: 0 0 6px; font-weight: 700; }
    .kpi-value { font-size: 18px; font-weight: 800; color: #0f172a; margin: 0; }
    .kpi-sub { font-size: 9.5px; color: #94a3b8; margin-top: 4px; }
    
    .section-title { font-size: 13px; font-weight: 700; color: #ffffff; background: #0f172a; padding: 8px 14px; border-radius: 6px; margin: 24px 0 12px; }
    .section-title:first-of-type { margin-top: 0; }
    
    table.data { width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 6px; }
    table.data thead th { background: #1e293b; color: #ffffff; text-align: left; padding: 7px 9px; font-weight: 600; }
    table.data thead th.num { text-align: right; }
    table.data tbody td { padding: 6px 9px; border-bottom: 1px solid #eef2f6; color: #334155; }
    table.data tbody td.num { text-align: right; }
    table.data tbody tr:nth-child(even) { background: #f8fafc; }
    table.data tfoot td { padding: 9px; font-weight: 700; background: #eef2ff; border-top: 2px solid #0ea5e9; color: #0f172a; }
    .empty-row td { text-align: center; color: #94a3b8; padding: 18px; font-style: italic; }
    
    .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 9px; font-weight: 700; }
    .badge.green { background: #dcfce7; color: #15803d; }
    .badge.red { background: #fee2e2; color: #b91c1c; }
    .badge.amber { background: #fef3c7; color: #b45309; }
    .badge.gray { background: #e2e8f0; color: #475569; }
    
    .profit-banner { text-align: center; padding: 18px; border-radius: 12px; margin: 4px 0 26px; }
    .profit-banner.positive { background: linear-gradient(135deg, #dcfce7, #bbf7d0); border: 1px solid #22c55e; }
    .profit-banner.negative { background: linear-gradient(135deg, #fee2e2, #fecaca); border: 1px solid #ef4444; }
    .profit-banner .label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; color: #475569; font-weight: 700; }
    .profit-banner .amount { font-size: 28px; font-weight: 800; margin-top: 6px; }
    .profit-banner.positive .amount { color: #15803d; }
    .profit-banner.negative .amount { color: #b91c1c; }
    
    .report-footer { margin-top: 30px; padding-top: 12px; border-top: 1px solid #e2e8f0; text-align: center; font-size: 9px; color: #94a3b8; }
    .divider-title { font-size: 16px; font-weight: 800; color: #0f172a; text-align: center; margin: 30px 0 4px; }
    .divider-sub { font-size: 10.5px; color: #64748b; text-align: center; margin: 0 0 18px; }
  `;
}

function buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType, metaExtra) {
  const now = new Date();
  const generatedAt = formatBusinessDateTime_(now);
  const safeBusinessName = escapeHtmlReport(businessName);
  const logoHTML = logoBase64
    ? '<img src="data:' + (logoContentType || 'image/png') + ';base64,' + logoBase64 + '" class="report-logo" alt="Logo">'
    : '<div class="report-logo-placeholder">' + escapeHtmlReport((businessName || 'S').charAt(0).toUpperCase()) + '</div>';
  
  return '<div class="report-header">' + logoHTML +
    '<div class="report-header-text">' +
    '<p class="report-business-name">' + safeBusinessName + '</p>' +
    '<p class="report-title">' + escapeHtmlReport(reportTitle) + '</p>' +
    '<p class="report-meta"><span>Generated ' + generatedAt + '</span>' + (metaExtra ? '<span>' + escapeHtmlReport(metaExtra) + '</span>' : '') + '</p>' +
    '</div></div>';
}

function buildReportFooterHTML() {
  return '<div class="report-footer">Generated by SmartTrade Pro &middot; This is an automatically generated report</div>';
}

function buildFiltersBarHTML(filters) {
  if (!filters) return '';
  const parts = [];
  if (filters.branch) parts.push('Branch: <strong>' + escapeHtmlReport(filters.branch) + '</strong>');
  if (filters.dateFrom) parts.push('From: <strong>' + escapeHtmlReport(filters.dateFrom) + '</strong>');
  if (filters.dateTo) parts.push('To: <strong>' + escapeHtmlReport(filters.dateTo) + '</strong>');
  if (filters.category) parts.push('Category: <strong>' + escapeHtmlReport(filters.category) + '</strong>');
  if (filters.status) parts.push('Status: <strong>' + escapeHtmlReport(filters.status) + '</strong>');
  if (filters.payment) parts.push('Payment: <strong>' + escapeHtmlReport(filters.payment) + '</strong>');
  if (filters.product) parts.push('Product: <strong>' + escapeHtmlReport(filters.product) + '</strong>');
  if (filters.supplier) parts.push('Supplier: <strong>' + escapeHtmlReport(filters.supplier) + '</strong>');
  if (parts.length === 0) return '';
  return '<div class="filters-bar">Filters applied: ' + parts.join(' &nbsp;&middot;&nbsp; ') + '</div>';
}

function buildKpiRowHTML(cards) {
  let html = '<table class="kpi-table"><tr>';
  cards.forEach(c => {
    html += '<td class="kpi-cell' + (c.accent ? ' accent-' + c.accent : '') + '">' +
      '<p class="kpi-label">' + escapeHtmlReport(c.label) + '</p>' +
      '<p class="kpi-value">' + escapeHtmlReport(c.value) + '</p>' +
      (c.sub ? '<p class="kpi-sub">' + escapeHtmlReport(c.sub) + '</p>' : '') +
      '</td>';
  });
  html += '</tr></table>';
  return html;
}

function createStockReportHTML(businessName, email, reportTitle, data, filters, logoBase64, logoContentType) {
  const stockItems = (data && data.items) || [];
  let totalStockValue = 0, totalStockQty = 0, lowStockCount = 0, outOfStockCount = 0;
  
  let rows = '';
  stockItems.forEach(item => {
    const qty = Number(item.quantity) || 0;
    totalStockValue += Number(item.value) || 0;
    totalStockQty += qty;
    let stockBadge = '<span class="badge green">In Stock</span>';
    if (qty === 0) { stockBadge = '<span class="badge red">Out of Stock</span>'; outOfStockCount++; }
    else if (qty < 10) { stockBadge = '<span class="badge amber">Low Stock</span>'; lowStockCount++; }
    
    rows += '<tr>' +
      '<td>' + escapeHtmlReport(item.date) + '</td>' +
      '<td>' + escapeHtmlReport(item.product) + '</td>' +
      '<td>' + escapeHtmlReport(item.branch) + '</td>' +
      '<td class="num">' + qty + '</td>' +
      '<td class="num">' + formatReportCurrency(item.buyPrice) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.sellPrice) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.value) + '</td>' +
      '<td>' + escapeHtmlReport(item.category || '-') + '</td>' +
      '<td>' + stockBadge + '</td>' +
      '</tr>';
  });
  
  const kpiRow = buildKpiRowHTML([
    { label: 'Unique Products', value: stockItems.length, accent: 'blue' },
    { label: 'Total Quantity', value: totalStockQty.toLocaleString(), accent: 'blue' },
    { label: 'Total Stock Value', value: formatReportCurrency(totalStockValue), accent: 'green' },
    { label: 'Low / Out of Stock', value: lowStockCount + outOfStockCount, sub: outOfStockCount + ' out of stock', accent: (outOfStockCount > 0 ? 'red' : 'amber') }
  ]);
  
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escapeHtmlReport(businessName) + ' - ' + escapeHtmlReport(reportTitle) + '</title>' +
    '<style>' + getModernReportCSS() + '</style></head><body><div class="page">' +
    buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType) +
    buildFiltersBarHTML(filters) +
    kpiRow +
    '<div class="section-title">Stock Inventory Details</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Product</th><th>Branch</th><th class="num">Qty</th><th class="num">Buy Price</th><th class="num">Sell Price</th><th class="num">Value</th><th>Category</th><th>Status</th></tr></thead>' +
    '<tbody>' + (rows || '<tr class="empty-row"><td colspan="9">No stock items found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="3">TOTAL</td><td class="num">' + totalStockQty.toLocaleString() + '</td><td></td><td></td><td class="num">' + formatReportCurrency(totalStockValue) + '</td><td colspan="2"></td></tr></tfoot>' +
    '</table>' +
    buildReportFooterHTML() +
    '</div></body></html>';
}

// Create Sales Report HTML
function createSalesReportHTML(businessName, email, reportTitle, data, filters, logoBase64, logoContentType) {
  const salesItems = (data && data.items) || [];
  let totalSales = 0, totalQty = 0;
  const paymentTotals = {};
  
  let rows = '';
  salesItems.forEach(item => {
    totalSales += Number(item.total) || 0;
    totalQty += Number(item.quantity) || 0;
    const pm = item.paymentMethod || 'cash';
    paymentTotals[pm] = (paymentTotals[pm] || 0) + (Number(item.total) || 0);
    
    rows += '<tr>' +
      '<td>' + escapeHtmlReport(item.date) + '</td>' +
      '<td>' + escapeHtmlReport(item.product) + '</td>' +
      '<td>' + escapeHtmlReport(item.seller) + '</td>' +
      '<td>' + escapeHtmlReport(item.branch) + '</td>' +
      '<td class="num">' + (item.quantity || 0) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.unitPrice) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.total) + '</td>' +
      '<td>' + escapeHtmlReport(item.customer || 'Walk-in') + '</td>' +
      '<td><span class="badge gray">' + escapeHtmlReport(item.paymentMethod || 'cash') + '</span></td>' +
      '</tr>';
  });
  
  const topPayment = Object.keys(paymentTotals).sort((a, b) => paymentTotals[b] - paymentTotals[a])[0];
  
  const kpiRow = buildKpiRowHTML([
    { label: 'Total Sales', value: formatReportCurrency(totalSales), accent: 'green' },
    { label: 'Transactions', value: salesItems.length },
    { label: 'Units Sold', value: totalQty.toLocaleString() },
    { label: 'Avg. per Sale', value: formatReportCurrency(salesItems.length ? totalSales / salesItems.length : 0), sub: topPayment ? 'Top method: ' + topPayment : '' }
  ]);
  
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escapeHtmlReport(businessName) + ' - ' + escapeHtmlReport(reportTitle) + '</title>' +
    '<style>' + getModernReportCSS() + '</style></head><body><div class="page">' +
    buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType) +
    buildFiltersBarHTML(filters) +
    kpiRow +
    '<div class="section-title">Sales Transactions</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Product</th><th>Seller</th><th>Branch</th><th class="num">Qty</th><th class="num">Unit Price</th><th class="num">Total</th><th>Customer</th><th>Payment</th></tr></thead>' +
    '<tbody>' + (rows || '<tr class="empty-row"><td colspan="9">No sales found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="4">TOTAL</td><td class="num">' + totalQty.toLocaleString() + '</td><td></td><td class="num">' + formatReportCurrency(totalSales) + '</td><td colspan="2"></td></tr></tfoot>' +
    '</table>' +
    buildReportFooterHTML() +
    '</div></body></html>';
}

// Create Purchases Report HTML
function createPurchasesReportHTML(businessName, email, reportTitle, data, filters, logoBase64, logoContentType) {
  const purchasesItems = (data && data.items) || [];
  let totalPurchases = 0, totalQty = 0;
  
  let rows = '';
  purchasesItems.forEach(item => {
    totalPurchases += Number(item.total) || 0;
    totalQty += Number(item.quantity) || 0;
    
    rows += '<tr>' +
      '<td>' + escapeHtmlReport(item.date) + '</td>' +
      '<td>' + escapeHtmlReport(item.product) + '</td>' +
      '<td>' + escapeHtmlReport(item.seller) + '</td>' +
      '<td>' + escapeHtmlReport(item.branch) + '</td>' +
      '<td class="num">' + (item.quantity || 0) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.unitCost) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.total) + '</td>' +
      '<td>' + escapeHtmlReport(item.supplier || '-') + '</td>' +
      '<td><span class="badge gray">' + escapeHtmlReport(item.paymentMethod || '-') + '</span></td>' +
      '</tr>';
  });
  
  const kpiRow = buildKpiRowHTML([
    { label: 'Total Purchases', value: formatReportCurrency(totalPurchases), accent: 'amber' },
    { label: 'Transactions', value: purchasesItems.length },
    { label: 'Units Bought', value: totalQty.toLocaleString() },
    { label: 'Avg. per Purchase', value: formatReportCurrency(purchasesItems.length ? totalPurchases / purchasesItems.length : 0) }
  ]);
  
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escapeHtmlReport(businessName) + ' - ' + escapeHtmlReport(reportTitle) + '</title>' +
    '<style>' + getModernReportCSS() + '</style></head><body><div class="page">' +
    buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType) +
    buildFiltersBarHTML(filters) +
    kpiRow +
    '<div class="section-title">Purchase Transactions</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Product</th><th>Seller</th><th>Branch</th><th class="num">Qty</th><th class="num">Unit Cost</th><th class="num">Total</th><th>Supplier</th><th>Payment</th></tr></thead>' +
    '<tbody>' + (rows || '<tr class="empty-row"><td colspan="9">No purchases found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="4">TOTAL</td><td class="num">' + totalQty.toLocaleString() + '</td><td></td><td class="num">' + formatReportCurrency(totalPurchases) + '</td><td colspan="2"></td></tr></tfoot>' +
    '</table>' +
    buildReportFooterHTML() +
    '</div></body></html>';
}

// Create Debts Report HTML
function createDebtsReportHTML(businessName, email, reportTitle, data, filters, logoBase64, logoContentType) {
  const debtsItems = (data && data.items) || [];
  let totalDebts = 0, totalPaid = 0, totalPending = 0;
  
  let rows = '';
  debtsItems.forEach(item => {
    const balance = (Number(item.totalAmount) || 0) - (Number(item.paidAmount) || 0);
    totalDebts += Number(item.totalAmount) || 0;
    totalPaid += Number(item.paidAmount) || 0;
    
    const status = (item.status || 'pending').toLowerCase();
    let badgeClass = 'amber', badgeText = 'Pending';
    if (status === 'paid') { badgeClass = 'green'; badgeText = 'Paid'; }
    else if (status === 'denied') { badgeClass = 'red'; badgeText = 'Denied'; }
    else { totalPending += balance; }
    
    rows += '<tr>' +
      '<td>' + escapeHtmlReport(item.date) + '</td>' +
      '<td>' + escapeHtmlReport(item.borrowerName) + '</td>' +
      '<td>' + escapeHtmlReport(item.borrowerPhone || '-') + '</td>' +
      '<td>' + escapeHtmlReport(item.product) + '</td>' +
      '<td class="num">' + (item.quantity || 0) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.totalAmount) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.paidAmount) + '</td>' +
      '<td class="num">' + formatReportCurrency(balance) + '</td>' +
      '<td><span class="badge ' + badgeClass + '">' + badgeText + '</span></td>' +
      '</tr>';
  });
  
  const kpiRow = buildKpiRowHTML([
    { label: 'Total Debts', value: formatReportCurrency(totalDebts) },
    { label: 'Paid Amount', value: formatReportCurrency(totalPaid), accent: 'green' },
    { label: 'Pending Amount', value: formatReportCurrency(totalPending), accent: (totalPending > 0 ? 'amber' : 'green') },
    { label: 'Debt Records', value: debtsItems.length }
  ]);
  
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escapeHtmlReport(businessName) + ' - ' + escapeHtmlReport(reportTitle) + '</title>' +
    '<style>' + getModernReportCSS() + '</style></head><body><div class="page">' +
    buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType) +
    buildFiltersBarHTML(filters) +
    kpiRow +
    '<div class="section-title">Debt Records</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Borrower</th><th>Phone</th><th>Product</th><th class="num">Qty</th><th class="num">Total</th><th class="num">Paid</th><th class="num">Balance</th><th>Status</th></tr></thead>' +
    '<tbody>' + (rows || '<tr class="empty-row"><td colspan="9">No debts found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="5">TOTAL</td><td class="num">' + formatReportCurrency(totalDebts) + '</td><td class="num">' + formatReportCurrency(totalPaid) + '</td><td class="num">' + formatReportCurrency(totalPending) + '</td><td></td></tr></tfoot>' +
    '</table>' +
    buildReportFooterHTML() +
    '</div></body></html>';
}

// Create Profit Report HTML
function createProfitReportHTML(businessName, email, reportTitle, data, filters, logoBase64, logoContentType) {
  const totalSales = data.totalSales || 0;
  const totalPurchases = data.totalPurchases || 0;
  const profit = data.profit || 0;
  const profitMargin = data.profitMargin || 0;
  const salesCount = data.salesCount || 0;
  const purchasesCount = data.purchasesCount || 0;
  const averageSaleValue = data.averageSaleValue || 0;
  const topProducts = data.topProducts || [];
  const isProfit = profit >= 0;
  
  let topProductsRows = '';
  topProducts.forEach((product, i) => {
    topProductsRows += '<tr><td class="num">' + (i + 1) + '</td><td>' + escapeHtmlReport(product.product) + '</td><td class="num">' + (product.quantity || 0) + '</td><td class="num">' + formatReportCurrency(product.amount) + '</td></tr>';
  });
  
  const kpiRow = buildKpiRowHTML([
    { label: 'Total Sales', value: formatReportCurrency(totalSales), accent: 'green' },
    { label: 'Total Purchases', value: formatReportCurrency(totalPurchases), accent: 'amber' },
    { label: 'Sales Transactions', value: salesCount },
    { label: 'Avg. Sale Value', value: formatReportCurrency(averageSaleValue) }
  ]);
  
  const profitBanner = '<div class="profit-banner ' + (isProfit ? 'positive' : 'negative') + '">' +
    '<div class="label">' + (isProfit ? 'Net Profit' : 'Net Loss') + '</div>' +
    '<div class="amount">' + (isProfit ? '' : '-') + formatReportCurrency(Math.abs(profit)) + '</div>' +
    '<div class="kpi-sub" style="margin-top:6px;">Profit margin: ' + Number(profitMargin).toFixed(1) + '%</div>' +
    '</div>';
  
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escapeHtmlReport(businessName) + ' - ' + escapeHtmlReport(reportTitle) + '</title>' +
    '<style>' + getModernReportCSS() + '</style></head><body><div class="page">' +
    buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType) +
    buildFiltersBarHTML(filters) +
    profitBanner +
    kpiRow +
    '<div class="section-title">Top Selling Products</div>' +
    '<table class="data"><thead><tr><th class="num">Rank</th><th>Product</th><th class="num">Qty Sold</th><th class="num">Total Sales</th></tr></thead>' +
    '<tbody>' + (topProductsRows || '<tr class="empty-row"><td colspan="4">No sales data for this period</td></tr>') + '</tbody></table>' +
    buildReportFooterHTML() +
    '</div></body></html>';
}


function createFullReportHTML(businessName, email, reportTitle, data, filters, logoBase64, logoContentType, logoFileName) {
  const stockData = data.stock || {};
  const salesData = data.sales || {};
  const purchasesData = data.purchases || {};
  const debtsData = data.debts || {};
  
  const stockItems = stockData.items || [];
  const salesItems = salesData.items || [];
  const purchasesItems = purchasesData.items || [];
  const debtsItems = debtsData.items || [];
  
  let totalStockValue = 0;
  stockItems.forEach(i => { totalStockValue += Number(i.value) || 0; });
  
  let totalSales = 0;
  salesItems.forEach(i => { totalSales += Number(i.total) || 0; });
  
  let totalPurchases = 0;
  purchasesItems.forEach(i => { totalPurchases += Number(i.total) || 0; });
  
  let totalDebtsOutstanding = 0;
  debtsItems.forEach(i => {
    const balance = (Number(i.totalAmount) || 0) - (Number(i.paidAmount) || 0);
    if ((i.status || 'pending').toLowerCase() !== 'paid') totalDebtsOutstanding += balance;
  });
  
  const netProfit = totalSales - totalPurchases;
  const isProfit = netProfit >= 0;
  
  // ---- Overview KPIs ----
  const overviewKpi = buildKpiRowHTML([
    { label: 'Total Sales', value: formatReportCurrency(totalSales), accent: 'green' },
    { label: 'Total Purchases', value: formatReportCurrency(totalPurchases), accent: 'amber' },
    { label: 'Stock Value', value: formatReportCurrency(totalStockValue) },
    { label: 'Outstanding Debts', value: formatReportCurrency(totalDebtsOutstanding), accent: (totalDebtsOutstanding > 0 ? 'red' : 'green') }
  ]);
  
  const profitBanner = '<div class="profit-banner ' + (isProfit ? 'positive' : 'negative') + '">' +
    '<div class="label">' + (isProfit ? 'Net Profit' : 'Net Loss') + ' (Sales - Purchases)</div>' +
    '<div class="amount">' + (isProfit ? '' : '-') + formatReportCurrency(Math.abs(netProfit)) + '</div>' +
    '</div>';
  
  // ---- Stock section ----
  let stockRows = '';
  stockItems.forEach(item => {
    const qty = Number(item.quantity) || 0;
    const badge = qty === 0 ? '<span class="badge red">Out</span>' : qty < 10 ? '<span class="badge amber">Low</span>' : '<span class="badge green">OK</span>';
    stockRows += '<tr><td>' + escapeHtmlReport(item.product) + '</td><td>' + escapeHtmlReport(item.branch) + '</td>' +
      '<td class="num">' + qty + '</td><td class="num">' + formatReportCurrency(item.value) + '</td><td>' + escapeHtmlReport(item.category || '-') + '</td><td>' + badge + '</td></tr>';
  });
  
  // ---- Sales section ----
  let salesRows = '';
  salesItems.forEach(item => {
    salesRows += '<tr><td>' + escapeHtmlReport(item.date) + '</td><td>' + escapeHtmlReport(item.product) + '</td>' +
      '<td>' + escapeHtmlReport(item.seller) + '</td><td class="num">' + (item.quantity || 0) + '</td><td class="num">' + formatReportCurrency(item.total) + '</td><td>' + escapeHtmlReport(item.customer || 'Walk-in') + '</td></tr>';
  });
  
  // ---- Purchases section ----
  let purchasesRows = '';
  purchasesItems.forEach(item => {
    purchasesRows += '<tr><td>' + escapeHtmlReport(item.date) + '</td><td>' + escapeHtmlReport(item.product) + '</td>' +
      '<td class="num">' + (item.quantity || 0) + '</td><td class="num">' + formatReportCurrency(item.total) + '</td><td>' + escapeHtmlReport(item.supplier || '-') + '</td></tr>';
  });
  
  // ---- Debts section ----
  let debtsRows = '';
  debtsItems.forEach(item => {
    const balance = (Number(item.totalAmount) || 0) - (Number(item.paidAmount) || 0);
    const status = (item.status || 'pending').toLowerCase();
    const badge = status === 'paid' ? '<span class="badge green">Paid</span>' : status === 'denied' ? '<span class="badge red">Denied</span>' : '<span class="badge amber">Pending</span>';
    debtsRows += '<tr><td>' + escapeHtmlReport(item.date) + '</td><td>' + escapeHtmlReport(item.borrowerName) + '</td>' +
      '<td class="num">' + formatReportCurrency(item.totalAmount) + '</td><td class="num">' + formatReportCurrency(balance) + '</td><td>' + badge + '</td></tr>';
  });
  
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + escapeHtmlReport(businessName) + ' - ' + escapeHtmlReport(reportTitle) + '</title>' +
    '<style>' + getModernReportCSS() + '</style></head><body><div class="page">' +
    buildReportHeaderHTML(businessName, reportTitle, logoBase64, logoContentType) +
    buildFiltersBarHTML(filters) +
    profitBanner +
    overviewKpi +
    
    '<div class="section-title">Stock Inventory (' + stockItems.length + ' products)</div>' +
    '<table class="data"><thead><tr><th>Product</th><th>Branch</th><th class="num">Qty</th><th class="num">Value</th><th>Category</th><th>Status</th></tr></thead>' +
    '<tbody>' + (stockRows || '<tr class="empty-row"><td colspan="6">No stock items found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="3">TOTAL VALUE</td><td class="num">' + formatReportCurrency(totalStockValue) + '</td><td colspan="2"></td></tr></tfoot></table>' +
    
    '<div class="section-title">Sales (' + salesItems.length + ' transactions)</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Product</th><th>Seller</th><th class="num">Qty</th><th class="num">Total</th><th>Customer</th></tr></thead>' +
    '<tbody>' + (salesRows || '<tr class="empty-row"><td colspan="6">No sales found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="4">TOTAL SALES</td><td class="num">' + formatReportCurrency(totalSales) + '</td><td></td></tr></tfoot></table>' +
    
    '<div class="section-title">Purchases (' + purchasesItems.length + ' transactions)</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Product</th><th class="num">Qty</th><th class="num">Total</th><th>Supplier</th></tr></thead>' +
    '<tbody>' + (purchasesRows || '<tr class="empty-row"><td colspan="5">No purchases found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="3">TOTAL PURCHASES</td><td class="num">' + formatReportCurrency(totalPurchases) + '</td><td></td></tr></tfoot></table>' +
    
    '<div class="section-title">Debts (' + debtsItems.length + ' records)</div>' +
    '<table class="data"><thead><tr><th>Date</th><th>Borrower</th><th class="num">Total</th><th class="num">Balance</th><th>Status</th></tr></thead>' +
    '<tbody>' + (debtsRows || '<tr class="empty-row"><td colspan="5">No debts found</td></tr>') + '</tbody>' +
    '<tfoot><tr><td colspan="3">OUTSTANDING</td><td class="num">' + formatReportCurrency(totalDebtsOutstanding) + '</td><td></td></tr></tfoot></table>' +
    
    buildReportFooterHTML() +
    '</div></body></html>';
}




/* ===== REPORT HTML GENERATION FUNCTIONS ===== */

/* ============================================
   MULTI-DEVICE SCAN LINKING ("Connect Device")
   Lets a phone/POS/tablet scan products and have them appear straight in
   a Sale/Debt form open on a Computer. Google Apps Script Web Apps have
   no WebSocket support - every call here is a plain HTTP request/response
   - so the Computer side polls this once a second rather than getting a
   true instant push. That's the fastest this platform can genuinely do;
   it feels close to instant for a retail scanning workflow, but it is
   not a true zero-latency connection.
   
   CacheService is used instead of a Sheet: it's built for exactly this -
   small, short-lived, auto-expiring data - and avoids adding a permanent
   sheet just to relay a few scanned items.
   ============================================ */

const SCAN_SESSION_TTL_SECONDS = 1800; // 30 minutes - plenty for a work session, then auto-expires

// Computer: generates a short pairing code + starts an empty scan queue
// for it. The Computer encodes this code into a QR (as a full URL) for a
// phone/tablet to scan, or shows it as plain text for a POS device to
// type in.
function createDeviceLinkSession(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const code = Utilities.getUuid().replace(/-/g, "").substring(0, 8).toUpperCase();
    const cache = CacheService.getScriptCache();
    
    const session = {
      databaseId: d.databaseId,
      branchId: d.branchId || "MAIN",
      userType: d.userType || "admin",
      sellerId: d.sellerId || null,
      createdAt: new Date().toISOString(),
      linkedDeviceName: null
    };
    
    cache.put("scanlink_" + code, JSON.stringify(session), SCAN_SESSION_TTL_SECONDS);
    cache.put("scanqueue_" + code, JSON.stringify([]), SCAN_SESSION_TTL_SECONDS);
    // Two more queues, alongside the scan queue above:
    // - removalqueue_: ids the phone has asked to be taken back OUT of
    //   whatever cart they ended up in on the Computer (it scanned
    //   something by mistake, or wants to undo one).
    // - confirmedqueue_: ids the Computer has actually applied to its
    //   cart (set in acknowledgeScannedItems below) - this is what lets
    //   the phone's own "Confirm" button show a real reconciliation
    //   ("this scan really made it into the cart on the other device"),
    //   not just "the network request succeeded".
    cache.put("removalqueue_" + code, JSON.stringify([]), SCAN_SESSION_TTL_SECONDS);
    cache.put("confirmedqueue_" + code, JSON.stringify([]), SCAN_SESSION_TTL_SECONDS);
    
    return output({ success: true, code: code, expiresInSeconds: SCAN_SESSION_TTL_SECONDS });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Phone/POS/Tablet: confirms the code is still valid and marks the
// session as linked, so the Computer's "Connect Device" screen can show
// "Connected" instead of just waiting.
function linkDeviceToSession(d) {
  try {
    if (!d.code) {
      return output({ success: false, error: "Code required" });
    }
    
    const cache = CacheService.getScriptCache();
    const raw = cache.get("scanlink_" + d.code);
    
    if (!raw) {
      return output({ success: false, error: "This code has expired or is invalid. Ask for a new one on the Computer." });
    }
    
    const session = JSON.parse(raw);
    session.linkedDeviceName = d.deviceName || "Mobile Device";
    session.linkedAt = new Date().toISOString();
    
    cache.put("scanlink_" + d.code, JSON.stringify(session), SCAN_SESSION_TTL_SECONDS);
    
    // Look up the business name for a friendly confirmation on the
    // phone's screen ("Connected to <Business Name>").
    let businessName = "";
    let stockSnapshot = [];
    let customerList = [];
    try {
      const masterDB = getMasterDatabase();
      const masterSheet = masterDB.getSheetByName("MASTER_USERS");
      const masterData = masterSheet.getDataRange().getValues();
      for (let i = 1; i < masterData.length; i++) {
        if (masterData[i][3] === session.databaseId) {
          businessName = masterData[i][0];
          break;
        }
      }
      
      // Give the scanning device its own snapshot of current stock for
      // this branch, so it can catch "insufficient stock" itself and
      // simply not send that item at all, instead of the Computer having
      // to reject it after the fact. This is a point-in-time snapshot
      // (taken once, at connection time) that the scanning device then
      // counts down locally as it scans - not a live, continuously
      // re-checked source of truth, so very long scanning sessions with
      // stock changing elsewhere at the same time could drift slightly.
      const userDB = SpreadsheetApp.openById(session.databaseId);
      const stockSheet = userDB.getSheetByName("STOCK");
      if (stockSheet) {
        const stockData = stockSheet.getDataRange().getValues();
        const totals = {};
        for (let i = 1; i < stockData.length; i++) {
          if (!stockData[i][0]) continue;
          const rowBranch = stockData[i][8] || "MAIN";
          if (rowBranch !== session.branchId) continue;
          const key = stockData[i][1] + "|||" + (stockData[i][9] || "");
          totals[key] = (totals[key] || 0) + (Number(stockData[i][2]) || 0);
        }
        stockSnapshot = Object.keys(totals).map(function(key) {
          const parts = key.split("|||");
          return [parts[0], parts[1], totals[key]];
        });
      }
      
      // Also give it a lightweight customer list (id + name + phone only
      // - never email/address/notes, which this device has no real need
      // for) so the "Customer name" field on the scanner can genuinely
      // search existing customers instead of only ever creating new ones
      // by typo-prone free text.
      const customersSheet = userDB.getSheetByName("CUSTOMERS");
      if (customersSheet) {
        const customersData = customersSheet.getDataRange().getValues();
        for (let i = 1; i < customersData.length; i++) {
          if (!customersData[i][0]) continue;
          customerList.push([customersData[i][0], customersData[i][1], customersData[i][2] || ""]);
        }
      }
    } catch (e) {
      // Non-critical - the phone can still scan without a stock snapshot;
      // the Computer still validates stock properly on its own side too.
    }
    
    return output({ success: true, businessName: businessName, branchId: session.branchId, stock: stockSnapshot, customers: customerList });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Phone/POS/Tablet: called every time it successfully scans a product
// label - queues that scan for the Computer to pick up on its next poll.
function pushScannedItemToSession(d) {
  try {
    if (!d.code || !d.product) {
      return output({ success: false, error: "Code and product required" });
    }
    
    const cache = CacheService.getScriptCache();
    const linkRaw = cache.get("scanlink_" + d.code);
    if (!linkRaw) {
      return output({ success: false, error: "This session has expired. Ask for a new QR code on the Computer." });
    }
    
    // Guard the read-modify-write on the shared queue with a lock. Fast,
    // back-to-back scans (sent without waiting for each other, so several
    // can be in flight at once) would otherwise race: two requests could
    // both read the same "current" queue, each add their own item, and
    // whichever writes back LAST would silently wipe out the other's
    // addition. The lock makes each push fully apply before the next one
    // reads the queue.
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    let newItemId = "";
    try {
      const queueRaw = cache.get("scanqueue_" + d.code);
      const queue = queueRaw ? JSON.parse(queueRaw) : [];
      
      newItemId = Utilities.getUuid();
      queue.push({
        id: newItemId,
        product: d.product,
        unit: d.unit || "",
        target: d.target === "debt" ? "debt" : "sale",
        customerName: d.customerName || "",
        customerSessionId: d.customerSessionId || "",
        scannedAt: new Date().toISOString()
      });
      
      cache.put("scanqueue_" + d.code, JSON.stringify(queue), SCAN_SESSION_TTL_SECONDS);
    } finally {
      lock.releaseLock();
    }
    
    // The id is returned so the phone can track this exact scan (for its
    // own "Sent this session" list, the remove-this-item feature, and the
    // Confirm button's reconciliation check) - all of which need to refer
    // to the SAME id the Computer will eventually see and acknowledge,
    // not a separately-generated one that would never match.
    return output({ success: true, itemId: newItemId });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Computer: polled roughly once a second while a Sale/Debt form is open
// with a device linked - returns whatever scans are currently queued.
// IMPORTANT: this is read-ONLY and never deletes anything itself anymore
// - see acknowledgeScannedItems() below for why. A plain "read then
// delete" here used to mean that if this response was ever lost in
// transit after the server had already cleared the queue (a timeout, a
// dropped connection, anything), those scanned items were gone forever
// even though they had scanned successfully - exactly the kind of
// silent, hard-to-reproduce data loss this whole feature must never have.
function pollScannedItems(d) {
  try {
    if (!d.code) {
      return output({ success: false, error: "Code required" });
    }
    
    const cache = CacheService.getScriptCache();
    const linkRaw = cache.get("scanlink_" + d.code);
    if (!linkRaw) {
      return output({ success: false, error: "expired", expired: true });
    }
    
    const session = JSON.parse(linkRaw);
    const queueRaw = cache.get("scanqueue_" + d.code);
    const queue = queueRaw ? JSON.parse(queueRaw) : [];
    
    // Removals the phone has requested since the last time the Computer
    // acknowledged them - items that were already applied to a cart on
    // this poll's previous round-trips and now need to be taken back out.
    const removalRaw = cache.get("removalqueue_" + d.code);
    const removals = removalRaw ? JSON.parse(removalRaw) : [];
    
    return output({
      success: true,
      items: queue,
      removals: removals,
      linkedDeviceName: session.linkedDeviceName
    });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Computer: called right after it has safely added one or more polled
// items to the cart - removes exactly those items (matched by id) from
// the shared queue, so they are never delivered again. Guarded by the
// same lock as pushScannedItemToSession(), so a scan pushed in the
// narrow window between a poll and this acknowledgement is never
// accidentally wiped out too.
//
// This "poll now, delete later once confirmed" two-step is what makes
// delivery reliable: if the acknowledgement call itself never reaches
// the server for any reason, the item simply stays in the queue and is
// delivered again on the very next poll - it can be delayed, but it can
// never quietly disappear. (The Computer keeps a short-lived local
// record of ids it has already acted on, so a re-delivered item is never
// added to the cart a second time either.)
function acknowledgeScannedItems(d) {
  try {
    if (!d.code || !Array.isArray(d.ids) || d.ids.length === 0) {
      return output({ success: false, error: "Code and ids required" });
    }
    
    const cache = CacheService.getScriptCache();
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const queueRaw = cache.get("scanqueue_" + d.code);
      const queue = queueRaw ? JSON.parse(queueRaw) : [];
      const idsToRemove = {};
      d.ids.forEach(function(id) { idsToRemove[id] = true; });
      
      const remaining = queue.filter(function(item) { return !idsToRemove[item.id]; });
      cache.put("scanqueue_" + d.code, JSON.stringify(remaining), SCAN_SESSION_TTL_SECONDS);
      
      // Record these ids as genuinely applied to the Computer's cart -
      // this is what the phone's "Confirm" button checks against, so it
      // can tell the difference between "sent" and "actually made it
      // into the cart on the other device", not just that the network
      // request itself succeeded.
      const confirmedRaw = cache.get("confirmedqueue_" + d.code);
      const confirmed = confirmedRaw ? JSON.parse(confirmedRaw) : [];
      const confirmedSet = {};
      confirmed.forEach(function(id) { confirmedSet[id] = true; });
      d.ids.forEach(function(id) { confirmedSet[id] = true; });
      cache.put("confirmedqueue_" + d.code, JSON.stringify(Object.keys(confirmedSet)), SCAN_SESSION_TTL_SECONDS);
    } finally {
      lock.releaseLock();
    }
    
    return output({ success: true });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Phone: called when the person removes a scanned item from their own
// "Sent this session" list - queues that item's id for the Computer to
// take back out of its cart too (if it's already there) on its next
// poll. This does NOT touch the scan queue itself (pushScannedItemToSession/
// acknowledgeScannedItems) - it's a separate, one-way "please undo this"
// signal, guarded by the same lock pattern as the other queue writes so
// a removal requested in the same instant as a fresh scan is never lost.
function removeScannedItemFromSession(d) {
  try {
    if (!d.code || !d.id) {
      return output({ success: false, error: "Code and id required" });
    }
    
    const cache = CacheService.getScriptCache();
    const linkRaw = cache.get("scanlink_" + d.code);
    if (!linkRaw) {
      return output({ success: false, error: "This session has expired." });
    }
    
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const removalRaw = cache.get("removalqueue_" + d.code);
      const removals = removalRaw ? JSON.parse(removalRaw) : [];
      if (removals.indexOf(d.id) === -1) {
        removals.push(d.id);
      }
      cache.put("removalqueue_" + d.code, JSON.stringify(removals), SCAN_SESSION_TTL_SECONDS);
    } finally {
      lock.releaseLock();
    }
    
    return output({ success: true });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Computer: called right after it has removed one or more items (matched
// by their original scan id) from its cart in response to a removal
// request - clears exactly those ids from the removal queue, same
// "acknowledge only what was actually applied" pattern as
// acknowledgeScannedItems above, for the same reason: if this call itself
// never reaches the server, the removal request simply stays queued and
// is retried on the next poll instead of being silently lost.
function acknowledgeRemovals(d) {
  try {
    if (!d.code || !Array.isArray(d.ids) || d.ids.length === 0) {
      return output({ success: false, error: "Code and ids required" });
    }
    
    const cache = CacheService.getScriptCache();
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const removalRaw = cache.get("removalqueue_" + d.code);
      const removals = removalRaw ? JSON.parse(removalRaw) : [];
      const idsToRemove = {};
      d.ids.forEach(function(id) { idsToRemove[id] = true; });
      const remaining = removals.filter(function(id) { return !idsToRemove[id]; });
      cache.put("removalqueue_" + d.code, JSON.stringify(remaining), SCAN_SESSION_TTL_SECONDS);
    } finally {
      lock.releaseLock();
    }
    
    return output({ success: true });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Phone: called by the "Confirm" button's reconciliation check - returns
// every scan id the Computer has actually applied to its cart so far
// (see acknowledgeScannedItems), so the phone can compare that against
// its own "Sent this session" list and show exactly which scans have
// genuinely landed on the other device versus which are still only
// "sent" (network-acknowledged) but not yet confirmed applied.
function getConfirmedScanIds(d) {
  try {
    if (!d.code) {
      return output({ success: false, error: "Code required" });
    }
    const cache = CacheService.getScriptCache();
    const linkRaw = cache.get("scanlink_" + d.code);
    if (!linkRaw) {
      return output({ success: false, error: "expired", expired: true });
    }
    const confirmedRaw = cache.get("confirmedqueue_" + d.code);
    const confirmed = confirmedRaw ? JSON.parse(confirmedRaw) : [];
    return output({ success: true, confirmedIds: confirmed });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Either side: explicitly ends the pairing (Computer closing the form,
// or the phone disconnecting) so the code can't be reused even before
// its 30-minute TTL would have expired it anyway.
function endDeviceLinkSession(d) {
  try {
    if (!d.code) {
      return output({ success: false, error: "Code required" });
    }
    const cache = CacheService.getScriptCache();
    cache.remove("scanlink_" + d.code);
    cache.remove("scanqueue_" + d.code);
    cache.remove("removalqueue_" + d.code);
    cache.remove("confirmedqueue_" + d.code);
    return output({ success: true });
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

function generateStockReportContentHTML(data, formatCurrency) {
  const items = data.items || [];
  const totalValue = data.totalValue || 0;
  const totalItems = data.totalItems || 0;
  
  let tableRows = '';
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    tableRows += '<tr>' +
      '<td>' + (item.date || '—') + '</td>' +
      '<td>' + (item.product || '—') + '</td>' +
      '<td>' + (item.branch || '—') + '</td>' +
      '<td>' + (item.quantity || 0) + '</td>' +
      '<td>' + formatCurrency(item.buyPrice || 0) + '</td>' +
      '<td>' + formatCurrency(item.sellPrice || 0) + '</td>' +
      '<td>' + formatCurrency(item.value || 0) + '</td>' +
      '<td>' + (item.category || '—') + '</td>' +
      '</tr>';
  }
  
  return '<div class="summary-cards">' +
    '<div class="card"><h3>Total Items</h3><div class="value">' + totalItems + '</div></div>' +
    '<div class="card"><h3>Unique Products</h3><div class="value">' + items.length + '</div></div>' +
    '<div class="card"><h3>Total Value</h3><div class="value">' + formatCurrency(totalValue) + '</div></div>' +
    '</div>' +
    '<div class="section"><h3>Stock Items (' + items.length + ')</h3>' +
    '<table><thead><tr><th>Date</th><th>Product</th><th>Branch</th><th>Quantity</th><th>Buy Price</th><th>Sell Price</th><th>Value</th><th>Category</th></tr></thead>' +
    '<tbody>' + (tableRows || '<tr><td colspan="8" style="text-align: center;">No items found</td></tr>') +
    '<tr class="total-row"><td colspan="3"><strong>Total</strong></td><td><strong>' + totalItems + '</strong></td><td colspan="2"></td><td><strong>' + formatCurrency(totalValue) + '</strong></td><td></td></tr>' +
    '</tbody></table></div>';
}

function generateSalesReportContentHTML(data, formatCurrency) {
  const items = data.items || [];
  const totalAmount = data.totalAmount || 0;
  
  let tableRows = '';
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    tableRows += '<tr>' +
      '<td>' + (item.date || '—') + '</td>' +
      '<td>' + (item.product || '—') + '</td>' +
      '<td>' + (item.seller || '—') + '</td>' +
      '<td>' + (item.branch || '—') + '</td>' +
      '<td>' + (item.quantity || 0) + '</td>' +
      '<td>' + formatCurrency(item.unitPrice || 0) + '</td>' +
      '<td>' + formatCurrency(item.total || 0) + '</td>' +
      '<td>' + (item.customer || '—') + '</td>' +
      '<td>' + (item.paymentMethod || '—') + '</td>' +
      '</tr>';
  }
  
  return '<div class="summary-cards">' +
    '<div class="card"><h3>Total Sales</h3><div class="value">' + formatCurrency(totalAmount) + '</div></div>' +
    '<div class="card"><h3>Transactions</h3><div class="value">' + items.length + '</div></div>' +
    '<div class="card"><h3>Average per Sale</h3><div class="value">' + (items.length > 0 ? formatCurrency(totalAmount / items.length) : '0') + '</div></div>' +
    '</div>' +
    '<div class="section"><h3>Sales Transactions (' + items.length + ')</h3>' +
    '<table><thead><tr><th>Date</th><th>Product</th><th>Seller</th><th>Branch</th><th>Qty</th><th>Unit Price</th><th>Total</th><th>Customer</th><th>Payment</th></tr></thead>' +
    '<tbody>' + (tableRows || '<tr><td colspan="9" style="text-align: center;">No transactions found</td></tr>') +
    '<tr class="total-row"><td colspan="6"><strong>Total</strong></td><td><strong>' + formatCurrency(totalAmount) + '</strong></td><td colspan="2"></td></tr>' +
    '</tbody></table></div>';
}

function generatePurchasesReportContentHTML(data, formatCurrency) {
  const items = data.items || [];
  const totalAmount = data.totalAmount || 0;
  
  let tableRows = '';
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    tableRows += '<tr>' +
      '<td>' + (item.date || '—') + '</td>' +
      '<td>' + (item.product || '—') + '</td>' +
      '<td>' + (item.seller || '—') + '</td>' +
      '<td>' + (item.branch || '—') + '</td>' +
      '<td>' + (item.quantity || 0) + '</td>' +
      '<td>' + formatCurrency(item.unitCost || 0) + '</td>' +
      '<td>' + formatCurrency(item.total || 0) + '</td>' +
      '<td>' + (item.supplier || '—') + '</td>' +
      '<td>' + (item.paymentMethod || '—') + '</td>' +
      '</tr>';
  }
  
  return '<div class="summary-cards">' +
    '<div class="card"><h3>Total Purchases</h3><div class="value">' + formatCurrency(totalAmount) + '</div></div>' +
    '<div class="card"><h3>Transactions</h3><div class="value">' + items.length + '</div></div>' +
    '<div class="card"><h3>Average per Purchase</h3><div class="value">' + (items.length > 0 ? formatCurrency(totalAmount / items.length) : '0') + '</div></div>' +
    '</div>' +
    '<div class="section"><h3>Purchase Transactions (' + items.length + ')</h3>' +
    '<table><thead><tr><th>Date</th><th>Product</th><th>Seller</th><th>Branch</th><th>Qty</th><th>Unit Cost</th><th>Total</th><th>Supplier</th><th>Payment</th></tr></thead>' +
    '<tbody>' + (tableRows || '<tr><td colspan="9" style="text-align: center;">No transactions found</td></tr>') +
    '<tr class="total-row"><td colspan="6"><strong>Total</strong></td><td><strong>' + formatCurrency(totalAmount) + '</strong></td><td colspan="2"></td></tr>' +
    '</tbody></table></div>';
}

function generateDebtsReportContentHTML(data, formatCurrency) {
  const items = data.items || [];
  const totalAmount = data.totalAmount || 0;
  const paidAmount = data.paidAmount || 0;
  const pendingAmount = data.pendingAmount || 0;
  
  let tableRows = '';
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const statusClass = item.status === 'paid' ? 'status-paid' : (item.status === 'denied' ? 'status-denied' : 'status-pending');
    tableRows += '<tr>' +
      '<td>' + (item.date || '—') + '</td>' +
      '<td>' + (item.borrowerName || '—') + '</td>' +
      '<td>' + (item.borrowerPhone || '—') + '</td>' +
      '<td>' + (item.product || '—') + '</td>' +
      '<td>' + (item.seller || '—') + '</td>' +
      '<td>' + (item.branch || '—') + '</td>' +
      '<td>' + (item.quantity || 0) + '</td>' +
      '<td>' + formatCurrency(item.totalAmount || 0) + '</td>' +
      '<td>' + formatCurrency(item.paidAmount || 0) + '</td>' +
      '<td>' + formatCurrency(item.balance || 0) + '</td>' +
      '<td class="' + statusClass + '">' + (item.status || 'PENDING').toUpperCase() + '</td>' +
      '</tr>';
  }
  
  return '<div class="summary-cards">' +
    '<div class="card"><h3>Total Debts</h3><div class="value">' + formatCurrency(totalAmount) + '</div></div>' +
    '<div class="card"><h3>Paid Amount</h3><div class="value">' + formatCurrency(paidAmount) + '</div></div>' +
    '<div class="card"><h3>Pending Amount</h3><div class="value">' + formatCurrency(pendingAmount) + '</div></div>' +
    '<div class="card"><h3>Records</h3><div class="value">' + items.length + '</div></div>' +
    '</div>' +
    '<div class="section"><h3>Debt Records (' + items.length + ')</h3>' +
    '<table><thead><tr><th>Date</th><th>Borrower</th><th>Phone</th><th>Product</th><th>Seller</th><th>Branch</th><th>Qty</th><th>Total</th><th>Paid</th><th>Balance</th><th>Status</th></tr></thead>' +
    '<tbody>' + (tableRows || '<tr><td colspan="11" style="text-align: center;">No records found</td></tr>') +
    '<tr class="total-row"><td colspan="7"><strong>Totals</strong></td><td><strong>' + formatCurrency(totalAmount) + '</strong></td><td><strong>' + formatCurrency(paidAmount) + '</strong></td><td><strong>' + formatCurrency(pendingAmount) + '</strong></td><td></td></tr>' +
    '</tbody></table></div>';
}

function generateProfitReportContentHTML(data, formatCurrency) {
  const totalSales = data.totalSales || 0;
  const totalPurchases = data.totalPurchases || 0;
  const profit = data.profit || 0;
  const profitMargin = data.profitMargin || 0;
  const salesCount = data.salesCount || 0;
  const purchasesCount = data.purchasesCount || 0;
  const averageSaleValue = data.averageSaleValue || 0;
  const topProducts = data.topProducts || [];
  
  let topProductsHtml = '';
  for (let i = 0; i < topProducts.length; i++) {
    const product = topProducts[i];
    topProductsHtml += '<tr><td>' + (i + 1) + '</td><td>' + (product.product || '—') + '</td><td>' + (product.quantity || 0) + '</td><td>' + formatCurrency(product.amount || 0) + '</td></tr>';
  }
  
  let salesRows = '';
  if (data.items && data.items.length > 0) {
    const itemsToShow = data.items.slice(0, 50);
    for (let i = 0; i < itemsToShow.length; i++) {
      const item = itemsToShow[i];
      salesRows += '<tr>' +
        '<td>' + (item.date || '—') + '</td>' +
        '<td>' + (item.product || '—') + '</td>' +
        '<td>' + (item.seller || '—') + '</td>' +
        '<td>' + (item.branch || '—') + '</td>' +
        '<td>' + (item.quantity || 0) + '</td>' +
        '<td>' + formatCurrency(item.unitPrice || 0) + '</td>' +
        '<td>' + formatCurrency(item.total || 0) + '</td>' +
        '<td>' + (item.customer || '—') + '</td>' +
        '</tr>';
    }
  }
  
  return '<div class="summary-cards">' +
    '<div class="card"><h3>Total Sales</h3><div class="value">' + formatCurrency(totalSales) + '</div></div>' +
    '<div class="card"><h3>Total Purchases</h3><div class="value">' + formatCurrency(totalPurchases) + '</div></div>' +
    '<div class="card"><h3>Gross Profit</h3><div class="value">' + formatCurrency(profit) + '</div></div>' +
    '<div class="card"><h3>Profit Margin</h3><div class="value">' + profitMargin.toFixed(2) + '%</div></div>' +
    '</div>' +
    '<div class="summary-cards">' +
    '<div class="card"><h3>Sales Transactions</h3><div class="value">' + salesCount + '</div></div>' +
    '<div class="card"><h3>Purchase Transactions</h3><div class="value">' + purchasesCount + '</div></div>' +
    '<div class="card"><h3>Average Sale Value</h3><div class="value">' + formatCurrency(averageSaleValue) + '</div></div>' +
    '</div>' +
    '<div class="section"><h3>Top 5 Selling Products</h3>' +
    '<table><thead><tr><th>Rank</th><th>Product</th><th>Quantity Sold</th><th>Total Sales</th></tr></thead>' +
    '<tbody>' + (topProductsHtml || '<tr><td colspan="4" style="text-align: center;">No sales data</td></tr>') + '</tbody></table></div>' +
    '<div class="section"><h3>Sales Transactions (' + salesCount + ')</h3>' +
    '<table><thead><tr><th>Date</th><th>Product</th><th>Seller</th><th>Branch</th><th>Qty</th><th>Unit Price</th><th>Total</th><th>Customer</th></tr></thead>' +
    '<tbody>' + (salesRows || '<tr><td colspan="8" style="text-align: center;">No sales found</td></tr>') + '</tbody></table></div>';
}

function generateFullReportContentHTML(data, formatCurrency) {
  return generateStockReportContentHTML(data.stock || { items: [] }, formatCurrency) +
         generateSalesReportContentHTML(data.sales || { items: [] }, formatCurrency) +
         generatePurchasesReportContentHTML(data.purchases || { items: [] }, formatCurrency) +
         generateDebtsReportContentHTML(data.debts || { items: [] }, formatCurrency);
}

function generateStockReportContent(data, formatCurrency, filterText) {
  const items = data.items || [];
  const totalValue = data.totalValue || 0;
  const totalItems = data.totalItems || 0;
  
  let tableRows = '';
  items.forEach(item => {
    tableRows += `
      <tr>
        <td>${item.date}</td>
        <td>${item.product}</td>
        <td>${item.branch}</td>
        <td>${item.quantity}</td>
        <td>${formatCurrency(item.buyPrice)}</td>
        <td>${formatCurrency(item.sellPrice)}</td>
        <td>${formatCurrency(item.value)}</td>
        <td>${item.category}</td>
      </tr>
    `;
  });
  
  return `
    <div class="summary-cards">
      <div class="card">
        <h3>Total Items</h3>
        <div class="value">${totalItems}</div>
      </div>
      <div class="card">
        <h3>Unique Products</h3>
        <div class="value">${items.length}</div>
      </div>
      <div class="card">
        <h3>Total Value</h3>
        <div class="value">${formatCurrency(totalValue)}</div>
      </div>
    </div>
    
    <div class="section">
      <h3>Stock Items (${items.length})</h3>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Branch</th>
            <th>Quantity</th>
            <th>Buy Price</th>
            <th>Sell Price</th>
            <th>Value</th>
            <th>Category</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows || '<tr><td colspan="8" style="text-align: center;">No items found</td></tr>'}
          <tr class="total-row">
            <td colspan="3"><strong>Total</strong></td>
            <td><strong>${totalItems}</strong></td>
            <td colspan="2"></td>
            <td><strong>${formatCurrency(totalValue)}</strong></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

function generateSalesReportContent(data, formatCurrency, filterText) {
  const items = data.items || [];
  const totalAmount = data.totalAmount || 0;
  
  let tableRows = '';
  items.forEach(item => {
    tableRows += `
      <tr>
        <td>${item.date}</td>
        <td>${item.product}</td>
        <td>${item.seller}</td>
        <td>${item.branch}</td>
        <td>${item.quantity}</td>
        <td>${formatCurrency(item.unitPrice)}</td>
        <td>${formatCurrency(item.total)}</td>
        <td>${item.customer}</td>
        <td>${item.paymentMethod}</td>
      </tr>
    `;
  });
  
  return `
    <div class="summary-cards">
      <div class="card">
        <h3>Total Sales</h3>
        <div class="value">${formatCurrency(totalAmount)}</div>
      </div>
      <div class="card">
        <h3>Transactions</h3>
        <div class="value">${items.length}</div>
      </div>
      <div class="card">
        <h3>Average per Sale</h3>
        <div class="value">${items.length > 0 ? formatCurrency(totalAmount / items.length) : '0'}</div>
      </div>
    </div>
    
    <div class="section">
      <h3>Sales Transactions (${items.length})</h3>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Unit Price</th>
            <th>Total</th>
            <th>Customer</th>
            <th>Payment</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows || '<tr><td colspan="9" style="text-align: center;">No transactions found</td></tr>'}
          <tr class="total-row">
            <td colspan="6"><strong>Total</strong></td>
            <td><strong>${formatCurrency(totalAmount)}</strong></td>
            <td colspan="2"></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

function generatePurchasesReportContent(data, formatCurrency, filterText) {
  const items = data.items || [];
  const totalAmount = data.totalAmount || 0;
  
  let tableRows = '';
  items.forEach(item => {
    tableRows += `
      <tr>
        <td>${item.date}</td>
        <td>${item.product}</td>
        <td>${item.seller}</td>
        <td>${item.branch}</td>
        <td>${item.quantity}</td>
        <td>${formatCurrency(item.unitCost)}</td>
        <td>${formatCurrency(item.total)}</td>
        <td>${item.supplier}</td>
        <td>${item.paymentMethod}</td>
      </tr>
    `;
  });
  
  return `
    <div class="summary-cards">
      <div class="card">
        <h3>Total Purchases</h3>
        <div class="value">${formatCurrency(totalAmount)}</div>
      </div>
      <div class="card">
        <h3>Transactions</h3>
        <div class="value">${items.length}</div>
      </div>
      <div class="card">
        <h3>Average per Purchase</h3>
        <div class="value">${items.length > 0 ? formatCurrency(totalAmount / items.length) : '0'}</div>
      </div>
    </div>
    
    <div class="section">
      <h3>Purchase Transactions (${items.length})</h3>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Unit Cost</th>
            <th>Total</th>
            <th>Supplier</th>
            <th>Payment</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows || '<tr><td colspan="9" style="text-align: center;">No transactions found</td></tr>'}
          <tr class="total-row">
            <td colspan="6"><strong>Total</strong></td>
            <td><strong>${formatCurrency(totalAmount)}</strong></td>
            <td colspan="2"></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

function generateDebtsReportContent(data, formatCurrency, filterText) {
  const items = data.items || [];
  const totalAmount = data.totalAmount || 0;
  const paidAmount = data.paidAmount || 0;
  const pendingAmount = data.pendingAmount || 0;
  
  let tableRows = '';
  items.forEach(item => {
    const statusClass = item.status === 'paid' ? 'status-paid' : (item.status === 'denied' ? 'status-denied' : 'status-pending');
    tableRows += `
      <tr>
        <td>${item.date}</td>
        <td>${item.borrowerName}</td>
        <td>${item.borrowerPhone}</td>
        <td>${item.product}</td>
        <td>${item.seller}</td>
        <td>${item.branch}</td>
        <td>${item.quantity}</td>
        <td>${formatCurrency(item.totalAmount)}</td>
        <td>${formatCurrency(item.paidAmount)}</td>
        <td>${formatCurrency(item.balance)}</td>
        <td class="${statusClass}">${item.status.toUpperCase()}</td>
      </tr>
    `;
  });
  
  return `
    <div class="summary-cards">
      <div class="card">
        <h3>Total Debts</h3>
        <div class="value">${formatCurrency(totalAmount)}</div>
      </div>
      <div class="card">
        <h3>Paid Amount</h3>
        <div class="value">${formatCurrency(paidAmount)}</div>
      </div>
      <div class="card">
        <h3>Pending Amount</h3>
        <div class="value">${formatCurrency(pendingAmount)}</div>
      </div>
      <div class="card">
        <h3>Records</h3>
        <div class="value">${items.length}</div>
      </div>
    </div>
    
    <div class="section">
      <h3>Debt Records (${items.length})</h3>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Borrower</th>
            <th>Phone</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Total</th>
            <th>Paid</th>
            <th>Balance</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows || '<tr><td colspan="11" style="text-align: center;">No records found</td></tr>'}
          <tr class="total-row">
            <td colspan="7"><strong>Totals</strong></td>
            <td><strong>${formatCurrency(totalAmount)}</strong></td>
            <td><strong>${formatCurrency(paidAmount)}</strong></td>
            <td><strong>${formatCurrency(pendingAmount)}</strong></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    </div>
  `;
}

function generateProfitReportContent(data, formatCurrency, filterText) {
  const totalSales = data.totalSales || 0;
  const totalPurchases = data.totalPurchases || 0;
  const profit = data.profit || 0;
  const profitMargin = data.profitMargin || 0;
  const salesCount = data.salesCount || 0;
  const purchasesCount = data.purchasesCount || 0;
  const averageSaleValue = data.averageSaleValue || 0;
  const topProducts = data.topProducts || [];
  
  let topProductsHtml = '';
  topProducts.forEach((product, index) => {
    topProductsHtml += `
      <tr>
        <td>${index + 1}</td>
        <td>${product.product}</td>
        <td>${product.quantity}</td>
        <td>${formatCurrency(product.amount)}</td>
      </tr>
    `;
  });
  
  return `
    <div class="summary-cards">
      <div class="card">
        <h3>Total Sales</h3>
        <div class="value">${formatCurrency(totalSales)}</div>
      </div>
      <div class="card">
        <h3>Total Purchases</h3>
        <div class="value">${formatCurrency(totalPurchases)}</div>
      </div>
      <div class="card">
        <h3>Gross Profit</h3>
        <div class="value">${formatCurrency(profit)}</div>
      </div>
      <div class="card">
        <h3>Profit Margin</h3>
        <div class="value">${profitMargin.toFixed(2)}%</div>
      </div>
    </div>
    
    <div class="summary-cards">
      <div class="card">
        <h3>Sales Transactions</h3>
        <div class="value">${salesCount}</div>
      </div>
      <div class="card">
        <h3>Purchase Transactions</h3>
        <div class="value">${purchasesCount}</div>
      </div>
      <div class="card">
        <h3>Average Sale Value</h3>
        <div class="value">${formatCurrency(averageSaleValue)}</div>
      </div>
    </div>
    
    <div class="section">
      <h3>Top 5 Selling Products</h3>
      <table>
        <thead>
          <tr>
            <th>Rank</th>
            <th>Product</th>
            <th>Quantity Sold</th>
            <th>Total Sales</th>
          </tr>
        </thead>
        <tbody>
          ${topProductsHtml || '<tr><td colspan="4" style="text-align: center;">No sales data</td></tr>'}
        </tbody>
      </table>
    </div>
    
    <div class="section">
      <h3>Sales Transactions (${salesCount})</h3>
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Product</th>
            <th>Seller</th>
            <th>Branch</th>
            <th>Qty</th>
            <th>Unit Price</th>
            <th>Total</th>
            <th>Customer</th>
          </tr>
        </thead>
        <tbody>
          ${data.items && data.items.length > 0 ? data.items.slice(0, 50).map(item => `
            <tr>
              <td>${item.date}</td>
              <td>${item.product}</td>
              <td>${item.seller}</td>
              <td>${item.branch}</td>
              <td>${item.quantity}</td>
              <td>${formatCurrency(item.unitPrice)}</td>
              <td>${formatCurrency(item.total)}</td>
              <td>${item.customer}</td>
            </tr>
          `).join('') : '<tr><td colspan="8" style="text-align: center;">No sales found</td></tr>'}
        </tbody>
      </table>
    </div>
  `;
}

function generateFullReportContent(data, formatCurrency, filterText) {
  return `
    ${generateStockReportContent(data.stock || { items: [] }, formatCurrency, '')}
    ${generateSalesReportContent(data.sales || { items: [] }, formatCurrency, '')}
    ${generatePurchasesReportContent(data.purchases || { items: [] }, formatCurrency, '')}
    ${generateDebtsReportContent(data.debts || { items: [] }, formatCurrency, '')}
  `;
}

// ============================================
// SUBSCRIPTION/VALIDITY MANAGEMENT FUNCTIONS (FIXED)
// ============================================

// Update user subscription expiry (with minutes support)
function updateUserSubscription(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    if (!d.targetEmail) {
      return output({ success: false, error: "Target email required" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    
    // Check if subscription sheet exists, create if not
    let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
    if (!subscriptionSheet) {
      subscriptionSheet = masterDB.insertSheet("USER_SUBSCRIPTIONS");
      subscriptionSheet.appendRow(["Email", "ExpiryDate", "Plan", "Status", "LastUpdated"]);
      subscriptionSheet.getRange("A1:E1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    }
    
    const subData = subscriptionSheet.getDataRange().getValues();
    let found = false;
    
    // Parse expiry date properly
    let expiryDateObj = null;
    if (d.expiryDate) {
      expiryDateObj = new Date(d.expiryDate);
      // Validate if date is valid
      if (isNaN(expiryDateObj.getTime())) {
        return output({ success: false, error: "Invalid expiry date format" });
      }
    }
    
    const expiryDateString = expiryDateObj ? expiryDateObj.toISOString() : null;
    
    for (let i = 1; i < subData.length; i++) {
      if (subData[i][0] === d.targetEmail) {
        subscriptionSheet.getRange(i + 1, 2).setValue(expiryDateString);
        subscriptionSheet.getRange(i + 1, 3).setValue(d.plan || "Custom");
        subscriptionSheet.getRange(i + 1, 4).setValue(d.status || "active");
        subscriptionSheet.getRange(i + 1, 5).setValue(new Date());
        found = true;
        break;
      }
    }
    
    if (!found) {
      subscriptionSheet.appendRow([
        d.targetEmail,
        expiryDateString,
        d.plan || "Custom",
        d.status || "active",
        new Date()
      ]);
    }
    
    return output({
      success: true,
      message: expiryDateObj ? `Subscription updated until ${formatBusinessDateTime_(expiryDateObj)}` : "Subscription set to Unlimited",
      expiryDate: expiryDateString,
      plan: d.plan
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Get user subscription info (with accurate expiry check)
function getUserSubscription(d) {
  try {
    if (!d.email) {
      return output({ success: false, error: "Email required" });
    }
    
    const masterDB = getMasterDatabase();
    let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
    
    if (!subscriptionSheet) {
      return output({
        success: true,
        email: d.email,
        expiryDate: null,
        expiryTimestamp: null,
        plan: "unlimited",
        status: "active",
        isExpired: false,
        remainingMinutes: null,
        remainingHours: null,
        remainingDays: null,
        remainingReadable: "Unlimited"
      });
    }
    
    const subData = subscriptionSheet.getDataRange().getValues();
    
    for (let i = 1; i < subData.length; i++) {
      if (subData[i][0] === d.email) {
        const expiryDateRaw = subData[i][1];
        const plan = subData[i][2] || "custom";
        const status = subData[i][3] || "active";
        
        let isExpired = false;
        let remainingMinutes = null;
        let remainingHours = null;
        let remainingDays = null;
        let remainingReadable = "Active";
        let expiryTimestamp = null;
        let expiryDateObj = null;
        
        if (plan === "unlimited") {
          remainingReadable = "Unlimited";
        } else if (expiryDateRaw && status === "active") {
          expiryDateObj = new Date(expiryDateRaw);
          expiryTimestamp = expiryDateObj.getTime();
          const now = new Date();
          const nowTimestamp = now.getTime();
          
          // Check if expired (using milliseconds for accuracy)
          isExpired = expiryTimestamp < nowTimestamp;
          
          if (!isExpired) {
            const diffMs = expiryTimestamp - nowTimestamp;
            remainingMinutes = Math.floor(diffMs / (1000 * 60));
            remainingHours = Math.floor(diffMs / (1000 * 60 * 60));
            remainingDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
            
            // Build readable string
            if (remainingDays > 0) {
              const hoursRemaining = remainingHours % 24;
              if (hoursRemaining > 0) {
                remainingReadable = `${remainingDays}d ${hoursRemaining}h`;
              } else {
                remainingReadable = `${remainingDays} day${remainingDays > 1 ? 's' : ''}`;
              }
            } else if (remainingHours > 0) {
              const minutesRemaining = remainingMinutes % 60;
              if (minutesRemaining > 0) {
                remainingReadable = `${remainingHours}h ${minutesRemaining}m`;
              } else {
                remainingReadable = `${remainingHours} hour${remainingHours > 1 ? 's' : ''}`;
              }
            } else if (remainingMinutes > 0) {
              remainingReadable = `${remainingMinutes} minute${remainingMinutes > 1 ? 's' : ''}`;
            } else {
              remainingReadable = "Less than a minute";
            }
          } else {
            remainingReadable = "Expired";
          }
        }
        
        return output({
          success: true,
          email: d.email,
          expiryDate: expiryDateRaw,
          expiryTimestamp: expiryTimestamp,
          expiryDateFormatted: expiryDateObj ? formatBusinessDateTime_(expiryDateObj) : null,
          plan: plan,
          status: status,
          isExpired: isExpired,
          remainingMinutes: remainingMinutes,
          remainingHours: remainingHours,
          remainingDays: remainingDays,
          remainingReadable: remainingReadable
        });
      }
    }
    
    // No subscription record found - return unlimited
    return output({
      success: true,
      email: d.email,
      expiryDate: null,
      expiryTimestamp: null,
      plan: "unlimited",
      status: "active",
      isExpired: false,
      remainingMinutes: null,
      remainingHours: null,
      remainingDays: null,
      remainingReadable: "Unlimited"
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Get all users with subscription info (accurate)
function getAllUsersWithSubscription(d) {
  try {
    // Verify developer email
    if (d.email !== DEVELOPER_EMAIL) {
      return output({ success: false, error: "Unauthorized access" });
    }
    
    const masterDB = getMasterDatabase();
    const masterSheet = masterDB.getSheetByName("MASTER_USERS");
    const data = masterSheet.getDataRange().getValues();
    
    // Get subscription sheet
    let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
    const subscriptionMap = {};
    
    if (subscriptionSheet) {
      const subData = subscriptionSheet.getDataRange().getValues();
      for (let i = 1; i < subData.length; i++) {
        if (subData[i][0]) {
          subscriptionMap[subData[i][0]] = {
            expiryDate: subData[i][1],
            plan: subData[i][2] || "custom",
            status: subData[i][3] || "active",
            lastUpdated: subData[i][4]
          };
        }
      }
    }
    
    const users = [];
    const now = new Date();
    const nowTimestamp = now.getTime();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) {
        const email = data[i][1];
        const subscription = subscriptionMap[email] || { plan: "unlimited", status: "active", expiryDate: null };
        
        let isExpired = false;
        let remainingMinutes = null;
        let remainingHours = null;
        let remainingDays = null;
        let remainingReadable = "Active";
        let expiryTimestamp = null;
        let expiryDateFormatted = null;
        
        if (subscription.plan === "unlimited") {
          remainingReadable = "Unlimited";
        } else if (subscription.expiryDate && subscription.status === "active") {
          const expiryDateObj = new Date(subscription.expiryDate);
          expiryTimestamp = expiryDateObj.getTime();
          expiryDateFormatted = formatBusinessDateTime_(expiryDateObj);
          
          // Check if expired using milliseconds
          isExpired = expiryTimestamp < nowTimestamp;
          
          if (!isExpired) {
            const diffMs = expiryTimestamp - nowTimestamp;
            remainingMinutes = Math.floor(diffMs / (1000 * 60));
            remainingHours = Math.floor(diffMs / (1000 * 60 * 60));
            remainingDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
            
            // Build readable string
            if (remainingDays > 0) {
              const hoursRemaining = remainingHours % 24;
              if (hoursRemaining > 0) {
                remainingReadable = `${remainingDays}d ${hoursRemaining}h`;
              } else {
                remainingReadable = `${remainingDays} day${remainingDays > 1 ? 's' : ''}`;
              }
            } else if (remainingHours > 0) {
              const minutesRemaining = remainingMinutes % 60;
              if (minutesRemaining > 0) {
                remainingReadable = `${remainingHours}h ${minutesRemaining}m`;
              } else {
                remainingReadable = `${remainingHours} hour${remainingHours > 1 ? 's' : ''}`;
              }
            } else if (remainingMinutes > 0) {
              remainingReadable = `${remainingMinutes} minute${remainingMinutes > 1 ? 's' : ''}`;
            } else {
              remainingReadable = "Less than a minute";
            }
          } else {
            remainingReadable = "Expired";
          }
        }
        
        users.push({
          business: data[i][0] || "",
          email: email,
          userType: data[i][7] || "admin",
          created: data[i][4] ? data[i][4].toString() : null,
          lastLogin: data[i][5] ? data[i][5].toString() : null,
          subscription: {
            plan: subscription.plan,
            expiryDate: subscription.expiryDate,
            expiryDateFormatted: expiryDateFormatted,
            expiryTimestamp: expiryTimestamp,
            status: subscription.status,
            isExpired: isExpired,
            remainingMinutes: remainingMinutes,
            remainingHours: remainingHours,
            remainingDays: remainingDays,
            remainingReadable: remainingReadable
          }
        });
      }
    }
    
    // Sort by created date (newest first)
    users.sort((a, b) => {
      const dateA = a.created ? new Date(a.created) : new Date(0);
      const dateB = b.created ? new Date(b.created) : new Date(0);
      return dateB - dateA;
    });
    
    return output({
      success: true,
      users: users,
      count: users.length
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Check user access (accurate to minutes)
function checkUserAccess(d) {
  try {
    if (!d.email) {
      return output({ success: false, error: "Email required" });
    }
    
    // Developer always has access
    if (d.email === DEVELOPER_EMAIL) {
      return output({ success: true, allowed: true, message: "Developer access granted" });
    }
    
    const masterDB = getMasterDatabase();
    let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
    
    if (!subscriptionSheet) {
      return output({ success: true, allowed: true, message: "Access granted (no subscription record)" });
    }
    
    const subData = subscriptionSheet.getDataRange().getValues();
    const now = new Date();
    const nowTimestamp = now.getTime();
    
    for (let i = 1; i < subData.length; i++) {
      if (subData[i][0] === d.email) {
        const expiryDateRaw = subData[i][1];
        const plan = subData[i][2] || "custom";
        const status = subData[i][3] || "active";
        
        // Unlimited plan - always allowed
        if (plan === "unlimited") {
          return output({ success: true, allowed: true, message: "Unlimited access" });
        }
        
        // Account suspended
        if (status !== "active") {
          return output({ 
            success: false, 
            allowed: false, 
            message: "Your account has been suspended. Please contact the administrator.",
            isExpired: false,
            isSuspended: true
          });
        }
        
        // Check expiry if exists
        if (expiryDateRaw) {
          const expiryDate = new Date(expiryDateRaw);
          const expiryTimestamp = expiryDate.getTime();
          
          // Check if expired (using milliseconds for accuracy)
          if (expiryTimestamp < nowTimestamp) {
            const diffMs = nowTimestamp - expiryTimestamp;
            const expiredMinutes = Math.floor(diffMs / (1000 * 60));
            const expiredHours = Math.floor(diffMs / (1000 * 60 * 60));
            const expiredDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
            
            let expiredText = "";
            if (expiredDays > 0) {
              expiredText = `${expiredDays} day${expiredDays > 1 ? 's' : ''} ago`;
            } else if (expiredHours > 0) {
              expiredText = `${expiredHours} hour${expiredHours > 1 ? 's' : ''} ago`;
            } else {
              expiredText = `${expiredMinutes} minute${expiredMinutes > 1 ? 's' : ''} ago`;
            }
            
            return output({ 
              success: false, 
              allowed: false, 
              message: `Your subscription expired ${expiredText}. Please renew to continue using the system.`,
              isExpired: true,
              expiryDate: expiryDateRaw,
              expiryDateFormatted: formatBusinessDateTime_(expiryDate),
              expiredDays: expiredDays,
              expiredHours: expiredHours,
              expiredMinutes: expiredMinutes
            });
          }
        }
        
        return output({ success: true, allowed: true, message: "Access granted" });
      }
    }
    
    // No subscription record found - allow access (default)
    return output({ success: true, allowed: true, message: "Access granted (no subscription record)" });
    
  } catch (error) {
    console.error("Error checking user access:", error);
    return output({ success: true, allowed: true, message: "Access granted (error)" });
  }
}

// Add this at the beginning of each action function
function checkUserAccessBeforeAction(d) {
  // Skip for developer
  if (d.email === DEVELOPER_EMAIL) return true;
  
  const masterDB = getMasterDatabase();
  let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
  if (!subscriptionSheet) return true;
  
  const subData = subscriptionSheet.getDataRange().getValues();
  for (let i = 1; i < subData.length; i++) {
    if (subData[i][0] === d.email) {
      const expiryDate = subData[i][1];
      const plan = subData[i][2];
      
      if (plan === "unlimited") return true;
      
      if (expiryDate) {
        const expiry = new Date(expiryDate);
        const now = new Date();
        if (expiry < now) {
          throw new Error("Your subscription has expired. Please renew to continue.");
        }
      }
      break;
    }
  }
  return true;
}

// ============================================
// OFFLINE SUBSCRIPTION CHECK FUNCTIONS
// ============================================

// Save subscription info to localStorage for offline access
function saveSubscriptionToLocal(d) {
  try {
    if (!d.email) {
      return output({ success: false, error: "Email required" });
    }
    
    const masterDB = getMasterDatabase();
    let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
    
    let expiryDate = null;
    let plan = "unlimited";
    let status = "active";
    
    if (subscriptionSheet) {
      const subData = subscriptionSheet.getDataRange().getValues();
      for (let i = 1; i < subData.length; i++) {
        if (subData[i][0] === d.email) {
          expiryDate = subData[i][1];
          plan = subData[i][2] || "custom";
          status = subData[i][3] || "active";
          break;
        }
      }
    }
    
    // Return subscription info to be saved locally
    return output({
      success: true,
      email: d.email,
      expiryDate: expiryDate,
      plan: plan,
      status: status,
      lastChecked: new Date().toISOString()
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Check subscription offline using cached data
function checkSubscriptionOffline(d) {
  try {
    if (!d.email) {
      return output({ success: false, error: "Email required" });
    }
    
    // This will be called from client-side with cached data
    const cachedData = d.cachedData;
    
    if (!cachedData) {
      return output({ 
        success: true, 
        allowed: true, 
        message: "No cached subscription data",
        isExpired: false 
      });
    }
    
    const now = new Date();
    const nowTimestamp = now.getTime();
    
    if (cachedData.plan === "unlimited") {
      return output({
        success: true,
        allowed: true,
        isExpired: false,
        message: "Unlimited access"
      });
    }
    
    if (cachedData.expiryDate) {
      const expiryTimestamp = new Date(cachedData.expiryDate).getTime();
      const isExpired = expiryTimestamp < nowTimestamp;
      
      if (isExpired) {
        const diffMs = nowTimestamp - expiryTimestamp;
        const expiredMinutes = Math.floor(diffMs / (1000 * 60));
        const expiredHours = Math.floor(diffMs / (1000 * 60 * 60));
        const expiredDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        
        return output({
          success: true,
          allowed: false,
          isExpired: true,
          expiryDate: cachedData.expiryDate,
          expiredDays: expiredDays,
          expiredHours: expiredHours,
          expiredMinutes: expiredMinutes,
          message: "Subscription expired"
        });
      }
    }
    
    return output({
      success: true,
      allowed: true,
      isExpired: false,
      message: "Access granted"
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}


// ============================================
// SERVER SIDE ACCESS CHECK FUNCTION
// ============================================

function checkUserAccessOnServer(email) {
  try {
    // Developer always has access
    if (email === DEVELOPER_EMAIL) {
      return { allowed: true, message: "Developer access granted" };
    }
    
    const masterDB = getMasterDatabase();
    let subscriptionSheet = masterDB.getSheetByName("USER_SUBSCRIPTIONS");
    
    if (!subscriptionSheet) {
      return { allowed: true, message: "No subscription record found" };
    }
    
    const subData = subscriptionSheet.getDataRange().getValues();
    const now = new Date();
    const nowTimestamp = now.getTime();
    
    for (let i = 1; i < subData.length; i++) {
      if (subData[i][0] === email) {
        const expiryDateRaw = subData[i][1];
        const plan = subData[i][2] || "custom";
        const status = subData[i][3] || "active";
        
        // Unlimited plan - always allowed
        if (plan === "unlimited") {
          return { allowed: true, message: "Unlimited access" };
        }
        
        // Account suspended
        if (status !== "active") {
          return { 
            allowed: false, 
            message: "Your account has been suspended. Please contact the administrator.",
            isExpired: false,
            isSuspended: true
          };
        }
        
        // Check expiry if exists
        if (expiryDateRaw) {
          const expiryDate = new Date(expiryDateRaw);
          const expiryTimestamp = expiryDate.getTime();
          
          if (expiryTimestamp < nowTimestamp) {
            return { 
              allowed: false, 
              message: "Your subscription has expired. Please renew to continue.",
              isExpired: true,
              expiryDate: expiryDateRaw
            };
          }
        }
        
        return { allowed: true, message: "Access granted" };
      }
    }
    
    // No subscription record found - allow access (default)
    return { allowed: true, message: "Access granted (no subscription record)" };
    
  } catch (error) {
    console.error("Error checking user access:", error);
    return { allowed: true, message: "Access granted (error)" };
  }
}

// ============================================
// WRAPPER FUNCTION FOR ACCESS CHECK
// ============================================

function withAccessCheck(actionFn, data) {
  try {
    // List of public actions that don't need access check
    const publicActions = [
      "register", "login", "forgot", 
      "checkUserAccess", "getAllUsers", "getAllUsersWithSubscription", 
      "deleteUser", "updateUserBusiness", "resetUserPassword", 
      "getUserDatabaseAccess", "updateUserSubscription", "getUserSubscription", 
      "saveSubscriptionToLocal", "checkSubscriptionOffline"
    ];
    
    // Skip access check for public actions
    if (publicActions.includes(data.action)) {
      return actionFn(data);
    }
    
    // Skip access check for developer
    if (data.email === DEVELOPER_EMAIL) {
      return actionFn(data);
    }
    
    // Check access for all other actions
    const accessCheck = checkUserAccessOnServer(data.email);
    if (!accessCheck.allowed) {
      return output({ 
        success: false, 
        error: accessCheck.message,
        accessDenied: true,
        isExpired: accessCheck.isExpired || false
      });
    }
    
    // Access granted, execute the action
    return actionFn(data);
    
  } catch(error) {
    return output({ 
      success: false, 
      error: error.toString(),
      stack: error.stack 
    });
  }
}


// ============================================
// DEBT HISTORY FUNCTIONS
// ============================================

// Ensure Debt History sheet exists
function ensureDebtHistorySheet(userDB) {
  try {
    let historySheet = userDB.getSheetByName("DEBT_HISTORY");
    if (!historySheet) {
      historySheet = userDB.insertSheet("DEBT_HISTORY");
      historySheet.appendRow([
        "HistoryID", "DebtID", "Action", "Amount", "PreviousStatus", "NewStatus",
        "PreviousPaid", "NewPaid", "Notes", "PerformedBy", "Timestamp"
      ]);
      historySheet.getRange("A1:K1").setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
    }
    return historySheet;
  } catch (error) {
    console.error("Error ensuring debt history sheet:", error);
    return null;
  }
}

function ensureDebtPhotoColumns_(debtsSheet) {
  if (!debtsSheet) return;
  if (debtsSheet.getLastColumn() < 18) {
    debtsSheet.getRange(1, 18).setValue("Photo Base64");
  }
  if (debtsSheet.getLastColumn() < 19) {
    debtsSheet.getRange(1, 19).setValue("Photo MIME Type");
  }
}

/* ===== ADD DEBT ===== */
function addDebt(d) {
  try {
    if (!d.databaseId) {
      return output({ success: false, error: "Database ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    const stockSheet = userDB.getSheetByName("STOCK");
    
    if (!debtsSheet) {
      return output({ success: false, error: "Debts sheet not found" });
    }
    ensureDebtPhotoColumns_(debtsSheet);
    
    const branchId = d.branchId || "MAIN";
    const timestamp = new Date(d.timestamp || Date.now());
    const debtId = d.debtId || "DEBT_" + Date.now() + "_" + Math.random().toString(36).substr(2, 8).toUpperCase();
    ensureDebtPhotoColumns_(debtsSheet);

    // Idempotency: an offline retry with the same debtId must not create a
    // second debt if the first request already reached the server.
    const existingDebtData = debtsSheet.getDataRange().getValues();
    for (let ei = 1; ei < existingDebtData.length; ei++) {
      if (String(existingDebtData[ei][0] || "") === String(debtId)) {
        return output({ success: true, alreadyProcessed: true, debtId: debtId, totalAmount: Number(existingDebtData[ei][9]) || 0 });
      }
    }
    
    // Check if this is a stock product by looking at STOCK sheet
    let isStockProduct = false;
    if (stockSheet && d.isStockProduct) {
      const stockData = stockSheet.getDataRange().getValues();
      for (let i = 1; i < stockData.length; i++) {
        if (stockData[i][1] === d.product) {
          isStockProduct = true;
          break;
        }
      }
    }
    
    if (d.checkStock && isStockProduct) {
      const stockCheck = checkBranchStockAvailability(d.databaseId, d.product, d.quantity, branchId, d.unit);
      if (!stockCheck.available) {
        return output({ 
          success: false, 
          error: `Insufficient stock at ${branchId} branch. Available: ${stockCheck.currentStock}, Requested: ${d.quantity}`,
          currentStock: stockCheck.currentStock
        });
      }
    }
    
    const totalAmount = d.quantity * d.unitPrice;
    
    var photoBase64 = String(d.photoBase64 || "");
    var photoMimeType = String(d.photoMimeType || "image/jpeg");
    if (photoBase64.indexOf("data:") === 0) {
      var commaPos = photoBase64.indexOf(",");
      if (commaPos !== -1) photoBase64 = photoBase64.substring(commaPos + 1);
    }
    if (photoBase64.length > 48000) {
      return output({ success: false, error: "Borrower photo is still too large after compression. Please choose another image." });
    }
    
    debtsSheet.appendRow([
      debtId, timestamp, d.borrowerName, d.borrowerPhone || "", d.borrowerNIDA || "",
      d.photoId || "", d.product, d.quantity, d.unitPrice, totalAmount,
      d.paidAmount || 0, d.status || DEBT_STATUS_PENDING, branchId,
      d.sellerId || null, d.notes || "", timestamp, d.unit || "", photoBase64, photoMimeType
    ]);
    
    // Add history entry
    addDebtHistory(d.databaseId, debtId, "CREATE", totalAmount, null, DEBT_STATUS_PENDING, 
                   0, 0, d.notes || "", d.sellerId || d.email);
    
    if (isStockProduct && d.checkStock) {
      updateStockAfterSale(d.databaseId, d.product, d.quantity, branchId, d.unit);
    }
    
    return output({
      success: true,
      message: `Debt recorded: ${d.borrowerName} owes ${totalAmount} for ${d.quantity} x ${d.product}`,
      debtId: debtId,
      totalAmount: totalAmount,
      timestamp: timestamp.toISOString()
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Add debt history entry
function addDebtHistory(databaseId, debtId, action, amount, previousStatus, newStatus, 
                        previousPaid, newPaid, notes, performedBy) {
  try {
    const userDB = SpreadsheetApp.openById(databaseId);
    const historySheet = ensureDebtHistorySheet(userDB);
    if (!historySheet) return false;
    
    const historyId = "HIST_" + Date.now() + "_" + Math.random().toString(36).substr(2, 6);
    const timestamp = new Date();
    
    historySheet.appendRow([
      historyId, debtId, action, amount || 0, previousStatus || "", newStatus || "",
      previousPaid || 0, newPaid || 0, notes || "", performedBy || "", timestamp
    ]);
    
    return true;
  } catch (error) {
    console.error("Error adding debt history:", error);
    return false;
  }
}

// Get debt history by debt ID
function getDebtHistory(d) {
  try {
    if (!d.databaseId || !d.debtId) {
      return output({ success: false, error: "Database ID and Debt ID required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const historySheet = userDB.getSheetByName("DEBT_HISTORY");
    
    if (!historySheet) {
      return output({ success: true, data: [], count: 0 });
    }
    
    const data = historySheet.getDataRange().getValues();
    const history = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === d.debtId) {
        history.push({
          historyId: data[i][0],
          debtId: data[i][1],
          action: data[i][2],
          amount: data[i][3],
          previousStatus: data[i][4],
          newStatus: data[i][5],
          previousPaid: data[i][6],
          newPaid: data[i][7],
          notes: data[i][8],
          performedBy: data[i][9],
          timestamp: data[i][10]
        });
      }
    }
    
    // Sort by timestamp (newest first)
    history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return output({
      success: true,
      data: history,
      count: history.length
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}

// Get borrower history (all debts for a borrower)
function getBorrowerDebtHistory(d) {
  try {
    if (!d.databaseId || !d.borrowerName) {
      return output({ success: false, error: "Database ID and Borrower Name required" });
    }
    
    const userDB = SpreadsheetApp.openById(d.databaseId);
    const debtsSheet = userDB.getSheetByName("DEBTS");
    const historySheet = userDB.getSheetByName("DEBT_HISTORY");
    
    if (!debtsSheet) {
      return output({ success: true, data: [], count: 0 });
    }
    
    const debtsData = debtsSheet.getDataRange().getValues();
    const borrowerDebts = [];
    
    // Get all debts for this borrower
    for (let i = 1; i < debtsData.length; i++) {
      if (debtsData[i][2] && debtsData[i][2].toLowerCase() === d.borrowerName.toLowerCase()) {
        borrowerDebts.push({
          id: debtsData[i][0],
          date: debtsData[i][1],
          borrowerName: debtsData[i][2],
          borrowerPhone: debtsData[i][3] || "",
          product: debtsData[i][6],
          quantity: debtsData[i][7],
          totalAmount: debtsData[i][9],
          paidAmount: debtsData[i][10] || 0,
          status: debtsData[i][11] || "pending",
          branchId: debtsData[i][12] || "MAIN",
          sellerId: debtsData[i][13] || null,
          lastUpdated: debtsData[i][15] || debtsData[i][1]
        });
      }
    }
    
    // Get history for each debt if history sheet exists
    if (historySheet) {
      const historyData = historySheet.getDataRange().getValues();
      const historyMap = {};
      
      for (let i = 1; i < historyData.length; i++) {
        const debtId = historyData[i][1];
        if (!historyMap[debtId]) {
          historyMap[debtId] = [];
        }
        historyMap[debtId].push({
          action: historyData[i][2],
          amount: historyData[i][3],
          previousStatus: historyData[i][4],
          newStatus: historyData[i][5],
          timestamp: historyData[i][10]
        });
      }
      
      // Sort history by timestamp for each debt
      for (const debtId in historyMap) {
        historyMap[debtId].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      }
      
      borrowerDebts.forEach(debt => {
        debt.history = historyMap[debt.id] || [];
      });
    }
    
    // Sort by date (newest first)
    borrowerDebts.sort((a, b) =>new Date(b.date) - new Date(a.date));
    
    // Calculate statistics
    let totalBorrowed = 0;
    let totalPaid = 0;
    let onTimePayments = 0;
    let latePayments = 0;
    
    borrowerDebts.forEach(debt => {
      totalBorrowed += Number(debt.totalAmount) || 0;
      totalPaid += Number(debt.paidAmount) || 0;
      
      // Check payment history for late payments
      if (debt.history && debt.history.length > 0) {
        debt.history.forEach(record => {
          if (record.action === "PAYMENT") {
            // Check if payment was late (custom logic based on due date)
            // For now, count all payments
            onTimePayments++;
          }
        });
      }
    });
    
    const reliabilityScore = totalBorrowed > 0 ? ((totalPaid / totalBorrowed) * 100).toFixed(0) : 100;
    
    return output({
      success: true,
      data: borrowerDebts,
      count: borrowerDebts.length,
      summary: {
        totalBorrowed: totalBorrowed,
        totalPaid: totalPaid,
        outstanding: totalBorrowed - totalPaid,
        reliabilityScore: reliabilityScore,
        onTimePayments: onTimePayments,
        latePayments: latePayments
      }
    });
    
  } catch (error) {
    return output({ success: false, error: error.toString() });
  }
}
