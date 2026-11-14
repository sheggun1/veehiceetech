/*
 * Veehicee Telecommunications and Services
 * Google Apps Script backend — Google Sheets database
 *
 * FIRST-TIME SETUP:
 * 1. Create a Google Sheet and open Extensions > Apps Script.
 * 2. Replace Code.gs with this file and save.
 * 3. Run setupVeehicee("your-google-email@gmail.com", "a-strong-password").
 * 4. Authorize the script when Google asks.
 * 5. Deploy > New deployment > Web app.
 *    Execute as: Me. Who has access: Anyone.
 * 6. Copy the /exec URL into index.html at API_URL.
 *
 * Keep this script project private. Do not publish passwords or secret keys.
 */

const APP = {
  NAME: "Veehicee Telecommunications and Services",
  VERSION: "1.0.0",
  SHEETS: {
    AdminUsers: ["id","email","salt","passwordHash","role","active","createdAt"],
    AdminSessions: ["tokenHash","email","expiresAt","createdAt"],
    WebsitePages: ["id","title","slug","body","published","updatedAt"],
    Enquiries: ["id","name","contact","message","status","createdAt"],
    Categories: ["id","name","active","createdAt"],
    Products: ["id","name","category","description","price","stock","imageUrl","active","createdAt","updatedAt"],
    Customers: ["id","name","email","phone","createdAt"],
    Orders: ["id","customerId","customerName","phone","town","address","itemsJson","subtotal","deliveryFee","total","paymentMethod","paymentStatus","orderStatus","createdAt"],
    Payments: ["id","orderId","reference","amount","status","verifiedAt","createdAt"],
    ShippingZones: ["town","active","deliveryFee"],
    SiteSettings: ["key","value","updatedAt"],
    ActivityLogs: ["id","email","action","details","createdAt"]
  }
};

function setupVeehicee(adminEmail, adminPassword) {
  if (!adminEmail || !adminPassword || String(adminPassword).length < 12) {
    throw new Error("Supply the administrator email and a password of at least 12 characters.");
  }
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty("VEEHICEE_SHEET_ID");
  let ss = existing ? SpreadsheetApp.openById(existing) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create("Veehicee Website Database");
  props.setProperty("VEEHICEE_SHEET_ID", ss.getId());

  // Create the website's Google Drive storage folder automatically.
  let folderId = props.getProperty("VEEHICEE_DRIVE_FOLDER_ID");
  let folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder("Veehicee Website Storage");
    props.setProperty("VEEHICEE_DRIVE_FOLDER_ID", folder.getId());
  }

  Object.keys(APP.SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = APP.SHEETS[name];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  const users = ss.getSheetByName("AdminUsers");
  const rows = users.getDataRange().getValues();
  const existsUser = rows.slice(1).some(r => String(r[1]).toLowerCase() === String(adminEmail).toLowerCase());
  if (!existsUser) {
    const salt = Utilities.getUuid() + Utilities.getUuid();
    users.appendRow([Utilities.getUuid(),String(adminEmail).toLowerCase(),salt,hashPassword_(adminPassword,salt),"owner",true,new Date()]);
  }
  const zones = ss.getSheetByName("ShippingZones");
  if (zones.getLastRow() === 1) {
    ["Ado-Ekiti","Iworoko","Ijan","Igbara-Odo","Ikere-Ekiti","Ilawe-Ekiti","Aramoko-Ekiti"].forEach(t => zones.appendRow([t,true,0]));
  }
  const pages = ss.getSheetByName("WebsitePages");
  if (pages.getLastRow() === 1) {
    [
      ["home","Home","home","Ideas into impact. We create useful digital products and technology experiences.",true,new Date()],
      ["about","About","about","Veehicee makes digital progress practical and accessible.",true,new Date()],
      ["services","Services","services","Software development, education technology, website design, ICT services, consulting and support.",true,new Date()]
    ].forEach(r => pages.appendRow(r));
  }
  return {ok:true,spreadsheetUrl:ss.getUrl(),driveFolderUrl:folder.getUrl(),message:"Veehicee database and Drive storage folder are ready. Deploy this project as a web app next."};
}

// One-click first-time setup for the administrator.
// Change the temporary password immediately after setup using a secure password-change feature.
function runVeehiceeSetup() {
  setupVeehicee("sheggun4@gmail.com", "sheggun2026!");
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";
  if (action === "health") return json_({ok:true,service:APP.NAME,version:APP.VERSION});
  return json_({ok:true,service:APP.NAME,version:APP.VERSION,message:"Veehicee API is running. Use POST for application actions."});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const input = JSON.parse(raw);
    const action = String(input.action || "");
    let result;
    switch (action) {
      case "adminLogin": result = adminLogin_(input); break;
      case "adminLogout": result = adminLogout_(input); break;
      case "adminListPages": result = adminListPages_(input); break;
      case "adminSavePage": lock.waitLock(15000); result = adminSavePage_(input); break;
      case "submitEnquiry": lock.waitLock(15000); result = submitEnquiry_(input); break;
      case "publicProducts": result = publicProducts_(); break;
      case "publicShippingZones": result = publicShippingZones_(); break;
      case "adminDashboard": result = adminDashboard_(input); break;
      case "adminSaveProduct": lock.waitLock(15000); result = adminSaveProduct_(input); break;
      case "adminUpdateOrder": lock.waitLock(15000); result = adminUpdateOrder_(input); break;
      default: throw new Error("Unsupported action: " + action);
    }
    return json_(Object.assign({ok:true},result || {}));
  } catch (err) {
    return json_({ok:false,error:String(err && err.message ? err.message : err)});
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
}

function adminLogin_(x) {
  const email = String(x.email || "").trim().toLowerCase();
  const password = String(x.password || "");
  if (!email || !password) throw new Error("Enter your email and password.");
  const sh = sheet_("AdminUsers"), data = sh.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    const r=data[i];
    if (String(r[1]).toLowerCase()===email && r[5]===true && hashPassword_(password,String(r[2]))===String(r[3])) {
      const token = Utilities.getUuid()+Utilities.getUuid();
      sheet_("AdminSessions").appendRow([hashToken_(token),email,new Date(Date.now()+6*60*60*1000),new Date()]);
      log_(email,"LOGIN","Administrator signed in");
      return {token:token,email:email,role:r[4]};
    }
  }
  throw new Error("Invalid login details.");
}
function adminLogout_(x) {
  const email = requireAdmin_(x.token);
  const sh=sheet_("AdminSessions"), d=sh.getDataRange().getValues(), h=hashToken_(x.token);
  for(let i=d.length-1;i>=1;i--) if(String(d[i][0])===h) sh.deleteRow(i+1);
  log_(email,"LOGOUT","Administrator signed out");
  return {message:"Signed out."};
}
function requireAdmin_(token) {
  if (!token) throw new Error("Please sign in again.");
  const h=hashToken_(token), d=sheet_("AdminSessions").getDataRange().getValues(), now=Date.now();
  for(let i=1;i<d.length;i++) {
    if(String(d[i][0])===h && new Date(d[i][2]).getTime()>now) {
      const email=String(d[i][1]);
      const users=sheet_("AdminUsers").getDataRange().getValues();
      if(users.slice(1).some(r=>String(r[1])===email && r[5]===true)) return email;
    }
  }
  throw new Error("Your session has expired. Sign in again.");
}
function adminListPages_(x) {
  requireAdmin_(x.token);
  const d=sheet_("WebsitePages").getDataRange().getValues();
  return {pages:d.slice(1).filter(r=>r[0]).map(r=>({id:r[0],title:r[1],slug:r[2],body:r[3],published:r[4]===true}))};
}
function adminSavePage_(x) {
  const email=requireAdmin_(x.token), title=String(x.title||"").trim(), body=String(x.body||"");
  if(!title) throw new Error("Page title is required.");
  const sh=sheet_("WebsitePages"), d=sh.getDataRange().getValues();
  const slug=slug_(title), id=slug;
  let row=-1;
  for(let i=1;i<d.length;i++) if(String(d[i][0])===id || String(d[i][2])===slug) {row=i+1;break;}
  const values=[id,title,slug,body,x.published===true,new Date()];
  if(row>0) sh.getRange(row,1,1,values.length).setValues([values]); else sh.appendRow(values);
  log_(email,"SAVE_PAGE",title);
  return {message:"Page saved."};
}
function submitEnquiry_(x) {
  const name=String(x.name||"").trim(), contact=String(x.contact||"").trim(), message=String(x.message||"").trim();
  if(!name || !contact || !message) throw new Error("Complete all enquiry fields.");
  if(name.length>120 || contact.length>160 || message.length>5000) throw new Error("One or more fields are too long.");
  sheet_("Enquiries").appendRow([Utilities.getUuid(),name,contact,message,"New",new Date()]);
  return {message:"Enquiry received."};
}
function publicProducts_() {
  const d=sheet_("Products").getDataRange().getValues();
  return {products:d.slice(1).filter(r=>r[0] && r[7]===true).map(r=>({id:r[0],name:r[1],category:r[2],description:r[3],price:Number(r[4]),stock:Number(r[5]),imageUrl:r[6]}))};
}
function publicShippingZones_() {
  const d=sheet_("ShippingZones").getDataRange().getValues();
  return {zones:d.slice(1).filter(r=>r[0] && r[1]===true).map(r=>({town:r[0],deliveryFee:Number(r[2])||0}))};
}
function adminDashboard_(x) {
  const email=requireAdmin_(x.token);
  return {email:email,enquiries:sheet_("Enquiries").getLastRow()-1,products:sheet_("Products").getLastRow()-1,orders:sheet_("Orders").getLastRow()-1};
}
function adminSaveProduct_(x) {
  const email=requireAdmin_(x.token);
  const name=String(x.name||"").trim(), price=Number(x.price), stock=Math.max(0,Number(x.stock)||0);
  if(!name || !isFinite(price) || price<0) throw new Error("Enter a product name and valid price.");
  const sh=sheet_("Products"), id=String(x.id||Utilities.getUuid()), d=sh.getDataRange().getValues();
  const vals=[id,name,String(x.category||"General"),String(x.description||""),price,stock,String(x.imageUrl||""),x.active!==false,new Date(),new Date()];
  let row=-1; for(let i=1;i<d.length;i++) if(String(d[i][0])===id){row=i+1;break;}
  if(row>0) sh.getRange(row,1,1,vals.length).setValues([vals]); else sh.appendRow(vals);
  log_(email,"SAVE_PRODUCT",name); return {message:"Product saved.",id:id};
}
function adminUpdateOrder_(x) {
  const email=requireAdmin_(x.token), id=String(x.id||"");
  const sh=sheet_("Orders"), d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++) if(String(d[i][0])===id) {
    if(x.orderStatus) sh.getRange(i+1,13).setValue(String(x.orderStatus));
    if(x.paymentStatus) sh.getRange(i+1,12).setValue(String(x.paymentStatus));
    log_(email,"UPDATE_ORDER",id); return {message:"Order updated."};
  }
  throw new Error("Order not found.");
}

function sheet_(name) {
  const id=PropertiesService.getScriptProperties().getProperty("VEEHICEE_SHEET_ID");
  if(!id) throw new Error("Run setupVeehicee(email,password) first.");
  const sh=SpreadsheetApp.openById(id).getSheetByName(name);
  if(!sh) throw new Error("Missing database sheet: "+name);
  return sh;
}
function hashPassword_(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(salt)+"|"+String(password),Utilities.Charset.UTF_8);
  return bytes.map(b=>(b+256).toString(16).slice(-2)).join("");
}
function hashToken_(token) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(token),Utilities.Charset.UTF_8);
  return bytes.map(b=>(b+256).toString(16).slice(-2)).join("");
}
function slug_(s) { return String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,""); }
function log_(email,action,details) { sheet_("ActivityLogs").appendRow([Utilities.getUuid(),email,action,details,new Date()]); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
