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
    ActivityLogs: ["id","email","action","details","createdAt"],
    OrderItems: ["id","orderId","productId","productName","quantity","unitPrice","lineTotal"],
    Addresses: ["id","customerId","name","phone","town","address","createdAt"],
    InventoryMovements: ["id","productId","productName","change","reason","orderId","createdAt"],
    Testimonials: ["id","name","role","message","rating","published","createdAt"],
    News: ["id","title","slug","excerpt","body","imageUrl","published","createdAt","updatedAt"],
    Newsletter: ["id","email","status","createdAt"],
    Media: ["id","fileName","mimeType","fileUrl","folder","createdAt"],
    Coupons: ["code","type","value","minimumOrder","active","expiresAt","usageLimit","usedCount"]
  }
};

function setupVeehicee(adminEmail, adminPassword) {
  adminEmail = String(adminEmail || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail) || !adminPassword || String(adminPassword).length < 12) {
    throw new Error("Supply a valid administrator email and a password of at least 12 characters.");
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
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    } else {
      // Extend existing sheets by appending missing columns; never clear or reorder existing data.
      const existingHeaders = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);
      headers.forEach(header => {
        if (existingHeaders.indexOf(header) === -1) {
          sh.getRange(1,sh.getLastColumn()+1).setValue(header);
          existingHeaders.push(header);
        }
      });
    }
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
      ["services","Services","services","Software development, education technology, website design, device setup, ICT support, consulting and support.",true,new Date()]
    ].forEach(r => pages.appendRow(r));
  }
  return {ok:true,spreadsheetUrl:ss.getUrl(),driveFolderUrl:folder.getUrl(),message:"Veehicee database and Drive storage folder are ready. Deploy this project as a web app next."};
}

// For first-time setup, run setupVeehicee(adminEmail, adminPassword)
// manually from the Apps Script editor using your own administrator details.
// Never place administrator credentials in source code or commit them to GitHub.

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";
  if (action === "health") return json_({ok:true,service:APP.NAME,version:APP.VERSION});
  return json_({ok:true,service:APP.NAME,version:APP.VERSION,message:"Veehicee API is running. Use POST for application actions."});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    if (raw.length > 100000) throw new Error("Request is too large.");
    const input = JSON.parse(raw);
    if (!input || typeof input !== "object" || Array.isArray(input)) throw new Error("Invalid request.");
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
      case "publicPages": result = publicPages_(); break;
      case "adminListProducts": result = adminListProducts_(input); break;
      case "adminListOrders": result = adminListOrders_(input); break;
      case "adminSaveCategory": lock.waitLock(15000); result = adminSaveCategory_(input); break;
      case "adminSaveShippingZone": lock.waitLock(15000); result = adminSaveShippingZone_(input); break;
      case "adminUploadMedia": result = adminUploadMedia_(input); break;
      case "createOrder": lock.waitLock(15000); result = createOrder_(input); break;
      case "verifyPayment": lock.waitLock(15000); result = verifyPayment_(input); break;
      case "customerOrderLookup": result = customerOrderLookup_(input); break;
      case "adminConfirmBankTransfer": lock.waitLock(15000); result = adminConfirmBankTransfer_(input); break;
      case "adminSaveTestimonial": lock.waitLock(15000); result = adminSaveTestimonial_(input); break;
      case "publicTestimonials": result = publicTestimonials_(); break;
      case "adminSaveNews": lock.waitLock(15000); result = adminSaveNews_(input); break;
      case "publicNews": result = publicNews_(); break;
      case "newsletterSubscribe": lock.waitLock(15000); result = newsletterSubscribe_(input); break;
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
  if (!email || !password || email.length > 254 || password.length > 256) throw new Error("Enter valid login details.");
  const cache = CacheService.getScriptCache();
  const attemptKey = "login_" + hashToken_(email);
  const attempts = Number(cache.get(attemptKey) || 0);
  if (attempts >= 5) throw new Error("Too many unsuccessful attempts. Please wait 15 minutes before trying again.");
  const sh = sheet_("AdminUsers"), data = sh.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    const r=data[i];
    if (String(r[1]).toLowerCase()===email && r[5]===true && hashPassword_(password,String(r[2]))===String(r[3])) {
      const token = Utilities.getUuid()+Utilities.getUuid();
      sheet_("AdminSessions").appendRow([hashToken_(token),email,new Date(Date.now()+6*60*60*1000),new Date()]);
      cache.remove(attemptKey);
      log_(email,"LOGIN","Administrator signed in");
      return {token:token,email:email,role:r[4]};
    }
  }
  cache.put(attemptKey, String(attempts + 1), 900);
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
  // The hidden website field is a honeypot. Silently accept bot submissions without storing them.
  if (String(x.website || "").trim()) return {message:"Enquiry received."};
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
  const name=String(x.name||"").trim(), price=Number(x.price), stock=Number(x.stock);
  if(!name || name.length>180 || !isFinite(price) || price<0 || price>1000000000 || !Number.isInteger(stock) || stock<0) {
    throw new Error("Enter a product name, a valid non-negative price and a whole-number stock quantity.");
  }
  if(String(x.category||"General").length>100 || String(x.description||"").length>5000 || String(x.imageUrl||"").length>2000) {
    throw new Error("Product details are too long.");
  }
  const sh=sheet_("Products"), id=String(x.id||Utilities.getUuid()), d=sh.getDataRange().getValues();
  const vals=[id,name,String(x.category||"General"),String(x.description||""),price,stock,String(x.imageUrl||""),x.active!==false,new Date(),new Date()];
  let row=-1; for(let i=1;i<d.length;i++) if(String(d[i][0])===id){row=i+1;break;}
  if(row>0) sh.getRange(row,1,1,vals.length).setValues([vals]); else sh.appendRow(vals);
  log_(email,"SAVE_PRODUCT",name); return {message:"Product saved.",id:id};
}
function adminUpdateOrder_(x) {
  const email=requireAdmin_(x.token), id=String(x.id||"");
  const sh=sheet_("Orders"), d=sh.getDataRange().getValues();
  const allowedOrderStatuses=["Pending","Processing","Ready for dispatch","Dispatched","Delivered","Cancelled","Refunded"];
  const allowedPaymentStatuses=["Pending","Awaiting verification","Confirmed","Paid","Failed","Refunded"];
  if(x.orderStatus && !allowedOrderStatuses.includes(String(x.orderStatus))) throw new Error("Invalid order status.");
  if(x.paymentStatus && !allowedPaymentStatuses.includes(String(x.paymentStatus))) throw new Error("Invalid payment status.");
  if(!id || id.length>120) throw new Error("A valid order ID is required.");
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


/* ───────── EXPANDED CONTENT, COMMERCE AND CUSTOMER SERVICES ───────── */

function publicPages_() {
  const d=sheet_("WebsitePages").getDataRange().getValues();
  return {pages:d.slice(1).filter(r=>r[0] && r[4]===true).map(r=>({id:r[0],title:r[1],slug:r[2],body:r[3]}))};
}
function adminListProducts_(x) {
  requireAdmin_(x.token);
  const d=sheet_("Products").getDataRange().getValues();
  return {products:d.slice(1).filter(r=>r[0]).map(r=>({id:r[0],name:r[1],category:r[2],description:r[3],price:Number(r[4]),stock:Number(r[5]),imageUrl:r[6],active:r[7]===true,createdAt:r[8],updatedAt:r[9]}))};
}
function adminListOrders_(x) {
  requireAdmin_(x.token);
  const d=sheet_("Orders").getDataRange().getValues();
  return {orders:d.slice(1).filter(r=>r[0]).map(r=>({id:r[0],customerId:r[1],customerName:r[2],phone:r[3],town:r[4],address:r[5],itemsJson:r[6],subtotal:Number(r[7]),deliveryFee:Number(r[8]),total:Number(r[9]),paymentMethod:r[10],paymentStatus:r[11],orderStatus:r[12],createdAt:r[13],deliveryStatus:r[14]||r[12],trackingCode:r[15]||""})).reverse()};
}
function adminSaveCategory_(x) {
  const email=requireAdmin_(x.token), name=String(x.name||"").trim();
  if(!name || name.length>100) throw new Error("Enter a category name (maximum 100 characters).");
  const sh=sheet_("Categories"), d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++) if(String(d[i][1]).toLowerCase()===name.toLowerCase()) {
    sh.getRange(i+1,3).setValue(x.active!==false); log_(email,"SAVE_CATEGORY",name); return {message:"Category updated."};
  }
  sh.appendRow([Utilities.getUuid(),name,x.active!==false,new Date()]);
  log_(email,"SAVE_CATEGORY",name); return {message:"Category saved."};
}
function adminSaveShippingZone_(x) {
  const email=requireAdmin_(x.token), town=String(x.town||"").trim(), fee=Number(x.deliveryFee);
  if(!town || town.length>120 || !isFinite(fee) || fee<0 || fee>10000000) throw new Error("Enter a valid delivery town and fee.");
  const sh=sheet_("ShippingZones"), d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++) if(String(d[i][0]).toLowerCase()===town.toLowerCase()) {
    sh.getRange(i+1,2,1,2).setValues([[x.active!==false,fee]]);
    log_(email,"SAVE_SHIPPING_ZONE",town); return {message:"Delivery zone updated."};
  }
  sh.appendRow([town,x.active!==false,fee]); log_(email,"SAVE_SHIPPING_ZONE",town); return {message:"Delivery zone saved."};
}
function adminUploadMedia_(x) {
  const email=requireAdmin_(x.token);
  const dataUrl=String(x.dataUrl||""), name=String(x.fileName||"upload");
  if(dataUrl.length>7000000) throw new Error("Image is too large. Maximum upload size is approximately 5 MB.");
  const m=dataUrl.match(/^data:(image\/(?:jpeg|png|webp|gif));base64,([A-Za-z0-9+/=]+)$/);
  if(!m) throw new Error("Only JPEG, PNG, WebP and GIF images are allowed.");
  const bytes=Utilities.base64Decode(m[2]), mime=m[1];
  if(bytes.length>5*1024*1024) throw new Error("Image exceeds the 5 MB limit.");
  const rootId=PropertiesService.getScriptProperties().getProperty("VEEHICEE_DRIVE_FOLDER_ID");
  if(!rootId) throw new Error("Run setupVeehicee first to create media storage.");
  const root=DriveApp.getFolderById(rootId), folderName=["portfolio","products","news","general"].includes(String(x.folder))?String(x.folder):"general";
  const folders=root.getFoldersByName(folderName), folder=folders.hasNext()?folders.next():root.createFolder(folderName);
  const safeName=name.replace(/[^\w.\- ]/g,"_").slice(0,100);
  const file=folder.createFile(Utilities.newBlob(bytes,mime,safeName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  const url="https://drive.google.com/uc?export=view&id="+file.getId();
  sheet_("Media").appendRow([Utilities.getUuid(),safeName,mime,url,folderName,new Date()]);
  log_(email,"UPLOAD_MEDIA",safeName);
  return {message:"Image uploaded.",fileId:file.getId(),fileUrl:url};
}
function createOrder_(x) {
  const customerName=String(x.customerName||"").trim(), phone=String(x.phone||"").trim();
  const email=String(x.email||"").trim().toLowerCase(), town=String(x.town||"").trim(), address=String(x.address||"").trim();
  if(!customerName || customerName.length>150 || !phone || phone.length>40 || !town || town.length>120 || !address || address.length>500) throw new Error("Complete your name, phone, town and delivery address.");
  if(email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Enter a valid email address.");
  const items=Array.isArray(x.items)?x.items:[];
  if(!items.length || items.length>30) throw new Error("Your cart is empty or contains too many different products.");
  const zoneRows=sheet_("ShippingZones").getDataRange().getValues();
  const zone=zoneRows.slice(1).find(r=>String(r[0]).toLowerCase()===town.toLowerCase() && r[1]===true);
  if(!zone) throw new Error("Delivery is not yet available in that town. Please contact us for assistance.");
  const products=sheet_("Products"), productRows=products.getDataRange().getValues();
  const lines=[], needed={}; let subtotal=0;
  items.forEach(item=>{
    const id=String(item.id||""), qty=Number(item.quantity);
    if(!id || !Number.isInteger(qty) || qty<1 || qty>50) throw new Error("A cart item has an invalid quantity.");
    const row=productRows.findIndex((r,i)=>i>0 && String(r[0])===id && r[7]===true);
    if(row<1) throw new Error("One of the selected products is no longer available.");
    const p=productRows[row], already=needed[id]||0;
    if(Number(p[5])<already+qty) throw new Error("Insufficient stock for "+p[1]+".");
    needed[id]=already+qty;
    const price=Number(p[4]), total=price*qty;
    subtotal+=total; lines.push({productId:id,name:String(p[1]),quantity:qty,unitPrice:price,lineTotal:total});
  });
  const deliveryFee=Number(zone[2])||0, total=subtotal+deliveryFee;
  const method=String(x.paymentMethod||"").toLowerCase();
  if(["paystack","flutterwave","bank_transfer"].indexOf(method)<0) throw new Error("Choose a supported payment method.");
  const orderId="VEE-"+Utilities.getUuid().replace(/-/g,"").slice(0,12).toUpperCase();
  let customerId="";
  const customers=sheet_("Customers"), cd=customers.getDataRange().getValues();
  const existing=cd.slice(1).find(r=>(email && String(r[2]).toLowerCase()===email) || String(r[3])===phone);
  if(existing) customerId=String(existing[0]);
  else {customerId=Utilities.getUuid();customers.appendRow([customerId,customerName,email,phone,new Date()]);}
  sheet_("Addresses").appendRow([Utilities.getUuid(),customerId,customerName,phone,town,address,new Date()]);
  sheet_("Orders").appendRow([orderId,customerId,customerName,phone,town,address,JSON.stringify(lines),subtotal,deliveryFee,total,method,method==="bank_transfer"?"Awaiting verification":"Pending","Pending",new Date(),"Pending",""]);
  const oi=sheet_("OrderItems"); lines.forEach(line=>oi.appendRow([Utilities.getUuid(),orderId,line.productId,line.name,line.quantity,line.unitPrice,line.lineTotal]));
  let payment={method:method,status:method==="bank_transfer"?"Awaiting verification":"Pending"};
  if(method==="paystack") payment=initializePaystack_(orderId,total,email);
  if(method==="flutterwave") payment=initializeFlutterwave_(orderId,total,email,customerName,phone);
  log_("public","CREATE_ORDER",orderId+" total "+total);
  return {message:"Order created.",orderId:orderId,subtotal:subtotal,deliveryFee:deliveryFee,total:total,payment:payment};
}
function initializePaystack_(orderId,amount,email) {
  const secret=PropertiesService.getScriptProperties().getProperty("VEEHICEE_PAYSTACK_SECRET");
  if(!secret) throw new Error("Paystack is not configured yet. Please choose bank transfer or contact the store.");
  const reference="VEE"+Utilities.getUuid().replace(/-/g,"").toUpperCase();
  const response=UrlFetchApp.fetch("https://api.paystack.co/transaction/initialize",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+secret},payload:JSON.stringify({email:email||"orders@veehicee.com",amount:Math.round(amount*100),reference:reference,metadata:{orderId:orderId}}),muteHttpExceptions:true});
  const data=JSON.parse(response.getContentText()||"{}");
  if(response.getResponseCode()<200 || response.getResponseCode()>=300 || !data.status || !data.data || !data.data.authorization_url) throw new Error("Could not initialize Paystack payment. Please retry or choose bank transfer.");
  sheet_("Payments").appendRow([Utilities.getUuid(),orderId,reference,amount,"Pending","",new Date()]);
  return {method:"paystack",status:"Pending",reference:reference,authorizationUrl:data.data.authorization_url};
}
function initializeFlutterwave_(orderId,amount,email,name,phone) {
  const secret=PropertiesService.getScriptProperties().getProperty("VEEHICEE_FLW_SECRET_KEY");
  if(!secret) throw new Error("Flutterwave is not configured yet. Please choose bank transfer or contact the store.");
  const reference="VEE"+Utilities.getUuid().replace(/-/g,"").toUpperCase();
  const response=UrlFetchApp.fetch("https://api.flutterwave.com/v3/payments",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+secret},payload:JSON.stringify({tx_ref:reference,amount:amount,currency:"NGN",redirect_url:ScriptApp.getService().getUrl(),customer:{email:email||"orders@veehicee.com",name:name,phonenumber:phone},customizations:{title:"Veehicee order",description:"Payment for "+orderId},meta:{orderId:orderId}}),muteHttpExceptions:true});
  const data=JSON.parse(response.getContentText()||"{}");
  if(response.getResponseCode()<200 || response.getResponseCode()>=300 || data.status!=="success" || !data.data || !data.data.link) throw new Error("Could not initialize Flutterwave payment. Please retry or choose bank transfer.");
  sheet_("Payments").appendRow([Utilities.getUuid(),orderId,reference,amount,"Pending","",new Date()]);
  return {method:"flutterwave",status:"Pending",reference:reference,authorizationUrl:data.data.link};
}
function verifyPayment_(x) {
  const orderId=String(x.orderId||""), reference=String(x.reference||"");
  if(!orderId || !reference || reference.length>160) throw new Error("Order ID and payment reference are required.");
  const orders=sheet_("Orders"), od=orders.getDataRange().getValues(), oi=od.findIndex((r,i)=>i>0 && String(r[0])===orderId);
  if(oi<1) throw new Error("Order not found.");
  const order=od[oi], method=String(order[10]), expected=Number(order[9]);
  const payRows=sheet_("Payments").getDataRange().getValues(), payment=payRows.slice(1).find(r=>String(r[1])===orderId && String(r[2])===reference);
  if(!payment) throw new Error("Payment reference does not match this order.");
  let verified=false, providerId="";
  if(method==="paystack") {
    const secret=PropertiesService.getScriptProperties().getProperty("VEEHICEE_PAYSTACK_SECRET");
    if(!secret) throw new Error("Paystack verification is not configured.");
    const resp=UrlFetchApp.fetch("https://api.paystack.co/transaction/verify/"+encodeURIComponent(reference),{headers:{Authorization:"Bearer "+secret},muteHttpExceptions:true});
    const data=JSON.parse(resp.getContentText()||"{}");
    verified=resp.getResponseCode()===200 && data.status===true && data.data && data.data.status==="success" && Number(data.data.amount)===Math.round(expected*100) && data.data.currency==="NGN" && data.data.reference===reference && data.data.metadata && String(data.data.metadata.orderId)===orderId;
    providerId=data.data&&data.data.id?String(data.data.id):"";
  } else if(method==="flutterwave") {
    const secret=PropertiesService.getScriptProperties().getProperty("VEEHICEE_FLW_SECRET_KEY");
    const transactionId=String(x.transactionId||"");
    if(!secret || !transactionId) throw new Error("Flutterwave transaction ID and verification configuration are required.");
    const resp=UrlFetchApp.fetch("https://api.flutterwave.com/v3/transactions/"+encodeURIComponent(transactionId)+"/verify",{headers:{Authorization:"Bearer "+secret},muteHttpExceptions:true});
    const data=JSON.parse(resp.getContentText()||"{}"), tx=data.data||{};
    verified=resp.getResponseCode()===200 && data.status==="success" && tx.status==="successful" && Number(tx.amount)>=expected && tx.currency==="NGN" && tx.tx_ref===reference;
    providerId=String(tx.id||"");
  } else {
    throw new Error("This order uses bank transfer. Payment must be confirmed by an administrator.");
  }
  if(!verified) throw new Error("Payment could not be verified. Do not treat this order as paid.");
  finalizePaidOrder_(orderId,reference,providerId);
  return {message:"Payment verified.",orderId:orderId,paymentStatus:"Paid"};
}
function finalizePaidOrder_(orderId,reference,providerId) {
  const orders=sheet_("Orders"), d=orders.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===orderId);
  if(idx<1) throw new Error("Order not found.");
  if(String(d[idx][11])==="Paid" || String(d[idx][11])==="Confirmed") return;
  const lines=JSON.parse(String(d[idx][6]||"[]")), products=sheet_("Products"), pd=products.getDataRange().getValues();
  lines.forEach(line=>{
    const pidx=pd.findIndex((r,i)=>i>0 && String(r[0])===String(line.productId));
    if(pidx<1 || Number(pd[pidx][5])<Number(line.quantity)) throw new Error("Stock changed before payment confirmation. Contact the store before dispatch.");
  });
  lines.forEach(line=>{
    const pidx=pd.findIndex((r,i)=>i>0 && String(r[0])===String(line.productId));
    const remaining=Number(pd[pidx][5])-Number(line.quantity);
    products.getRange(pidx+1,6).setValue(remaining);
    sheet_("InventoryMovements").appendRow([Utilities.getUuid(),line.productId,line.name,-Number(line.quantity),"Order paid",orderId,new Date()]);
  });
  orders.getRange(idx+1,12).setValue("Paid");
  orders.getRange(idx+1,13).setValue("Processing");
  const payments=sheet_("Payments"), pay=payments.getDataRange().getValues();
  for(let i=1;i<pay.length;i++) if(String(pay[i][1])===orderId && String(pay[i][2])===reference) {
    payments.getRange(i+1,5).setValue("Paid"); payments.getRange(i+1,6).setValue(new Date());
  }
  log_("system","PAYMENT_VERIFIED",orderId+" provider transaction "+providerId);
}
function adminConfirmBankTransfer_(x) {
  const email=requireAdmin_(x.token), id=String(x.orderId||"");
  const orders=sheet_("Orders"), d=orders.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  if(idx<1 || String(d[idx][10])!=="bank_transfer") throw new Error("Bank-transfer order not found.");
  const reference="BANK-"+id;
  const payments=sheet_("Payments");
  payments.appendRow([Utilities.getUuid(),id,reference,Number(d[idx][9]),"Confirmed",new Date(),new Date()]);
  finalizePaidOrder_(id,reference,"manual-bank-confirmation");
  log_(email,"CONFIRM_BANK_TRANSFER",id);
  return {message:"Bank transfer confirmed and stock updated."};
}
function customerOrderLookup_(x) {
  const id=String(x.orderId||"").trim(), phone=String(x.phone||"").trim();
  if(!id || !phone) throw new Error("Enter your order number and phone number.");
  const d=sheet_("Orders").getDataRange().getValues(), row=d.slice(1).find(r=>String(r[0])===id && String(r[3])===phone);
  if(!row) throw new Error("We could not find an order matching those details.");
  return {order:{id:row[0],customerName:row[2],town:row[4],items:JSON.parse(String(row[6]||"[]")),subtotal:Number(row[7]),deliveryFee:Number(row[8]),total:Number(row[9]),paymentMethod:row[10],paymentStatus:row[11],orderStatus:row[12],createdAt:row[13],deliveryStatus:row[14]||row[12],trackingCode:row[15]||""}};
}
function adminSaveTestimonial_(x) {
  const email=requireAdmin_(x.token), name=String(x.name||"").trim(), message=String(x.message||"").trim(), rating=Number(x.rating||5);
  if(!name || !message || name.length>120 || message.length>1500 || !Number.isInteger(rating) || rating<1 || rating>5) throw new Error("Enter a name, message and rating from 1 to 5.");
  const id=String(x.id||Utilities.getUuid()), sh=sheet_("Testimonials"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  const row=[id,name,String(x.role||"Customer").slice(0,100),message,rating,x.published===true,new Date()];
  if(idx>0) sh.getRange(idx+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  log_(email,"SAVE_TESTIMONIAL",name); return {message:"Testimonial saved.",id:id};
}
function publicTestimonials_() {
  const d=sheet_("Testimonials").getDataRange().getValues();
  return {testimonials:d.slice(1).filter(r=>r[0] && r[5]===true).map(r=>({id:r[0],name:r[1],role:r[2],message:r[3],rating:Number(r[4])}))};
}
function adminSaveNews_(x) {
  const email=requireAdmin_(x.token), title=String(x.title||"").trim(), body=String(x.body||"");
  if(!title || title.length>180 || body.length>15000) throw new Error("Enter a news title and body (maximum 15,000 characters).");
  const id=String(x.id||slug_(title)||Utilities.getUuid()), sh=sheet_("News"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  const row=[id,title,slug_(title),String(x.excerpt||"").slice(0,500),body,String(x.imageUrl||"").slice(0,2000),x.published===true,new Date(),new Date()];
  if(idx>0) sh.getRange(idx+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  log_(email,"SAVE_NEWS",title); return {message:"News item saved.",id:id};
}
function publicNews_() {
  const d=sheet_("News").getDataRange().getValues();
  return {news:d.slice(1).filter(r=>r[0] && r[6]===true).map(r=>({id:r[0],title:r[1],slug:r[2],excerpt:r[3],body:r[4],imageUrl:r[5],createdAt:r[7]}))};
}
function newsletterSubscribe_(x) {
  const email=String(x.email||"").trim().toLowerCase();
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length>254) throw new Error("Enter a valid email address.");
  const sh=sheet_("Newsletter"), d=sh.getDataRange().getValues();
  if(d.slice(1).some(r=>String(r[1]).toLowerCase()===email && r[2]==="Subscribed")) return {message:"You are already subscribed."};
  sh.appendRow([Utilities.getUuid(),email,"Subscribed",new Date()]);
  return {message:"Subscription received."};
}
