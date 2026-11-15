/*
 * Veehicee Telecommunications and Services
 * Google Apps Script backend — Google Sheets database
 *
 * FIRST-TIME SETUP:
 * 1. Create a Google Sheet and open Extensions > Apps Script.
 * 2. Replace Code.gs with this file and save.
 * 3. Add VEEHICEE_SETUP_ADMIN_EMAIL to Script Properties (and, for a new
 *    administrator, VEEHICEE_SETUP_ADMIN_PASSWORD), then run runVeehiceeSetup.
 * 4. Authorize the script when Google asks.
 * 5. Deploy > New deployment > Web app.
 *    Execute as: Me. Who has access: Anyone.
 * 6. Copy the /exec URL into index.html at API_URL.
 *
 * Keep this script project private. Do not publish passwords or secret keys.
 */

const APP = {
  NAME: "Veehicee Telecommunications and Services",
  VERSION: "2.2.0",
  SHEETS: {
    AdminUsers: ["id","email","salt","passwordHash","role","active","createdAt"],
    AdminSessions: ["tokenHash","email","expiresAt","createdAt"],
    WebsitePages: ["id","title","slug","body","published","updatedAt"],
    Portfolio: ["id","title","category","description","imageUrls","projectUrl","featured","published","createdAt","updatedAt","projectGithubUrl","projectStatus"],
    Enquiries: ["id","name","contact","message","status","createdAt"],
    Categories: ["id","name","active","createdAt"],
    Products: ["id","name","category","description","price","stock","imageUrl","active","createdAt","updatedAt","costPrice","wholesalePrice"],
    Customers: ["id","name","email","phone","createdAt"],
    CustomerAccounts: ["id","name","email","phone","salt","passwordHash","active","createdAt"],
    CustomerSessions: ["tokenHash","customerId","email","expiresAt","createdAt"],
    Orders: ["id","customerId","customerName","phone","town","address","itemsJson","subtotal","deliveryFee","total","paymentMethod","paymentStatus","orderStatus","createdAt","deliveryStatus","trackingCode","stockReserved","updatedAt","couponCode","discountAmount","loyaltyPointsEarned","loyaltyPointsRedeemed"],
    Payments: ["id","orderId","reference","amount","status","verifiedAt","createdAt"],
    ShippingZones: ["town","active","deliveryFee"],
    SiteSettings: ["key","value","updatedAt"],
    ActivityLogs: ["id","email","action","details","createdAt"],
    OrderItems: ["id","orderId","productId","productName","quantity","unitPrice","lineTotal","costPrice"],
    Addresses: ["id","customerId","name","phone","town","address","createdAt"],
    InventoryMovements: ["id","productId","productName","change","reason","orderId","createdAt"],
    Testimonials: ["id","name","role","message","rating","published","createdAt"],
    News: ["id","title","slug","excerpt","body","imageUrl","published","createdAt","updatedAt"],
    Newsletter: ["id","email","status","createdAt"],
    Media: ["id","fileName","mimeType","fileUrl","folder","createdAt"],
    Backups: ["id","fileId","fileName","fileUrl","createdAt","createdBy"],
    Coupons: ["code","type","value","minimumOrder","active","expiresAt","usageLimit","usedCount"],
    Expenses: ["id","category","description","amount","expenseDate","createdBy","createdAt"],
    ProductReviews: ["id","productId","customerId","customerName","rating","review","approved","createdAt"],
    Wishlists: ["customerId","productId","createdAt"],
    LoyaltyLedger: ["id","customerId","orderId","points","amount","reason","createdAt"],
    AbandonedCarts: ["id","customerId","email","cartJson","status","updatedAt","reminderSentAt"],
    CouponRedemptions: ["id","code","orderId","customerId","createdAt"],
    PriceAlerts: ["id","customerId","email","productId","targetPrice","active","createdAt"],
    WebsiteVisits: ["id","visitedAt","pagePath","pageTitle","referrer","referrerHost","deviceType","browser","operatingSystem","screenSize","language","campaignSource","campaignMedium","campaignName"]
  }
};

function setupVeehicee(adminEmail, adminPassword) {
  adminEmail = String(adminEmail || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail)) {
    throw new Error("Supply a valid administrator email.");
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
    if (!adminPassword || String(adminPassword).length < 12) throw new Error("For a new administrator, supply a password of at least 12 characters.");
    const salt = Utilities.getUuid() + Utilities.getUuid();
    users.appendRow([Utilities.getUuid(),String(adminEmail).toLowerCase(),salt,"PBKDF2$"+strongPasswordHash_(adminPassword,salt),"owner",true,new Date()]);
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

// Run this no-argument function from the Apps Script editor after setting
// VEEHICEE_SETUP_ADMIN_EMAIL in Script Properties. For a new admin only,
// also set VEEHICEE_SETUP_ADMIN_PASSWORD; it is deleted after successful setup.
function runVeehiceeSetup() {
  const props=PropertiesService.getScriptProperties();
  const email=props.getProperty("VEEHICEE_SETUP_ADMIN_EMAIL");
  const password=props.getProperty("VEEHICEE_SETUP_ADMIN_PASSWORD")||"";
  if(!email) throw new Error("Add VEEHICEE_SETUP_ADMIN_EMAIL in Script Properties first.");
  const result=setupVeehicee(email,password);
  props.deleteProperty("VEEHICEE_SETUP_ADMIN_PASSWORD");
  return result;
}

// Never place administrator credentials in source code or commit them to GitHub.

function doGet(e) {
  const params=e && e.parameter ? e.parameter : {};
  const action=params.action||"";
  if(params.reference || params.trxref) {
    try {
      const reference=String(params.reference||params.trxref||"");
      const rows=sheet_("Payments").getDataRange().getValues();
      const payment=rows.slice(1).find(r=>String(r[2])===reference);
      if(!payment) throw new Error("Payment reference was not found.");
      verifyPayment_({orderId:String(payment[1]),reference:reference,transactionId:String(params.transaction_id||"")});
      return HtmlService.createHtmlOutput("<main style='font:16px Arial;max-width:620px;margin:10vh auto;padding:24px'><h2>Payment confirmed</h2><p>Your Veehicee order payment has been verified successfully.</p><p>Order: "+escapeHtml_(String(payment[1]))+"</p><p>You can close this page and return to the Veehicee website.</p></main>");
    } catch(callbackError) {
      return HtmlService.createHtmlOutput("<main style='font:16px Arial;max-width:620px;margin:10vh auto;padding:24px'><h2>Payment verification pending</h2><p>"+escapeHtml_(String(callbackError.message||callbackError))+"</p><p>Please keep your order number and contact Veehicee for assistance.</p></main>");
    }
  }
  if (action === "health") return json_({ok:true,service:APP.NAME,version:APP.VERSION});
  return json_({ok:true,service:APP.NAME,version:APP.VERSION,message:"Veehicee API is running. Use POST for application actions."});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    if (raw.length > 7500000) throw new Error("Request is too large.");
    const input = JSON.parse(raw);
    if (!input || typeof input !== "object" || Array.isArray(input)) throw new Error("Invalid request.");
    const action = String(input.action || "");
    let result;
    switch (action) {
      case "adminLogin": result = adminLogin_(input); break;
      case "adminLogout": result = adminLogout_(input); break;
      case "adminChangePassword": lock.waitLock(15000); result = adminChangePassword_(input); break;
      case "adminListPages": result = adminListPages_(input); break;
      case "adminSavePage": lock.waitLock(15000); result = adminSavePage_(input); break;
      case "submitEnquiry": lock.waitLock(15000); result = submitEnquiry_(input); break;
      case "publicProducts": result = publicProducts_(); break;
      case "publicShippingZones": result = publicShippingZones_(); break;
      case "adminDashboard": result = adminDashboard_(input); break;
      case "adminSaveProduct": lock.waitLock(15000); result = adminSaveProduct_(input); break;
      case "adminSetProductStatus": lock.waitLock(15000); result = adminSetProductStatus_(input); break;
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
      case "customerRegister": lock.waitLock(15000); result = customerRegister_(input); break;
      case "customerLogin": result = customerLogin_(input); break;
      case "customerLogout": result = customerLogout_(input); break;
      case "customerOrders": result = customerOrders_(input); break;
      case "adminSavePortfolio": lock.waitLock(15000); result = adminSavePortfolio_(input); break;
      case "adminSetPortfolioStatus": lock.waitLock(15000); result = adminSetPortfolioStatus_(input); break;
      case "adminListPortfolio": result = adminListPortfolio_(input); break;
      case "publicPortfolio": result = publicPortfolio_(); break;
      case "adminSaveSetting": lock.waitLock(15000); result = adminSaveSetting_(input); break;
      case "publicSettings": result = publicSettings_(); break;
      case "logWebsiteVisit": lock.waitLock(15000); result = logWebsiteVisit_(input); break;
      case "adminListWebsiteVisits": result = adminListWebsiteVisits_(input); break;
      case "adminInventoryReport": result = adminInventoryReport_(input); break;
      case "adminCreateBackup": result = adminCreateBackup_(input); break;
      case "adminSaveExpense": lock.waitLock(15000); result = adminSaveExpense_(input); break;
      case "adminListExpenses": result = adminListExpenses_(input); break;
      case "adminListCustomers": result = adminListCustomers_(input); break;
      case "adminBusinessReport": result = adminBusinessReport_(input); break;
      case "adminSaveCoupon": lock.waitLock(15000); result = adminSaveCoupon_(input); break;
      case "adminListCoupons": result = adminListCoupons_(input); break;
      case "publicValidateCoupon": result = publicValidateCoupon_(input); break;
      case "customerWishlist": result = customerWishlist_(input); break;
      case "customerSaveWishlist": lock.waitLock(15000); result = customerSaveWishlist_(input); break;
      case "publicProductReviews": result = publicProductReviews_(input); break;
      case "customerSubmitReview": lock.waitLock(15000); result = customerSubmitReview_(input); break;
      case "adminListReviews": result = adminListReviews_(input); break;
      case "adminSetReviewStatus": lock.waitLock(15000); result = adminSetReviewStatus_(input); break;
      case "customerLoyalty": result = customerLoyalty_(input); break;
      case "customerSaveAbandonedCart": lock.waitLock(15000); result = customerSaveAbandonedCart_(input); break;
      case "adminSendAbandonedCartReminder": result = adminSendAbandonedCartReminder_(input); break;
      case "adminListAbandonedCarts": result = adminListAbandonedCarts_(input); break;
      case "customerSavePriceAlert": lock.waitLock(15000); result = customerSavePriceAlert_(input); break;
      case "adminSendLowStockAlert": result = adminSendLowStockAlert_(input); break;
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
    if(String(r[1]).toLowerCase()!==email || r[5]!==true) continue;
    const storedHash=String(r[3]||""), salt=String(r[2]||"");
    const validPassword=storedHash.indexOf("PBKDF2$")===0
      ? strongPasswordHash_(password,salt)===storedHash.substring(7)
      : hashPassword_(password,salt)===storedHash;
    if (validPassword) {
      // Upgrade older salted SHA-256 credentials to PBKDF2 after a successful sign-in.
      if(storedHash.indexOf("PBKDF2$")!==0) {
        const newSalt=Utilities.getUuid()+Utilities.getUuid();
        sh.getRange(i+1,3,1,2).setValues([[newSalt,"PBKDF2$"+strongPasswordHash_(password,newSalt)]]);
      }
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
  const spamKey="enquiry_"+hashToken_(contact.toLowerCase());
  const cache=CacheService.getScriptCache();
  if(cache.get(spamKey)) throw new Error("Please wait a little before sending another enquiry.");
  cache.put(spamKey,"1",60);
  sheet_("Enquiries").appendRow([Utilities.getUuid(),name,contact,message,"New",new Date()]);
  return {message:"Enquiry received."};
}
function publicProducts_() {
  const d=sheet_("Products").getDataRange().getValues();
  return {products:d.slice(1).filter(r=>r[0] && r[7]===true).map(r=>({id:r[0],name:r[1],category:r[2],description:r[3],price:Number(r[4]),stock:Number(r[5]),imageUrl:r[6],wholesalePrice:Number(r[11])||0}))};
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
  const costPrice=Number(x.costPrice||0), wholesalePrice=Number(x.wholesalePrice||0);
  if(!isFinite(costPrice)||costPrice<0||!isFinite(wholesalePrice)||wholesalePrice<0) throw new Error("Enter valid non-negative cost and wholesale prices.");
  const vals=[id,name,String(x.category||"General"),String(x.description||""),price,stock,String(x.imageUrl||""),x.active!==false,new Date(),new Date(),costPrice,wholesalePrice];
  let row=-1, oldStock=0, oldPrice=null; for(let i=1;i<d.length;i++) if(String(d[i][0])===id){row=i+1;oldStock=Number(d[i][5])||0;oldPrice=Number(d[i][4])||0;break;}
  if(row>0) sh.getRange(row,1,1,vals.length).setValues([vals]); else sh.appendRow(vals);
  if(row>0 && oldPrice!==null && price<oldPrice) notifyPriceDrop_(id,name,price);
  if(row>0 && oldStock!==stock) sheet_("InventoryMovements").appendRow([Utilities.getUuid(),id,name,stock-oldStock,"Admin stock adjustment","",new Date()]);
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
    if(x.paymentStatus && ["Paid","Confirmed"].includes(String(x.paymentStatus))) {
      throw new Error("Do not manually mark an order as paid here. Use the bank-transfer confirmation action or payment verification.");
    }
    if(x.orderStatus) {
      const previous=String(d[i][12]), next=String(x.orderStatus);
      if(next==="Cancelled" && previous!=="Cancelled" && d[i][16]===true && !["Paid","Confirmed"].includes(String(d[i][11]))) restoreReservedStock_(id,d[i]);
      sh.getRange(i+1,13).setValue(next);
    }
    if(x.paymentStatus) sh.getRange(i+1,12).setValue(String(x.paymentStatus));
    if(x.deliveryStatus) {
      const allowedDelivery=["Pending","Processing","Ready for dispatch","Dispatched","Delivered","Cancelled"];
      if(!allowedDelivery.includes(String(x.deliveryStatus))) throw new Error("Invalid delivery status.");
      sh.getRange(i+1,15).setValue(String(x.deliveryStatus));
    }
    if(x.trackingCode!==undefined) sh.getRange(i+1,16).setValue(String(x.trackingCode).slice(0,120));
    sh.getRange(i+1,18).setValue(new Date());
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
  return {products:d.slice(1).filter(r=>r[0]).map(r=>({id:r[0],name:r[1],category:r[2],description:r[3],price:Number(r[4]),stock:Number(r[5]),imageUrl:r[6],active:r[7]===true,createdAt:r[8],updatedAt:r[9],costPrice:Number(r[10])||0,wholesalePrice:Number(r[11])||0}))};
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
  const root=DriveApp.getFolderById(rootId), folderName=["portfolio","products","news","general","site-images","service-gallery"].includes(String(x.folder))?String(x.folder):"general";
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
  let customerName=String(x.customerName||"").trim(), phone=String(x.phone||"").trim();
  let email=String(x.email||"").trim().toLowerCase();
  const town=String(x.town||"").trim(), address=String(x.address||"").trim();
  const signedCustomer=x.customerToken ? requireCustomer_(x.customerToken) : null;
  if(signedCustomer) { customerName=signedCustomer.name; phone=signedCustomer.phone; email=signedCustomer.email; }
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
    const retailPrice=Number(p[4]), wholesalePrice=Number(p[11])||0;
    const price=(qty>=5 && wholesalePrice>0)?wholesalePrice:retailPrice;
    const costPrice=Number(p[10])||0, lineTotal=price*qty;
    subtotal+=lineTotal; lines.push({productId:id,name:String(p[1]),quantity:qty,unitPrice:price,costPrice:costPrice,lineTotal:lineTotal});
  });
  const deliveryFee=Number(zone[2])||0;
  const couponCode=String(x.couponCode||"").trim().toUpperCase();
  let discountAmount=0;
  if(couponCode) discountAmount=validateCouponAmount_(couponCode,subtotal);
  const loyaltyPointsToRedeem=Math.max(0,Math.min(100000,Math.floor(Number(x.redeemPoints)||0)));
  let loyaltyDiscount=0;
  if(loyaltyPointsToRedeem) {
    if(!signedCustomer) throw new Error("Sign in to redeem loyalty points.");
    const balance=customerLoyaltyBalance_(signedCustomer.id);
    if(loyaltyPointsToRedeem>balance) throw new Error("You do not have enough loyalty points.");
    loyaltyDiscount=Math.min(loyaltyPointsToRedeem,subtotal-discountAmount);
  }
  discountAmount+=loyaltyDiscount;
  const total=Math.max(0,subtotal-discountAmount)+deliveryFee;
  const method=String(x.paymentMethod||"").toLowerCase();
  if(["paystack","flutterwave","bank_transfer"].indexOf(method)<0) throw new Error("Choose a supported payment method.");
  if(["paystack","flutterwave"].includes(method) && !email) throw new Error("Enter your email address for online payment.");
  const orderId="VEE-"+Utilities.getUuid().replace(/-/g,"").slice(0,12).toUpperCase();
  let customerId="";
  const customers=sheet_("Customers"), cd=customers.getDataRange().getValues();
  let account=null;
  if(signedCustomer) account=signedCustomer;
  const existing=cd.slice(1).find(r=>(email && String(r[2]).toLowerCase()===email) || String(r[3])===phone);
  if(account) customerId=String(account.id);
  else if(existing) customerId=String(existing[0]);
  else {customerId=Utilities.getUuid();customers.appendRow([customerId,customerName,email,phone,new Date()]);}
  sheet_("Addresses").appendRow([Utilities.getUuid(),customerId,customerName,phone,town,address,new Date()]);
  // Reserve stock atomically so two customers cannot buy the same final unit.
  lines.forEach(line=>{
    const pidx=productRows.findIndex((r,i)=>i>0 && String(r[0])===line.productId);
    const currentStock=Number(products.getRange(pidx+1,6).getValue());
    products.getRange(pidx+1,6).setValue(currentStock-line.quantity);
    sheet_("InventoryMovements").appendRow([Utilities.getUuid(),line.productId,line.name,-line.quantity,"Reserved for order",orderId,new Date()]);
  });
  const earnedPoints=Math.floor(Math.max(0,total)/1000);
  sheet_("Orders").appendRow([orderId,customerId,customerName,phone,town,address,JSON.stringify(lines),subtotal,deliveryFee,total,method,method==="bank_transfer"?"Awaiting verification":"Pending","Pending",new Date(),"Pending","",true,new Date(),couponCode,discountAmount,earnedPoints,loyaltyPointsToRedeem]);
  const oi=sheet_("OrderItems"); lines.forEach(line=>oi.appendRow([Utilities.getUuid(),orderId,line.productId,line.name,line.quantity,line.unitPrice,line.lineTotal,line.costPrice||0]));
  
  let payment={method:method,status:method==="bank_transfer"?"Awaiting verification":"Pending"};
  try {
    if(method==="paystack") payment=initializePaystack_(orderId,total,email);
    if(method==="flutterwave") payment=initializeFlutterwave_(orderId,total,email,customerName,phone);
  } catch(paymentError) {
    const rows=sheet_("Orders").getDataRange().getValues(), idx=rows.findIndex((r,i)=>i>0 && String(r[0])===orderId);
    if(idx>0) { restoreReservedStock_(orderId,rows[idx]); sheet_("Orders").getRange(idx+1,13).setValue("Cancelled"); }
    throw paymentError;
  }
  if(email) {
    try { MailApp.sendEmail(email,"Veehicee order "+orderId,"Thank you for your order. Order total: NGN "+total+". We will update you as it is processed."); } catch(mailError) {}
  }
  const notify=PropertiesService.getScriptProperties().getProperty("VEEHICEE_NOTIFY_EMAIL");
  if(notify) { try { MailApp.sendEmail(notify,"New Veehicee order "+orderId,"A new order was placed. Customer: "+customerName+"; Phone: "+phone+"; Total: NGN "+total+"; Payment: "+method+"."); } catch(mailError) {} }
  if(signedCustomer) {
    const saved=sheet_("AbandonedCarts"), carts=saved.getDataRange().getValues();
    for(let i=1;i<carts.length;i++) if(String(carts[i][1])===String(signedCustomer.id)&&String(carts[i][4])==="Open") saved.getRange(i+1,5).setValue("Converted");
  }
  log_("public","CREATE_ORDER",orderId+" total "+total);
  return {message:"Order created.",orderId:orderId,subtotal:subtotal,deliveryFee:deliveryFee,total:total,payment:payment};
}
function initializePaystack_(orderId,amount,email) {
  const secret=PropertiesService.getScriptProperties().getProperty("VEEHICEE_PAYSTACK_SECRET");
  if(!secret) throw new Error("Paystack is not configured yet. Please choose bank transfer or contact the store.");
  const reference="VEE"+Utilities.getUuid().replace(/-/g,"").toUpperCase();
  const response=UrlFetchApp.fetch("https://api.paystack.co/transaction/initialize",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+secret},payload:JSON.stringify({email:email,amount:Math.round(amount*100),reference:reference,callback_url:ScriptApp.getService().getUrl(),metadata:{orderId:orderId}}),muteHttpExceptions:true});
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
    verified=resp.getResponseCode()===200 && data.status==="success" && tx.status==="successful" && Number(tx.amount)===expected && tx.currency==="NGN" && tx.tx_ref===reference;
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
  const lines=JSON.parse(String(d[idx][6]||"[]"));
  // Stock was reserved when the order was created; do not deduct it twice.
  if(d[idx][16]!==true) {
    const products=sheet_("Products"), pd=products.getDataRange().getValues();
    lines.forEach(line=>{
      const pidx=pd.findIndex((r,i)=>i>0 && String(r[0])===String(line.productId));
      if(pidx<1 || Number(pd[pidx][5])<Number(line.quantity)) throw new Error("Stock changed before payment confirmation. Contact the store before dispatch.");
    });
    lines.forEach(line=>{
      const pidx=pd.findIndex((r,i)=>i>0 && String(r[0])===String(line.productId));
      const currentStock=Number(products.getRange(pidx+1,6).getValue());
      products.getRange(pidx+1,6).setValue(currentStock-Number(line.quantity));
      sheet_("InventoryMovements").appendRow([Utilities.getUuid(),line.productId,line.name,-Number(line.quantity),"Order paid",orderId,new Date()]);
    });
    orders.getRange(idx+1,17).setValue(true);
  }
  orders.getRange(idx+1,12).setValue("Paid");
  orders.getRange(idx+1,13).setValue("Processing");
  const payments=sheet_("Payments"), pay=payments.getDataRange().getValues();
  for(let i=1;i<pay.length;i++) if(String(pay[i][1])===orderId && String(pay[i][2])===reference) {
    payments.getRange(i+1,5).setValue("Paid"); payments.getRange(i+1,6).setValue(new Date());
  }
  const refreshed=orders.getDataRange().getValues(), paidRow=refreshed.find((r,i)=>i>0 && String(r[0])===orderId);
  if(paidRow && paidRow[1]) {
    const ledger=sheet_("LoyaltyLedger").getDataRange().getValues();
    const earned=Math.max(0,Math.floor(Number(paidRow[20])||0));
    const earnedExists=ledger.slice(1).some(r=>String(r[2])===orderId && String(r[5])==="Earned on paid order");
    if(earned>0 && !earnedExists) sheet_("LoyaltyLedger").appendRow([Utilities.getUuid(),paidRow[1],orderId,earned,Number(paidRow[9])||0,"Earned on paid order",new Date()]);
    const redeemed=Math.max(0,Math.floor(Number(paidRow[21])||0));
    const redeemedExists=ledger.slice(1).some(r=>String(r[2])===orderId && String(r[5])==="Redeemed on paid order");
    if(redeemed>0 && !redeemedExists) sheet_("LoyaltyLedger").appendRow([Utilities.getUuid(),paidRow[1],orderId,-redeemed,-redeemed,"Redeemed on paid order",new Date()]);
    const coupon=String(paidRow[18]||"");
    if(coupon) {
      const redemptions=sheet_("CouponRedemptions").getDataRange().getValues();
      if(!redemptions.slice(1).some(r=>String(r[2])===orderId)) recordCouponUse_(coupon,orderId,paidRow[1]);
    }
  }
  log_("system","PAYMENT_VERIFIED",orderId+" provider transaction "+providerId);
}
function adminConfirmBankTransfer_(x) {
  const email=requireAdmin_(x.token), id=String(x.orderId||"");
  const orders=sheet_("Orders"), d=orders.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  if(idx<1 || String(d[idx][10])!=="bank_transfer") throw new Error("Bank-transfer order not found.");
  if(["Paid","Confirmed"].includes(String(d[idx][11]))) throw new Error("This bank transfer has already been confirmed.");
  if(String(d[idx][12])==="Cancelled") throw new Error("This order has been cancelled.");
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


function restoreReservedStock_(orderId,orderRow) {
  const lines=JSON.parse(String(orderRow[6]||"[]")), products=sheet_("Products"), pd=products.getDataRange().getValues();
  lines.forEach(line=>{
    const pidx=pd.findIndex((r,i)=>i>0 && String(r[0])===String(line.productId));
    if(pidx>0) {
      products.getRange(pidx+1,6).setValue(Number(pd[pidx][5])+Number(line.quantity));
      sheet_("InventoryMovements").appendRow([Utilities.getUuid(),line.productId,line.name,Number(line.quantity),"Cancelled order stock release",orderId,new Date()]);
    }
  });
  const orders=sheet_("Orders"), d=orders.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===orderId);
  if(idx>0) orders.getRange(idx+1,17).setValue(false);
}
function customerRegister_(x) {
  const name=String(x.name||"").trim(), email=String(x.email||"").trim().toLowerCase(), phone=String(x.phone||"").trim(), password=String(x.password||"");
  if(!name || name.length>150 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length>254 || !phone || phone.length>40 || password.length<12 || password.length>256) {
    throw new Error("Enter your name, valid email, phone number and a password of at least 12 characters.");
  }
  const sh=sheet_("CustomerAccounts"), d=sh.getDataRange().getValues();
  if(d.slice(1).some(r=>String(r[2]).toLowerCase()===email)) throw new Error("An account with this email already exists.");
  const salt=Utilities.getUuid()+Utilities.getUuid(), id=Utilities.getUuid();
  sh.appendRow([id,name,email,phone,salt,strongPasswordHash_(password,salt),true,new Date()]);
  const token=issueCustomerSession_(id,email);
  return {message:"Account created.",token:token,customer:{id:id,name:name,email:email,phone:phone}};
}
function customerLogin_(x) {
  const email=String(x.email||"").trim().toLowerCase(), password=String(x.password||"");
  if(!email || !password || email.length>254 || password.length>256) throw new Error("Enter your email and password.");
  const cache=CacheService.getScriptCache(), key="customer_login_"+hashToken_(email), attempts=Number(cache.get(key)||0);
  if(attempts>=5) throw new Error("Too many unsuccessful attempts. Please wait 15 minutes.");
  const d=sheet_("CustomerAccounts").getDataRange().getValues();
  for(let i=1;i<d.length;i++) if(String(d[i][2]).toLowerCase()===email && d[i][6]===true && strongPasswordHash_(password,String(d[i][4]))===String(d[i][5])) {
    cache.remove(key);
    const token=issueCustomerSession_(d[i][0],email);
    return {token:token,customer:{id:d[i][0],name:d[i][1],email:d[i][2],phone:d[i][3]}};
  }
  cache.put(key,String(attempts+1),900);
  throw new Error("Invalid email or password.");
}
function issueCustomerSession_(id,email) {
  const token=Utilities.getUuid()+Utilities.getUuid();
  sheet_("CustomerSessions").appendRow([hashToken_(token),id,email,new Date(Date.now()+30*24*60*60*1000),new Date()]);
  return token;
}
function requireCustomer_(token) {
  if(!token) throw new Error("Please sign in to your customer account.");
  const h=hashToken_(token), d=sheet_("CustomerSessions").getDataRange().getValues(), now=Date.now();
  for(let i=1;i<d.length;i++) if(String(d[i][0])===h && new Date(d[i][3]).getTime()>now) {
    const a=sheet_("CustomerAccounts").getDataRange().getValues().slice(1).find(r=>String(r[0])===String(d[i][1]) && r[6]===true);
    if(a) return {id:a[0],name:a[1],email:a[2],phone:a[3]};
  }
  throw new Error("Your customer session has expired. Sign in again.");
}
function customerLogout_(x) {
  const h=hashToken_(String(x.token||"")), sh=sheet_("CustomerSessions"), d=sh.getDataRange().getValues();
  for(let i=d.length-1;i>=1;i--) if(String(d[i][0])===h) sh.deleteRow(i+1);
  return {message:"Signed out."};
}
function customerOrders_(x) {
  const customer=requireCustomer_(x.token), d=sheet_("Orders").getDataRange().getValues();
  return {orders:d.slice(1).filter(r=>String(r[1])===String(customer.id)).map(r=>({id:r[0],town:r[4],items:JSON.parse(String(r[6]||"[]")),subtotal:Number(r[7]),deliveryFee:Number(r[8]),total:Number(r[9]),paymentMethod:r[10],paymentStatus:r[11],orderStatus:r[12],createdAt:r[13],deliveryStatus:r[14]||r[12],trackingCode:r[15]||""})).reverse()};
}
function strongPasswordHash_(password,salt) {
  // PBKDF2-HMAC-SHA256 with 12,000 iterations; customer passwords are never stored in plain text.
  const key=Utilities.newBlob(String(password)).getBytes(), saltBytes=Utilities.newBlob(String(salt)+"\u0000\u0000\u0000\u0001").getBytes();
  let u=Utilities.computeHmacSha256Signature(saltBytes,key), result=u.slice();
  for(let i=1;i<12000;i++) {
    u=Utilities.computeHmacSha256Signature(u,key);
    for(let j=0;j<result.length;j++) result[j]=(result[j]^u[j])&255;
  }
  return u8hex_(result);
}
function u8hex_(bytes) { return bytes.map(b=>(b+256).toString(16).slice(-2)).join(""); }


function isSafeWebUrl_(value) {
  const url=String(value||"").trim();
  return !url || (url.length<=1000 && /^https?:\/\/[A-Za-z0-9.-]+(?::[0-9]{1,5})?(?:[/?#][^\s"'<>]*)?$/i.test(url));
}
function adminSavePortfolio_(x) {
  const email=requireAdmin_(x.token), title=String(x.title||"").trim(), description=String(x.description||"");
  const imageUrls=Array.isArray(x.imageUrls)?x.imageUrls.map(String):[];
  const projectUrl=String(x.projectUrl||"").trim(), projectGithubUrl=String(x.projectGithubUrl||"").trim();
  const projectStatus=String(x.projectStatus|| (projectUrl?"Live":"In Development"));
  if(!title || title.length>180 || description.length>5000 || imageUrls.length>12 || imageUrls.some(u=>!/^https:\/\/drive\.google\.com\//.test(u)) || !isSafeWebUrl_(projectUrl) || !isSafeWebUrl_(projectGithubUrl) || !["In Development","Live","Maintenance"].includes(projectStatus)) throw new Error("Enter a valid project title, description, safe HTTP/HTTPS links, status and uploaded Drive image links.");
  const id=String(x.id||Utilities.getUuid()), sh=sheet_("Portfolio"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  const row=[id,title,String(x.category||"General").slice(0,100),description,JSON.stringify(imageUrls),projectUrl,x.featured===true,x.published===true,new Date(),new Date(),projectGithubUrl,projectStatus];
  if(idx>0) sh.getRange(idx+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  log_(email,"SAVE_PORTFOLIO",title); return {message:"Portfolio project saved.",id:id};
}
function adminListPortfolio_(x) {
  requireAdmin_(x.token); const d=sheet_("Portfolio").getDataRange().getValues();
  return {projects:d.slice(1).filter(r=>r[0]).map(r=>({id:r[0],title:r[1],category:r[2],description:r[3],imageUrls:JSON.parse(String(r[4]||"[]")),projectUrl:r[5]||"",featured:r[6]===true,published:r[7]===true,projectGithubUrl:r[10]||"",projectStatus:r[11]||(r[5]?"Live":"In Development")}))};
}
function publicPortfolio_() {
  const d=sheet_("Portfolio").getDataRange().getValues();
  const projects=d.slice(1).filter(r=>r[0] && r[7]===true).map(r=>({id:r[0],title:r[1],category:r[2],description:r[3],imageUrls:JSON.parse(String(r[4]||"[]")),projectUrl:r[5]||"",featured:r[6]===true,projectGithubUrl:r[10]||"",projectStatus:r[11]||(r[5]?"Live":"In Development")}));
  const builtIns=[
    {id:"vicora-public-practice",title:"VICORA Examination Study and Practice",category:"Education Technology",description:"An examination study and practice platform for WAEC, NECO, JAMB and general CBT preparation. It supports student accounts, subject and examination practice, timed attempts, scoring, explanations after submission and practice history.",imageUrls:[],projectUrl:"",featured:true,projectGithubUrl:"https://github.com/sheggun1/vicora-public-practice",projectStatus:"In Development"},
    {id:"vicora-teacher",title:"VICORA TEACHER",category:"Education Technology",description:"A multi-school teaching and planning platform for Nigerian Primary, JSS and SSS levels, supporting teacher and school accounts, AI-assisted lesson notes, schemes of work, instructional materials, curriculum resources and professional PDF/DOCX documents.",imageUrls:[],projectUrl:"",featured:true,projectGithubUrl:"https://github.com/sheggun1/vicora-teacher",projectStatus:"In Development"}
  ];
  builtIns.forEach(item=>{
    const legacyTitles=item.id==="vicora-public-practice"?["vicora public practice"]:["vicora teacher"];
    const alreadyListed=projects.some(p=>String(p.id)===item.id || String(p.title).toLowerCase()===item.title.toLowerCase() || legacyTitles.includes(String(p.title).toLowerCase()));
    if(!alreadyListed) projects.push(item);
  });
  return {projects:projects};
}
function adminSaveSetting_(x) {
  const email=requireAdmin_(x.token), key=String(x.key||"").trim(), value=String(x.value||"");
  if(!/^[A-Za-z0-9_.-]{1,100}$/.test(key) || value.length>10000) throw new Error("Invalid setting key or value too long.");
  const sh=sheet_("SiteSettings"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===key);
  const row=[key,value,new Date()];
  if(idx>0) sh.getRange(idx+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  log_(email,"SAVE_SETTING",key); return {message:"Setting saved."};
}
function publicSettings_() {
  const allowed=["siteImages","serviceShowcaseImages","companyName","businessPhone","whatsapp","businessEmail","businessAddress","mission","vision","seoTitle","seoDescription","instagram","facebook","openingHours","bankName","bankAccountName","bankAccountNumber","siteTheme","siteLayout","sitePattern","storefrontUrl","lowStockThreshold","clientNames","vicoraTeacherUrl","vicoraPracticeUrl"];
  const d=sheet_("SiteSettings").getDataRange().getValues(), settings={};
  d.slice(1).forEach(r=>{if(allowed.indexOf(String(r[0]))>=0) settings[String(r[0])]=String(r[1]||"");});
  return {settings:settings};
}


function websiteVisitsSheet_() {
  const id=PropertiesService.getScriptProperties().getProperty("VEEHICEE_SHEET_ID");
  if(!id) throw new Error("Website database is not configured.");
  const ss=SpreadsheetApp.openById(id);
  let sh=ss.getSheetByName("WebsiteVisits");
  if(!sh){sh=ss.insertSheet("WebsiteVisits");sh.getRange(1,1,1,APP.SHEETS.WebsiteVisits.length).setValues([APP.SHEETS.WebsiteVisits]);sh.setFrozenRows(1);}
  return sh;
}
function logWebsiteVisit_(x) {
  const pagePath=String(x.pagePath||"/").slice(0,300);
  if(pagePath.charAt(0)!=="/" || pagePath.indexOf("..")>=0) throw new Error("Invalid page path.");
  const row=[Utilities.getUuid(),new Date(),pagePath,String(x.pageTitle||"").slice(0,200),
    String(x.referrer||"").slice(0,300),String(x.referrerHost||"").slice(0,160),
    String(x.deviceType||"Unknown").slice(0,40),String(x.browser||"Unknown").slice(0,80),
    String(x.operatingSystem||"Unknown").slice(0,80),String(x.screenSize||"").slice(0,40),
    String(x.language||"").slice(0,40),String(x.campaignSource||"").slice(0,120),
    String(x.campaignMedium||"").slice(0,120),String(x.campaignName||"").slice(0,160)];
  websiteVisitsSheet_().appendRow(row);
  return {message:"Visit recorded."};
}
function adminListWebsiteVisits_(x) {
  requireAdmin_(x.token);
  const rows=websiteVisitsSheet_().getDataRange().getValues().slice(1).filter(r=>r[0]);
  const recent=rows.slice(-300).reverse().map(r=>({
    visitedAt:r[1] instanceof Date?r[1].toISOString():String(r[1]||""),
    pagePath:String(r[2]||""),pageTitle:String(r[3]||""),referrerHost:String(r[5]||""),
    deviceType:String(r[6]||""),browser:String(r[7]||""),operatingSystem:String(r[8]||""),screenSize:String(r[9]||"")
  }));
  const byPage={},byDevice={};
  rows.forEach(r=>{const p=String(r[2]||"/"),d=String(r[6]||"Unknown");byPage[p]=(byPage[p]||0)+1;byDevice[d]=(byDevice[d]||0)+1;});
  return {totalVisits:rows.length,recent:recent,byPage:byPage,byDevice:byDevice};
}

function adminInventoryReport_(x) {
  requireAdmin_(x.token);
  const d=sheet_("Products").getDataRange().getValues(), products=d.slice(1).filter(r=>r[0]);
  const low=products.filter(r=>Number(r[5])<=5 && r[7]===true).map(r=>({id:r[0],name:r[1],stock:Number(r[5])}));
  const stockValue=products.reduce((sum,r)=>sum+(Number(r[4])||0)*(Number(r[5])||0),0);
  const orders=sheet_("Orders").getDataRange().getValues().slice(1).filter(r=>r[0]);
  const paid=orders.filter(r=>["Paid","Confirmed"].includes(String(r[11])));
  const revenue=paid.reduce((sum,r)=>sum+(Number(r[9])||0),0);
  return {totalProducts:products.length,lowStock:low,stockValue:stockValue,paidOrders:paid.length,revenue:revenue,totalOrders:orders.length};
}

function adminSetProductStatus_(x) {
  const email=requireAdmin_(x.token), id=String(x.id||"");
  const sh=sheet_("Products"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  if(idx<1) throw new Error("Product not found.");
  sh.getRange(idx+1,8).setValue(x.active===true);
  sh.getRange(idx+1,10).setValue(new Date());
  log_(email,"SET_PRODUCT_STATUS",id+" active="+(x.active===true));
  return {message:x.active===true?"Product published.":"Product hidden."};
}
function adminSetPortfolioStatus_(x) {
  const email=requireAdmin_(x.token), id=String(x.id||"");
  const sh=sheet_("Portfolio"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[0])===id);
  if(idx<1) throw new Error("Portfolio project not found.");
  sh.getRange(idx+1,8).setValue(x.published===true);
  sh.getRange(idx+1,10).setValue(new Date());
  log_(email,"SET_PORTFOLIO_STATUS",id+" published="+(x.published===true));
  return {message:x.published===true?"Project published.":"Project hidden."};
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[ch]));
}

function adminChangePassword_(x) {
  const email=requireAdmin_(x.token), current=String(x.currentPassword||""), next=String(x.newPassword||"");
  if(next.length<12 || next.length>256) throw new Error("Your new password must contain at least 12 characters.");
  const sh=sheet_("AdminUsers"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0 && String(r[1]).toLowerCase()===email && r[5]===true);
  if(idx<1) throw new Error("Administrator account not found.");
  const oldHash=String(d[idx][3]), salt=String(d[idx][2]);
  const valid=oldHash.indexOf("PBKDF2$")===0 ? strongPasswordHash_(current,salt)===oldHash.substring(7) : hashPassword_(current,salt)===oldHash;
  if(!valid) throw new Error("Current password is incorrect.");
  const newSalt=Utilities.getUuid()+Utilities.getUuid();
  sh.getRange(idx+1,3,1,2).setValues([[newSalt,"PBKDF2$"+strongPasswordHash_(next,newSalt)]]);
  const sessions=sheet_("AdminSessions"), sd=sessions.getDataRange().getValues();
  for(let i=sd.length-1;i>=1;i--) if(String(sd[i][1]).toLowerCase()===email) sessions.deleteRow(i+1);
  const token=Utilities.getUuid()+Utilities.getUuid();
  sessions.appendRow([hashToken_(token),email,new Date(Date.now()+6*60*60*1000),new Date()]);
  log_(email,"CHANGE_PASSWORD","Administrator password changed");
  return {message:"Password changed. Other sessions have been signed out.",token:token};
}

function adminCreateBackup_(x) {
  const email=requireAdmin_(x.token);
  const sheetId=PropertiesService.getScriptProperties().getProperty("VEEHICEE_SHEET_ID");
  const folderId=PropertiesService.getScriptProperties().getProperty("VEEHICEE_DRIVE_FOLDER_ID");
  if(!sheetId || !folderId) throw new Error("Database setup is incomplete.");
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmmss");
  const source=DriveApp.getFileById(sheetId), folder=DriveApp.getFolderById(folderId);
  const copy=source.makeCopy("Veehicee-Database-Backup-"+stamp,folder);
  const url=copy.getUrl();
  sheet_("Backups").appendRow([Utilities.getUuid(),copy.getId(),copy.getName(),url,new Date(),email]);
  log_(email,"CREATE_BACKUP",copy.getName());
  return {message:"Database backup created.",fileName:copy.getName(),fileUrl:url};
}

function adminListAbandonedCarts_(x) {
  requireAdmin_(x.token);
  const d=sheet_("AbandonedCarts").getDataRange().getValues();
  return {carts:d.slice(1).filter(r=>r[0]&&r[4]==="Open").reverse().map(r=>({id:r[0],email:r[2],items:r[3],updatedAt:r[5],reminderSentAt:r[6]}))};
}

function customerSavePriceAlert_(x) {
  const c=requireCustomer_(x.token), id=String(x.productId||""), target=Number(x.targetPrice);
  if(!id||!isFinite(target)||target<=0) throw new Error("Enter a product and a valid target price.");
  const p=sheet_("Products").getDataRange().getValues().slice(1).find(r=>String(r[0])===id&&r[7]===true);
  if(!p) throw new Error("Product is unavailable.");
  const sh=sheet_("PriceAlerts"), d=sh.getDataRange().getValues();
  if(d.slice(1).some(r=>String(r[1])===String(c.id)&&String(r[3])===id&&r[5]===true)) throw new Error("You already have an active alert for this product.");
  sh.appendRow([Utilities.getUuid(),c.id,c.email,id,target,true,new Date()]);
  return {message:"Price alert saved. We will email you if the price reaches your target."};
}
function notifyPriceDrop_(productId,name,newPrice) {
  const sh=sheet_("PriceAlerts"), d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++) {
    if(String(d[i][3])===String(productId)&&d[i][5]===true&&Number(d[i][4])>=newPrice&&String(d[i][2])) {
      try { MailApp.sendEmail(String(d[i][2]),"Price drop: "+name,"Good news! "+name+" is now ₦"+newPrice+". Visit the Veehicee Marketplace to check availability."); sh.getRange(i+1,6).setValue(false); }
      catch(e) {}
    }
  }
}



// ───── VEEHICEE BUSINESS MANAGEMENT EXTENSIONS ─────
function adminSaveExpense_(x) {
  const email=requireAdmin_(x.token), category=String(x.category||"").trim(), description=String(x.description||"").trim(), amount=Number(x.amount);
  if(!category||category.length>100||description.length>500||!isFinite(amount)||amount<=0||amount>1000000000) throw new Error("Enter a category, description and valid positive expense amount.");
  const date=x.expenseDate?new Date(x.expenseDate):new Date();
  if(isNaN(date.getTime())) throw new Error("Enter a valid expense date.");
  sheet_("Expenses").appendRow([Utilities.getUuid(),category,description,amount,date,email,new Date()]);
  log_(email,"ADD_EXPENSE",category+" "+amount);
  return {message:"Expense recorded."};
}
function adminListExpenses_(x) {
  requireAdmin_(x.token);
  const d=sheet_("Expenses").getDataRange().getValues();
  return {expenses:d.slice(1).filter(r=>r[0]).reverse().slice(0,500).map(r=>({id:r[0],category:r[1],description:r[2],amount:Number(r[3]),expenseDate:r[4]}))};
}
function adminBusinessReport_(x) {
  requireAdmin_(x.token);
  const orders=sheet_("Orders").getDataRange().getValues().slice(1).filter(r=>r[0]&&["Paid","Confirmed"].includes(String(r[11])));
  const items=sheet_("OrderItems").getDataRange().getValues().slice(1);
  const expenses=sheet_("Expenses").getDataRange().getValues().slice(1).filter(r=>r[0]);
  const revenue=orders.reduce((s,r)=>s+(Number(r[9])||0),0);
  const cogs=items.filter(r=>orders.some(o=>String(o[0])===String(r[1]))).reduce((s,r)=>s+(Number(r[7])||0)*(Number(r[4])||0),0);
  const expenseTotal=expenses.reduce((s,r)=>s+(Number(r[3])||0),0);
  const monthly={};
  orders.forEach(r=>{const key=Utilities.formatDate(new Date(r[13]),Session.getScriptTimeZone(),"yyyy-MM");monthly[key]=(monthly[key]||0)+(Number(r[9])||0)});
  return {paidOrders:orders.length,revenue:revenue,costOfGoods:cogs,expenses:expenseTotal,estimatedProfit:revenue-cogs-expenseTotal,monthlyRevenue:monthly,expenseCount:expenses.length,profitNote:"Estimated profit uses saved product cost prices and recorded expenses. Older orders without saved cost prices may understate costs."};
}
function adminSaveCoupon_(x) {
  const email=requireAdmin_(x.token), code=String(x.code||"").trim().toUpperCase(), type=String(x.type||"percent"), value=Number(x.value), minimum=Number(x.minimumOrder||0), limit=Number(x.usageLimit||0);
  if(!/^[A-Z0-9_-]{3,30}$/.test(code)||!["percent","fixed"].includes(type)||!isFinite(value)||value<=0||(type==="percent"&&value>100)||minimum<0||!isFinite(minimum)||!Number.isInteger(limit)||limit<0) throw new Error("Enter a valid coupon code, discount and usage limit.");
  const expires=x.expiresAt?new Date(x.expiresAt):"";
  if(expires!==""&&isNaN(expires.getTime())) throw new Error("Invalid coupon expiry date.");
  const sh=sheet_("Coupons"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0&&String(r[0]).toUpperCase()===code);
  const row=[code,type,value,minimum,x.active!==false,expires,limit,idx>0?(Number(d[idx][7])||0):0];
  if(idx>0) sh.getRange(idx+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  log_(email,"SAVE_COUPON",code); return {message:"Coupon saved."};
}
function adminListCoupons_(x) {
  requireAdmin_(x.token);
  const d=sheet_("Coupons").getDataRange().getValues();
  return {coupons:d.slice(1).filter(r=>r[0]).map(r=>({code:r[0],type:r[1],value:Number(r[2]),minimumOrder:Number(r[3]),active:r[4]===true,expiresAt:r[5],usageLimit:Number(r[6])||0,usedCount:Number(r[7])||0}))};
}
function validateCouponAmount_(code,subtotal) {
  const d=sheet_("Coupons").getDataRange().getValues(), r=d.slice(1).find(v=>String(v[0]).toUpperCase()===code);
  if(!r||r[4]!==true) throw new Error("This coupon is invalid or inactive.");
  if(r[5]&&new Date(r[5]).getTime()<Date.now()) throw new Error("This coupon has expired.");
  if(Number(r[6])>0&&Number(r[7])>=Number(r[6])) throw new Error("This coupon has reached its usage limit.");
  if(subtotal<Number(r[3]||0)) throw new Error("Your order does not meet this coupon's minimum spend.");
  return Math.min(subtotal,r[1]==="percent"?subtotal*Number(r[2])/100:Number(r[2]));
}
function publicValidateCoupon_(x) {
  const code=String(x.code||"").trim().toUpperCase(), subtotal=Number(x.subtotal);
  if(!code||!isFinite(subtotal)||subtotal<0) throw new Error("Enter a coupon code and valid cart subtotal.");
  const amount=validateCouponAmount_(code,subtotal);
  return {valid:true,discount:amount,message:"Coupon applied."};
}
function recordCouponUse_(code,orderId,customerId) {
  const sh=sheet_("Coupons"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0&&String(r[0]).toUpperCase()===code);
  if(idx>0) sh.getRange(idx+1,8).setValue((Number(d[idx][7])||0)+1);
  sheet_("CouponRedemptions").appendRow([Utilities.getUuid(),code,orderId,customerId,new Date()]);
}
function customerWishlist_(x) {
  const c=requireCustomer_(x.token), rows=sheet_("Wishlists").getDataRange().getValues().slice(1), ids=rows.filter(r=>String(r[0])===String(c.id)).map(r=>String(r[1]));
  const products=sheet_("Products").getDataRange().getValues().slice(1);
  return {items:ids.map(id=>products.find(p=>String(p[0])===id)).filter(p=>p&&p[7]===true).map(p=>({id:p[0],name:p[1],price:Number(p[4]),stock:Number(p[5]),imageUrl:p[6]}))};
}
function customerSaveWishlist_(x) {
  const c=requireCustomer_(x.token), id=String(x.productId||""), action=String(x.action||"add");
  if(!id||!["add","remove"].includes(action)) throw new Error("Invalid wishlist request.");
  const sh=sheet_("Wishlists"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0&&String(r[0])===String(c.id)&&String(r[1])===id);
  if(action==="add"&&idx<1) {
    const p=sheet_("Products").getDataRange().getValues().slice(1).find(r=>String(r[0])===id&&r[7]===true);
    if(!p) throw new Error("Product is unavailable.");
    sh.appendRow([c.id,id,new Date()]);
  } else if(action==="remove"&&idx>0) sh.deleteRow(idx+1);
  return {message:action==="add"?"Saved to wishlist.":"Removed from wishlist."};
}
function publicProductReviews_(x) {
  const id=String(x.productId||"");
  if(!id) throw new Error("Product ID is required.");
  const d=sheet_("ProductReviews").getDataRange().getValues();
  const reviews=d.slice(1).filter(r=>String(r[1])===id&&r[6]===true).map(r=>({name:r[3],rating:Number(r[4]),review:r[5],createdAt:r[7]}));
  return {reviews:reviews};
}
function customerSubmitReview_(x) {
  const c=requireCustomer_(x.token), id=String(x.productId||""), rating=Number(x.rating), review=String(x.review||"").trim();
  if(!id||!Number.isInteger(rating)||rating<1||rating>5||!review||review.length>1500) throw new Error("Enter a product, rating from 1 to 5 and review up to 1,500 characters.");
  const orders=sheet_("Orders").getDataRange().getValues().slice(1).filter(r=>String(r[1])===String(c.id)&&["Paid","Confirmed"].includes(String(r[11])));
  const purchased=orders.some(o=>JSON.parse(String(o[6]||"[]")).some(line=>String(line.productId)===id));
  if(!purchased) throw new Error("Only customers with a paid order containing this product can review it.");
  const sh=sheet_("ProductReviews"), d=sh.getDataRange().getValues();
  if(d.slice(1).some(r=>String(r[1])===id&&String(r[2])===String(c.id))) throw new Error("You have already reviewed this product.");
  sh.appendRow([Utilities.getUuid(),id,c.id,c.name,rating,review,false,new Date()]);
  return {message:"Review submitted for admin approval."};
}
function adminListReviews_(x) {
  requireAdmin_(x.token);
  const d=sheet_("ProductReviews").getDataRange().getValues(), products=sheet_("Products").getDataRange().getValues();
  return {reviews:d.slice(1).filter(r=>r[0]).reverse().map(r=>({id:r[0],productId:r[1],productName:(products.find(p=>String(p[0])===String(r[1]))||[])[1]||"",customerName:r[3],rating:Number(r[4]),review:r[5],approved:r[6]===true}))};
}
function adminSetReviewStatus_(x) {
  const email=requireAdmin_(x.token), id=String(x.id||""), sh=sheet_("ProductReviews"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0&&String(r[0])===id);
  if(idx<1) throw new Error("Review not found.");
  sh.getRange(idx+1,7).setValue(x.approved===true); log_(email,"MODERATE_REVIEW",id);
  return {message:x.approved===true?"Review published.":"Review hidden."};
}
function customerLoyaltyBalance_(customerId) {
  const d=sheet_("LoyaltyLedger").getDataRange().getValues().slice(1);
  const earnedAndRedeemed=d.filter(r=>String(r[1])===String(customerId)).reduce((s,r)=>s+(Number(r[3])||0),0);
  const pending=sheet_("Orders").getDataRange().getValues().slice(1).filter(r=>String(r[1])===String(customerId)&&!["Paid","Confirmed","Failed"].includes(String(r[11]))&&String(r[12])!=="Cancelled").reduce((s,r)=>s+(Number(r[21])||0),0);
  return Math.max(0,earnedAndRedeemed-pending);
}
function customerLoyalty_(x) {
  const c=requireCustomer_(x.token), d=sheet_("LoyaltyLedger").getDataRange().getValues().slice(1);
  const rows=d.filter(r=>String(r[1])===String(c.id));
  return {points:rows.reduce((s,r)=>s+(Number(r[3])||0),0),history:rows.reverse().slice(0,50).map(r=>({points:Number(r[3]),reason:r[5],date:r[6]})),redemptionRate:"1 point = ₦1"};
}
function customerSaveAbandonedCart_(x) {
  const c=requireCustomer_(x.token), items=Array.isArray(x.items)?x.items:[];
  if(!items.length||items.length>30) throw new Error("Cart is empty or too large.");
  const products=sheet_("Products").getDataRange().getValues(), safe=[];
  items.forEach(it=>{
    const p=products.slice(1).find(r=>String(r[0])===String(it.id)&&r[7]===true);
    const q=Number(it.quantity);
    if(!p||!Number.isInteger(q)||q<1||q>50) throw new Error("Invalid saved cart item.");
    safe.push({id:p[0],name:p[1],quantity:q,price:Number(p[4])});
  });
  const sh=sheet_("AbandonedCarts"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0&&String(r[1])===String(c.id)&&String(r[4])==="Open");
  const row=[idx>0?d[idx][0]:Utilities.getUuid(),c.id,c.email,JSON.stringify(safe),"Open",new Date(),""];
  if(idx>0) sh.getRange(idx+1,1,1,row.length).setValues([row]); else sh.appendRow(row);
  return {message:"Your cart has been saved to your account."};
}
function adminSendAbandonedCartReminder_(x) {
  const email=requireAdmin_(x.token), id=String(x.id||""), sh=sheet_("AbandonedCarts"), d=sh.getDataRange().getValues(), idx=d.findIndex((r,i)=>i>0&&String(r[0])===id);
  if(idx<1) throw new Error("Saved cart not found.");
  const r=d[idx]; if(String(r[4])!=="Open") throw new Error("This cart is no longer open.");
  if(!String(r[2])) throw new Error("Customer has no email address.");
  const base=String(x.storefrontUrl||"https://sheggun1.github.io/veehiceetech/marketplace.html");
  const body="You left items in your Veehicee cart. You can return to the marketplace to continue shopping: "+base+"\n\nItems: "+String(r[3]).slice(0,4000);
  MailApp.sendEmail(String(r[2]),"Your Veehicee shopping cart",body);
  sh.getRange(idx+1,7).setValue(new Date()); log_(email,"ABANDONED_CART_REMINDER",id);
  return {message:"Cart reminder email sent."};
}
function adminSendLowStockAlert_(x) {
  const email=requireAdmin_(x.token), threshold=Math.max(0,Math.min(100,Number(x.threshold===undefined?5:x.threshold)||0));
  const d=sheet_("Products").getDataRange().getValues(), low=d.slice(1).filter(r=>r[0]&&r[7]===true&&Number(r[5])<=threshold);
  const body=low.length?low.map(r=>String(r[1])+" — stock "+Number(r[5])).join("\n"):"No products are at or below the selected threshold.";
  MailApp.sendEmail(email,"Veehicee low-stock report",body);
  return {message:"Low-stock report emailed to "+email+".",count:low.length};
}

function adminListCustomers_(x) {
  requireAdmin_(x.token);
  const base=sheet_("Customers").getDataRange().getValues().slice(1);
  const accounts=sheet_("CustomerAccounts").getDataRange().getValues().slice(1);
  const orders=sheet_("Orders").getDataRange().getValues().slice(1);
  const byEmail={};
  base.concat(accounts).forEach(r=>{
    const email=String(r[2]||"").toLowerCase(), phone=String(r[3]||"");
    if(!email&&!phone)return;
    const key=email||phone;
    if(!byEmail[key]) byEmail[key]={name:String(r[1]||""),email:email,phone:phone,hasAccount:false};
    if(r.length>=8 && r[6]===true) byEmail[key].hasAccount=true;
  });
  const customers=Object.values(byEmail).map(c=>{
    const matched=orders.filter(o=>(c.email&&String(o[2]||"").toLowerCase()===c.email)||(!c.email&&c.phone&&String(o[3])===c.phone));
    const paid=matched.filter(o=>["Paid","Confirmed"].includes(String(o[11])));
    const last=matched.reduce((v,o)=>!v||new Date(o[13])>new Date(v)?o[13]:v,null);
    return {name:c.name,email:c.email,phone:c.phone,hasAccount:c.hasAccount,orderCount:matched.length,paidOrderCount:paid.length,totalPaid:paid.reduce((s,o)=>s+(Number(o[9])||0),0),lastOrderAt:last};
  }).sort((a,b)=>b.orderCount-a.orderCount);
  return {customers:customers};
}
function setupVeehiceeCartReminderTrigger() {
  const triggers=ScriptApp.getProjectTriggers();
  if(!triggers.some(t=>t.getHandlerFunction()==="sendDueVeehiceeCartReminders")) {
    ScriptApp.newTrigger("sendDueVeehiceeCartReminders").timeBased().everyDays(1).atHour(10).create();
  }
  return {ok:true,message:"Daily abandoned-cart reminder check scheduled for approximately 10:00 in the Apps Script project time zone."};
}
function sendDueVeehiceeCartReminders() {
  const sh=sheet_("AbandonedCarts"), d=sh.getDataRange().getValues(), cutoff=Date.now()-24*60*60*1000;
  let sent=0;
  for(let i=1;i<d.length;i++) {
    const r=d[i], updated=new Date(r[5]).getTime();
    if(!r[0]||String(r[4])!=="Open"||r[6]||!String(r[2])||!isFinite(updated)||updated>cutoff) continue;
    const url="https://sheggun1.github.io/veehiceetech/marketplace.html";
    const body="You saved items in your Veehicee shopping cart more than 24 hours ago. If you are still interested, visit our marketplace to continue: "+url+"\n\nSaved items: "+String(r[3]).slice(0,3000)+"\n\nThank you for shopping with Veehicee Telecommunications and Services.";
    try { MailApp.sendEmail(String(r[2]),"Your saved Veehicee shopping cart",body); sh.getRange(i+1,7).setValue(new Date()); sent++; } catch(e) {}
  }
  return {sent:sent};
}
