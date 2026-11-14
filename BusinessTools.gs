
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
  return d.filter(r=>String(r[1])===String(customerId)).reduce((s,r)=>s+(Number(r[3])||0),0);
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
