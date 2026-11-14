/* Veehicee extras – load after the main site script. */
(function(){
const K="veehiceeAdminSession",g=id=>document.getElementById(id);
const store={get(){try{return JSON.parse(sessionStorage.getItem(K)||"null")}catch(e){return null}},set(v){try{sessionStorage.setItem(K,JSON.stringify(v))}catch(e){}},clear(){try{sessionStorage.removeItem(K)}catch(e){}}};

/* Add a process section and FAQ without replacing existing content. */
const process=document.createElement("section");process.className="section";process.id="process";
process.innerHTML='<div class="wrap"><span class="section-kicker">HOW WE WORK</span><h2>A clear process from idea to launch.</h2><div class="steps">'+
[["Discover","We listen, study your goals and agree the scope."],["Design","Layouts and visuals are shaped around your brand."],["Build","Clean, tested development with regular check-ins."],["Support","Launch, training and ongoing technical help."]].map(s=>'<div class="step"><h3>'+s[0]+'</h3><p>'+s[1]+'</p></div>').join("")+'</div></div>';
const about=g("about");if(about)about.before(process);
const faq=document.createElement("section");faq.className="section";faq.id="faq";
faq.innerHTML='<div class="wrap"><span class="section-kicker">QUESTIONS</span><h2>Frequently asked questions.</h2><div class="faq">'+
[["How long does a website take?","Most business websites take 1 to 3 weeks depending on pages and content. We agree a timeline before we start."],["Do you build systems for schools?","Yes. We build result systems, CBT, lesson trackers and school websites tailored to how your school works."],["Can I update the site myself?","Yes. We can set up easy content management and train your team."],["Do you offer support after launch?","Yes. We provide ongoing technical support and updates."],["How do I get a quote?","Use the enquiry form below with a short description of what you need."]].map(q=>'<details><summary>'+q[0]+'</summary><p>'+q[1]+'</p></details>').join("")+'</div></div>';
const contact=g("contact");if(contact)contact.before(faq);

/* Admin style bar. */
const bar=document.createElement("div");bar.id="adminBar";
bar.innerHTML='<div class="ab-t" id="abThemes"></div><button class="btn" id="abSave" type="button">Save style</button><button class="ghost" id="abWork" type="button">+ Add work</button><button class="ghost" id="abPanel" type="button">Admin panel</button><button class="ghost" id="abOut" type="button">Sign out</button><small id="abMsg"></small>';
document.body.append(bar);
const NAMES={midnight:"Professional Blue",agency:"Agency Violet",forest:"Emerald",sunrise:"Warm Coral",royal:"Royal Navy & Gold",graphite:"Graphite Minimal"};
const DOT={midnight:"#174ea6",agency:"#5b3df5",forest:"#0f7a55",sunrise:"#d9481f",royal:"#e0b44c",graphite:"#111"};
function paintThemes(){const host=g("abThemes");if(!host)return;host.innerHTML="";THEMES.forEach(t=>{const b=document.createElement("button");b.type="button";b.title=NAMES[t[0]]||t[1];b.style.background=DOT[t[0]]||t[2];b.setAttribute("aria-pressed",t[0]===activeDesign.theme);b.onclick=()=>{applyDesign({...activeDesign,theme:t[0]});paintThemes();g("abMsg").textContent=(NAMES[t[0]]||t[1])+" (preview)"};host.append(b)})}
g("abSave").onclick=async()=>{g("abMsg").textContent="Saving…";await saveSiteDesign();g("abMsg").textContent=g("designStatus").textContent};
g("abPanel").onclick=()=>openAdmin();
g("abWork").onclick=()=>{openAdmin();setTimeout(()=>{const d=[...document.querySelectorAll("#dashboard details")].find(x=>/portfolio project/i.test(x.textContent.slice(0,60)));if(d){d.open=true;d.scrollIntoView({behavior:"smooth",block:"start"})}},120)};
g("abOut").onclick=()=>adminLogout();

/* Keep the admin session for this browser tab only; do not persist the token across browser restarts. */
const _login=window.adminLogin,_logout=window.adminLogout,_pw=window.changeAdminPassword;
window.adminLogin=async function(){await _login();if(adminToken){store.set({t:adminToken,e:(g("adminStatus").textContent||"").replace("Signed in as ","")});document.body.classList.add("is-admin");paintThemes()}};
window.adminLogout=async function(){store.clear();document.body.classList.remove("is-admin");await _logout()};
window.changeAdminPassword=async function(){await _pw();const s=store.get();if(adminToken&&s)store.set({...s,t:adminToken})};
async function loadAll(){for(const n of ["loadPages","loadOrders","loadPortfolioAdmin","loadProductsAdmin","loadInventoryReport","loadBusinessSettings","loadBusinessReport","loadCoupons","loadReviews","loadAbandonedCarts","loadCustomers","loadExpenses"]){try{if(typeof window[n]==="function")await window[n]()}catch(e){}}}
(async function restore(){const s=store.get();if(!s||!s.t)return;adminToken=s.t;try{await api("adminListPages",{token:adminToken})}catch(e){store.clear();adminToken="";return}openAdmin();g("loginPanel").hidden=true;g("dashboard").hidden=false;g("adminStatus").textContent="Signed in"+(s.e?" as "+s.e:"");document.body.classList.add("is-admin");paintThemes();loadAll()})();
paintThemes();
})();

/* Trust strip – client names are entered by the administrator only. */
(function(){
const g=id=>document.getElementById(id);
const strip=document.createElement("section");strip.id="trust";
strip.innerHTML='<div class="wrap"><span class="t-label" id="tLabel">Built for</span><div class="t-items" id="tItems"></div></div>';
const services=g("services");if(services)services.before(strip);
const DEFAULT=["Schools","Businesses","Organisations","Entrepreneurs","Communities"];
function paint(names){const real=(names||"").split(",").map(s=>s.trim()).filter(Boolean);g("tLabel").textContent=real.length?"Trusted by":"Built for";g("tItems").innerHTML="";(real.length?real:DEFAULT).forEach(n=>{const s=document.createElement("span");s.textContent=n;g("tItems").append(s)})}
paint("");
api("publicSettings").then(r=>paint((r.settings||{}).clientNames)).catch(()=>{});
const bar=g("adminBar");if(bar){const b=document.createElement("button");b.className="ghost";b.type="button";b.textContent="Client names";b.onclick=async()=>{const cur=[...g("tItems").children].map(x=>x.textContent).join(", ");const v=prompt("Enter your real client names, separated by commas (leave empty to show the default strip):",cur===DEFAULT.join(", ")?"":cur);if(v===null)return;g("abMsg").textContent="Saving…";try{await api("adminSaveSetting",{token:adminToken,key:"clientNames",value:v.trim()});paint(v);g("abMsg").textContent="Client names saved."}catch(e){g("abMsg").textContent=e.message}};g("abWork").after(b)}
})();
