/* MILO CASHIER - Entire server backend in one Apps Script file. Do NOT also add BackendOps.gs. */
/**
 * MILO CASHIER 1.0 — Google Apps Script / Google Sheets backend.
 * Deploy ONLY as Web App: Execute as Me. Never share the database with store users.
 * All store scoping is derived from the authenticated session, never from the request.
 * Copyright (c) 2026 MILO CASHIER project. See README.
 */
const MILO_VERSION = '1.1.0-beta';
const SESSION_HOURS = 12;
const LOCK_WAIT_MS = 25000;
const BATCH = 400;
const HEADERS = {
  Stores: ['StoreId','Name','Address','Phone','VatNumber','TaxRate','RestaurantMode','InvoicePrefix','NextInvoice','LogoUrl','ReceiptFooter','Active','CreatedAt','UpdatedAt','Version'],
  Users: ['UserId','StoreId','Username','DisplayName','Role','PasswordSalt','PasswordHash','Active','FailedAttempts','LockedUntil','CreatedAt','UpdatedAt','Version'],
  Subscriptions: ['SubscriptionId','StoreId','Plan','StartsAt','EndsAt','Status','CreatedAt','UpdatedAt','Version'],
  Sessions: ['SessionId','StoreId','UserId','TokenHash','ExpiresAt','RevokedAt','DeviceId','CreatedAt'],
  Devices: ['DeviceId','StoreId','Name','Active','CreatedAt','UpdatedAt','Version'],
  Categories: ['CategoryId','StoreId','Name','SortOrder','Active','CreatedAt','UpdatedAt','Version'],
  Products: ['ProductId','StoreId','CategoryId','Name','Barcode','Price','WholesalePrice','Cost','Stock','MinStock','Unlimited','Active','ImageUrl','SizeRequired','CreatedAt','UpdatedAt','Version'],
  ProductOptions: ['OptionId','StoreId','ProductId','Type','Name','Price','Active','SortOrder','CreatedAt','UpdatedAt','Version'],
  Customers: ['CustomerId','StoreId','Name','Phone','Email','Address','CreatedAt','UpdatedAt','Version'],
  Suppliers: ['SupplierId','StoreId','Name','Phone','Email','Address','CreatedAt','UpdatedAt','Version'],
  Sales: ['SaleId','StoreId','InvoiceNo','OrderNo','OperationId','DeviceId','DeviceName','CashierId','CashierName','CustomerId','CustomerName','ShiftId','Subtotal','Discount','Tax','Total','CashPaid','CardPaid','ChangeDue','Status','Note','CreatedAt','UpdatedAt','Version'],
  SaleItems: ['SaleItemId','StoreId','SaleId','ProductId','ProductName','Barcode','Quantity','UnitPrice','UnitCost','SizeName','AddonsJson','Note','LineTotal','CreatedAt'],
  Payments: ['PaymentId','StoreId','SaleId','Method','Amount','CreatedAt'],
  Operations: ['OperationId','StoreId','SaleId','Status','PayloadHash','Error','CreatedAt','UpdatedAt'],
  InventoryMovements: ['MovementId','StoreId','ProductId','Type','Quantity','ReferenceId','Note','UserId','CreatedAt'],
  Purchases: ['PurchaseId','StoreId','SupplierId','SupplierName','TotalCost','Note','CreatedAt','UserId'],
  PurchaseItems: ['PurchaseItemId','StoreId','PurchaseId','ProductId','Quantity','UnitCost','LineTotal'],
  Shifts: ['ShiftId','StoreId','DeviceId','UserId','OpenCash','CloseCash','ExpectedCash','OpenedAt','ClosedAt','Status','Note'],
  HeldOrders: ['HoldId','StoreId','DeviceId','UserId','PayloadJson','Status','CreatedAt','UpdatedAt'],
  Vouchers: ['VoucherId','StoreId','Type','AccountName','Amount','Method','Reference','Note','CreatedAt','UserId'],
  ManualOrders: ['ManualOrderId','StoreId','CustomerName','Description','Amount','Status','Note','CreatedAt','UserId'],
  Returns: ['ReturnId','StoreId','SaleId','Amount','Reason','UserId','CreatedAt'],
  ReturnItems: ['ReturnItemId','StoreId','ReturnId','SaleItemId','ProductId','Quantity','Amount'],
  ActivityLogs: ['LogId','StoreId','UserId','Action','Entity','EntityId','Details','CreatedAt'],
  SyncLog: ['SyncId','StoreId','Entity','EntityId','Action','ChangedAt'],
  Reports: ['ReportId','StoreId','Name','PayloadJson','CreatedAt'],
  Errors: ['ErrorId','Action','StoreId','Message','CreatedAt']
};
const ROLE = { owner: 3, manager: 2, cashier: 1 };
const WRITE_ADMIN = ['owner','manager'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('MILO CASHIER | نظام نقاط البيع')
    .addMetaTag('viewport','width=device-width,initial-scale=1,viewport-fit=cover');
}
function include(file) { return HtmlService.createHtmlOutputFromFile(file).getContent(); }
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) throw new Error('لم يتم إعداد قاعدة البيانات. شغّل setupDatabase من محرر السكربت.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sheet = db_().getSheetByName(name);
  if (!sheet) throw new Error('ورقة غير موجودة: ' + name + '. شغّل setupDatabase.');
  return sheet;
}
function now_(){ return new Date().toISOString(); }
function id_(){ return Utilities.getUuid(); }
function text_(v,max){ return String(v==null?'':v).trim().slice(0,max||300); }
function num_(v){ const n=Number(v); if(!Number.isFinite(n)) throw new Error('قيمة رقمية غير صحيحة'); return n; }
function money_(n){ return Math.round((Number(n)+Number.EPSILON)*100)/100; }
function require_(ok,msg){ if(!ok) throw new Error(msg); }
function safeCell_(v){ if (typeof v==='string' && /^[\s]*[=+@\-]/.test(v)) return "'"+v; return v==null?'':v; }
function mapRow_(name,values){const h=HEADERS[name], obj={}; for(let i=0;i<h.length;i++) obj[h[i]]=values[i] instanceof Date?values[i].toISOString():values[i]; return obj;}
function sheetRows_(name,from,limit){const sh=sheet_(name), last=sh.getLastRow(); if(last<2||from>last)return[]; return sh.getRange(from,1,Math.min(limit,last-from+1),HEADERS[name].length).getValues().map(r=>mapRow_(name,r));}
function allForStore_(name,storeId,limit){let results=[], sh=sheet_(name), last=sh.getLastRow(); for(let row=2;row<=last;row+=BATCH){for(const r of sheetRows_(name,row,BATCH)){if(String(r.StoreId)===String(storeId)) results.push(r); if(limit&&results.length>=limit)return results;}}return results;}
function byId_(name,key,id,storeId){if(!id)return null; const sh=sheet_(name), c=HEADERS[name].indexOf(key)+1; require_(c>0,'معرّف غير معروف');const found=sh.getRange(1,c,sh.getLastRow(),1).createTextFinder(String(id)).matchEntireCell(true).findNext();if(!found)return null;const row=found.getRow(); if(row<2)return null; const data=mapRow_(name,sh.getRange(row,1,1,HEADERS[name].length).getValues()[0]);if(storeId!==undefined&&String(data.StoreId)!==String(storeId))return null;return {row,data};}
function insert_(name,data){const sh=sheet_(name), h=HEADERS[name]; const row=h.map(k=>safeCell_(data[k])); const nr=sh.getLastRow()+1;if(nr>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),Math.max(10,nr-sh.getMaxRows())); sh.getRange(nr,1,1,h.length).setValues([row]);return mapRow_(name,row);}
function update_(name,row,data){const sh=sheet_(name),old=mapRow_(name,sh.getRange(row,1,1,HEADERS[name].length).getValues()[0]);const updated=Object.assign({},old,data);const values=HEADERS[name].map(k=>safeCell_(updated[k]));sh.getRange(row,1,1,values.length).setValues([values]);return mapRow_(name,values);}
function scriptLock_(fn){const lock=LockService.getScriptLock(); if(!lock.tryLock(LOCK_WAIT_MS))throw new Error('النظام مشغول حاليًا. أعد المحاولة بنفس العملية.');try{return fn();}finally{lock.releaseLock();}}
function writeSync_(storeId,entity,entityId,action){insert_('SyncLog',{SyncId:id_(),StoreId:storeId,Entity:entity,EntityId:entityId,Action:action,ChangedAt:now_()});}
function audit_(s,action,entity,eid,details){try{insert_('ActivityLogs',{LogId:id_(),StoreId:s.storeId,UserId:s.userId,Action:action,Entity:entity,EntityId:eid,Details:JSON.stringify(details||{}).slice(0,1200),CreatedAt:now_()});}catch(e){console.error(e);}}
function appendError_(action,storeId,err){try{insert_('Errors',{ErrorId:id_(),Action:action,StoreId:storeId||'',Message:String(err&&err.message||err).slice(0,600),CreatedAt:now_()});}catch(e){console.error(e);}}
function validateRole_(s,min){require_(ROLE[s.role]>=ROLE[min],'ليس لديك صلاحية تنفيذ هذه العملية.');}
function createColumnsSafe_(db,name,h){let sh=db.getSheetByName(name);if(!sh){sh=db.insertSheet(name);if(sh.getMaxColumns()<h.length)sh.insertColumnsAfter(sh.getMaxColumns(),h.length-sh.getMaxColumns());sh.getRange(1,1,1,h.length).setValues([h]);sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setBackground('#122c3d').setFontColor('#ffffff').setFontWeight('bold');const extra=sh.getMaxRows()-2;if(extra>0)sh.deleteRows(3,extra);const extraC=sh.getMaxColumns()-h.length;if(extraC>0)sh.deleteColumns(h.length+1,extraC);}else{const actual=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),h.length)).getValues()[0];for(let i=0;i<h.length;i++)require_(actual[i]===h[i],`مخطط ${name} مختلف عند العمود ${i+1}: لا تُعدل البيانات تلقائيًا.`);}return sh;}
/** Run this ONE TIME manually from Apps Script editor. It is non-destructive. */
function setupDatabase_(){let p=PropertiesService.getScriptProperties(),id=p.getProperty('DB_ID'),db=id?SpreadsheetApp.openById(id):SpreadsheetApp.getActiveSpreadsheet();if(!db){db=SpreadsheetApp.create('MILO CASHIER • Private Database');p.setProperty('DB_ID',db.getId());}else if(!id){p.setProperty('DB_ID',db.getId());}Object.keys(HEADERS).forEach(name=>createColumnsSafe_(db,name,HEADERS[name]));if(!p.getProperty('PASSWORD_PEPPER'))p.setProperty('PASSWORD_PEPPER',id_()+id_());console.log('DB_ID='+db.getId()+' | URL='+db.getUrl());return {spreadsheetId:db.getId(),url:db.getUrl(),sheets:Object.keys(HEADERS).length};}
function sha_(value){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value).map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');}
function pwdHash_(pwd,salt){const pepper=PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER');require_(pepper,'إعداد أمني ناقص'); let v=String(pwd)+':'+salt+':'+pepper; // server-side salted, peppered repeated SHA-256; review/upgrade KDF for commercial release
  for(let i=0;i<1200;i++)v=sha_(v+':'+pepper);return v;}
function tokenHash_(token){return sha_(token+':'+PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER'));}
function strongPassword_(pwd){require_(typeof pwd==='string'&&pwd.length>=12&&pwd.length<=128,'كلمة المرور يجب أن تتكون من 12 حرفًا على الأقل.');}
/** Owner-provisioning only; never callable from the web. Set INITIAL_STORE_NAME, INITIAL_USERNAME and INITIAL_PASSWORD in script properties, then run manually. */
function provisionInitialStore_(){return scriptLock_(()=>{const p=PropertiesService.getScriptProperties(),name=text_(p.getProperty('INITIAL_STORE_NAME'),90),username=text_(p.getProperty('INITIAL_USERNAME'),60).toLowerCase(),pwd=p.getProperty('INITIAL_PASSWORD');require_(name&&username&&pwd,'ضع INITIAL_STORE_NAME وINITIAL_USERNAME وINITIAL_PASSWORD في خصائص السكربت');strongPassword_(pwd);require_(/^[a-z0-9_.-]{3,60}$/.test(username),'اسم المستخدم يجب أن يكون إنجليزيًا بدون مسافات');require_(!allForStore_('Stores','no-such-store').length,''); // harmless schema check
  require_(sheet_('Stores').getLastRow()===1,'الدالة الأولى مخصصة لأول متجر فقط؛ استخدم provisionStoreAdmin للمتاجر اللاحقة');const sid=id_(),uid=id_(),t=now_(),salt=id_();insert_('Stores',{StoreId:sid,Name:name,Address:'',Phone:'',VatNumber:'',TaxRate:15,RestaurantMode:false,InvoicePrefix:'MLO',NextInvoice:1,LogoUrl:'',ReceiptFooter:'شكرًا لزيارتكم',Active:true,CreatedAt:t,UpdatedAt:t,Version:1});insert_('Users',{UserId:uid,StoreId:sid,Username:username,DisplayName:'المالك',Role:'owner',PasswordSalt:salt,PasswordHash:pwdHash_(pwd,salt),Active:true,FailedAttempts:0,LockedUntil:'',CreatedAt:t,UpdatedAt:t,Version:1});insert_('Subscriptions',{SubscriptionId:id_(),StoreId:sid,Plan:'trial',StartsAt:t,EndsAt:new Date(Date.now()+30*86400000).toISOString(),Status:'active',CreatedAt:t,UpdatedAt:t,Version:1});p.deleteProperty('INITIAL_PASSWORD');console.log('Created StoreId='+sid);return sid;});}
/** Admin-only editor function, not registered in API. Set ADMIN_STORE_NAME, ADMIN_USERNAME, ADMIN_PASSWORD first. */
function provisionStoreAdmin_(){return scriptLock_(()=>{const p=PropertiesService.getScriptProperties(),name=text_(p.getProperty('ADMIN_STORE_NAME'),90),username=text_(p.getProperty('ADMIN_USERNAME'),60).toLowerCase(),pwd=p.getProperty('ADMIN_PASSWORD');require_(name&&username&&pwd,'خصائص المتجر الجديد ناقصة');require_(/^[a-z0-9_.-]{3,60}$/.test(username),'اسم المستخدم غير صالح');strongPassword_(pwd);require_(!findUserByUsername_(username),'اسم المستخدم مستخدم');const t=now_(),sid=id_(),salt=id_();insert_('Stores',{StoreId:sid,Name:name,Address:'',Phone:'',VatNumber:'',TaxRate:15,RestaurantMode:false,InvoicePrefix:'MLO',NextInvoice:1,ReceiptFooter:'شكرًا لزيارتكم',Active:true,CreatedAt:t,UpdatedAt:t,Version:1});insert_('Users',{UserId:id_(),StoreId:sid,Username:username,DisplayName:'المالك',Role:'owner',PasswordSalt:salt,PasswordHash:pwdHash_(pwd,salt),Active:true,FailedAttempts:0,LockedUntil:'',CreatedAt:t,UpdatedAt:t,Version:1});insert_('Subscriptions',{SubscriptionId:id_(),StoreId:sid,Plan:'trial',StartsAt:t,EndsAt:new Date(Date.now()+30*86400000).toISOString(),Status:'active',CreatedAt:t,UpdatedAt:t,Version:1});p.deleteProperty('ADMIN_PASSWORD');console.log('New StoreId='+sid);return sid;});}
/** Renew from the script editor only (example ISO end date 2027-01-01T00:00:00.000Z). */
function renewStoreSubscription_(storeId,endsAt){return scriptLock_(()=>{require_(storeId&&endsAt&&Number.isFinite(Date.parse(endsAt)),'معرف متجر أو تاريخ غير صالح');let r=allForStore_('Subscriptions',storeId).sort((a,b)=>String(b.EndsAt).localeCompare(String(a.EndsAt)))[0];require_(r,'الاشتراك غير موجود');const ref=byId_('Subscriptions','SubscriptionId',r.SubscriptionId,storeId),t=now_();return update_('Subscriptions',ref.row,{EndsAt:new Date(endsAt).toISOString(),Status:'active',UpdatedAt:t,Version:Number(r.Version||0)+1});});}
function findUserByUsername_(username){const sh=sheet_('Users'),i=HEADERS.Users.indexOf('Username')+1,m=sh.getRange(1,i,sh.getLastRow(),1).createTextFinder(username).matchEntireCell(true).findNext();return m&&m.getRow()>1?{row:m.getRow(),data:mapRow_('Users',sh.getRange(m.getRow(),1,1,HEADERS.Users.length).getValues()[0])}:null;}
function login(payload){return scriptLock_(()=>{const u=text_(payload&&payload.username,60).toLowerCase(),pwd=payload&&payload.password,dev=text_(payload&&payload.deviceId,60);require_(u&&typeof pwd==='string','أدخل اسم المستخدم وكلمة المرور');const ref=findUserByUsername_(u);require_(ref&&ref.data.Active===true,'بيانات الدخول غير صحيحة.');const user=ref.data,t=now_();require_(!user.LockedUntil||String(user.LockedUntil)<=t,'الحساب مقفل مؤقتًا بعد محاولات متكررة.');if(pwdHash_(pwd,String(user.PasswordSalt))!==String(user.PasswordHash)){const fails=Number(user.FailedAttempts||0)+1;update_('Users',ref.row,{FailedAttempts:fails,LockedUntil:fails>=5?new Date(Date.now()+15*60000).toISOString():'',UpdatedAt:t});throw new Error('بيانات الدخول غير صحيحة.');}update_('Users',ref.row,{FailedAttempts:0,LockedUntil:'',UpdatedAt:t});const store=byId_('Stores','StoreId',user.StoreId);require_(store&&store.data.Active===true,'المتجر غير مفعل.');const chosen=dev?byId_('Devices','DeviceId',dev,user.StoreId):null;const verifiedDevice=chosen&&chosen.data.Active===true?dev:'';const token=id_()+id_();insert_('Sessions',{SessionId:id_(),StoreId:user.StoreId,UserId:user.UserId,TokenHash:tokenHash_(token),ExpiresAt:new Date(Date.now()+SESSION_HOURS*3600000).toISOString(),RevokedAt:'',DeviceId:verifiedDevice,CreatedAt:t});audit_({storeId:user.StoreId,userId:user.UserId},'LOGIN','Users',user.UserId,{deviceId:dev});return {token,profile:{userId:user.UserId,name:user.DisplayName,role:user.Role,storeId:user.StoreId,storeName:store.data.Name},expired:!activeSubscription_(user.StoreId),version:MILO_VERSION};});}
function activeSubscription_(storeId){const t=now_();return allForStore_('Subscriptions',storeId).some(s=>s.Status==='active'&&String(s.StartsAt)<=t&&String(s.EndsAt)>t);}
function session_(token,allowExpired){require_(typeof token==='string'&&token.length>30,'الرجاء تسجيل الدخول.');const hash=tokenHash_(token);let sh=sheet_('Sessions'),last=sh.getLastRow(),ref=null,col=HEADERS.Sessions.indexOf('TokenHash')+1;const hit=sh.getRange(1,col,last,1).createTextFinder(hash).matchEntireCell(true).findNext();if(hit&&hit.getRow()>1)ref=mapRow_('Sessions',sh.getRange(hit.getRow(),1,1,HEADERS.Sessions.length).getValues()[0]);require_(ref&&!ref.RevokedAt&&String(ref.ExpiresAt)>now_(),'انتهت الجلسة. سجل الدخول مجددًا.');const user=byId_('Users','UserId',ref.UserId,ref.StoreId);require_(user&&user.data.Active===true,'المستخدم غير نشط.');const store=byId_('Stores','StoreId',ref.StoreId);require_(store&&store.data.Active===true,'المتجر غير نشط.');const subscriptionActive=activeSubscription_(ref.StoreId);if(!allowExpired)require_(subscriptionActive,'انتهى الاشتراك. يرجى التواصل مع الإدارة للتجديد.');return {storeId:ref.StoreId,userId:ref.UserId,role:user.data.Role,name:user.data.DisplayName,username:user.data.Username,deviceId:ref.DeviceId,store:store.data,subscriptionActive,sessionId:ref.SessionId};}
function publicUser_(u){return {UserId:u.UserId,Username:u.Username,DisplayName:u.DisplayName,Role:u.Role,Active:u.Active,Version:u.Version};}
/** SINGLE external RPC. Always authorize on the server. */
function api(action,payload,token){let s=null;try{require_(typeof action==='string'&&action.length<50,'طلب غير صالح');s=session_(token,action==='auth.me'||action==='auth.logout');const p=payload||{};switch(action){
  case 'auth.me':return {profile:{name:s.name,role:s.role,userId:s.userId,storeId:s.storeId,storeName:s.store.Name},subscriptionActive:s.subscriptionActive,store:publicStore_(s.store),subscription:allForStore_('Subscriptions',s.storeId),version:MILO_VERSION};
  case 'auth.logout': {const r=byId_('Sessions','SessionId',s.sessionId,s.storeId);if(r)update_('Sessions',r.row,{RevokedAt:now_()});return {ok:true};}
  case 'bootstrap':return bootstrap_(s,p);
  case 'sync.pull':return syncPull_(s,p);
  case 'catalog.list':return catalog_(s,p);
  case 'products.save':return saveProduct_(s,p);
  case 'products.options':return allForStore_('ProductOptions',s.storeId).filter(x=>x.ProductId===p.productId&&x.Active===true);
  case 'products.archive':return archiveCatalog_(s,'Products',p);
  case 'categories.list':return allForStore_('Categories',s.storeId).filter(x=>x.Active===true);
  case 'categories.save':return saveGeneric_(s,'Categories',p);
  case 'categories.archive':return archiveCatalog_(s,'Categories',p);
  case 'customers.list':return allForStore_('Customers',s.storeId,1000);
  case 'customers.save':return saveGeneric_(s,'Customers',p);
  case 'suppliers.list':validateRole_(s,'manager');return allForStore_('Suppliers',s.storeId,1000);
  case 'suppliers.save':validateRole_(s,'manager');return saveGeneric_(s,'Suppliers',p);
  case 'users.list':validateRole_(s,'manager');return allForStore_('Users',s.storeId).map(publicUser_);
  case 'users.save':return saveUser_(s,p);
  case 'devices.list':return allForStore_('Devices',s.storeId);
  case 'devices.select':return bindDevice_(s,p);
  case 'devices.save':return saveGeneric_(s,'Devices',p);
  case 'stores.save':return saveStore_(s,p);
  case 'sales.create':return createSale_(s,p);
  case 'sales.list':return saleList_(s,p);
  case 'sales.detail':return saleDetail_(s,p.saleId);
  case 'sales.operation':return saleOperation_(s,p);
  case 'sales.cancel':return cancelSale_(s,p);
  case 'sales.return':return returnSale_(s,p);
  case 'held.save':return holdOrder_(s,p);
  case 'held.list':return allForStore_('HeldOrders',s.storeId).filter(r=>r.Status==='held'&&(!p.deviceId||r.DeviceId===p.deviceId));
  case 'held.remove':return deleteHold_(s,p);
  case 'inventory.adjust':return adjustStock_(s,p);
  case 'inventory.movements':return allForStore_('InventoryMovements',s.storeId,1000).reverse().slice(0,200);
  case 'purchases.create':return createPurchase_(s,p);
  case 'purchases.list':validateRole_(s,'manager');return allForStore_('Purchases',s.storeId,500).reverse();
  case 'shifts.open':return openShift_(s,p);
  case 'shifts.close':return closeShift_(s,p);
  case 'shifts.list':return allForStore_('Shifts',s.storeId,500).reverse();
  case 'subscriptions.current':return allForStore_('Subscriptions',s.storeId).map(x=>({Plan:x.Plan,StartsAt:x.StartsAt,EndsAt:x.EndsAt,Status:x.Status}));
  case 'vouchers.save':return createVoucher_(s,p);
  case 'vouchers.list':validateRole_(s,'manager');return allForStore_('Vouchers',s.storeId,500).reverse();
  case 'manual.save':return saveManual_(s,p);
  case 'manual.list':return allForStore_('ManualOrders',s.storeId,500).reverse();
  case 'reports.overview':validateRole_(s,'manager');return reports_(s,p);
  case 'audit.list':validateRole_(s,'owner');return allForStore_('ActivityLogs',s.storeId,300).reverse().slice(0,100);
  case 'system.capacity':validateRole_(s,'owner');return capacity_();
  default:throw new Error('الأمر المطلوب غير مدعوم.');
}}catch(e){appendError_(action,s&&s.storeId,e);throw new Error(e.message||'حصل خطأ في العملية.');}}
function publicStore_(st){const {StoreId,Name,Address,Phone,VatNumber,TaxRate,RestaurantMode,InvoicePrefix,LogoUrl,ReceiptFooter,Version}=st;return {StoreId,Name,Address,Phone,VatNumber,TaxRate,RestaurantMode,InvoicePrefix,LogoUrl,ReceiptFooter,Version};}
function bootstrap_(s,p){const data=catalog_(s,{offset:0,limit:300});return {store:publicStore_(s.store),user:{name:s.name,role:s.role,userId:s.userId},devices:allForStore_('Devices',s.storeId).filter(x=>x.Active===true),customers:allForStore_('Customers',s.storeId,500),catalog:data,cursor:sheet_('SyncLog').getLastRow()-1,subscriptionActive:s.subscriptionActive};}
function catalog_(s,p){const offset=Math.max(0,Math.floor(Number(p.offset)||0)),limit=Math.min(300,Math.max(10,Number(p.limit)||100));const products=allForStore_('Products',s.storeId).filter(r=>r.Active===true);return {products:products.slice(offset,offset+limit),total:products.length,nextOffset:offset+limit<products.length?offset+limit:null,categories:allForStore_('Categories',s.storeId).filter(x=>x.Active===true),options:allForStore_('ProductOptions',s.storeId).filter(x=>x.Active===true)};}
function syncPull_(s,p){const sh=sheet_('SyncLog'),total=sh.getLastRow()-1,cursor=Math.max(0,Math.floor(Number(p.cursor)||0));require_(cursor<=total,'مؤشر المزامنة غير صالح');const limit=Math.min(200,Math.max(10,Number(p.limit)||100)),changes=[];let row=cursor+2,examined=0;while(row<=total+1&&examined<limit){const chunk=Math.min(BATCH,total+2-row,limit-examined);for(const e of sheetRows_('SyncLog',row,chunk)){if(e.StoreId===s.storeId){let rec=null;const idKey={Products:'ProductId',Categories:'CategoryId',ProductOptions:'OptionId',Customers:'CustomerId',Stores:'StoreId',Devices:'DeviceId'}[e.Entity];if(idKey){const found=byId_(e.Entity,idKey,e.EntityId,s.storeId);rec=found?found.data:null;}changes.push({entity:e.Entity,action:e.Action,id:e.EntityId,record:rec});}}row+=chunk;examined+=chunk;}return {changes,cursor:Math.min(total,row-2),hasMore:row<=total+1};}
function safeListText_(value,max){return text_(value,max).replace(/[\u0000-\u001F]/g,'');}
function saveGeneric_(s,table,p){validateRole_(s,'manager');const spec={Categories:{key:'CategoryId',fields:['Name','SortOrder','Active']},Customers:{key:'CustomerId',fields:['Name','Phone','Email','Address']},Suppliers:{key:'SupplierId',fields:['Name','Phone','Email','Address']},Devices:{key:'DeviceId',fields:['Name','Active']}}[table];require_(spec,'نوع سجل غير صالح');return scriptLock_(()=>{const id=p[spec.key],ref=id?byId_(table,spec.key,id,s.storeId):null;require_(!id||ref,'السجل غير موجود في المتجر');if(ref)require_(Number(ref.data.Version)===Number(p.Version),'السجل تغير على جهاز آخر. حدّث الصفحة.');const row={StoreId:s.storeId,[spec.key]:id||id_(),CreatedAt:ref?ref.data.CreatedAt:now_(),UpdatedAt:now_(),Version:ref?Number(ref.data.Version)+1:1};for(const f of spec.fields){if(f==='Active')row.Active=p.Active!==false;else if(f==='SortOrder')row.SortOrder=Math.max(0,Number(p.SortOrder)||0);else row[f]=safeListText_(p[f],f==='Name'?100:200);}require_(row.Name,'الاسم مطلوب');if(table==='Devices'&&p.Name.length>60)throw new Error('اسم الكاشير طويل');const saved=ref?update_(table,ref.row,row):insert_(table,row);writeSync_(s.storeId,table,saved[spec.key],ref?'update':'create');audit_(s,ref?'UPDATE':'CREATE',table,saved[spec.key],{});return saved;});}
function saveUser_(s,p){validateRole_(s,'owner');return scriptLock_(()=>{const username=text_(p.Username,60).toLowerCase();require_(/^[a-z0-9_.-]{3,60}$/.test(username),'اسم المستخدم بالإنجليزية 3 أحرف على الأقل');require_(['owner','manager','cashier'].includes(p.Role),'صلاحية غير صحيحة');const ref=p.UserId?byId_('Users','UserId',p.UserId,s.storeId):null;require_(!p.UserId||ref,'المستخدم غير موجود');const found=findUserByUsername_(username);require_(!found||found.data.UserId===(p.UserId||''),'اسم المستخدم مستخدم');require_(!ref||Number(ref.data.Version)===Number(p.Version),'المستخدم تغير على جهاز آخر');if(ref&&ref.data.UserId===s.userId)require_(p.Active!==false&&p.Role==='owner','لا يمكن تعطيل أو تخفيض حسابك بنفسك');if(!ref||p.Password)strongPassword_(p.Password);const salt=p.Password?id_():ref&&ref.data.PasswordSalt;const data={StoreId:s.storeId,UserId:p.UserId||id_(),Username:username,DisplayName:safeListText_(p.DisplayName,90),Role:p.Role,PasswordSalt:salt,PasswordHash:p.Password?pwdHash_(p.Password,salt):ref.data.PasswordHash,Active:p.Active!==false,FailedAttempts:0,LockedUntil:'',CreatedAt:ref?ref.data.CreatedAt:now_(),UpdatedAt:now_(),Version:ref?Number(ref.data.Version)+1:1};require_(data.DisplayName,'اسم المستخدم الظاهر مطلوب');const out=ref?update_('Users',ref.row,data):insert_('Users',data);audit_(s,'SAVE_USER','Users',out.UserId,{role:out.Role});return publicUser_(out);});}
function saveStore_(s,p){validateRole_(s,'owner');return scriptLock_(()=>{const ref=byId_('Stores','StoreId',s.storeId),old=ref.data;require_(Number(p.Version)===Number(old.Version),'إعدادات المتجر تغيرت على جهاز آخر');let vat=safeListText_(p.VatNumber,30);require_(!vat||/^3\d{13}3$/.test(vat),'الرقم الضريبي السعودي يجب أن يكون 15 رقمًا ويبدأ وينتهي بـ3');const tax=Number(p.TaxRate);require_(Number.isFinite(tax)&&tax>=0&&tax<=25,'نسبة الضريبة غير صحيحة');const next={Name:safeListText_(p.Name,90),Address:safeListText_(p.Address,160),Phone:safeListText_(p.Phone,30),VatNumber:vat,TaxRate:tax,RestaurantMode:p.RestaurantMode===true,InvoicePrefix:safeListText_(p.InvoicePrefix,10).replace(/[^A-Za-z0-9-]/g,'')||'MLO',LogoUrl:safeListText_(p.LogoUrl,300),ReceiptFooter:safeListText_(p.ReceiptFooter,120),UpdatedAt:now_(),Version:Number(old.Version)+1};require_(next.Name,'اسم المحل مطلوب');const r=update_('Stores',ref.row,next);writeSync_(s.storeId,'Stores',s.storeId,'update');audit_(s,'STORE_SETTINGS','Stores',s.storeId,{restaurantMode:next.RestaurantMode});return publicStore_(r);});}
function optionSafe_(o){require_(o&&['size','addon'].includes(o.Type),'نوع الخيار غير صالح');const price=num_(o.Price||0);require_(price>=0&&price<=100000,'سعر الخيار غير صالح');return {OptionId:o.OptionId||id_(),Type:o.Type,Name:safeListText_(o.Name,80),Price:money_(price),Active:o.Active!==false,SortOrder:Number(o.SortOrder)||0};}
function saveProduct_(s,p){validateRole_(s,'manager');return scriptLock_(()=>{const ref=p.ProductId?byId_('Products','ProductId',p.ProductId,s.storeId):null;require_(!p.ProductId||ref,'المنتج غير موجود');if(ref)require_(Number(ref.data.Version)===Number(p.Version),'المنتج عُدّل في كاشير آخر');const price=num_(p.Price),wholesale=num_(p.WholesalePrice||0),cost=num_(p.Cost||0),min=num_(p.MinStock||0);require_(price>=0&&wholesale>=0&&cost>=0&&min>=0,'القيم النقدية لا يمكن أن تكون سالبة');const category=p.CategoryId?byId_('Categories','CategoryId',p.CategoryId,s.storeId):null;require_(!p.CategoryId||category,'التصنيف غير موجود');const barcode=safeListText_(p.Barcode,60);if(barcode){const dupe=allForStore_('Products',s.storeId).find(x=>x.Barcode===barcode&&x.ProductId!==p.ProductId);require_(!dupe,'الباركود مسجل على منتج آخر');}const newOptions=Array.isArray(p.Options)?p.Options.map(optionSafe_):[];require_(newOptions.length<=80,'عدد الخيارات أكبر من المسموح');require_(newOptions.every(o=>o.Name),'يجب تسمية جميع الأحجام والإضافات');const productId=p.ProductId||id_(),row={ProductId:productId,StoreId:s.storeId,CategoryId:p.CategoryId||'',Name:safeListText_(p.Name,120),Barcode:barcode,Price:money_(price),WholesalePrice:money_(wholesale),Cost:money_(cost),Stock:ref?ref.data.Stock:Math.floor(num_(p.Stock||0)),MinStock:Math.floor(min),Unlimited:p.Unlimited===true,Active:p.Active!==false,ImageUrl:safeListText_(p.ImageUrl,350),SizeRequired:p.SizeRequired===true,CreatedAt:ref?ref.data.CreatedAt:now_(),UpdatedAt:now_(),Version:ref?Number(ref.data.Version)+1:1};require_(row.Name,'اسم المنتج مطلوب');require_(Number(row.Stock)>=0,'المخزون لا يمكن أن يكون سالبًا');const out=ref?update_('Products',ref.row,row):insert_('Products',row);if(!ref&&row.Stock>0)insert_('InventoryMovements',{MovementId:id_(),StoreId:s.storeId,ProductId:productId,Type:'opening',Quantity:row.Stock,ReferenceId:productId,Note:'رصيد افتتاحي',UserId:s.userId,CreatedAt:now_()}); // Upsert, archive removed options; never delete history.
  const old=allForStore_('ProductOptions',s.storeId).filter(x=>x.ProductId===productId);const ids=new Set();for(const o of newOptions){const prev=old.find(z=>z.OptionId===o.OptionId);if(!prev)o.OptionId=id_();const saved={...o,StoreId:s.storeId,ProductId:productId,CreatedAt:prev?prev.CreatedAt:now_(),UpdatedAt:now_(),Version:prev?Number(prev.Version)+1:1};if(prev){const oldRow=byId_('ProductOptions','OptionId',prev.OptionId,s.storeId);update_('ProductOptions',oldRow.row,saved);}else{insert_('ProductOptions',saved);}ids.add(o.OptionId);writeSync_(s.storeId,'ProductOptions',o.OptionId,'update');}for(const prev of old){if(!ids.has(prev.OptionId)&&prev.Active===true){const oldRow=byId_('ProductOptions','OptionId',prev.OptionId,s.storeId);update_('ProductOptions',oldRow.row,{Active:false,UpdatedAt:now_(),Version:Number(prev.Version)+1});writeSync_(s.storeId,'ProductOptions',prev.OptionId,'update');}}writeSync_(s.storeId,'Products',productId,ref?'update':'create');audit_(s,'SAVE_PRODUCT','Products',productId,{});return {product:out,options:allForStore_('ProductOptions',s.storeId).filter(x=>x.ProductId===productId&&x.Active===true)};});}

/** MILO CASHIER — verified central sale and commercial operations. */
function appendBatch_(name,records){
  if(!records.length)return;
  const sh=sheet_(name), headers=HEADERS[name],start=sh.getLastRow()+1;
  if(start+records.length-1>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),Math.max(10,start+records.length-1-sh.getMaxRows()));
  sh.getRange(start,1,records.length,headers.length).setValues(records.map(r=>headers.map(h=>safeCell_(r[h]))));
}
function positiveInt_(value,max){const n=Number(value);require_(Number.isSafeInteger(n)&&n>=1&&n<=max,'الكمية غير صحيحة');return n;}
function exactMoney_(v,max){const n=Number(v);require_(Number.isFinite(n)&&n>=0&&n<=max,'مبلغ غير صحيح');return money_(n);}
function receiptNo_(store,sequence){return (store.InvoicePrefix||'MLO')+'-'+String(sequence).padStart(7,'0');}
function saleOperation_(s,p){const op=text_(p.OperationId,80);require_(op,'معرّف العملية مطلوب');const found=byId_('Operations','OperationId',op,s.storeId);return found?{status:found.data.Status,saleId:found.data.SaleId||'',error:found.data.Error||''}:{status:'not_found'};}
function createSale_(s,p){
 return scriptLock_(()=>{
  const op=text_(p.OperationId,80);
  require_(/^[0-9a-f-]{36}$/i.test(op),'معرّف العملية غير صالح');
  const prev=byId_('Operations','OperationId',op,s.storeId);
  if(prev){
   if(prev.data.Status==='completed')return {duplicate:true,...saleDetail_(s,prev.data.SaleId)};
   throw new Error('عملية بيع قيد التحقق، لا تنشئ عملية جديدة. راجع رقم العملية في المبيعات أو اتصل بالدعم: '+op);
  }
  const lines=Array.isArray(p.Items)?p.Items:[];
  require_(lines.length>0&&lines.length<=40,'الفاتورة تتطلب من 1 إلى 40 صنفًا');
  const device=byId_('Devices','DeviceId',text_(p.DeviceId,80),s.storeId);
  require_(device&&device.data.Active===true,'اختر كاشيرًا مسجلًا ومفعّلًا');
  require_(s.deviceId===device.data.DeviceId,'الكاشير غير مرتبط بالجلسة الحالية؛ أعد اختياره');
  const active=byId_('Shifts','ShiftId',text_(p.ShiftId,80),s.storeId);
  require_(active&&active.data.Status==='open'&&active.data.DeviceId===device.data.DeviceId,'يجب فتح وردية على الكاشير المختار');
  const productRefs=new Map(),priced=[],need=new Map(),options=allForStore_('ProductOptions',s.storeId).filter(x=>x.Active===true);
  for(const input of lines){
   require_(input&&typeof input==='object','تفاصيل صنف غير صحيحة');
   const id=text_(input.ProductId,80),qty=positiveInt_(input.Quantity,9999);
   let ref=productRefs.get(id);
   if(!ref){ref=byId_('Products','ProductId',id,s.storeId);require_(ref&&ref.data.Active===true,'منتج غير موجود أو غير نشط');productRefs.set(id,ref);}
   const prod=ref.data,chosen=Array.isArray(input.OptionIds)?input.OptionIds:[];
   require_(chosen.length<=25&&new Set(chosen).size===chosen.length,'الخيارات غير صالحة');
   const available=options.filter(o=>o.ProductId===id), selected=chosen.map(x=>available.find(o=>o.OptionId===x));
   require_(selected.every(Boolean),'تم تغيير خيارات أحد المنتجات؛ حدّث الأصناف');
   require_(!selected.length||s.store.RestaurantMode===true,'خيارات المطعم غير مفعّلة');
   const sizes=selected.filter(o=>o.Type==='size'),addons=selected.filter(o=>o.Type==='addon');
   require_(sizes.length<=1,'يمكن اختيار حجم واحد فقط');
   if(prod.SizeRequired===true)require_(sizes.length===1,'هذا الصنف يتطلب اختيار حجم');
   const wholesale=input.PriceMode==='wholesale';
   if(wholesale)validateRole_(s,'manager');
   let unit=wholesale?Number(prod.WholesalePrice):Number(prod.Price);
   if(sizes.length)unit=Number(sizes[0].Price);
   unit+=addons.reduce((sum,x)=>sum+Number(x.Price),0);
   require_(Number.isFinite(unit)&&unit>=0&&unit<=100000,'سعر الصنف غير صالح');
   const total=money_(unit*qty);need.set(id,(need.get(id)||0)+qty);
   priced.push({ProductId:id,ProductName:prod.Name,Barcode:prod.Barcode,Quantity:qty,UnitPrice:money_(unit),UnitCost:money_(prod.Cost),SizeName:sizes[0]?.Name||'',AddonsJson:JSON.stringify(addons.map(a=>({name:a.Name,price:Number(a.Price)}))),Note:text_(input.Note,180),LineTotal:total});
  }
  for(const [id,q] of need){const prod=productRefs.get(id).data;require_(prod.Unlimited===true||Number(prod.Stock)>=q,'المخزون غير كافٍ للصنف: '+prod.Name);}
  const subtotal=money_(priced.reduce((sum,x)=>sum+x.LineTotal,0));
  const discount=exactMoney_(p.Discount||0,1000000);require_(discount<=subtotal,'الخصم أعلى من إجمالي الأصناف');
  const total=money_(subtotal-discount),rate=Number(s.store.TaxRate||0);
  if(p.ExpectedTotal!==undefined)require_(Math.abs(Number(p.ExpectedTotal)-total)<0.011,'تغيّرت أسعار المنتجات أو الإجمالي. حدث السلة قبل إتمام البيع');
  require_(rate>=0&&rate<=25,'إعداد نسبة الضريبة غير صحيح');
  const tax=money_(rate?total*rate/(100+rate):0);
  const cash=exactMoney_(p.CashPaid||0,10000000),card=exactMoney_(p.CardPaid||0,10000000);
  require_(card<=total+0.001&&cash+card+0.001>=total,'المبلغ المدفوع لا يغطي الإجمالي أو الشبكة تتجاوز المستحق');
  const change=money_(Math.max(0,cash+card-total));
  const storeRef=byId_('Stores','StoreId',s.storeId),sequence=positiveInt_(storeRef.data.NextInvoice||1,999999999);
  const saleId=id_(),created=now_(),invoiceNo=receiptNo_(s.store,sequence);
  // Google Sheets offers no multi-tab ACID transaction. A pending operation is intentionally
  // NOT acknowledged as a sale, and must be reviewed/reconciled before any retry with a new ID.
  insert_('Operations',{OperationId:op,StoreId:s.storeId,SaleId:saleId,Status:'pending',PayloadHash:sha_(JSON.stringify(p)),Error:'',CreatedAt:created,UpdatedAt:created});
  update_('Stores',storeRef.row,{NextInvoice:sequence+1,UpdatedAt:created,Version:Number(storeRef.data.Version)+1});
  const header={SaleId:saleId,StoreId:s.storeId,InvoiceNo:invoiceNo,OrderNo:invoiceNo,OperationId:op,DeviceId:device.data.DeviceId,DeviceName:device.data.Name,
   CashierId:s.userId,CashierName:s.name,CustomerId:'',CustomerName:text_(p.CustomerName,120),ShiftId:active.data.ShiftId,
   Subtotal:subtotal,Discount:discount,Tax:tax,Total:total,CashPaid:cash,CardPaid:card,ChangeDue:change,Status:'pending',Note:text_(p.Note,200),CreatedAt:created,UpdatedAt:created,Version:1};
  insert_('Sales',header);
  appendBatch_('SaleItems',priced.map(it=>({...it,SaleItemId:id_(),StoreId:s.storeId,SaleId:saleId,CreatedAt:created})));
  appendBatch_('Payments',[...(cash>0?[{PaymentId:id_(),StoreId:s.storeId,SaleId:saleId,Method:'cash',Amount:money_(cash-change),CreatedAt:created}]:[]),
   ...(card>0?[{PaymentId:id_(),StoreId:s.storeId,SaleId:saleId,Method:'card',Amount:card,CreatedAt:created}]:[])]);
  for(const [id,q] of need){const ref=productRefs.get(id);if(ref.data.Unlimited===true)continue;
   update_('Products',ref.row,{Stock:Number(ref.data.Stock)-q,Version:Number(ref.data.Version)+1,UpdatedAt:created});
   insert_('InventoryMovements',{MovementId:id_(),StoreId:s.storeId,ProductId:id,Type:'sale',Quantity:-q,ReferenceId:saleId,Note:invoiceNo,UserId:s.userId,CreatedAt:created});
   writeSync_(s.storeId,'Products',id,'update');
  }
  const saved=byId_('Sales','SaleId',saleId,s.storeId);update_('Sales',saved.row,{Status:'completed',UpdatedAt:now_(),Version:2});
  const opr=byId_('Operations','OperationId',op,s.storeId);update_('Operations',opr.row,{Status:'completed',UpdatedAt:now_()});
  audit_(s,'SALE_CREATED','Sales',saleId,{invoiceNo,total,deviceId:device.data.DeviceId});
  SpreadsheetApp.flush();
  return {duplicate:false,...saleDetail_(s,saleId)};
 });
}
function saleDetail_(s,saleId){const ref=byId_('Sales','SaleId',saleId,s.storeId);require_(ref,'الفاتورة غير موجودة');const sale=ref.data;
 const items=allForStore_('SaleItems',s.storeId).filter(x=>x.SaleId===saleId),payments=allForStore_('Payments',s.storeId).filter(x=>x.SaleId===saleId);
 return {sale,items,payments};}
function saleList_(s,p){
 const from=text_(p.from,10),to=text_(p.to,10),dev=text_(p.deviceId,80),cashier=text_(p.cashierId,80),query=text_(p.query,120).toLowerCase(),method=text_(p.method,12);
 require_((!from||/^\d{4}-\d\d-\d\d$/.test(from))&&(!to||/^\d{4}-\d\d-\d\d$/.test(to)),'التاريخ غير صحيح');
 return allForStore_('Sales',s.storeId).filter(x=>x.Status!=='pending'&&(!from||x.CreatedAt.slice(0,10)>=from)&&(!to||x.CreatedAt.slice(0,10)<=to)&&(!dev||x.DeviceId===dev)&&(!cashier||x.CashierId===cashier)&&(!query||`${x.InvoiceNo} ${x.CustomerName} ${x.CashierName}`.toLowerCase().includes(query))&&(!method||(method==='cash'&&Number(x.CashPaid)>0&&Number(x.CardPaid)===0)||(method==='card'&&Number(x.CardPaid)>0&&Number(x.CashPaid)===0)||(method==='mixed'&&Number(x.CardPaid)>0&&Number(x.CashPaid)>0))).reverse().slice(0,Math.min(500,Math.max(20,Number(p.limit)||100)));
}
function archiveCatalog_(s,table,p){validateRole_(s,'manager');return scriptLock_(()=>{
 const field=table==='Products'?'ProductId':'CategoryId',id=text_(p.id,80),ref=byId_(table,field,id,s.storeId);require_(ref,'السجل غير موجود');
 const saved=update_(table,ref.row,{Active:false,UpdatedAt:now_(),Version:Number(ref.data.Version)+1});
 writeSync_(s.storeId,table,id,'update');audit_(s,'ARCHIVE',table,id,{});return saved;
});}
function adjustStock_(s,p){validateRole_(s,'manager');return scriptLock_(()=>{
 const ref=byId_('Products','ProductId',text_(p.ProductId,80),s.storeId);require_(ref,'المنتج غير موجود');
 const target=Number(p.NewStock);require_(Number.isSafeInteger(target)&&target>=0&&target<100000000,'الكمية الجديدة غير صحيحة');
 const delta=target-Number(ref.data.Stock),t=now_();const row=update_('Products',ref.row,{Stock:target,Version:Number(ref.data.Version)+1,UpdatedAt:t});
 insert_('InventoryMovements',{MovementId:id_(),StoreId:s.storeId,ProductId:row.ProductId,Type:'adjust',Quantity:delta,ReferenceId:row.ProductId,Note:text_(p.Note,160),UserId:s.userId,CreatedAt:t});
 writeSync_(s.storeId,'Products',row.ProductId,'update');audit_(s,'STOCK_ADJUST','Products',row.ProductId,{delta});return row;
});}
function openShift_(s,p){return scriptLock_(()=>{
 const dev=byId_('Devices','DeviceId',text_(p.DeviceId,80),s.storeId);require_(dev&&dev.data.Active===true,'الجهاز غير مسجل');
 const open=allForStore_('Shifts',s.storeId).find(x=>x.DeviceId===dev.data.DeviceId&&x.Status==='open');require_(!open,'يوجد بالفعل وردية مفتوحة لهذا الكاشير');
 const opening=exactMoney_(p.OpenCash||0,1000000),row={ShiftId:id_(),StoreId:s.storeId,DeviceId:dev.data.DeviceId,UserId:s.userId,OpenCash:opening,CloseCash:'',ExpectedCash:'',OpenedAt:now_(),ClosedAt:'',Status:'open',Note:''};
 insert_('Shifts',row);audit_(s,'SHIFT_OPEN','Shifts',row.ShiftId,{});return row;
});}
function closeShift_(s,p){return scriptLock_(()=>{
 const ref=byId_('Shifts','ShiftId',text_(p.ShiftId,80),s.storeId);require_(ref&&ref.data.Status==='open','الوردية مغلقة أو غير موجودة');
 require_(ref.data.UserId===s.userId||ROLE[s.role]>=2,'ليس لديك صلاحية إغلاق هذه الوردية');
 const invoices=allForStore_('Sales',s.storeId).filter(x=>x.ShiftId===ref.data.ShiftId&&x.Status==='completed');
 const cash=money_(invoices.reduce((t,x)=>t+Number(x.CashPaid||0)-Number(x.ChangeDue||0),0));
 const expected=money_(Number(ref.data.OpenCash)+cash),closing=exactMoney_(p.CloseCash,10000000);
 const row=update_('Shifts',ref.row,{CloseCash:closing,ExpectedCash:expected,ClosedAt:now_(),Status:'closed',Note:text_(p.Note,180)});
 audit_(s,'SHIFT_CLOSE','Shifts',row.ShiftId,{expected,closing,variance:money_(closing-expected)});return row;
});}
function holdOrder_(s,p){return scriptLock_(()=>{
 const json=JSON.stringify(p.Cart||{});require_(json.length>2&&json.length<=30000,'الطلب المعلق غير صالح');
 const id=text_(p.HoldId,80),ref=id?byId_('HeldOrders','HoldId',id,s.storeId):null;
 const row={HoldId:id||id_(),StoreId:s.storeId,DeviceId:text_(p.DeviceId,80),UserId:s.userId,PayloadJson:json,Status:'held',CreatedAt:ref?ref.data.CreatedAt:now_(),UpdatedAt:now_()};
 return ref?update_('HeldOrders',ref.row,row):insert_('HeldOrders',row);
});}
function deleteHold_(s,p){return scriptLock_(()=>{const ref=byId_('HeldOrders','HoldId',text_(p.HoldId,80),s.storeId);require_(ref,'الطلب غير موجود');return update_('HeldOrders',ref.row,{Status:'retrieved',UpdatedAt:now_()});});}
function createPurchase_(s,p){validateRole_(s,'manager');return scriptLock_(()=>{
 const lines=Array.isArray(p.Items)?p.Items:[];require_(lines.length>0&&lines.length<=50,'عدد أصناف المشتريات غير صالح');
 const supplier=p.SupplierId?byId_('Suppliers','SupplierId',p.SupplierId,s.storeId):null;require_(!p.SupplierId||supplier,'المورد غير موجود');
 const purchaseId=id_(),items=[],refs=new Map(),t=now_();let total=0;
 for(const x of lines){const ref=byId_('Products','ProductId',x.ProductId,s.storeId);require_(ref,'منتج مشتريات غير موجود');const qty=positiveInt_(x.Quantity,100000),cost=exactMoney_(x.UnitCost,1000000);refs.set(x.ProductId,ref);const line=money_(qty*cost);total=money_(total+line);items.push({PurchaseItemId:id_(),StoreId:s.storeId,PurchaseId:purchaseId,ProductId:x.ProductId,Quantity:qty,UnitCost:cost,LineTotal:line});}
 insert_('Purchases',{PurchaseId:purchaseId,StoreId:s.storeId,SupplierId:p.SupplierId||'',SupplierName:supplier?.data.Name||'',TotalCost:total,Note:text_(p.Note,180),CreatedAt:t,UserId:s.userId});
 appendBatch_('PurchaseItems',items);
 for(const item of items){const ref=byId_('Products','ProductId',item.ProductId,s.storeId);update_('Products',ref.row,{Stock:Number(ref.data.Stock)+item.Quantity,Cost:item.UnitCost,Version:Number(ref.data.Version)+1,UpdatedAt:t});
  insert_('InventoryMovements',{MovementId:id_(),StoreId:s.storeId,ProductId:item.ProductId,Type:'purchase',Quantity:item.Quantity,ReferenceId:purchaseId,Note:'شراء',UserId:s.userId,CreatedAt:t});writeSync_(s.storeId,'Products',item.ProductId,'update');}
 audit_(s,'PURCHASE','Purchases',purchaseId,{total});return {purchaseId,total};
});}
function createVoucher_(s,p){validateRole_(s,'manager');return scriptLock_(()=>{
 const amount=exactMoney_(p.Amount,10000000);require_(amount>0,'المبلغ مطلوب');const row={VoucherId:id_(),StoreId:s.storeId,Type:['receipt','payment'].includes(p.Type)?p.Type:'payment',AccountName:text_(p.AccountName,150),Amount:amount,Method:['cash','card','transfer'].includes(p.Method)?p.Method:'cash',Reference:text_(p.Reference,120),Note:text_(p.Note,240),CreatedAt:now_(),UserId:s.userId};
 insert_('Vouchers',row);audit_(s,'VOUCHER','Vouchers',row.VoucherId,{amount});return row;
});}
function saveManual_(s,p){return scriptLock_(()=>{
 const row={ManualOrderId:id_(),StoreId:s.storeId,CustomerName:text_(p.CustomerName,100),Description:text_(p.Description,900),Amount:exactMoney_(p.Amount||0,10000000),Status:'new',Note:text_(p.Note,200),CreatedAt:now_(),UserId:s.userId};
 require_(row.Description,'وصف الطلب مطلوب');insert_('ManualOrders',row);audit_(s,'MANUAL_ORDER','ManualOrders',row.ManualOrderId,{});return row;
});}
function cancelSale_(s,p){validateRole_(s,'manager');return scriptLock_(()=>{
 const ref=byId_('Sales','SaleId',text_(p.saleId,80),s.storeId);require_(ref&&ref.data.Status==='completed','الفاتورة غير موجودة أو سبق إلغاؤها');
 require_(!allForStore_('Returns',s.storeId).some(x=>x.SaleId===ref.data.SaleId),'لا يمكن إلغاء فاتورة تحتوي على مرتجعات؛ راجع الإدارة');
 const items=allForStore_('SaleItems',s.storeId).filter(x=>x.SaleId===ref.data.SaleId),sum=new Map();for(const it of items)sum.set(it.ProductId,(sum.get(it.ProductId)||0)+Number(it.Quantity));
 const t=now_();for(const [pid,q] of sum){const pr=byId_('Products','ProductId',pid,s.storeId);if(!pr||pr.data.Unlimited===true)continue;update_('Products',pr.row,{Stock:Number(pr.data.Stock)+q,Version:Number(pr.data.Version)+1,UpdatedAt:t});insert_('InventoryMovements',{MovementId:id_(),StoreId:s.storeId,ProductId:pid,Type:'cancel',Quantity:q,ReferenceId:ref.data.SaleId,Note:'إلغاء فاتورة',UserId:s.userId,CreatedAt:t});writeSync_(s.storeId,'Products',pid,'update');}
 const saved=update_('Sales',ref.row,{Status:'cancelled',UpdatedAt:t,Version:Number(ref.data.Version)+1});audit_(s,'SALE_CANCEL','Sales',saved.SaleId,{reason:text_(p.reason,180)});return saved;
});}
function returnSale_(s,p){validateRole_(s,'manager');return scriptLock_(()=>{
 const sale=byId_('Sales','SaleId',text_(p.SaleId,80),s.storeId);require_(sale&&sale.data.Status==='completed','الفاتورة غير قابلة للاسترجاع');
 const sold=allForStore_('SaleItems',s.storeId).filter(x=>x.SaleId===sale.data.SaleId),returns=allForStore_('Returns',s.storeId).filter(x=>x.SaleId===sale.data.SaleId),previous=allForStore_('ReturnItems',s.storeId).filter(x=>returns.some(r=>r.ReturnId===x.ReturnId));
 const entries=Array.isArray(p.Items)?p.Items:[];require_(entries.length>0&&entries.length<=40,'اختر منتجات المرتجع');
 let amount=0;const rows=[],seen=new Set();for(const it of entries){require_(!seen.has(it.SaleItemId),'لا تكرر صنفًا مرتجعًا في نفس الطلب');seen.add(it.SaleItemId);const orig=sold.find(x=>x.SaleItemId===it.SaleItemId);require_(orig,'صنف المرتجع غير موجود');const qty=positiveInt_(it.Quantity,9999),already=previous.filter(x=>x.SaleItemId===it.SaleItemId).reduce((a,b)=>a+Number(b.Quantity),0);
 require_(qty+already<=Number(orig.Quantity),'كمية المرتجع تتجاوز الكمية المباعة');const gross=money_(Number(orig.LineTotal)/Number(orig.Quantity)*qty);const line=money_(gross*(Number(sale.data.Subtotal)>0?Number(sale.data.Total)/Number(sale.data.Subtotal):0));amount=money_(amount+line);rows.push({ReturnItemId:id_(),StoreId:s.storeId,SaleItemId:it.SaleItemId,ProductId:orig.ProductId,Quantity:qty,Amount:line});}
 const id=id_(),t=now_();insert_('Returns',{ReturnId:id,StoreId:s.storeId,SaleId:sale.data.SaleId,Amount:amount,Reason:text_(p.Reason,180),UserId:s.userId,CreatedAt:t});appendBatch_('ReturnItems',rows.map(x=>({...x,ReturnId:id})));
 for(const x of rows){const pr=byId_('Products','ProductId',x.ProductId,s.storeId);if(!pr||pr.data.Unlimited===true)continue;update_('Products',pr.row,{Stock:Number(pr.data.Stock)+x.Quantity,Version:Number(pr.data.Version)+1,UpdatedAt:t});insert_('InventoryMovements',{MovementId:id_(),StoreId:s.storeId,ProductId:x.ProductId,Type:'return',Quantity:x.Quantity,ReferenceId:id,Note:'مرتجع',UserId:s.userId,CreatedAt:t});writeSync_(s.storeId,'Products',x.ProductId,'update');}
 audit_(s,'SALE_RETURN','Sales',sale.data.SaleId,{amount,returnId:id});return {ReturnId:id,Amount:amount};
});}
function reports_(s,p){validateRole_(s,'manager');const from=text_(p.from,10),to=text_(p.to,10),dev=text_(p.deviceId,80),cashier=text_(p.cashierId,80);
 require_((!from||/^\d{4}-\d\d-\d\d$/.test(from))&&(!to||/^\d{4}-\d\d-\d\d$/.test(to)),'نطاق التاريخ غير صحيح');
 const all=allForStore_('Sales',s.storeId).filter(x=>x.Status==='completed'&&(!from||x.CreatedAt.slice(0,10)>=from)&&(!to||x.CreatedAt.slice(0,10)<=to)&&(!dev||x.DeviceId===dev)&&(!cashier||x.CashierId===cashier));
 const total=(field)=>money_(all.reduce((sum,r)=>sum+Number(r[field]||0),0));
 const daily=new Map(),cashiers=new Map(),products=new Map(),idSet=new Set(all.map(x=>x.SaleId));
 all.forEach(x=>{const day=x.CreatedAt.slice(0,10);daily.set(day,money_((daily.get(day)||0)+Number(x.Total)));const key=x.DeviceName||x.CashierName;cashiers.set(key,money_((cashiers.get(key)||0)+Number(x.Total)));});
 const items=allForStore_('SaleItems',s.storeId).filter(x=>idSet.has(x.SaleId));items.forEach(x=>{const old=products.get(x.ProductId)||{name:x.ProductName,qty:0,total:0};old.qty+=Number(x.Quantity);old.total=money_(old.total+Number(x.LineTotal));products.set(x.ProductId,old);});
 const returns=allForStore_('Returns',s.storeId).filter(x=>idSet.has(x.SaleId)),returnTotal=money_(returns.reduce((a,b)=>a+Number(b.Amount||0),0));
 const cost=money_(items.reduce((a,b)=>a+Number(b.UnitCost||0)*Number(b.Quantity||0),0));
 return {count:all.length,total:total('Total'),subtotal:total('Subtotal'),discount:total('Discount'),vat:total('Tax'),cash:money_(all.reduce((a,x)=>a+Number(x.CashPaid)-Number(x.ChangeDue),0)),card:total('CardPaid'),returns:returnTotal,net:money_(total('Total')-returnTotal),estimatedProfit:money_(total('Total')-total('Tax')-cost-returnTotal),cost,
 daily:Array.from(daily,([date,total])=>({date,total})).sort((a,b)=>a.date.localeCompare(b.date)),cashiers:Array.from(cashiers,([name,total])=>({name,total})),topProducts:Array.from(products.values()).sort((a,b)=>b.qty-a.qty).slice(0,15)};
}
function capacity_(){const db=db_(),tabs=db.getSheets(),sheets=tabs.map(sh=>({name:sh.getName(),dataRows:Math.max(0,sh.getLastRow()-1),allocated:sh.getMaxRows()*sh.getMaxColumns()}));return {allocated:sheets.reduce((a,b)=>a+b.allocated,0),limit:20000000,sheets};}
/** Administrator-only manual actions: names end with '_' so google.script.run cannot call them. */
function backupDatabase_(){const file=DriveApp.getFileById(db_().getId()),copy=file.makeCopy('MILO BACKUP '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH-mm'));console.log(copy.getUrl());return copy.getUrl();}

function bindDevice_(s,p){return scriptLock_(()=>{const device=byId_('Devices','DeviceId',text_(p.DeviceId,80),s.storeId);require_(device&&device.data.Active===true,'الجهاز غير مسجل أو موقوف');const session=byId_('Sessions','SessionId',s.sessionId,s.storeId);require_(session,'جلسة غير موجودة');update_('Sessions',session.row,{DeviceId:device.data.DeviceId});return {deviceId:device.data.DeviceId,name:device.data.Name};});}
/** Optional daily backup. Activate explicitly from editor after verifying Drive quota. */
function dailyBackup_(){
 const props=PropertiesService.getScriptProperties();
 if(props.getProperty('BACKUP_ENABLED')!=='1'){console.log('MILO backup disabled');return;}
 const report=capacity_();if(report.allocated>report.limit*0.8)console.warn('MILO capacity exceeds 80%');
 const link=backupDatabase_();console.log('MILO daily backup: '+link);
}
function installDailyBackup_(){
 const props=PropertiesService.getScriptProperties();
 require_(props.getProperty('BACKUP_ENABLED')==='1','اضبط BACKUP_ENABLED=1 قبل تفعيل النسخ الاحتياطي');
 const old=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='dailyBackup_');
 old.forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('dailyBackup_').timeBased().everyDays(1).atHour(3).create();
 console.log('نسخ احتياطي يومي مفعل. راقب مساحة Drive والنسخ القديمة.');
}
