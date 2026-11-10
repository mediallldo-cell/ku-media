/**
 * KU Media Booking — Google Apps Script backend
 * First run: set Script Properties ADMIN_EMAILS and ADMIN_PASSWORD_HASH.
 * Then run setup() once. It creates a new Google Spreadsheet and Drive folder
 * automatically and stores their IDs in Script Properties.
 *
 * ADMIN_PASSWORD_HASH = lowercase SHA-256 hex digest of a unique password.
 * Example (Mac/Linux): printf %s 'YOUR-UNIQUE-PASSWORD' | shasum -a 256
 */
const HEADERS = ['id','createdAt','name','department','phone','email','date','event','project','location','start','end','requirements','services','serviceNames','estimate','payment','paymentDetail','taxInvoice','taxName','taxId','taxAddress','status','adminNote','driveFolderUrl','receiptUrl'];
const STATUSES = ['Pending','Approved / Quoted','Rejected','In Progress','Completed'];
const SERVICE_LABELS = {studio_pro:'Studio ETO Full HD',studio_basic:'Studio ETO',photo:'ห้องถ่ายภาพ STUDIO',audio:'บันทึกเสียงคำบรรยาย',single_cam:'ถ่ายทำนอกสถานที่ กล้องเดี่ยว',ob2:'OB 2 กล้อง',ob3:'OB 3 กล้อง',live:'ถ่ายทอดสดออนไลน์',editing:'ตัดต่อวีดิทัศน์'};
const PORTFOLIO_CATEGORIES = {studio:'ห้องสตูดิโอ (Studio ETO) · ชั้น 3 อาคารวิทยบริการ',location:'บันทึกภาพนอกสถานที่',livestream:'บริการถ่ายทอดสดออนไลน์ (Facebook Live / YouTube Live)',editing:'บริการตัดต่อวีดิทัศน์ (Video Editing)'};
const PORTFOLIO_CATEGORY_BY_SERVICE = {studio_pro:'studio',studio_basic:'studio',photo:'studio',audio:'studio',single_cam:'location',ob2:'location',ob3:'location',live:'livestream',editing:'editing'};
const PORTFOLIO_HEADERS = ['id','createdAt','category','title','description','fileId','active','position'];
const MAX_PORTFOLIO_IMAGE_BYTES = 5 * 1024 * 1024;
const PROP = PropertiesService.getScriptProperties();

function setup() {
  const admins = PROP.getProperty('ADMIN_EMAILS');
  const hash = PROP.getProperty('ADMIN_PASSWORD_HASH');
  if (!admins || !hash) throw new Error('ตั้งค่า Script Properties: ADMIN_EMAILS และ ADMIN_PASSWORD_HASH ก่อนเรียก setup()');
  let ssId = PROP.getProperty('SHEET_ID');
  if (!ssId) {
    const ss = SpreadsheetApp.create('KU Media Booking — ฐานข้อมูลคำขอใช้บริการ');
    ssId = ss.getId(); PROP.setProperty('SHEET_ID', ssId);
    const sh = ss.getSheets()[0]; sh.setName('Bookings');
    sh.appendRow(HEADERS); sh.setFrozenRows(1);
    sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#102d4f').setFontColor('#ffffff');
  } else {
    const ss = SpreadsheetApp.openById(ssId);
    if (!ss.getSheetByName('Bookings')) { const sh=ss.insertSheet('Bookings'); sh.appendRow(HEADERS); }
  }
  let folderId = PROP.getProperty('DRIVE_FOLDER_ID');
  if (!folderId) {
    const folder = DriveApp.createFolder('KU Media Booking — เอกสารคำขอใช้บริการ');
    folderId=folder.getId(); PROP.setProperty('DRIVE_FOLDER_ID',folderId);
  }
  const ssObj = SpreadsheetApp.openById(ssId);
  if (!ssObj.getSheetByName('MailQueue')) { const q=ssObj.insertSheet('MailQueue'); q.appendRow(['queuedAt','bookingId','recipient','kind','state','lastError']); q.setFrozenRows(1); q.getRange(1,1,1,6).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  if (!ssObj.getSheetByName('Portfolio')) { const p=ssObj.insertSheet('Portfolio'); p.appendRow(PORTFOLIO_HEADERS); p.setFrozenRows(1); p.getRange(1,1,1,PORTFOLIO_HEADERS.length).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  let galleryFolderId=PROP.getProperty('GALLERY_FOLDER_ID');
  if (!galleryFolderId) { galleryFolderId=DriveApp.getFolderById(folderId).createFolder('คลังผลงาน').getId(); PROP.setProperty('GALLERY_FOLDER_ID',galleryFolderId); }
  // Install one lightweight worker. Email/PDF work runs after the booking API has responded.
  if (!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='processMailQueue_')) ScriptApp.newTrigger('processMailQueue_').timeBased().everyMinutes(1).create();
  const ssUrl = ssObj.getUrl();
  const folderUrl = DriveApp.getFolderById(folderId).getUrl();
  Logger.log('SETUP COMPLETE. Spreadsheet: '+ssUrl+' | Drive folder: '+folderUrl);
  return {spreadsheetUrl:ssUrl,driveFolderUrl:folderUrl};
}

function doGet(e) {
  return json_({ok:true,service:'KU Media Booking API',version:'1.0',configured:!!(PROP.getProperty('SHEET_ID')&&PROP.getProperty('DRIVE_FOLDER_ID'))});
}
function doPost(e) {
  try {
    const q=JSON.parse((e.postData&&e.postData.contents)||'{}');
    return json_(route_(q));
  } catch(err) { return json_({error:String(err.message||err)}); }
}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function sheet_(){const id=PROP.getProperty('SHEET_ID');if(!id)throw new Error('ยังไม่ได้ตั้งค่า กรุณารัน setup()');return SpreadsheetApp.openById(id).getSheetByName('Bookings');}
function all_(){
  const sh=sheet_(),v=sh.getDataRange().getValues();
  return v.slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]),createdAt:r[1],name:r[2],department:r[3],phone:phoneString_(r[4]),email:r[5],date:dateString_(r[6]),
    event:r[7],project:r[8],location:r[9],start:timeString_(r[10]),end:timeString_(r[11]),requirements:r[12],
    services:parse_(r[13]),serviceNames:parse_(r[14]),estimate:Number(r[15]||0),payment:r[16],
    paymentDetail:r[17],taxInvoice:r[18],taxName:r[19],taxId:r[20],taxAddress:r[21],
    status:r[22],adminNote:r[23],driveFolderUrl:r[24],receiptUrl:r[25]
  }));
}
function parse_(v){try{return JSON.parse(v||'[]')}catch(_){return []}}
function dateString_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v||'')}
function timeString_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'HH:mm');return String(v||'')}
function phoneString_(v){
  if(v==null||v==='')return '';
  let phone=String(v).trim();
  if(typeof v==='number'&&Number.isFinite(v)){phone=String(Math.trunc(v));if(/^\d{8,9}$/.test(phone))phone='0'+phone}
  return phone;
}
function route_(q){
  switch(q.action){
    case 'createBooking': return create_(q.booking);
    case 'getBooking': {const b=all_().find(x=>x.id===q.id);if(!b)throw Error('ไม่พบรหัสคำขอ');return {booking:{id:b.id,event:b.event,date:b.date,start:b.start,end:b.end,status:b.status,adminNote:b.adminNote,receiptUrl:b.receiptUrl}}}
    case 'publicCalendar': return {bookings:all_().filter(b=>b.date.slice(0,7)===q.month&&['Approved / Quoted','In Progress'].includes(b.status)).map(b=>({id:b.id,date:b.date,start:b.start,end:b.end,serviceNames:b.serviceNames}))};
    case 'publicPortfolio': return {items:portfolioItems_(true)};
    case 'adminLogin': return login_(q);
    case 'forgotPassword': return forgotPassword_(q);
    case 'resetPassword': return resetPassword_(q);
    case 'adminList': auth_(q.token); return {bookings:all_(),adminEmails:adminEmails_()};
    case 'updateStatus': return updateStatus_(q);
    case 'addAdminEmail': return addAdminEmail_(q);
    case 'removeAdminEmail': return removeAdminEmail_(q);
    case 'adminPortfolio': auth_(q.token); return {items:portfolioItems_(false)};
    case 'savePortfolioItem': return savePortfolioItem_(q);
    case 'reorderPortfolioItems': return reorderPortfolioItems_(q);
    case 'deletePortfolioItem': return deletePortfolioItem_(q);
    case 'removeBooking': return removeBooking_(q);
    case 'health': return {ok:true,spreadsheet:!!PROP.getProperty('SHEET_ID'),drive:!!PROP.getProperty('DRIVE_FOLDER_ID')};
    default: throw Error('ไม่รู้จักคำสั่ง API');
  }
}
function create_(b){
  if(!b||!b.name||!b.email||!b.phone||!b.date||!b.start||!b.end||!b.services||!b.services.length)throw Error('ข้อมูลคำขอไม่ครบ');
  if(b.end<=b.start)throw Error('เวลาสิ้นสุดต้องอยู่หลังเวลาเริ่มต้น');
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try{
    const sh=sheet_(),lastRow=sh.getLastRow(),clashes=[];
    if(lastRow>1){
      const rows=sh.getRange(2,7,lastRow-1,17).getValues();
      rows.forEach(r=>{
        const status=r[16],services=parse_(r[7]);
        if(dateString_(r[0])===b.date&&['Approved / Quoted','In Progress'].includes(status)&&timeString_(r[4])<b.end&&timeString_(r[5])>b.start&&services.some(s=>b.services.includes(s)))clashes.push(r);
      });
    }
    if(clashes.length)throw Error('ช่วงเวลานี้มีรายการที่อนุมัติแล้ว กรุณาติดต่อเจ้าหน้าที่หรือเลือกเวลาใหม่');
    b.status='Pending';
    const newRow=lastRow+1,phone=String(b.phone).trim();
    sh.getRange(newRow,5).setNumberFormat('@');
    sh.appendRow([b.id,b.createdAt,b.name,b.department||'',phone,b.email,b.date,b.event,b.project||'',b.location||'',b.start,b.end,b.requirements||'',JSON.stringify(b.services),JSON.stringify(b.serviceNames||[]),b.estimate||0,b.payment||'',b.paymentDetail||'',!!b.taxInvoice,b.taxName||'',b.taxId||'',b.taxAddress||'',b.status,'','','']);
    sh.getRange(newRow,5).setNumberFormat('@').setValue(phone);
    // Queue the PDF and email; do not make the applicant wait for Gmail/Drive.
    const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));
    ss.getSheetByName('MailQueue').appendRow([new Date(),b.id,b.email,'newBooking','QUEUED','']);
    return {ok:true,id:b.id,emailSent:false,notificationQueued:true};
  } finally {lock.releaseLock()}
}

function login_(q){
  const email=String(q.email||'').trim().toLowerCase();
  const allowed=(PROP.getProperty('ADMIN_EMAILS')||'').toLowerCase().split(',').map(x=>x.trim()).filter(Boolean);
  const password=String(q.password||'');
  const saved=String(PROP.getProperty('ADMIN_PASSWORD_HASH')||'').trim().toLowerCase();
  // Accept the common terminal newline variants used when generating a SHA-256 hash.
  const candidates=[sha256_(password),sha256_(password+'\n'),sha256_(password+'\r\n')].map(x=>x.toLowerCase());
  if(!allowed.includes(email)||!candidates.includes(saved))throw Error('อีเมลหรือรหัสผ่านไม่ถูกต้อง');
  const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put('admin:'+token,email,21600);
  return {token,expiresIn:21600};
}


// Run manually from the Apps Script editor when an administrator needs a password reset.
// The password is entered in a prompt and only its SHA-256 hash is stored.
function resetAdminPassword(){
  const ui=SpreadsheetApp.getUi();
  const response=ui.prompt('ตั้งรหัสผ่านผู้ดูแลใหม่','กรอกรหัสผ่านใหม่ (อย่างน้อย 12 ตัวอักษร)',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK)return;
  const password=response.getResponseText();
  if(password.length<12)throw new Error('รหัสผ่านต้องมีอย่างน้อย 12 ตัวอักษร');
  PROP.setProperty('ADMIN_PASSWORD_HASH',sha256_(password));
  ui.alert('บันทึก hash รหัสผ่านใหม่แล้ว กรุณาทดลองเข้าสู่ระบบ');
}


// Password recovery: a one-time code is sent only to an email already listed in ADMIN_EMAILS.
// OTP expires after 5 minutes, is single-use, and allows at most 5 verification attempts.
function forgotPassword_(q){
  const email=String(q.email||'').trim().toLowerCase();
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email))throw new Error('กรุณากรอกอีเมลให้ถูกต้อง');
  const allowed=(PROP.getProperty('ADMIN_EMAILS')||'').toLowerCase().split(',').map(x=>x.trim()).filter(Boolean);
  // Do not reveal whether an address is registered.
  if(!allowed.includes(email))return {ok:true};
  const cache=CacheService.getScriptCache();
  const key=sha256_('ku-media-reset:'+email);
  if(cache.get('reset-rate:'+key))throw new Error('กรุณารอ 1 นาทีก่อนขอรหัส OTP อีกครั้ง');
  const otp=String(Math.floor(100000+Math.random()*900000));
  cache.put('reset-otp:'+key,sha256_(email+'|'+otp),300);
  cache.put('reset-attempts:'+key,'0',300);
  cache.put('reset-rate:'+key,'1',60);
  MailApp.sendEmail({
    to:email,
    subject:'รหัส OTP สำหรับรีเซ็ตรหัสผ่านผู้ดูแลระบบ KU Media',
    body:'รหัส OTP ของคุณคือ '+otp+'\\nรหัสนี้มีอายุ 5 นาที ใช้ได้ครั้งเดียว และห้ามเปิดเผยแก่ผู้อื่น หากคุณไม่ได้ร้องขอ สามารถละเว้นอีเมลนี้ได้',
    htmlBody:'<div style="font-family:Arial,Tahoma,sans-serif;background:#f3f6fa;padding:24px;color:#172b45"><div style="max-width:520px;margin:auto;background:#fff;border-radius:14px;overflow:hidden;border:1px solid #dce4ec"><div style="padding:20px 24px;background:#102d4f;color:#fff;border-bottom:4px solid #d9ad55"><b>KU Media Service</b><div style="font-size:13px;margin-top:5px">รีเซ็ตรหัสผ่านผู้ดูแลระบบ</div></div><div style="padding:24px"><p>รหัส OTP สำหรับยืนยันการรีเซ็ตรหัสผ่านของคุณ</p><div style="font-size:32px;letter-spacing:8px;font-weight:bold;color:#174b78;background:#eef4fa;padding:16px;text-align:center;border-radius:10px">'+otp+'</div><p style="line-height:1.7">รหัสนี้มีอายุ <b>5 นาที</b> ใช้ได้ครั้งเดียว และห้ามเปิดเผยแก่ผู้อื่น หากคุณไม่ได้ร้องขอ สามารถละเว้นอีเมลนี้ได้</p></div></div></div>'
  });
  return {ok:true};
}

function resetPassword_(q){
  const email=String(q.email||'').trim().toLowerCase();
  const otp=String(q.otp||'').trim();
  const password=String(q.password||'');
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email))throw new Error('อีเมลไม่ถูกต้อง');
  if(!/^\d{6}$/.test(otp))throw new Error('กรุณากรอกรหัส OTP 6 หลัก');
  if(password.length<12)throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 12 ตัวอักษร');
  const allowed=(PROP.getProperty('ADMIN_EMAILS')||'').toLowerCase().split(',').map(x=>x.trim()).filter(Boolean);
  if(!allowed.includes(email))throw new Error('ไม่สามารถยืนยันคำขอได้ กรุณาขอรหัส OTP ใหม่');
  const cache=CacheService.getScriptCache();
  const key=sha256_('ku-media-reset:'+email);
  const otpKey='reset-otp:'+key;
  const saved=cache.get(otpKey);
  if(!saved)throw new Error('รหัส OTP หมดอายุหรือยังไม่ได้ขอรหัส กรุณาขอรหัสใหม่');
  const attempts=Number(cache.get('reset-attempts:'+key)||0);
  if(attempts>=5){cache.remove(otpKey);throw new Error('กรอกรหัส OTP ผิดเกินจำนวนที่กำหนด กรุณาขอรหัสใหม่');}
  if(sha256_(email+'|'+otp)!==saved){
    cache.put('reset-attempts:'+key,String(attempts+1),300);
    throw new Error('รหัส OTP ไม่ถูกต้อง');
  }
  PROP.setProperty('ADMIN_PASSWORD_HASH',sha256_(password));
  cache.remove(otpKey);cache.remove('reset-attempts:'+key);cache.remove('reset-rate:'+key);
  return {ok:true};
}

function auth_(token){const email=CacheService.getScriptCache().get('admin:'+token);if(!email)throw Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');return email}
function updateStatus_(q){
  const actor=auth_(q.token);if(!STATUSES.includes(q.status))throw Error('สถานะไม่ถูกต้อง');
  const sh=sheet_(),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&String(r[0])===String(q.id));if(i<1)throw Error('ไม่พบคำขอ');
  sh.getRange(i+1,23).setValue(q.status);sh.getRange(i+1,24).setValue(q.adminNote||'');
  const b=all_().find(x=>x.id===q.id);if(!b)throw Error('ไม่พบคำขอ');
  let receiptUrl='';
  if(q.status==='Completed'){
    const pdf=makePdf_(b,'สรุปผลการให้บริการ');receiptUrl=pdf.getUrl();sh.getRange(i+1,26).setValue(receiptUrl);b.receiptUrl=receiptUrl;
  }
  const statusLabel=statusLabel_(q.status),note=String(q.adminNote||'');
  MailApp.sendEmail({to:b.email,subject:'อัปเดตสถานะคำขอ '+q.id+' | บริการผลิตและบริการสื่อ มก.',body:'เรียน คุณ'+b.name+'\\nระบบได้อัปเดตสถานะคำขอของท่านแล้ว\\nหมายเลขคำขอ: '+q.id+'\\nสถานะ: '+statusLabel+(note?'\\nข้อความจากเจ้าหน้าที่: '+note:'')+(receiptUrl?'\\nเอกสารสรุปผล: '+receiptUrl:''),htmlBody:emailHtml_(b,receiptUrl,false,{status:q.status,adminNote:note}),name:'KU Media Service'});
  notifyAdmins_('อัปเดตสถานะ '+q.id,'โดย '+actor+' เป็น '+statusLabel);
  return {ok:true,receiptUrl};
}
function processMailQueue_(){
  const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));const q=ss.getSheetByName('MailQueue');if(!q)return;
  const rows=q.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][4]!=='QUEUED')continue;
    const row=i+1,id=String(rows[i][1]),recipient=String(rows[i][2]);
    try{
      const b=all_().find(x=>x.id===id);if(!b)throw new Error('ไม่พบข้อมูลคำขอ '+id);
      sendBookingEmail_(b);
      notifyAdminsHtml_(b);
      q.getRange(row,5).setValue('SENT');
    }catch(err){q.getRange(row,5).setValue('ERROR');q.getRange(row,6).setValue(String(err.message||err).slice(0,500));}
  }
}
function statusLabel_(status){
  const labels={'Pending':'รอตรวจสอบ','Approved / Quoted':'อนุมัติแล้ว / เสนอราคา','Rejected':'ไม่อนุมัติ','In Progress':'กำลังดำเนินการ','Completed':'เสร็จสิ้น'};
  return labels[status]||String(status||'ไม่ระบุ');
}
function emailHtml_(b,pdfUrl,adminCopy,update){
  const esc=escapeHtml_,service=(b.serviceNames||[]).map(esc).join(' · ')||'—';
  const status=update?update.status:(b.status||'Pending'),statusLabel=statusLabel_(status);
  const statusStyle=status==='Rejected'?'background:#fde8e8;color:#a52d35':status==='Completed'?'background:#e5f5ec;color:#167247':status==='Approved / Quoted'||status==='In Progress'?'background:#e4f0fa;color:#174b78':'background:#fff2d6;color:#8a6113';
  const field=(label,value)=>'<tr><td style="padding:11px 14px;color:#64748b;width:34%;border-bottom:1px solid #e8edf3">'+label+'</td><td style="padding:11px 14px;color:#172b45;font-weight:600;border-bottom:1px solid #e8edf3">'+esc(value==null||value===''?'—':value)+'</td></tr>';
  const greeting=update?'อัปเดตสถานะคำขอใช้บริการ':adminCopy?'มีคำขอใช้บริการผลิตและบริการสื่อรายการใหม่':'ระบบได้รับคำขอใช้บริการของท่านเรียบร้อยแล้ว';
  const intro=update?'เรียน คุณ'+esc(b.name)+' ระบบได้อัปเดตสถานะคำขอของท่านแล้ว':adminCopy?'โปรดตรวจสอบรายละเอียดและดำเนินการตามขั้นตอน':'เรียน คุณ'+esc(b.name)+' ขอบคุณที่ส่งคำขอผ่านระบบออนไลน์';
  const note=update&&update.adminNote?'<div style="background:#fff9eb;border:1px solid #f1dfb8;border-radius:10px;padding:14px 16px;margin:18px 0"><div style="font-weight:700;color:#765615">ข้อความจากเจ้าหน้าที่</div><div style="line-height:1.8;margin-top:5px">'+esc(update.adminNote)+'</div></div>':'';
  const pdfButton=update&&pdfUrl?'<div style="text-align:center;margin:24px 0"><a href="'+esc(pdfUrl)+'" style="display:inline-block;background:#d9ad55;color:#172b45;text-decoration:none;font-weight:700;padding:12px 22px;border-radius:8px">เปิดเอกสารสรุปผล</a></div>':'';
  const noteText=update?'สถานะและรายละเอียดคำขอของท่านได้รับการอัปเดตโดยเจ้าหน้าที่ กรุณาเก็บหมายเลขคำขอนี้ไว้เพื่อติดตามสถานะ':'หมายเหตุ: ยอดดังกล่าวเป็นเพียงประมาณการ การให้บริการและค่าใช้จ่ายจะได้รับการยืนยันจากเจ้าหน้าที่อีกครั้ง กรุณาเก็บหมายเลขคำขอนี้ไว้เพื่อติดตามสถานะ';
  return '<div style="margin:0;background:#f2f5f9;padding:28px 12px;font-family:Arial,Tahoma,sans-serif;color:#172b45"><table role="presentation" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;margin:auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e1e8f0"><tr><td style="padding:26px 30px;background:linear-gradient(120deg,#102d4f,#1d5688);border-bottom:5px solid #d9ad55;color:#fff"><div style="font-size:12px;letter-spacing:1px;color:#dbe8f4">KASETSART UNIVERSITY</div><div style="font-size:22px;font-weight:700;margin-top:7px">บริการผลิตและบริการสื่อ</div><div style="font-size:13px;color:#e3edf6;margin-top:5px">สำนักพัฒนาการเรียนรู้ตลอดชีวิต มหาวิทยาลัยเกษตรศาสตร์</div></td></tr><tr><td style="padding:28px 30px 12px"><div style="font-size:17px;font-weight:700;color:#102d4f">'+greeting+'</div><p style="font-size:14px;line-height:1.8;color:#526276">'+intro+'</p><div style="background:#eef4fa;border:1px solid #dbe7f2;border-radius:12px;padding:17px 20px;margin:18px 0"><div style="font-size:12px;color:#61758a">หมายเลขคำขอ</div><div style="font-size:25px;letter-spacing:1px;font-weight:700;color:#174b78;margin-top:4px">'+esc(b.id)+'</div><span style="display:inline-block;margin-top:10px;'+statusStyle+';padding:5px 12px;border-radius:20px;font-size:13px;font-weight:700">'+esc(statusLabel)+'</span></div><div style="font-size:15px;font-weight:700;color:#102d4f;margin:20px 0 8px">รายละเอียดคำขอ</div><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border:1px solid #e8edf3;border-radius:10px;overflow:hidden">'+field('ชื่อผู้ขอ',b.name)+field('หน่วยงาน',b.department)+field('โทรศัพท์',b.phone)+field('อีเมล',b.email)+field('กิจกรรม',b.event)+field('วันที่ใช้บริการ',b.date)+field('เวลา',b.start+' – '+b.end)+field('สถานที่',b.location)+field('บริการที่เลือก',service)+field('ประมาณการค่าใช้จ่าย',Number(b.estimate||0).toLocaleString('th-TH')+' บาท')+'</table>'+note+pdfButton+'<p style="font-size:13px;line-height:1.8;color:#718096">'+noteText+'</p></td></tr><tr><td style="padding:18px 30px;background:#f7f9fc;border-top:1px solid #e8edf3;font-size:12px;line-height:1.7;color:#718096">สำนักพัฒนาการเรียนรู้ตลอดชีวิต · มหาวิทยาลัยเกษตรศาสตร์<br>อีเมลฉบับนี้ส่งจากระบบอัตโนมัติ กรุณาอย่าตอบกลับอีเมลนี้โดยตรง</td></tr></table></div>';
}
function sendBookingEmail_(b){
  MailApp.sendEmail({to:b.email,subject:'ยืนยันรับคำขอ '+b.id+' | บริการผลิตและบริการสื่อ มก.',body:'ระบบได้รับคำขอ '+b.id+' แล้ว สถานะรอเจ้าหน้าที่ตรวจสอบ',htmlBody:emailHtml_(b,'',false),name:'KU Media Service'});
}
function notifyAdminsHtml_(b){
  adminEmails_().forEach(to=>MailApp.sendEmail({to,subject:'🔔 คำขอใช้บริการใหม่ '+b.id,body:'มีคำขอใหม่ '+b.id+' จาก '+b.name,htmlBody:emailHtml_(b,'',true),name:'KU Media Service'}));
}
function makePdf_(b,title){
  const folder=DriveApp.getFolderById(PROP.getProperty('DRIVE_FOLDER_ID'));
  const rows=[['รหัสคำขอ',b.id],['ชื่อผู้ขอ',b.name],['หน่วยงาน',b.department||'-'],['โทรศัพท์',b.phone],['อีเมล',b.email],['กิจกรรม',b.event],['วันที่',b.date],['เวลา',b.start+'–'+b.end],['บริการ',(b.serviceNames||[]).join(', ')],['ยอดประมาณการ',Number(b.estimate||0).toLocaleString()+' บาท'],['สถานะ',b.status||'Pending']];
  const html='<html><meta charset="utf-8"><body style="font-family:Arial;padding:32px"><h2>'+title+'</h2><h3>สำนักพัฒนาการเรียนรู้ตลอดชีวิต มก.</h3><table border="1" cellpadding="8" cellspacing="0">'+rows.map(r=>'<tr><td>'+escapeHtml_(r[0])+'</td><td>'+escapeHtml_(r[1])+'</td></tr>').join('')+'</table><p>เอกสารนี้เป็นเอกสารสรุปคำขอ ไม่ใช่ใบเสร็จรับเงินหรือใบกำกับภาษี</p></body></html>';
  const blob=Utilities.newBlob(html,'text/html','summary.html').getAs('application/pdf').setName(b.id+'-'+title+'.pdf');
  const file=folder.createFile(blob);file.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.VIEW);return file;
}
function escapeHtml_(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function notifyAdmins_(subject,body){(PROP.getProperty('ADMIN_EMAILS')||'').split(',').map(x=>x.trim()).filter(Boolean).forEach(to=>MailApp.sendEmail(to,subject,body))}
function adminEmails_(){return (PROP.getProperty('ADMIN_EMAILS')||'').split(',').map(x=>x.trim().toLowerCase()).filter(Boolean)}
function addAdminEmail_(q){
  auth_(q.token);
  const email=String(q.email||'').trim().toLowerCase();
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email))throw Error('อีเมลไม่ถูกต้อง');
  const emails=adminEmails_();
  if(!emails.includes(email))emails.push(email);
  PROP.setProperty('ADMIN_EMAILS',emails.join(','));
  return {ok:true,emails};
}
function removeAdminEmail_(q){
  auth_(q.token);
  const email=String(q.email||'').trim().toLowerCase(),emails=adminEmails_();
  if(!emails.includes(email))throw Error('ไม่พบอีเมลนี้ในรายชื่อผู้รับแจ้งเตือน');
  if(emails.length<=1)throw Error('ต้องมีอีเมลผู้ดูแลอย่างน้อย 1 รายการ');
  const remaining=emails.filter(item=>item!==email);
  PROP.setProperty('ADMIN_EMAILS',remaining.join(','));
  return {ok:true,emails:remaining};
}
function portfolioSheet_(){
  const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID')),sheet=ss.getSheetByName('Portfolio');
  if(!sheet)throw new Error('ไม่พบชีต Portfolio กรุณารัน setup() ใน Apps Script ก่อน');
  portfolioOrderColumn_(sheet);
  return sheet;
}
function portfolioOrderColumn_(sheet){
  const width=sheet.getLastColumn(),headers=sheet.getRange(1,1,1,Math.max(1,width)).getValues()[0].map(value=>String(value||'').trim().toLowerCase());
  let column=headers.indexOf('position')+1;
  if(!column){column=width+1;sheet.getRange(1,column).setValue('position')}
  const lastRow=sheet.getLastRow();
  if(lastRow>1){
    const range=sheet.getRange(2,column,lastRow-1,1),values=range.getValues();
    let changed=false;
    values.forEach((row,index)=>{if(row[0]===''||row[0]==null){row[0]=index;changed=true}});
    if(changed)range.setValues(values);
  }
  return column;
}
function portfolioCategory_(category){return PORTFOLIO_CATEGORIES[category]?category:(PORTFOLIO_CATEGORY_BY_SERVICE[category]||category)}
function portfolioImageUrl_(fileId){return 'https://drive.google.com/thumbnail?id='+encodeURIComponent(fileId)+'&sz=w1200'}
function portfolioItems_(publicOnly){
  const sheet=portfolioSheet_(),orderColumn=portfolioOrderColumn_(sheet),values=sheet.getDataRange().getValues();
  return values.slice(1).filter(row=>row[0]&&(!publicOnly||row[6]!==false)).map(row=>({
    id:String(row[0]),createdAt:row[1],category:portfolioCategory_(String(row[2]||'')),title:String(row[3]||''),
    description:String(row[4]||''),fileId:String(row[5]||''),
    imageUrl:portfolioImageUrl_(String(row[5]||'')),
    active:row[6]!==false,position:Number(row[orderColumn-1])
  })).filter(item=>item.fileId).sort((a,b)=>a.position-b.position);
}
function savePortfolioItem_(q){
  auth_(q.token);
  const id=String(q.id||''),category=String(q.category||''),title=String(q.title||'').trim(),description=String(q.description||'').trim();
  if(!PORTFOLIO_CATEGORIES[category])throw new Error('กรุณาเลือกประเภทงานที่ถูกต้อง');
  if(!title)throw new Error('กรุณาระบุชื่อผลงาน');
  if(title.length>120||description.length>1000)throw new Error('ชื่อผลงานต้องไม่เกิน 120 ตัวอักษร และรายละเอียดไม่เกิน 1,000 ตัวอักษร');
  const sheet=portfolioSheet_(),values=sheet.getDataRange().getValues();
  const rowIndex=id?values.findIndex((row,index)=>index>0&&String(row[0])===id):-1;
  if(id&&rowIndex<1)throw new Error('ไม่พบผลงานที่ต้องการแก้ไข');
  const existing=rowIndex>0?values[rowIndex]:null;
  let fileId=existing?String(existing[5]||''):'',newFile=null;
  if(q.imageData){
    const mimeType=String(q.mimeType||''),fileName=String(q.fileName||'ผลงาน');
    if(!['image/jpeg','image/png','image/webp'].includes(mimeType))throw new Error('รองรับเฉพาะไฟล์ JPG, PNG หรือ WebP');
    if(typeof q.imageData!=='string'||q.imageData.length>Math.ceil(MAX_PORTFOLIO_IMAGE_BYTES*4/3)+8||!/^[A-Za-z0-9+/]+={0,2}$/.test(q.imageData))throw new Error('ไฟล์รูปไม่ถูกต้องหรือมีขนาดใหญ่เกิน 5 MB');
    const bytes=Utilities.base64Decode(q.imageData);
    if(!bytes.length||bytes.length>MAX_PORTFOLIO_IMAGE_BYTES)throw new Error('ไฟล์รูปต้องมีขนาดไม่เกิน 5 MB');
    const safeName=fileName.replace(/[\\/:*?"<>|]/g,'-').slice(0,100)||'portfolio-image';
    const folderId=PROP.getProperty('GALLERY_FOLDER_ID');
    if(!folderId)throw new Error('ยังไม่ได้ตั้งค่าโฟลเดอร์คลังผลงาน กรุณารัน setup() ใน Apps Script');
    newFile=DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(bytes,mimeType,safeName));
    try{newFile.setSharing(DriveApp.Access.ANYONE,DriveApp.Permission.VIEW)}
    catch(error){newFile.setTrashed(true);throw new Error('ไม่สามารถเปิดสิทธิ์ดูรูปผลงานสาธารณะได้: '+String(error.message||error))}
    fileId=newFile.getId();
  }else if(!fileId){
    throw new Error('กรุณาเลือกรูปผลงาน');
  }
  const orderColumn=portfolioOrderColumn_(sheet),orderValues=sheet.getRange(2,orderColumn,Math.max(1,sheet.getLastRow()-1),1).getValues();
  const position=existing?Number(existing[orderColumn-1]):Math.max(-1,...orderValues.map(value=>Number(value[0])||0))+1;
  const row=[id||Utilities.getUuid(),existing?existing[1]:new Date(),category,title,description,fileId,true];
  if(existing)sheet.getRange(rowIndex+1,1,1,row.length).setValues([row]);
  else sheet.appendRow(row);
  sheet.getRange(existing?rowIndex+1:sheet.getLastRow(),orderColumn).setValue(position);
  if(existing&&existing[5]&&existing[5]!==fileId)DriveApp.getFileById(String(existing[5])).setTrashed(true);
  return {ok:true,item:{id:row[0],createdAt:row[1],category,title,description,fileId,imageUrl:portfolioImageUrl_(fileId),active:true,position}};
}
function reorderPortfolioItems_(q){
  auth_(q.token);
  const ids=q.ids;
  if(!Array.isArray(ids)||ids.some(id=>typeof id!=='string'||!id)||new Set(ids).size!==ids.length)throw new Error('ลำดับผลงานไม่ถูกต้อง');
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try{
    const sheet=portfolioSheet_(),values=sheet.getDataRange().getValues(),existingIds=values.slice(1).filter(row=>row[0]).map(row=>String(row[0]));
    if(ids.length!==existingIds.length||ids.some(id=>!existingIds.includes(id)))throw new Error('รายการผลงานมีการเปลี่ยนแปลง กรุณาโหลดรายการใหม่แล้วลองอีกครั้ง');
    const orderColumn=portfolioOrderColumn_(sheet),positions=new Map(ids.map((id,index)=>[id,index]));
    const ordered=values.slice(1).filter(row=>row[0]).map(row=>[positions.get(String(row[0]))]);
    if(ordered.length)sheet.getRange(2,orderColumn,ordered.length,1).setValues(ordered);
    return {ok:true};
  }finally{lock.releaseLock()}
}
function deletePortfolioItem_(q){
  auth_(q.token);
  const id=String(q.id||''),sheet=portfolioSheet_(),values=sheet.getDataRange().getValues();
  const rowIndex=values.findIndex((row,index)=>index>0&&String(row[0])===id);
  if(rowIndex<1)throw new Error('ไม่พบผลงานที่ต้องการลบ');
  const fileId=String(values[rowIndex][5]||'');
  if(fileId)DriveApp.getFileById(fileId).setTrashed(true);
  sheet.deleteRow(rowIndex+1);
  return {ok:true};
}
function removeBooking_(q){
  auth_(q.token);
  const id=String(q.id||''),sh=sheet_(),values=sh.getDataRange().getValues(),index=values.findIndex((row,n)=>n>0&&String(row[0])===id);
  if(index<1)throw Error('ไม่พบคำขอ');
  const queue=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID')).getSheetByName('MailQueue');
  if(queue){
    const queued=queue.getDataRange().getValues();
    for(let row=queued.length-1;row>0;row--)if(String(queued[row][1])===id&&queued[row][4]==='QUEUED')queue.deleteRow(row+1);
  }
  sh.deleteRow(index+1);
  return {ok:true};
}
function sha256_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>(b+256).toString(16).slice(-2)).join('')}
