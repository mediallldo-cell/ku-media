/**
 * KU Media Booking — Google Apps Script backend
 * First run: set Script Properties ADMIN_EMAILS and ADMIN_PASSWORD_HASH.
 * Then run setup() once. It creates a new Google Spreadsheet and Drive folder
 * automatically and stores their IDs in Script Properties.
 * Each administrator must set a unique password through the OTP reset flow.
 *
 * ADMIN_PASSWORD_HASH = lowercase SHA-256 hex digest of a unique password.
 * Example (Mac/Linux): printf %s 'YOUR-UNIQUE-PASSWORD' | shasum -a 256
 */
const HEADERS = ['id','createdAt','name','department','phone','email','date','event','project','location','start','end','requirements','services','serviceNames','estimate','payment','paymentDetail','taxInvoice','taxName','taxId','taxAddress','status','adminNote','driveFolderUrl','receiptUrl','graphicsDetails'];
const STATUSES = ['Pending','Approved / Quoted','Rejected','Completed'];
const PRIMARY_ADMIN_EMAIL = 'media.llldo@ku.th';
const ADMIN_AUDIT_HEADERS = ['timestamp','actor','action','target','details'];
const ADMIN_ACTIVITY_HEADERS = ['sessionHash','email','startedAt','lastSeenAt','state'];
const SURVEY_HEADERS = ['submittedAt','category','rating','recommendation'];
const SURVEY_RATINGS = {5:'ดีมาก',4:'ดี',3:'ปานกลาง',2:'พอใช้',1:'ปรับปรุง'};
const GRAPHICS_ONLY_ADMIN_EMAILS = ['tassawan.d@ku.th'];
const SERVICE_LABELS = {studio_pro:'Studio LLLdo Full HD',studio_basic:'Studio LLLdo',photo:'ห้องถ่ายภาพ STUDIO',audio:'บันทึกเสียงคำบรรยาย',single_cam:'ถ่ายทำนอกสถานที่ กล้องเดี่ยว',still_photo:'บันทึกภาพนิ่ง',ob2:'OB 2 กล้อง',ob3:'OB 3 กล้อง',live:'ถ่ายทอดสดออนไลน์',editing:'ตัดต่อวีดิทัศน์',e_poster:'ออกแบบกราฟิก e-Poster',poster_a0_lam:'พิมพ์โปสเตอร์ A0 เคลือบสติกเกอร์',poster_a0_plain:'พิมพ์โปสเตอร์ A0 ไม่เคลือบสติกเกอร์',poster_a1_lam:'พิมพ์โปสเตอร์ A1 เคลือบสติกเกอร์',poster_a1_plain:'พิมพ์โปสเตอร์ A1 ไม่เคลือบสติกเกอร์',poster_a2_lam:'พิมพ์โปสเตอร์ A2 เคลือบสติกเกอร์',poster_a2_plain:'พิมพ์โปสเตอร์ A2 ไม่เคลือบสติกเกอร์',vinyl_plain:'พิมพ์ป้ายไวนิล ไม่ซีลขอบ',vinyl_eyelet:'พิมพ์ป้ายไวนิล ซีลขอบและตอกตาไก่',sticker_basic:'สติกเกอร์ผิวด้าน-มัน',sticker_matte:'สติกเกอร์เคลือบด้าน',foam_3:'ติดสติกเกอร์ลงฟิวเจอร์บอร์ด 3 มม.',foam_4:'ติดสติกเกอร์ลงฟิวเจอร์บอร์ด 4 มม.',foam_5:'ติดสติกเกอร์ลงโฟมบอร์ด 5 มม.',plaswood_3:'ติดสติกเกอร์ลงพลาสวูด 3 มม.',plaswood_4:'ติดสติกเกอร์ลงพลาสวูด 4 มม.'};
const PORTFOLIO_CATEGORIES = {studio:'ห้องสตูดิโอ (Studio LLLdo) · ชั้น 3 อาคารวิทยบริการ',location:'บันทึกภาพนอกสถานที่',livestream:'บริการถ่ายทอดสดออนไลน์ (Facebook Live / YouTube Live)',editing:'บริการตัดต่อวีดิทัศน์ (Video Editing)',graphics:'งานกราฟิกสื่อสิ่งพิมพ์'};
const PORTFOLIO_CATEGORY_BY_SERVICE = {studio_pro:'studio',studio_basic:'studio',photo:'studio',audio:'studio',single_cam:'location',still_photo:'location',ob2:'location',ob3:'location',live:'livestream',editing:'editing',e_poster:'graphics',poster_a0_lam:'graphics',poster_a0_plain:'graphics',poster_a1_lam:'graphics',poster_a1_plain:'graphics',poster_a2_lam:'graphics',poster_a2_plain:'graphics',vinyl_plain:'graphics',vinyl_eyelet:'graphics',sticker_basic:'graphics',sticker_matte:'graphics',foam_3:'graphics',foam_4:'graphics',foam_5:'graphics',plaswood_3:'graphics',plaswood_4:'graphics'};
const PORTFOLIO_HEADERS = ['id','createdAt','category','title','description','fileId','active','position'];
const MAX_PORTFOLIO_IMAGE_BYTES = 5 * 1024 * 1024;
const MAX_SIGNATURE_BYTES = 1024 * 1024;
const PROP = PropertiesService.getScriptProperties();

function setup() {
  const admins = PROP.getProperty('ADMIN_EMAILS');
  const passwordConfigured=!!PROP.getProperty('ADMIN_PASSWORD_HASH')||Object.keys(adminPasswordHashes_()).length>0;
  if (!admins || !passwordConfigured) throw new Error('ตั้งค่า Script Properties: ADMIN_EMAILS และรหัสผ่านผู้ดูแลก่อนเรียก setup()');
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
  const bookingsSheet=ssObj.getSheetByName('Bookings');
  if(bookingsSheet){
    const headerValues=bookingsSheet.getRange(1,1,1,Math.max(1,bookingsSheet.getLastColumn())).getValues()[0].map(value=>String(value||'').trim());
    if(!headerValues.includes('graphicsDetails')){
      const newColumn=bookingsSheet.getLastColumn()+1;
      bookingsSheet.getRange(1,newColumn).setValue('graphicsDetails').setFontWeight('bold').setBackground('#102d4f').setFontColor('#ffffff');
    }
  }
  if (!ssObj.getSheetByName('MailQueue')) { const q=ssObj.insertSheet('MailQueue'); q.appendRow(['queuedAt','bookingId','recipient','kind','state','lastError']); q.setFrozenRows(1); q.getRange(1,1,1,6).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  if (!ssObj.getSheetByName('AdminAudit')) { const a=ssObj.insertSheet('AdminAudit'); a.appendRow(ADMIN_AUDIT_HEADERS); a.setFrozenRows(1); a.getRange(1,1,1,ADMIN_AUDIT_HEADERS.length).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  if (!ssObj.getSheetByName('AdminActivity')) { const a=ssObj.insertSheet('AdminActivity'); a.appendRow(ADMIN_ACTIVITY_HEADERS); a.setFrozenRows(1); a.getRange(1,1,1,ADMIN_ACTIVITY_HEADERS.length).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  if (!ssObj.getSheetByName('SatisfactionSurvey')) { const s=ssObj.insertSheet('SatisfactionSurvey'); s.appendRow(SURVEY_HEADERS); s.setFrozenRows(1); s.getRange(1,1,1,SURVEY_HEADERS.length).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  if (!ssObj.getSheetByName('Portfolio')) { const p=ssObj.insertSheet('Portfolio'); p.appendRow(PORTFOLIO_HEADERS); p.setFrozenRows(1); p.getRange(1,1,1,PORTFOLIO_HEADERS.length).setFontWeight('bold').setBackground('#174b78').setFontColor('#ffffff'); }
  calendarEventsSheet_(true);
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
    status:r[22],adminNote:r[23],driveFolderUrl:r[24],receiptUrl:r[25],graphicsDetails:String(r[26]||'')
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
function sheetSafeValue_(value){
  if(typeof value==='string'&&/^[\s\u200B\uFEFF]*[=+\-@]/.test(value))return "'"+value;
  return value;
}
function newBookingId_(){
  const sh=sheet_(),lastRow=sh.getLastRow(),existing=new Set(lastRow>1?sh.getRange(2,1,lastRow-1,1).getValues().map(row=>String(row[0])):[]);
  for(let attempt=0;attempt<100;attempt++){
    const id=String(Math.floor(Math.random()*900000)+100000);
    if(!existing.has(id))return id;
  }
  throw new Error('ไม่สามารถสร้างรหัสคำขอ 6 หลักที่ไม่ซ้ำได้ กรุณาลองอีกครั้ง');
}
function getBooking_(q){
  const id=String(q.id||'').trim(),email=String(q.email||'').trim().toLowerCase();
  if(!id||!email)throw new Error('กรุณากรอกรหัสคำขอและอีเมลที่ใช้ยื่นคำขอ');
  const booking=all_().find(item=>item.id===id&&String(item.email||'').trim().toLowerCase()===email);
  if(!booking)throw new Error('ไม่พบคำขอหรือข้อมูลยืนยันไม่ถูกต้อง');
  return {booking:{id:booking.id,event:booking.event,date:booking.date,start:booking.start,end:booking.end,status:booking.status,adminNote:booking.adminNote}};
}
function surveySheet_(){
  const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));
  let sheet=ss.getSheetByName('SatisfactionSurvey');
  if(!sheet){sheet=ss.insertSheet('SatisfactionSurvey');sheet.appendRow(SURVEY_HEADERS);sheet.setFrozenRows(1)}
  return sheet;
}
function submitSurvey_(q){
  const category=String(q.category||'').trim(),rating=Number(q.rating),recommendation=String(q.recommendation||'').trim();
  if(!Object.prototype.hasOwnProperty.call(PORTFOLIO_CATEGORIES,category))throw new Error('กรุณาเลือกประเภทบริการที่ต้องการประเมิน');
  if(!Number.isInteger(rating)||!Object.prototype.hasOwnProperty.call(SURVEY_RATINGS,rating))throw new Error('กรุณาเลือกคะแนนความพึงพอใจ');
  if(recommendation.length>1000)throw new Error('ข้อแนะนำต้องไม่เกิน 1,000 ตัวอักษร');
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try{
    surveySheet_().appendRow([new Date(),category,rating,sheetSafeValue_(recommendation)]);
    return {ok:true};
  }finally{lock.releaseLock()}
}
function adminSurveyStats_(){
  const sheet=surveySheet_(),rows=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,SURVEY_HEADERS.length).getValues():[];
  const responses=rows.filter(row=>row[0]).map(row=>({
    submittedAt:row[0],category:String(row[1]||''),rating:Number(row[2]),ratingLabel:SURVEY_RATINGS[Number(row[2])]||'',recommendation:String(row[3]||'').replace(/^'(?=[\s\u200B\uFEFF]*[=+\-@])/,'')
  }));
  const categories=Object.entries(PORTFOLIO_CATEGORIES).map(([id,label])=>{
    const categoryResponses=responses.filter(response=>response.category===id),ratingCounts={};
    Object.keys(SURVEY_RATINGS).forEach(rating=>{ratingCounts[rating]=categoryResponses.filter(response=>response.rating===Number(rating)).length});
    const average=categoryResponses.length?categoryResponses.reduce((total,response)=>total+response.rating,0)/categoryResponses.length:0;
    return {id,label,count:categoryResponses.length,average,ratingCounts};
  });
  return {total:responses.length,categories,responses};
}
function route_(q){
  switch(q.action){
    case 'createBooking': return create_(q.booking);
    case 'getBooking': return getBooking_(q);
    case 'publicCalendar': return {bookings:publicCalendarItems_(q.month)};
    case 'publicPortfolio': return {items:portfolioItems_(true)};
    case 'submitSurvey': return submitSurvey_(q);
    case 'adminLogin': return login_(q);
    case 'forgotPassword': return forgotPassword_(q);
    case 'resetPassword': return resetPassword_(q);
    case 'adminList': { const actor=auth_(q.token); return {bookings:all_(),adminEmails:adminEmails_(),passwordSetupRequired:adminPasswordSetupRequired_(),passwordConfiguredEmails:Object.keys(adminPasswordHashes_()),isPrimaryAdmin:actor===PRIMARY_ADMIN_EMAIL,primaryAdminEmail:PRIMARY_ADMIN_EMAIL}; }
    case 'adminAudit': return adminAudit_(q);
    case 'adminSurveyStats': auth_(q.token); return adminSurveyStats_();
    case 'adminHeartbeat': auth_(q.token); return {ok:true};
    case 'adminLogout': return adminLogout_(q);
    case 'adminPdf': return adminPdf_(q);
    case 'adminSignature': return getAdminSignature_(q);
    case 'saveAdminSignature': return saveAdminSignature_(q);
    case 'deleteAdminSignature': return deleteAdminSignature_(q);
    case 'updateStatus': return updateStatus_(q);
    case 'addAdminEmail': return addAdminEmail_(q);
    case 'removeAdminEmail': return removeAdminEmail_(q);
    case 'createCalendarEvent': return createCalendarEvent_(q);
    case 'deleteCalendarEvent': return deleteCalendarEvent_(q);
    case 'adminPortfolio': auth_(q.token); return {items:portfolioItems_(false)};
    case 'savePortfolioItem': return savePortfolioItem_(q);
    case 'reorderPortfolioItems': return reorderPortfolioItems_(q);
    case 'deletePortfolioItem': return deletePortfolioItem_(q);
    case 'removeBooking': return removeBooking_(q);
    case 'health': return {ok:true,spreadsheet:!!PROP.getProperty('SHEET_ID'),drive:!!PROP.getProperty('DRIVE_FOLDER_ID')};
    default: throw Error('ไม่รู้จักคำสั่ง API');
  }
}
function signatureFolder_(){
  let folderId=PROP.getProperty('SIGNATURE_FOLDER_ID');
  if(!folderId){
    const rootId=PROP.getProperty('DRIVE_FOLDER_ID');
    if(!rootId)throw new Error('ยังไม่ได้ตั้งค่าโฟลเดอร์จัดเก็บ กรุณารัน setup() ใน Apps Script');
    folderId=DriveApp.getFolderById(rootId).createFolder('ลายเซ็นผู้อนุมัติ (ส่วนตัว)').getId();
    PROP.setProperty('SIGNATURE_FOLDER_ID',folderId);
  }
  return DriveApp.getFolderById(folderId);
}
function getAdminSignature_(q){
  auth_(q.token);
  const fileId=PROP.getProperty('APPROVER_SIGNATURE_FILE_ID');
  if(!fileId)return {ok:true,signature:null};
  const blob=DriveApp.getFileById(fileId).getBlob(),mimeType=blob.getContentType(),bytes=blob.getBytes();
  if(mimeType!=='image/png'||bytes.length>MAX_SIGNATURE_BYTES)throw new Error('ไฟล์ลายเซ็นที่จัดเก็บไม่ถูกต้องหรือมีขนาดใหญ่เกิน 1 MB');
  return {ok:true,signature:{mimeType,dataUrl:'data:'+mimeType+';base64,'+Utilities.base64Encode(bytes)}};
}
function saveAdminSignature_(q){
  const actor=auth_(q.token);
  if(q.authorizationConfirmed!==true)throw new Error('กรุณายืนยันว่าได้รับอนุญาตให้ใช้ลายเซ็นนี้ในเอกสาร');
  if(q.mimeType!=='image/png'||typeof q.imageData!=='string'||q.imageData.length>Math.ceil(MAX_SIGNATURE_BYTES*4/3)+8||!/^[A-Za-z0-9+/]+={0,2}$/.test(q.imageData))throw new Error('กรุณาอัปโหลดไฟล์ลายเซ็น PNG ที่ถูกต้อง ขนาดไม่เกิน 1 MB');
  const bytes=Utilities.base64Decode(q.imageData);
  if(!bytes.length||bytes.length>MAX_SIGNATURE_BYTES||bytes.length<8||[137,80,78,71,13,10,26,10].some((value,index)=>(bytes[index]+256)%256!==value))throw new Error('ไฟล์ลายเซ็น PNG ไม่ถูกต้องหรือมีขนาดใหญ่เกิน 1 MB');
  const oldFileId=PROP.getProperty('APPROVER_SIGNATURE_FILE_ID');
  const file=signatureFolder_().createFile(Utilities.newBlob(bytes,'image/png','approver-signature.png'));
  PROP.setProperty('APPROVER_SIGNATURE_FILE_ID',file.getId());
  PROP.setProperty('APPROVER_SIGNATURE_AUTHORIZED_AT',new Date().toISOString());
  PROP.setProperty('APPROVER_SIGNATURE_AUTHORIZED_BY',actor);
  let warning='';
  if(oldFileId&&oldFileId!==file.getId()){
    try{DriveApp.getFileById(oldFileId).setTrashed(true)}
    catch(error){warning='เปลี่ยนลายเซ็นแล้ว แต่ไม่สามารถนำไฟล์ลายเซ็นเดิมไปถังขยะได้: '+String(error.message||error)}
  }
  recordAdminAudit_(actor,'เปลี่ยนลายเซ็นผู้อนุมัติ',file.getId(),'');
  return {ok:true,warning};
}
function deleteAdminSignature_(q){
  const actor=auth_(q.token);
  const fileId=PROP.getProperty('APPROVER_SIGNATURE_FILE_ID');
  if(!fileId)return {ok:true};
  PROP.deleteProperty('APPROVER_SIGNATURE_FILE_ID');
  PROP.deleteProperty('APPROVER_SIGNATURE_AUTHORIZED_AT');
  PROP.deleteProperty('APPROVER_SIGNATURE_AUTHORIZED_BY');
  recordAdminAudit_(actor,'ลบลายเซ็นผู้อนุมัติ',fileId,'');
  try{DriveApp.getFileById(fileId).setTrashed(true)}
  catch(error){return {ok:true,warning:'ลบลายเซ็นออกจากระบบแล้ว แต่ไม่สามารถนำไฟล์เดิมไปถังขยะได้: '+String(error.message||error)}}
  return {ok:true};
}
function create_(b){
  return createWithStatus_(b,'Pending',true);
}
function createWithStatus_(b,status,queueNotification){
  if(!b||!b.name||!b.email||!b.phone||!b.date||!b.start||!b.end||!b.services||!b.services.length)throw Error('ข้อมูลคำขอไม่ครบ');
  if(b.end<=b.start)throw Error('เวลาสิ้นสุดต้องอยู่หลังเวลาเริ่มต้น');
  if(!['Pending','Approved / Quoted'].includes(status))throw Error('สถานะสำหรับสร้างรายการไม่ถูกต้อง');
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
    b.status=status;
    b.id=newBookingId_();
    const newRow=lastRow+1,phone=String(b.phone).trim();
    sh.getRange(newRow,5).setNumberFormat('@');
    sh.appendRow([b.id,new Date(),b.name,b.department||'',phone,b.email,b.date,b.event,b.project||'',b.location||'',b.start,b.end,b.requirements||'',JSON.stringify(b.services),JSON.stringify(b.serviceNames||[]),b.estimate||0,b.payment||'',b.paymentDetail||'',!!b.taxInvoice,b.taxName||'',b.taxId||'',b.taxAddress||'',b.status,'','','',String(b.graphicsDetails||'').trim()].map(sheetSafeValue_));
    sh.getRange(newRow,5).setNumberFormat('@').setValue(sheetSafeValue_(phone));
    // Queue the PDF and email; do not make the applicant wait for Gmail/Drive.
    if(queueNotification){
      const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));
      ss.getSheetByName('MailQueue').appendRow([new Date(),b.id,b.email,'newBooking','QUEUED',''].map(sheetSafeValue_));
    }
    return {ok:true,id:b.id,emailSent:false,notificationQueued:queueNotification};
  } finally {lock.releaseLock()}
}

function calendarEventsSheet_(createIfMissing){
  const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));
  let sheet=ss.getSheetByName('CalendarEvents');
  if(!sheet&&createIfMissing){sheet=ss.insertSheet('CalendarEvents');sheet.appendRow(['id','date','start','end','title','notes','createdBy','createdAt']);sheet.setFrozenRows(1)}
  return sheet;
}
function publicCalendarItems_(month){
  month=String(month||'');
  if(!/^\d{4}-\d{2}$/.test(month))throw new Error('กรุณาเลือกเดือนให้ถูกต้อง');
  const bookings=all_().filter(b=>b.date.slice(0,7)===month&&['Approved / Quoted','In Progress'].includes(b.status)).map(b=>({id:b.id,date:b.date,start:b.start,end:b.end,serviceNames:b.serviceNames,services:b.services,eventType:'booking'}));
  const sheet=calendarEventsSheet_(false);
  if(!sheet||sheet.getLastRow()<2)return bookings;
  const events=sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues().filter(row=>String(row[1]||'').slice(0,7)===month).map(row=>({id:String(row[0]),date:dateString_(row[1]),start:timeString_(row[2]),end:timeString_(row[3]),title:String(row[4]||''),notes:String(row[5]||''),eventType:'block'}));
  return bookings.concat(events);
}
function createCalendarEvent_(q){
  const actor=auth_(q.token),date=String(q.date||''),start=String(q.start||''),end=String(q.end||''),title=String(q.title||'').trim(),notes=String(q.notes||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||isNaN(new Date(date+'T00:00:00').getTime()))throw new Error('กรุณาระบุวันที่ให้ถูกต้อง');
  if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(start)||!/^([01]\d|2[0-3]):[0-5]\d$/.test(end)||end<=start)throw new Error('กรุณาระบุช่วงเวลาให้ถูกต้อง');
  if(!title||title.length>120||notes.length>500)throw new Error('ชื่อรายการต้องมี 1–120 ตัวอักษร และหมายเหตุไม่เกิน 500 ตัวอักษร');
  const sheet=calendarEventsSheet_(true),id='EV-'+Utilities.getUuid();
  sheet.appendRow([id,date,start,end,title,notes,actor,new Date()].map(sheetSafeValue_));
  recordAdminAudit_(actor,'เพิ่มรายการกันคิว',id,title+' · '+date+' '+start+'-'+end);
  return {ok:true,event:{id,date,start,end,title,notes,eventType:'block'}};
}
function deleteCalendarEvent_(q){
  const actor=auth_(q.token);
  const id=String(q.id||''),sheet=calendarEventsSheet_(false);
  if(!sheet||sheet.getLastRow()<2)throw new Error('ไม่พบรายการกันคิว');
  const row=sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().findIndex(value=>String(value[0])===id);
  if(row<0)throw new Error('ไม่พบรายการกันคิว');
  const removed=sheet.getRange(row+2,1,1,8).getValues()[0];
  sheet.deleteRow(row+2);
  recordAdminAudit_(actor,'ลบรายการกันคิว',id,String(removed[4]||''));
  return {ok:true};
}

function login_(q){
  const email=String(q.email||'').trim().toLowerCase();
  const allowed=adminEmails_();
  const password=String(q.password||'');
  const hashes=adminPasswordHashes_(),legacyHash=String(PROP.getProperty('ADMIN_PASSWORD_HASH')||'').trim().toLowerCase();
  const usingLegacyPrimaryPassword=email===PRIMARY_ADMIN_EMAIL&&!hashes[email]&&!!legacyHash;
  const saved=String(hashes[email]||(usingLegacyPrimaryPassword?legacyHash:'')).trim().toLowerCase();
  // Accept the common terminal newline variants used when generating a SHA-256 hash.
  const candidates=[sha256_(password),sha256_(password+'\n'),sha256_(password+'\r\n')].map(x=>x.toLowerCase());
  if(!allowed.includes(email)||!candidates.includes(saved))throw Error('อีเมลหรือรหัสผ่านไม่ถูกต้อง');
  if(usingLegacyPrimaryPassword){
    hashes[email]=legacyHash;
    PROP.setProperty('ADMIN_PASSWORD_HASHES',JSON.stringify(hashes));
  }
  const token=Utilities.getUuid()+Utilities.getUuid();CacheService.getScriptCache().put('admin:'+token,JSON.stringify({email,version:adminSessionVersion_(email)}),21600);
  touchAdminSession_(email,token);recordAdminAudit_(email,usingLegacyPrimaryPassword?'เข้าสู่ระบบด้วยรหัสผ่านเดิม':'เข้าสู่ระบบ','',usingLegacyPrimaryPassword?'กรุณาตั้งรหัสผ่านส่วนตัวผ่าน OTP เพื่อยกเลิกรหัสผ่านเดิม':'');
  return {token,expiresIn:21600,passwordMigrationRequired:usingLegacyPrimaryPassword};
}


// Run manually from the Apps Script editor when an administrator needs a password reset.
// The password is entered in a prompt and only its SHA-256 hash is stored.
function resetAdminPassword(){
  const ui=SpreadsheetApp.getUi();
  const response=ui.prompt('ตั้งรหัสผ่านผู้ดูแลใหม่','กรอกรหัสผ่านใหม่ (อย่างน้อย 6 ตัวอักษร)',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK)return;
  const password=response.getResponseText();
  if(password.length<6)throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
  const hashes=adminPasswordHashes_();hashes[PRIMARY_ADMIN_EMAIL]=sha256_(password);
  PROP.setProperty('ADMIN_PASSWORD_HASHES',JSON.stringify(hashes));
  PROP.setProperty('ADMIN_PASSWORD_SETUP_REQUIRED',adminPasswordSetupRequired_().filter(email=>email!==PRIMARY_ADMIN_EMAIL).join(','));
  invalidateAdminSessions_(PRIMARY_ADMIN_EMAIL);
  ui.alert('ตั้งรหัสผ่านส่วนตัวของอีเมลหลักแล้ว กรุณาทดลองเข้าสู่ระบบ');
}


// Password recovery: a one-time code is sent only to an email already listed in ADMIN_EMAILS.
// OTP expires after 5 minutes, is single-use, and allows at most 5 verification attempts.
function forgotPassword_(q){
  const email=String(q.email||'').trim().toLowerCase();
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email))throw new Error('กรุณากรอกอีเมลให้ถูกต้อง');
  const allowed=adminEmails_();
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
  if(password.length<6)throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
  const allowed=adminEmails_();
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
  const hashes=adminPasswordHashes_();hashes[email]=sha256_(password);
  PROP.setProperty('ADMIN_PASSWORD_HASHES',JSON.stringify(hashes));
  PROP.setProperty('ADMIN_PASSWORD_SETUP_REQUIRED',adminPasswordSetupRequired_().filter(item=>item!==email).join(','));
  invalidateAdminSessions_(email);
  if(adminEmails_().every(item=>hashes[item]))PROP.deleteProperty('ADMIN_PASSWORD_HASH');
  cache.remove(otpKey);cache.remove('reset-attempts:'+key);cache.remove('reset-rate:'+key);
  recordAdminAudit_(email,'เปลี่ยนรหัสผ่าน','', 'ยืนยันผ่าน OTP');
  return {ok:true};
}

function adminSessionVersion_(email){
  return Number(PROP.getProperty('ADMIN_SESSION_VERSION_'+sha256_(email))||1);
}
function invalidateAdminSessions_(email){
  const key='ADMIN_SESSION_VERSION_'+sha256_(email),version=adminSessionVersion_(email);
  if(!Number.isSafeInteger(version)||version<1)throw new Error('ข้อมูลเวอร์ชันเซสชันผู้ดูแลไม่ถูกต้อง');
  PROP.setProperty(key,String(version+1));
}
function auth_(token){
  const cached=CacheService.getScriptCache().get('admin:'+token);
  if(!cached)throw Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  let session;
  try{session=JSON.parse(cached)}catch(_){throw Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่')}
  const email=String(session.email||'').toLowerCase(),allowed=adminEmails_();
  if(!email||!allowed.includes(email)||session.version!==adminSessionVersion_(email))throw Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  if(!adminPasswordHashes_()[email])throw Error('บัญชีนี้ต้องตั้งรหัสผ่านส่วนตัวผ่าน OTP ก่อนใช้งาน');
  touchAdminSession_(email,token);
  return email;
}
function adminAuditSheet_(){
  const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));
  let sheet=ss.getSheetByName('AdminAudit');
  if(!sheet){sheet=ss.insertSheet('AdminAudit');sheet.appendRow(ADMIN_AUDIT_HEADERS);sheet.setFrozenRows(1)}
  return sheet;
}
function adminActivitySheet_(){
  const ss=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID'));
  let sheet=ss.getSheetByName('AdminActivity');
  if(!sheet){sheet=ss.insertSheet('AdminActivity');sheet.appendRow(ADMIN_ACTIVITY_HEADERS);sheet.setFrozenRows(1)}
  return sheet;
}
function recordAdminAudit_(actor,action,target,details){
  adminAuditSheet_().appendRow([new Date(),actor,action,target,details||'']);
}
function touchAdminSession_(email,token){
  const sheet=adminActivitySheet_(),hash=sha256_(token),lastRow=sheet.getLastRow();
  const row=lastRow>1?sheet.getRange(2,1,lastRow-1,1).getValues().findIndex(value=>String(value[0])===hash):-1;
  const now=new Date();
  if(row<0)sheet.appendRow([hash,email,now,now,'ACTIVE']);
  else sheet.getRange(row+2,2,1,4).setValues([[email,sheet.getRange(row+2,3).getValue(),now,'ACTIVE']]);
}
function revokeAdminActivity_(email){
  const sheet=adminActivitySheet_(),lastRow=sheet.getLastRow();
  if(lastRow<2)return;
  const rows=sheet.getRange(2,1,lastRow-1,5).getValues();
  rows.forEach((row,index)=>{if(String(row[1]||'').toLowerCase()===email)sheet.getRange(index+2,5).setValue('REVOKED')});
}
function adminAudit_(q){
  const actor=auth_(q.token);requirePrimaryAdmin_(actor);
  const audit=adminAuditSheet_(),rows=audit.getDataRange().getValues().slice(1).filter(row=>row[0]);
  const events=rows.slice(-500).reverse().map(row=>({timestamp:row[0],actor:String(row[1]||''),action:String(row[2]||''),target:String(row[3]||''),details:String(row[4]||'')}));
  const activity=adminActivitySheet_(),now=Date.now(),activeSessions=activity.getDataRange().getValues().slice(1)
    .filter(row=>row[0]&&row[4]==='ACTIVE'&&now-new Date(row[3]).getTime()<=150000)
    .map(row=>({email:String(row[1]||''),startedAt:row[2],lastSeenAt:row[3]}));
  return {events,activeSessions};
}
function requirePrimaryAdmin_(actor){
  if(actor!==PRIMARY_ADMIN_EMAIL)throw new Error('เฉพาะอีเมลหลัก '+PRIMARY_ADMIN_EMAIL+' เท่านั้นที่ดำเนินการนี้ได้');
}
function adminLogout_(q){
  const actor=auth_(q.token),sheet=adminActivitySheet_(),hash=sha256_(q.token),lastRow=sheet.getLastRow();
  if(lastRow>1){
    const row=sheet.getRange(2,1,lastRow-1,1).getValues().findIndex(value=>String(value[0])===hash);
    if(row>=0)sheet.getRange(row+2,5).setValue('LOGGED_OUT');
  }
  recordAdminAudit_(actor,'ออกจากระบบ','', '');
  CacheService.getScriptCache().remove('admin:'+q.token);
  return {ok:true};
}
function updateStatus_(q){
  const actor=auth_(q.token);if(!STATUSES.includes(q.status))throw Error('สถานะไม่ถูกต้อง');
  const sh=sheet_(),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&String(r[0])===String(q.id));if(i<1)throw Error('ไม่พบคำขอ');
  sh.getRange(i+1,23).setValue(q.status);sh.getRange(i+1,24).setValue(q.adminNote||'');
  recordAdminAudit_(actor,'เปลี่ยนสถานะคำขอ',String(q.id),String(v[i][22]||'')+' → '+q.status);
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
  const labels={'Pending':'รอตรวจสอบ','Approved / Quoted':'อนุมัติแล้ว','Rejected':'ไม่อนุมัติ','In Progress':'อนุมัติแล้ว','Completed':'เสร็จสิ้น'};
  return labels[status]||String(status||'ไม่ระบุ');
}
function emailHtml_(b,pdfUrl,adminCopy,update){
  const esc=escapeHtml_,service=(b.serviceNames||[]).map(esc).join(' · ')||'—';
  const compact=Boolean(update);
  const status=update?update.status:(b.status||'Pending'),statusLabel=statusLabel_(status);
  const statusStyle=status==='Rejected'?'background:#fde8e8;color:#a52d35':status==='Completed'?'background:#e5f5ec;color:#167247':status==='Approved / Quoted'||status==='In Progress'?'background:#e4f0fa;color:#174b78':'background:#fff2d6;color:#8a6113';
  const fieldPadding=compact?'8px 12px':'11px 14px';
  const field=(label,value)=>'<tr><td style="padding:'+fieldPadding+';color:#64748b;width:34%;border-bottom:1px solid #e8edf3">'+label+'</td><td style="padding:'+fieldPadding+';color:#172b45;font-weight:600;border-bottom:1px solid #e8edf3">'+esc(value==null||value===''?'—':value)+'</td></tr>';
  const greeting=update?'อัปเดตสถานะคำขอใช้บริการ':adminCopy?'มีคำขอใช้บริการผลิตและบริการสื่อรายการใหม่':'ระบบได้รับคำขอใช้บริการของท่านเรียบร้อยแล้ว';
  const intro=update?'เรียน คุณ'+esc(b.name)+' ระบบได้อัปเดตสถานะคำขอของท่านแล้ว':adminCopy?'โปรดตรวจสอบรายละเอียดและดำเนินการตามขั้นตอน':'เรียน คุณ'+esc(b.name)+' ขอบคุณที่ส่งคำขอผ่านระบบออนไลน์';
  const note=update&&update.adminNote?'<div style="background:#fff9eb;border:1px solid #f1dfb8;border-radius:10px;padding:'+(compact?'11px 14px':'14px 16px')+';margin:'+(compact?'14px 0':'18px 0')+'"><div style="font-weight:700;color:#765615">ข้อความจากเจ้าหน้าที่</div><div style="line-height:1.8;margin-top:5px">'+esc(update.adminNote)+'</div></div>':'';
  const graphicsDetails=b.graphicsDetails?'<div style="background:#f7f9fc;border:1px solid #e8edf3;border-radius:10px;padding:12px 14px;margin:14px 0;white-space:pre-wrap"><strong>รายละเอียดงานกราฟิก/สื่อสิ่งพิมพ์</strong><br>'+esc(b.graphicsDetails)+'</div>':'';
  const pdfButton=update&&pdfUrl?'<div style="text-align:center;margin:18px 0"><a href="'+esc(pdfUrl)+'" style="display:inline-block;background:#d9ad55;color:#172b45;text-decoration:none;font-weight:700;padding:10px 18px;border-radius:8px">เปิดเอกสารสรุปผล</a></div>':'';
  const noteText=update?'สถานะและรายละเอียดคำขอของท่านได้รับการอัปเดตโดยเจ้าหน้าที่ กรุณาเก็บหมายเลขคำขอนี้ไว้เพื่อติดตามสถานะ':'หมายเหตุ: ยอดดังกล่าวเป็นเพียงประมาณการ การให้บริการและค่าใช้จ่ายจะได้รับการยืนยันจากเจ้าหน้าที่อีกครั้ง กรุณาเก็บหมายเลขคำขอนี้ไว้เพื่อติดตามสถานะ';
  const header=compact
    ?'<tr><td style="padding:16px 24px;background:linear-gradient(120deg,#102d4f,#1d5688);border-bottom:4px solid #d9ad55;color:#fff;text-align:center"><div style="font-size:19px;font-weight:700">แบบขอรับบริการ (ส.1)</div><div style="font-size:13px;font-weight:600;margin-top:2px">Audio Visual Services Request Form (S1)</div><div style="font-size:11px;color:#e3edf6;margin-top:6px">ฝ่ายพัฒนาสื่อส่งเสริมการเรียนรู้ · สำนักพัฒนาการเรียนรู้ตลอดชีวิต มก.</div><div style="font-size:11px;color:#e3edf6;margin-top:3px">โทร. 0 2942 8820 – 4 · ภายใน 61 1336-9 ต่อ 500 · http://www.llldo.ku.ac.th</div></td></tr>'
    :'<tr><td style="padding:24px 30px;background:linear-gradient(120deg,#102d4f,#1d5688);border-bottom:5px solid #d9ad55;color:#fff;text-align:center"><div style="font-size:22px;font-weight:700">แบบขอรับบริการ (ส.1)</div><div style="font-size:16px;font-weight:600;margin-top:3px">Audio Visual Services Request Form (S1)</div><div style="font-size:13px;color:#e3edf6;margin-top:8px">ฝ่ายพัฒนาสื่อส่งเสริมการเรียนรู้</div><div style="font-size:13px;color:#e3edf6">สำนักพัฒนาการเรียนรู้ตลอดชีวิต มก.</div><div style="font-size:12px;color:#e3edf6;margin-top:5px">โทร. 0 2942 8820 – 4 · ภายใน 61 1336-9 ต่อ 500</div><div style="font-size:12px;color:#e3edf6">http://www.llldo.ku.ac.th</div></td></tr>';
  const width=compact?'640px':'680px',outerPadding=compact?'22px 12px':'28px 12px';
  const contentPadding=compact?'20px 24px 10px':'28px 30px 12px';
  const summaryPadding=compact?'13px 16px':'17px 20px',requestIdSize=compact?'21px':'25px';
  return '<div style="margin:0;background:#f2f5f9;padding:'+outerPadding+';font-family:Arial,Tahoma,sans-serif;color:#172b45"><table role="presentation" cellpadding="0" cellspacing="0" style="max-width:'+width+';width:100%;margin:auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e1e8f0">'+header+'<tr><td style="padding:'+contentPadding+'"><div style="font-size:17px;font-weight:700;color:#102d4f">'+greeting+'</div><p style="font-size:14px;line-height:1.8;color:#526276">'+intro+'</p><div style="background:#eef4fa;border:1px solid #dbe7f2;border-radius:12px;padding:'+summaryPadding+';margin:16px 0"><div style="font-size:12px;color:#61758a">หมายเลขคำขอ</div><div style="font-size:'+requestIdSize+';letter-spacing:1px;font-weight:700;color:#174b78;margin-top:4px">'+esc(b.id)+'</div><span style="display:inline-block;margin-top:8px;'+statusStyle+';padding:5px 12px;border-radius:20px;font-size:13px;font-weight:700">'+esc(statusLabel)+'</span></div><div style="font-size:15px;font-weight:700;color:#102d4f;margin:18px 0 8px">รายละเอียดคำขอ</div><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border:1px solid #e8edf3;border-radius:10px;overflow:hidden">'+field('ชื่อผู้ขอ',b.name)+field('หน่วยงาน',b.department)+field('โทรศัพท์',b.phone)+field('อีเมล',b.email)+field('กิจกรรม',b.event)+field('วันที่ใช้บริการ',b.date)+field('เวลา',b.start+' – '+b.end)+field('สถานที่',b.location)+field('บริการที่เลือก',service)+field('ประมาณการค่าใช้จ่าย',Number(b.estimate||0).toLocaleString('th-TH')+' บาท')+'</table>'+graphicsDetails+note+pdfButton+'<p style="font-size:13px;line-height:1.8;color:#718096">'+noteText+'</p></td></tr><tr><td style="padding:'+(compact?'14px 24px':'18px 30px')+';background:#f7f9fc;border-top:1px solid #e8edf3;font-size:12px;line-height:1.7;color:#718096">สำนักพัฒนาการเรียนรู้ตลอดชีวิต · มหาวิทยาลัยเกษตรศาสตร์<br>อีเมลฉบับนี้ส่งจากระบบอัตโนมัติ กรุณาอย่าตอบกลับอีเมลนี้โดยตรง</td></tr></table></div>';
}
function sendBookingEmail_(b){
  MailApp.sendEmail({to:b.email,subject:'ยืนยันรับคำขอ '+b.id+' | บริการผลิตและบริการสื่อ มก.',body:'ระบบได้รับคำขอ '+b.id+' แล้ว สถานะรอเจ้าหน้าที่ตรวจสอบ',htmlBody:emailHtml_(b,'',false),name:'KU Media Service'});
}
function notifyAdminsHtml_(b){
  adminEmails_().filter(to=>shouldNotifyAdmin_(to,b)).forEach(to=>MailApp.sendEmail({to,subject:'🔔 คำขอใช้บริการใหม่ '+b.id,body:'มีคำขอใหม่ '+b.id+' จาก '+b.name,htmlBody:emailHtml_(b,'',true),name:'KU Media Service'}));
}
function shouldNotifyAdmin_(email,b){
  if(!GRAPHICS_ONLY_ADMIN_EMAILS.includes(String(email||'').trim().toLowerCase()))return true;
  return (b.services||[]).some(serviceId=>PORTFOLIO_CATEGORY_BY_SERVICE[serviceId]==='graphics');
}
function makePdf_(b,title){
  const folder=DriveApp.getFolderById(PROP.getProperty('DRIVE_FOLDER_ID'));
  const rows=[['รหัสคำขอ',b.id],['ชื่อผู้ขอ',b.name],['หน่วยงาน',b.department||'-'],['โทรศัพท์',b.phone],['อีเมล',b.email],['กิจกรรม',b.event],['วันที่',b.date],['เวลา',b.start+'–'+b.end],['บริการ',(b.serviceNames||[]).join(', ')],['รายละเอียดงานกราฟิก/สื่อสิ่งพิมพ์',b.graphicsDetails||'-'],['ยอดประมาณการ',Number(b.estimate||0).toLocaleString()+' บาท'],['สถานะ',b.status||'Pending']];
  const html='<html><meta charset="utf-8"><body style="font-family:Arial;padding:32px"><h2>'+title+'</h2><h3>สำนักพัฒนาการเรียนรู้ตลอดชีวิต มก.</h3><table border="1" cellpadding="8" cellspacing="0">'+rows.map(r=>'<tr><td>'+escapeHtml_(r[0])+'</td><td>'+escapeHtml_(r[1])+'</td></tr>').join('')+'</table><p>เอกสารนี้เป็นเอกสารสรุปคำขอ ไม่ใช่ใบเสร็จรับเงินหรือใบกำกับภาษี</p></body></html>';
  const blob=Utilities.newBlob(html,'text/html','summary.html').getAs('application/pdf').setName(b.id+'-'+title+'.pdf');
  const file=folder.createFile(blob);file.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.VIEW);return file;
}
function adminPdf_(q){
  auth_(q.token);
  const id=String(q.id||''),b=all_().find(item=>item.id===id);
  if(!b)throw new Error('ไม่พบคำขอ');
  const html=String(q.html||'');
  if(!html.startsWith('<!doctype html>')||html.length>1500000)throw new Error('ไฟล์ PDF ไม่ถูกต้องหรือมีขนาดใหญ่เกินกำหนด');
  const filename='KU-Media-Request-'+id.replace(/[^A-Za-z0-9._-]/g,'_')+'.pdf';
  const pdf=Utilities.newBlob(html,'text/html','request.html').getAs('application/pdf').setName(filename);
  return {filename,mimeType:'application/pdf',data:Utilities.base64Encode(pdf.getBytes())};
}
function escapeHtml_(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function notifyAdmins_(subject,body){adminEmails_().forEach(to=>MailApp.sendEmail(to,subject,body))}
function adminEmails_(){
  const configured=(PROP.getProperty('ADMIN_EMAILS')||'').split(',').map(x=>x.trim().toLowerCase()).filter(Boolean);
  return [PRIMARY_ADMIN_EMAIL,...configured.filter(email=>email!==PRIMARY_ADMIN_EMAIL)];
}
function adminPasswordHashes_(){
  const value=PROP.getProperty('ADMIN_PASSWORD_HASHES');
  if(!value)return {};
  const hashes=JSON.parse(value);
  if(!hashes||Array.isArray(hashes)||typeof hashes!=='object')throw new Error('ข้อมูลรหัสผ่านผู้ดูแลไม่ถูกต้อง');
  return hashes;
}
function adminPasswordSetupRequired_(){return (PROP.getProperty('ADMIN_PASSWORD_SETUP_REQUIRED')||'').split(',').map(email=>email.trim().toLowerCase()).filter(Boolean)}
function addAdminEmail_(q){
  const actor=auth_(q.token);requirePrimaryAdmin_(actor);
  const email=String(q.email||'').trim().toLowerCase();
  if(!/^[^@\s]+@ku\.th$/.test(email))throw Error('เพิ่มผู้ดูแลได้เฉพาะอีเมล @ku.th');
  const emails=adminEmails_();
  const added=!emails.includes(email);
  if(added){
    emails.push(email);
    PROP.setProperty('ADMIN_PASSWORD_SETUP_REQUIRED',[...new Set(adminPasswordSetupRequired_().concat(email))].join(','));
  }
  PROP.setProperty('ADMIN_EMAILS',emails.join(','));
  if(added)recordAdminAudit_(actor,'เพิ่มอีเมลผู้ดูแล',email,'');
  return {ok:true,emails};
}
function removeAdminEmail_(q){
  const actor=auth_(q.token);requirePrimaryAdmin_(actor);
  const email=String(q.email||'').trim().toLowerCase(),emails=adminEmails_();
  if(email===PRIMARY_ADMIN_EMAIL)throw Error('ไม่สามารถนำอีเมลหลักออกจากระบบได้');
  if(!emails.includes(email))throw Error('ไม่พบอีเมลนี้ในรายชื่อผู้รับแจ้งเตือน');
  if(emails.length<=1)throw Error('ต้องมีอีเมลผู้ดูแลอย่างน้อย 1 รายการ');
  const remaining=emails.filter(item=>item!==email);
  invalidateAdminSessions_(email);
  PROP.setProperty('ADMIN_EMAILS',remaining.join(','));
  const hashes=adminPasswordHashes_();delete hashes[email];PROP.setProperty('ADMIN_PASSWORD_HASHES',JSON.stringify(hashes));
  PROP.setProperty('ADMIN_PASSWORD_SETUP_REQUIRED',adminPasswordSetupRequired_().filter(item=>item!==email).join(','));
  revokeAdminActivity_(email);
  recordAdminAudit_(actor,'นำอีเมลผู้ดูแลออก',email,'');
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
  const actor=auth_(q.token);
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
  recordAdminAudit_(actor,existing?'แก้ไขผลงาน':'เพิ่มผลงาน',String(row[0]),title);
  return {ok:true,item:{id:row[0],createdAt:row[1],category,title,description,fileId,imageUrl:portfolioImageUrl_(fileId),active:true,position}};
}
function reorderPortfolioItems_(q){
  const actor=auth_(q.token);
  const ids=q.ids;
  if(!Array.isArray(ids)||ids.some(id=>typeof id!=='string'||!id)||new Set(ids).size!==ids.length)throw new Error('ลำดับผลงานไม่ถูกต้อง');
  const lock=LockService.getScriptLock();lock.waitLock(12000);
  try{
    const sheet=portfolioSheet_(),values=sheet.getDataRange().getValues(),existingIds=values.slice(1).filter(row=>row[0]).map(row=>String(row[0]));
    if(ids.length!==existingIds.length||ids.some(id=>!existingIds.includes(id)))throw new Error('รายการผลงานมีการเปลี่ยนแปลง กรุณาโหลดรายการใหม่แล้วลองอีกครั้ง');
    const orderColumn=portfolioOrderColumn_(sheet),positions=new Map(ids.map((id,index)=>[id,index]));
    const ordered=values.slice(1).filter(row=>row[0]).map(row=>[positions.get(String(row[0]))]);
    if(ordered.length)sheet.getRange(2,orderColumn,ordered.length,1).setValues(ordered);
    recordAdminAudit_(actor,'เปลี่ยนลำดับผลงาน','',ids.length+' รายการ');
    return {ok:true};
  }finally{lock.releaseLock()}
}
function deletePortfolioItem_(q){
  const actor=auth_(q.token);
  const id=String(q.id||''),sheet=portfolioSheet_(),values=sheet.getDataRange().getValues();
  const rowIndex=values.findIndex((row,index)=>index>0&&String(row[0])===id);
  if(rowIndex<1)throw new Error('ไม่พบผลงานที่ต้องการลบ');
  const fileId=String(values[rowIndex][5]||'');
  if(fileId)DriveApp.getFileById(fileId).setTrashed(true);
  sheet.deleteRow(rowIndex+1);
  recordAdminAudit_(actor,'ลบผลงาน',id,String(values[rowIndex][3]||''));
  return {ok:true};
}
function removeBooking_(q){
  const actor=auth_(q.token);
  const id=String(q.id||''),sh=sheet_(),values=sh.getDataRange().getValues(),index=values.findIndex((row,n)=>n>0&&String(row[0])===id);
  if(index<1)throw Error('ไม่พบคำขอ');
  const queue=SpreadsheetApp.openById(PROP.getProperty('SHEET_ID')).getSheetByName('MailQueue');
  if(queue){
    const queued=queue.getDataRange().getValues();
    for(let row=queued.length-1;row>0;row--)if(String(queued[row][1])===id&&queued[row][4]==='QUEUED')queue.deleteRow(row+1);
  }
  sh.deleteRow(index+1);
  recordAdminAudit_(actor,'ลบคำขอใช้บริการ',id,'');
  return {ok:true};
}
function sha256_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>(b+256).toString(16).slice(-2)).join('')}
