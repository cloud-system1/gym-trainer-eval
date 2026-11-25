// Set Script Properties: SPREADSHEET_ID, SHEET_NAME, ADMIN_PASSWORD.
const NAMES = ['Bhagya','Sameera','Kavindu','Dilshan','Sasanka','Ravija','Thilini','Hiruka','Shalani','Pavani','Eran','Methum'];
const POINTS = {'Excellent':3,'Good':2,'Average':1,'Needs Improvement':0,'Poor':-1,'Not Familiar':0};
function json(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);}
function sheet(){const p=PropertiesService.getScriptProperties(); const s=SpreadsheetApp.openById(p.getProperty('SPREADSHEET_ID')).getSheetByName(p.getProperty('SHEET_NAME')); if(!s) throw Error('Sheet not found'); return s;}
function doGet(){return run(null);}
function doPost(e){try{return run(JSON.parse(e.postData.contents));}catch(err){return json({ok:false,error:String(err.message)});}}
function run(body){
 const lock=LockService.getScriptLock();
 try {
  lock.waitLock(30000);
  const s=sheet();
  if(s.getLastColumn()<5 || s.getLastRow()<1) throw Error('Expected existing five-column sheet with headers');
  const rows=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,5).getValues():[];
  if(!body){
   const stats=new Map();
   rows.forEach(r=>{const name=String(r[2]).trim(), rating=r[3]; if(!name||!Object.prototype.hasOwnProperty.call(POINTS,rating))return;
    if(!stats.has(name)){const t={name,clients:0,marks:0};Object.keys(POINTS).forEach(k=>t[k]=0);stats.set(name,t);}
    const t=stats.get(name);t[rating]++;if(rating!=='Not Familiar'){t.clients++;t.marks+=Number(r[4])||0;}
   });
   const data=Array.from(stats.values()).sort((a,b)=>(b.clients>0)-(a.clients>0)||b.marks-a.marks);let rank=0;data.forEach(t=>t.rank=t.clients?++rank:'—');return json({ok:true,data});
  }
  if(body.action){
   if(!['backupAndClear','clear'].includes(body.action))throw Error('Unsupported action');
   const password=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');if(!password||body.password!==password)throw Error('Invalid password');
   if(!rows.length)throw Error('No data to back up');
   const name='Trainer Evaluation - '+Utilities.formatDate(new Date(),'Asia/Colombo','MMMM dd, yyyy - HH-mm-ss');
   const backup=SpreadsheetApp.create(name);backup.setSpreadsheetTimeZone(s.getParent().getSpreadsheetTimeZone());backup.setSpreadsheetLocale(s.getParent().getSpreadsheetLocale());
   const copy=s.copyTo(backup);copy.setName(s.getName());backup.getSheets().filter(x=>x.getSheetId()!==copy.getSheetId()).forEach(x=>backup.deleteSheet(x));SpreadsheetApp.flush();
   if(JSON.stringify(s.getDataRange().getValues())!==JSON.stringify(copy.getDataRange().getValues()))throw Error('Backup verification failed; original retained');
   s.getRange(2,1,rows.length,5).clearContent();SpreadsheetApp.flush();return json({ok:true,backupName:name,backupUrl:backup.getUrl()});
  }
  if(!body.ratings||typeof body.ratings!=='object')throw Error('Missing ratings');
  const timestamp=new Date(), id='CUST-'+Utilities.getUuid();const entries=[];
  Object.keys(body.ratings).forEach(name=>{const rating=body.ratings[name];if(!rating||rating==='null')return;if(!NAMES.includes(name)||!Object.prototype.hasOwnProperty.call(POINTS,rating))throw Error('Invalid trainer or rating');entries.push([timestamp,id,name,rating,rating==='Not Familiar'?'':POINTS[rating]]);});
  if(!entries.length)throw Error('No ratings submitted');s.getRange(s.getLastRow()+1,1,entries.length,5).setValues(entries);return json({ok:true});
 }catch(err){return json({ok:false,error:String(err.message)});}finally{if(lock.hasLock())lock.releaseLock();}
}
