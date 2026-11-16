/**
 * Cooper's Cricket Journey — Google Sheets backend
 * Bind this script to a NEW spreadsheet. Run setupDatabase once.
 * Deploy two web apps:
 *   PUBLIC: Execute as me, Anyone (read-only JSONP)
 *   EDITOR: Execute as user accessing, Anyone with Google account
 * Set EDITOR_EMAIL in Script Properties to the permitted Google email.
 *
 * IMPORTANT: Never put editing credentials in GitHub Pages.
 */
const TABS = {
 Players:['id','name','club','createdAt'],
 Seasons:['id','label','ageGroup','notes'],
 Matches:['id','opponent','matchDate','season','venue','teamOvers','ballAllocation','runs','balls','dismissals','fours','sixes','wickets','conceded','bowlingBalls','catches','runouts','result','notes','playCricketUrl','recordStatus','unknownFields','fieldSources','scorecardSnapshot','updatedAt'],
 Deliveries:['id','matchId','innings','overNumber','ballNumber','batterId','bowlerId','runs','extrasType','extrasRuns','wicketType','createdAt'],
 ScorecardSources:['id','matchId','url','sourceType','status','checkedAt','snapshot'],
 AuditLog:['id','matchId','action','editor','at']
};
function db_(){return SpreadsheetApp.getActiveSpreadsheet()}
function setupDatabase(){
 const ss=db_();
 Object.entries(TABS).forEach(([name,head])=>{
  let sh=ss.getSheetByName(name)||ss.insertSheet(name);
  if(sh.getLastRow()===0){sh.getRange(1,1,1,head.length).setValues([head]);sh.setFrozenRows(1)}
 });
 const players=ss.getSheetByName('Players');
 if(players.getLastRow()===1)players.appendRow(['cooper','Cooper','Donnybrook Cricket Club',new Date().toISOString()]);
 return 'Database ready. Set EDITOR_EMAIL in Project Settings → Script Properties.';
}
function readMatches_(){
 const sh=db_().getSheetByName('Matches');
 if(!sh||sh.getLastRow()<2)return [];
 const [head,...rows]=sh.getDataRange().getValues();
 return rows.filter(r=>r[0]).map(r=>{
  const o={};head.forEach((k,i)=>{let v=r[i];if(['unknownFields','fieldSources','scorecardSnapshot'].includes(k)){try{v=JSON.parse(v|| (k==='unknownFields'?'[]':'{}'))}catch{v=k==='unknownFields'?[]:{}}}else if(v instanceof Date)v=Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');o[k]=v===''?null:v});return o;
 });
}
function output_(obj,callback){
 const json=JSON.stringify(obj);
 // JSONP callback restricted to a simple identifier (read-only public endpoint)
 if(callback&&/^[a-zA-Z_$][\w$]{0,60}$/.test(callback))return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
 return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e){
 const p=e&&e.parameter||{};
 if(p.action==='publicMatches')return output_({ok:true,matches:readMatches_()},p.callback);
 // Editor UI is accessible only on the separate editor deployment.
 if(p.action==='editor'){
  assertEditor_();
  return HtmlService.createHtmlOutputFromFile('Editor').setTitle("Cooper's Cricket Journey — Editor").addMetaTag('viewport','width=device-width,initial-scale=1');
 }
 return output_({ok:false,error:'Unknown action'});
}
function assertEditor_(){
 const allowed=PropertiesService.getScriptProperties().getProperty('EDITOR_EMAIL');
 const email=Session.getActiveUser().getEmail();
 if(!allowed||!email||email.toLowerCase()!==allowed.toLowerCase())throw Error('Not authorised to edit Cooper’s cricket records.');
}
function editorMatches(){assertEditor_();return readMatches_()}
function saveMatch(payload){
 assertEditor_();
 if(!payload||typeof payload!=='object'||Array.isArray(payload))throw Error('Invalid match');
 const id=String(payload.id||Utilities.getUuid());
 const sh=db_().getSheetByName('Matches'),head=TABS.Matches;
 if(!String(payload.opponent||'').trim()||!/^\d{4}-\d{2}-\d{2}$/.test(String(payload.matchDate||'')))throw Error('Opponent and match date required');
 const allowed=['manual','awaiting','review','verified','complete','incomplete','inprogress'];
 if(payload.recordStatus&&!allowed.includes(payload.recordStatus))throw Error('Invalid status');
 if(payload.playCricketUrl){const u=String(payload.playCricketUrl);if(!/^https:\/\/(www\.)?play\.cricket\.com\.au\/match\/[a-zA-Z0-9-]+\/?$/.test(u))throw Error('Invalid PlayCricket URL')}
 const existing=readMatches_().find(m=>String(m.id)===id);
 const rowObj=Object.assign({},existing||{},payload,{id,updatedAt:new Date().toISOString()});
 const row=head.map(k=>{
  let v=rowObj[k]??'';
  if(['unknownFields','fieldSources','scorecardSnapshot'].includes(k))return JSON.stringify(v|| (k==='unknownFields'?[]:{}));
  if(typeof v==='string'&&/^[=+@\-\t\r]/.test(v)&&!['matchDate'].includes(k))v="'"+v; // spreadsheet formula injection guard
  return v;
 });
 const lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
  const ids=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat():[];
  const idx=ids.findIndex(x=>String(x)===id);
  if(idx>=0)sh.getRange(idx+2,1,1,head.length).setValues([row]);
  else sh.appendRow(row);
  db_().getSheetByName('AuditLog').appendRow([Utilities.getUuid(),id,idx>=0?'update':'create',Session.getActiveUser().getEmail(),new Date().toISOString()]);
 }finally{lock.releaseLock()}
 return {ok:true,id};
}
