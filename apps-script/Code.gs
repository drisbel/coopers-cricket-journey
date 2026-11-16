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
 Matches:['id','opponent','matchDate','season','venue','teamOvers','ballAllocation','runs','balls','dismissals','fours','sixes','wickets','conceded','bowlingBalls','catches','runouts','result','notes','playCricketUrl','recordStatus','unknownFields','fieldSources','scorecardSnapshot','updatedAt','ourScore','ourWickets','opponentScore','opponentWickets','scorecardAvailability'],
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
  return HtmlService.createHtmlOutputFromFile('Editor').setTitle("Cooper's Cricket Journey — Editor").addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
 if(payload.playCricketUrl){const u=String(payload.playCricketUrl);if(!/^https:\/\/(?:www\.)?(?:play\.cricket\.com\.au\/match\/[a-zA-Z0-9-]+\/?|playhq\.com\/cricket-australia\/[a-zA-Z0-9/_-]+\/?)(?:[?#][^\\s]*)?$/i.test(u))throw Error('Enter a valid PlayCricket or PlayHQ match URL')}
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

/** Private live scoring. A complete event snapshot is written atomically under a script lock.
 * The client sends one request at a time; retrying the same snapshot is safe.
 */
function liveSession(matchId){
 assertEditor_();
 const id=String(matchId||'');
 if(!readMatches_().some(m=>String(m.id)===id))throw Error('Match not found');
 const sh=db_().getSheetByName('Deliveries');
 if(sh.getLastRow()<2)return [];
 const rows=sh.getRange(2,1,sh.getLastRow()-1,12).getValues();
 return rows.filter(r=>String(r[1])===id && String(r[2])==='cooper').map(r=>{
  try{return JSON.parse(String(r[9]||'{}'))}catch{return null}
 }).filter(Boolean);
}
function syncLiveSession(matchId,events){
 assertEditor_();
 const id=String(matchId||'');
 if(!Array.isArray(events)||events.length>1000)throw Error('Invalid scoring session');
 if(!readMatches_().some(m=>String(m.id)===id))throw Error('Save the match details first');
 const clean=events.map((e,i)=>{
  if(!e||!['bat','bowl'].includes(e.side)||!(typeof e.value==='number'&&[0,1,2,3,4,5,6].includes(e.value)||['out','extra','wide','noball','wicket','bye','legbye'].includes(e.value)))throw Error('Invalid delivery '+i);
  return {id:String(e.id||Utilities.getUuid()).slice(0,100),side:e.side,value:e.value};
 });
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const sh=db_().getSheetByName('Deliveries');
  const all=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,12).getValues():[];
  const keep=all.filter(r=>!(String(r[1])===id&&String(r[2])==='cooper'));
  if(all.length)sh.getRange(2,1,all.length,12).clearContent();
  const now=new Date().toISOString();
  const added=clean.map((e,i)=>[e.id,id,'cooper',Math.floor(i/6)+1,i%6+1,e.side==='bat'?'cooper':'',e.side==='bowl'?'cooper':'',typeof e.value==='number'?e.value:0,e.side,JSON.stringify(e),'',now]);
  const rows=keep.concat(added);
  if(rows.length)sh.getRange(2,1,rows.length,12).setValues(rows);
  const t={runs:0,balls:0,dismissals:0,fours:0,sixes:0,wickets:0,conceded:0,bowlingBalls:0};
  const sides=new Set();
  clean.forEach(e=>{
   sides.add(e.side);
   if(e.side==='bat'){
    if(typeof e.value==='number'){t.runs+=e.value;t.balls++;if(e.value===4)t.fours++;if(e.value===6)t.sixes++}
    else if(e.value==='out'){t.dismissals++;t.balls++}
   }else{
    if(typeof e.value==='number'){t.conceded+=e.value;t.bowlingBalls++}
    else if(e.value==='wicket'){t.wickets++;t.bowlingBalls++}
    else if(e.value==='wide'||e.value==='noball')t.conceded++;
    else if(e.value==='bye'||e.value==='legbye')t.bowlingBalls++;
   }
  });
  const matchSheet=db_().getSheetByName('Matches');
  const ids=matchSheet.getRange(2,1,matchSheet.getLastRow()-1,1).getValues().flat();
  const index=ids.findIndex(x=>String(x)===id);
  if(index<0)throw Error('Match not found');
  const head=matchSheet.getRange(1,1,1,matchSheet.getLastColumn()).getValues()[0];
  const row=index+2;
  const unknownCol=head.indexOf('unknownFields')+1;
  let unknown=[];
  try{unknown=JSON.parse(matchSheet.getRange(row,unknownCol).getValue()||'[]')}catch{}
  const sourcesCol=head.indexOf('fieldSources')+1;
  let sources={};
  try{sources=JSON.parse(matchSheet.getRange(row,sourcesCol).getValue()||'{}')}catch{}
  Object.keys(t).forEach(k=>{
   if((sides.has('bat')&&['runs','balls','dismissals','fours','sixes'].includes(k))||(sides.has('bowl')&&['wickets','conceded','bowlingBalls'].includes(k))||sources[k]==='live'){
    const c=head.indexOf(k)+1;if(c>0)matchSheet.getRange(row,c).setValue(t[k]);
    unknown=unknown.filter(x=>x!==k);sources[k]='live';
   }
  });
  matchSheet.getRange(row,unknownCol).setValue(JSON.stringify(unknown));
  matchSheet.getRange(row,sourcesCol).setValue(JSON.stringify(sources));
  matchSheet.getRange(row,head.indexOf('updatedAt')+1).setValue(now);
  return {ok:true,count:clean.length,totals:t};
 }finally{lock.releaseLock()}
}
