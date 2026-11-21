/** Live Song Requests — add as a separate .gs file to the existing bound Apps Script project.
 * Your existing PDF functions remain untouched.
 */
const SR = {
  spreadsheetId: '1pv3g2oRsBX6JuWqVTNfz-GIYV3TTTtckD1YURVSc7OU',
  playlist: 'Playlist',
  sessions: 'Sessions',
  requests: 'Requests',
  votes: 'Votes'
};

// Run once from the bound spreadsheet's Apps Script editor.
// A browser visitor cannot use this setup dialog. No password is stored in code.
function setupSongRequests() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set Song Requests admin password',
    'Use a unique password of at least 16 characters. It will be visible while typing.',
    ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const password = result.getResponseText();
  if (password.length < 16) {
    throw new Error('Use a unique admin password of at least 16 characters.');
  }
  const p = PropertiesService.getScriptProperties();
  const salt = Utilities.getUuid();
  p.setProperties({ SR_SALT: salt, SR_HASH: srHash_(salt + password) });
  srEnsureSheets_();
  Logger.log('Setup complete. Your password is stored as a salted hash.');
}

// If your project already has a doGet function from the earlier search page,
// REPLACE that function with this one (Apps Script must have only one doGet).
function doGet(e) {
  const admin = e && e.parameter && e.parameter.admin === '1';
  return HtmlService.createHtmlOutputFromFile(admin ? 'Admin' : 'Public')
    .setTitle(admin ? 'Song Requests — Admin' : 'Request a Song')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function srBook_() { return SpreadsheetApp.openById(SR.spreadsheetId); }
function srSheet_(name) { return srBook_().getSheetByName(name); }
function srRows_(name) {
  const sh = srSheet_(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run setupSongRequests first.');
  return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
}
function srEnsureSheets_() {
  const specs = {
    Sessions: ['ID','Date','Venue','State','StartedAt','ClosedAt','ArchiveURL'],
    Requests: ['ID','SessionID','Song','Artist','Requester','Status','Order','OwnerToken','CreatedAt','PlayedAt'],
    Votes: ['SessionID','RequestID','VisitorID','CreatedAt']
  };
  const ss = srBook_();
  Object.keys(specs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(specs[name]); sh.setFrozenRows(1);
      sh.getRange(1,1,1,specs[name].length).setFontWeight('bold');
    }
  });
}
function srHash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function srLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function srCurrent_() {
  const rows = srRows_(SR.sessions);
  const active = rows.filter(r => r[3] === 'OPEN' || r[3] === 'NO_REQUESTS');
  if (active.length > 1) throw new Error('Multiple open sessions; fix Sessions sheet.');
  if (!active.length) return null;
  const r = active[0];
  return { id:String(r[0]), date:String(r[1]), venue:String(r[2]), state:String(r[3]) };
}
function srPlaylist_() {
  const sh = srSheet_(SR.playlist);
  if (!sh) throw new Error('Playlist sheet not found.');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2,2,sh.getLastRow()-1,2).getDisplayValues()
    .filter(r => String(r[0]).trim())
    .map(r=>({song:String(r[0]).trim(),artist:String(r[1]).trim()}))
    .sort((a,b)=>a.song.localeCompare(b.song,undefined,{sensitivity:'base'}));
}
function srKey_(song,artist) { return [song,artist].map(x=>String(x).trim().toLocaleLowerCase()).join('\u0001'); }
function srRequests_(sessionId) {
  return srRows_(SR.requests).filter(r=>String(r[1])===sessionId);
}
function srVotes_(sessionId) {
  return srRows_(SR.votes).filter(r=>String(r[0])===sessionId);
}
function srPublicPayload_() {
  const session = srCurrent_();
  const rows = session ? srRequests_(session.id) : [];
  const counts = {};
  if (session) srVotes_(session.id).forEach(r=>counts[String(r[1])] = (counts[String(r[1])]||0)+1);
  const requests = rows.map(r=>({id:String(r[0]),song:String(r[2]),artist:String(r[3]),status:String(r[5]),order:r[6]===''?null:Number(r[6]),votes:counts[String(r[0])]||0}));
  return {session,requests,playlist:srPlaylist_()};
}
function getPublicData() { return srPublicPayload_(); }
function srFindRequest_(id, sid) {
  const sh = srSheet_(SR.requests);
  const rows = srRequests_(sid);
  const row = rows.find(r=>String(r[0])===String(id));
  if (!row) throw new Error('Request not found in current session.');
  const all = srRows_(SR.requests);
  const sheetRow = all.findIndex(r=>String(r[0])===String(id) && String(r[1])===sid)+2;
  return {row, sheetRow, sh};
}
function srValidVisitor_(visitor) {
  if (!/^[a-f0-9-]{30,80}$/i.test(String(visitor||''))) throw new Error('Invalid visitor identifier.');
  return String(visitor);
}
// First request for a song creates one item; subsequent visitors vote on it.
function requestSong(song, artist, requester, visitor) {
  return srLock_(()=>{
    srValidVisitor_(visitor);
    const session = srCurrent_();
    if (!session || session.state !== 'OPEN') throw new Error('Requests are closed.');
    const match = srPlaylist_().find(x=>srKey_(x.song,x.artist)===srKey_(song,artist));
    if (!match) throw new Error('Song is not in the repertoire.');
    const existing = srRequests_(session.id).find(r=>
      srKey_(r[2],r[3])===srKey_(song,artist) && String(r[5])!=='CANC');
    if (existing) return {alreadyExists:true,id:String(existing[0])};
    const id = Utilities.getUuid();
    const ownerToken = Utilities.getUuid()+Utilities.getUuid();
    srSheet_(SR.requests).appendRow([
      id,session.id,match.song,match.artist,String(requester||'').slice(0,50),
      'REQU','',ownerToken,new Date(),''
    ]);
    return {id,ownerToken,alreadyExists:false};
  });
}
function voteSong(id, visitor) {
  return srLock_(()=>{
    srValidVisitor_(visitor);
    const s = srCurrent_();
    if (!s || s.state!=='OPEN') throw new Error('Voting is closed.');
    const req = srFindRequest_(id,s.id).row;
    if (!['REQU','QUED'].includes(String(req[5]))) throw new Error('Voting closed for this song.');
    const v = srSheet_(SR.votes);
    const rows = srRows_(SR.votes);
    const index = rows.findIndex(r=>String(r[0])===s.id && String(r[1])===String(id) && String(r[2])===visitor);
    if (index>=0) {
      v.deleteRow(index+2); // Toggle off
      return {voted:false};
    }
    v.appendRow([s.id,id,visitor,new Date()]);
    return {voted:true};
  });
}
function cancelOwnRequest(id, ownerToken) {
  return srLock_(()=>{
    const s = srCurrent_();
    if (!s || s.state!=='OPEN') throw new Error('Requests are closed.');
    const f = srFindRequest_(id,s.id);
    if (f.row[5]!=='REQU' || String(f.row[7])!==String(ownerToken)) throw new Error('Cannot cancel this request.');
    f.sh.getRange(f.sheetRow,6).setValue('CANC');
    return true;
  });
}
// IMPORTANT: Admin identity must be checked on EVERY privileged operation.
// This is a secret-token admin session, not an unprotected ?admin=1 URL.
function adminLogin(password) {
  const p = PropertiesService.getScriptProperties();
  const hash = p.getProperty('SR_HASH');
  if (!hash || srHash_(p.getProperty('SR_SALT')+String(password||''))!==hash)
    throw new Error('Incorrect admin password.');
  const token = Utilities.getUuid()+Utilities.getUuid();
  CacheService.getScriptCache().put('SR_ADMIN_'+srHash_(token),'ok',21600);
  return token;
}
function srAuth_(token) {
  if (!/^[0-9a-f-]{70,80}$/i.test(String(token||'')) ||
      CacheService.getScriptCache().get('SR_ADMIN_'+srHash_(token))!=='ok') {
    throw new Error('Admin login required or expired.');
  }
}
function adminData(token) {
  srAuth_(token);
  return srPublicPayload_();
}
function adminStartSession(token, venue) {
  srAuth_(token);
  return srLock_(()=>{
    if (srCurrent_()) throw new Error('Close the current session first.');
    const id = 'S-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,6);
    srSheet_(SR.sessions).appendRow([id,new Date(),String(venue||'Live performance').slice(0,100),'OPEN',new Date(),'','']);
    return id;
  });
}
function srSessionSheetRow_(sid) {
  const n = srRows_(SR.sessions).findIndex(r=>String(r[0])===sid);
  if(n<0) throw new Error('Session missing.');
  return n+2;
}
function adminCloseRequests(token) {
  srAuth_(token);
  return srLock_(()=>{
    const s=srCurrent_();
    if(!s || s.state!=='OPEN') throw new Error('No session accepting requests.');
    srSheet_(SR.sessions).getRange(srSessionSheetRow_(s.id),4).setValue('NO_REQUESTS');
    return true;
  });
}
function adminReopenRequests(token) {
  srAuth_(token);
  return srLock_(()=>{
    const s=srCurrent_();
    if(!s || s.state!=='NO_REQUESTS') throw new Error('Requests already open or session closed.');
    srSheet_(SR.sessions).getRange(srSessionSheetRow_(s.id),4).setValue('OPEN');
    return true;
  });
}
function adminSetRequest(token,id,status,order) {
  srAuth_(token);
  return srLock_(()=>{
    const s=srCurrent_(); if(!s) throw new Error('No active session.');
    if(!['REQU','PARK','QUED','PLYD'].includes(String(status))) throw new Error('Invalid status.');
    const f=srFindRequest_(id,s.id);
    if (f.row[5]==='CANC') throw new Error('Cancelled requests cannot be changed.');
    const sh=f.sh;
    if(status==='QUED') {
      const n=Number(order);
      if(!Number.isInteger(n) || n<1 || n>9999) throw new Error('Order must be an integer from 1 to 9999.');
      sh.getRange(f.sheetRow,6,1,2).setValues([['QUED',n]]);
      sh.getRange(f.sheetRow,10).clearContent();
    } else {
      sh.getRange(f.sheetRow,6,1,2).setValues([[status,'']]);
      sh.getRange(f.sheetRow,10).setValue(status==='PLYD'?new Date():'');
    }
    return true;
  });
}
function srCsv_(rows) {
  return rows.map(row=>row.map(x=>'"'+String(x==null?'':x).replace(/"/g,'""')+'"').join(',')).join('\r\n');
}
function adminCloseSession(token) {
  srAuth_(token);
  return srLock_(()=>{
    const s=srCurrent_(); if(!s) throw new Error('No active session.');
    const requests = srRequests_(s.id), votes=srVotes_(s.id);
    const counts={}; votes.forEach(v=>counts[String(v[1])]=(counts[String(v[1])]||0)+1);
    const rows=[['Session','Song','Artist','Requester','Status','Order','Votes','RequestedAt','PlayedAt']];
    requests.forEach(r=>rows.push([s.id,r[2],r[3],r[4],r[5],r[6],counts[String(r[0])]||0,r[8],r[9]]));
    // Preserve detailed individual-vote data in a separate file.
    const folder=DriveApp.createFolder('Song Requests '+s.id);
    folder.createFile('Requests.csv',srCsv_(rows),MimeType.CSV);
    folder.createFile('Votes.csv',srCsv_([['SessionID','RequestID','VisitorID','CreatedAt'],...votes]),MimeType.CSV);
    const sh=srSheet_(SR.sessions), n=srSessionSheetRow_(s.id);
    sh.getRange(n,4).setValue('CLOSED');
    sh.getRange(n,6).setValue(new Date());
    sh.getRange(n,7).setValue(folder.getUrl());
    return folder.getUrl();
  });
}
