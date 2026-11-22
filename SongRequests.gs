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

const SR_COLS = { Sessions: 7, Requests: 10, Votes: 4 };
const SR_HEADERS = {
  Sessions: ['ID', 'Date', 'Venue', 'State', 'StartedAt', 'ClosedAt', 'ArchiveURL'],
  Requests: ['ID', 'SessionID', 'Song', 'Artist', 'Requester', 'Status', 'Order', 'OwnerToken', 'CreatedAt', 'PlayedAt'],
  Votes: ['SessionID', 'RequestID', 'VisitorID', 'CreatedAt']
};

// Run once from the bound spreadsheet's Apps Script editor.
// A browser visitor cannot use this setup dialog. No password is stored in code.
function setupSongRequests() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SR_HASH')) {
    const warn = ui.alert(
      'Replace admin password?',
      'A Song Requests admin password is already set. Replacing it signs every admin out of the web app. Existing requests stay in the sheet.',
      ui.ButtonSet.OK_CANCEL
    );
    if (warn !== ui.Button.OK) return;
  }
  const result = ui.prompt(
    'Set Song Requests admin password',
    'Use a unique password of at least 16 characters. It will be visible while typing.',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const password = result.getResponseText();
  if (String(password || '').length < 16) {
    throw new Error('Use a unique admin password of at least 16 characters.');
  }
  const salt = Utilities.getUuid();
  props.setProperties({ SR_SALT: salt, SR_HASH: srHash_(salt + String(password)) });
  srEnsureSheets_();
  Logger.log('Setup complete. The password is stored as a salted hash.');
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
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run setupSongRequests in Apps Script.');
  const width = SR_COLS[name];
  if (sh.getLastRow() < 2) return [];
  if (sh.getLastColumn() < width) {
    throw new Error(name + ' sheet is missing columns. Run setupSongRequests in Apps Script.');
  }
  return sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
}

function srEnsureSheets_() {
  const ss = srBook_();
  Object.keys(SR_HEADERS).forEach(name => {
    const headers = SR_HEADERS[name];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      return;
    }
    if (sh.getLastRow() === 1) {
      const width = Math.max(sh.getLastColumn(), 1);
      const current = sh.getRange(1, 1, 1, width).getValues()[0];
      const same = headers.every((header, index) => String(current[index] || '') === header);
      if (!same) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
      return;
    }
    if (sh.getLastColumn() < headers.length) {
      throw new Error(name + ' sheet is missing columns. Expected: ' + headers.join(', '));
    }
    const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const same = headers.every((header, index) => String(current[index] || '') === header);
    if (!same) {
      throw new Error(name + ' headers do not match Song Requests. Expected: ' + headers.join(', '));
    }
  });
}

function srHash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function srHashEquals_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  let diff = a.length ^ b.length;
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function srLock_(fn) {
  const lock = LockService.getScriptLock();
  let got = false;
  try {
    got = lock.tryLock(10000);
  } catch (ignore) {
    got = false;
  }
  if (!got) throw new Error('The server is busy. Please try again.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// Shared snapshot so a room of phones does not each read the whole spreadsheet.
function srInvalidatePublic_() {
  try {
    const token = String(Date.now()) + '-' + Utilities.getUuid().slice(0, 8);
    CacheService.getScriptCache().put('SR_PUBLIC_GEN', token, 21600);
  } catch (ignore) {}
}

function srPublicState_() {
  const cache = CacheService.getScriptCache();
  const gen = cache.get('SR_PUBLIC_GEN') || '0';
  const key = 'SR_STATE_' + gen;
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (ignore) {}
  }
  const fresh = srBuildPublicState_();
  try {
    let stored = fresh;
    let json = JSON.stringify(stored);
    if (json.length >= 90000) {
      stored = { session: fresh.session, requests: fresh.requests, voteHashes: null };
      json = JSON.stringify(stored);
    }
    if (json.length < 90000 && (cache.get('SR_PUBLIC_GEN') || '0') === gen) cache.put(key, json, 12);
  } catch (ignore) {}
  return fresh;
}

function srCurrent_() {
  const rows = srRows_(SR.sessions);
  const active = rows.filter(r => r[3] === 'OPEN' || r[3] === 'NO_REQUESTS');
  if (active.length > 1) throw new Error('Multiple open sessions; fix the Sessions sheet.');
  if (!active.length) return null;
  const r = active[0];
  return { id: String(r[0]), venue: String(r[2]), state: String(r[3]) };
}

// Playlist columns A and D+ stay on the sheet. The public API only receives B (song) and C (artist).
function srPlaylist_() {
  const sh = srSheet_(SR.playlist);
  if (!sh) throw new Error('Playlist sheet not found.');
  if (sh.getLastRow() < 2) return [];
  const seen = {};
  const list = [];
  sh.getRange(2, 2, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
    const song = String(r[0] || '').trim();
    if (!song) return;
    const artist = String(r[1] || '').trim();
    const key = srKey_(song, artist);
    if (seen[key]) return;
    seen[key] = true;
    list.push({ song: song, artist: artist });
  });
  list.sort((a, b) => a.song.localeCompare(b.song, undefined, { sensitivity: 'base' })
    || a.artist.localeCompare(b.artist, undefined, { sensitivity: 'base' }));
  return list;
}

function srPlaylistFor_(clientHash) {
  const cache = CacheService.getScriptCache();
  const known = cache.get('SR_PLAYLIST_HASH');
  const client = typeof clientHash === 'string' ? clientHash : '';
  if (known && client && client === known) return { hash: known, songs: null };
  if (known) {
    const body = cache.get('SR_PLAYLIST_BODY');
    if (body) {
      try {
        return { hash: known, songs: JSON.parse(body) };
      } catch (ignore) {}
    }
  }
  const songs = srPlaylist_();
  const next = srHash_(songs.map(s => s.song + '\u0001' + s.artist).join('\n')).slice(0, 20);
  try {
    cache.put('SR_PLAYLIST_HASH', next, 120);
    const json = JSON.stringify(songs);
    if (json.length < 90000) cache.put('SR_PLAYLIST_BODY', json, 120);
  } catch (ignore) {}
  if (client && client === next) return { hash: next, songs: null };
  return { hash: next, songs: songs };
}

function srCanon_(value) {
  return String(value || '').normalize('NFC').trim().toLowerCase();
}

function srKey_(song, artist) {
  return srCanon_(song) + '\u0001' + srCanon_(artist);
}

function srRequests_(sessionId) {
  return srRows_(SR.requests).filter(r => String(r[1]) === String(sessionId));
}

function srVotes_(sessionId) {
  return srRows_(SR.votes).filter(r => String(r[0]) === String(sessionId));
}

function srOrder_(value) {
  if (value === '' || value == null) return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function srTime_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value.getTime();
  return null;
}

function srParseOrder_(order) {
  const text = String(order == null ? '' : order).trim();
  if (!text) return null;
  if (!/^\d+$/.test(text)) {
    throw new Error('Order must be a whole number from 1 to 9999, or blank to add the song at the end.');
  }
  const n = Number(text);
  if (n < 1 || n > 9999) throw new Error('Order must be a whole number from 1 to 9999.');
  return n;
}

function srNextOrder_(sid) {
  let max = 0;
  srRequests_(sid).forEach(row => {
    if (String(row[5]) !== 'QUED') return;
    const n = Number(row[6]);
    if (Number.isInteger(n) && n > max) max = n;
  });
  if (max >= 9999) throw new Error('The queue is full. Reuse an order number.');
  return max + 1;
}

function srVisitorOrNull_(visitor) {
  const value = String(visitor || '').trim().toLowerCase();
  return /^[a-f0-9-]{30,80}$/.test(value) ? value : '';
}

function srValidVisitor_(visitor) {
  const value = srVisitorOrNull_(visitor);
  if (!value) throw new Error('Invalid visitor identifier.');
  return value;
}

function srFindRequest_(id, sid) {
  const sh = srSheet_(SR.requests);
  const all = srRows_(SR.requests);
  const index = all.findIndex(r => String(r[0]) === String(id) && String(r[1]) === String(sid));
  if (index < 0) throw new Error('Request not found in current session.');
  return { row: all[index], sheetRow: index + 2, sh: sh };
}

function srBuildPublicState_() {
  const session = srCurrent_();
  if (!session) return { session: null, requests: [], voteHashes: {} };
  const counts = {};
  const voteHashes = {};
  srVotes_(session.id).forEach(row => {
    const id = String(row[1]);
    counts[id] = (counts[id] || 0) + 1;
    const visitor = String(row[2] || '').toLowerCase();
    if (!visitor) return;
    if (!voteHashes[id]) voteHashes[id] = [];
    voteHashes[id].push(srHash_(visitor).slice(0, 16));
  });
  const requests = srRequests_(session.id).map(row => ({
    id: String(row[0]),
    song: String(row[2]),
    artist: String(row[3]),
    status: String(row[5]),
    order: srOrder_(row[6]),
    votes: counts[String(row[0])] || 0,
    createdAt: srTime_(row[8]),
    playedAt: srTime_(row[9])
  }));
  return { session: session, requests: requests, voteHashes: voteHashes };
}

function srVotesForVisitor_(state, visitorId) {
  const out = {};
  if (!state || !state.session || !visitorId) return out;
  if (state.voteHashes) {
    const hash = srHash_(visitorId).slice(0, 16);
    Object.keys(state.voteHashes).forEach(id => {
      const list = state.voteHashes[id] || [];
      for (let i = 0; i < list.length; i++) {
        if (list[i] === hash) {
          out[id] = true;
          break;
        }
      }
    });
    return out;
  }
  srVotes_(state.session.id).forEach(row => {
    if (String(row[2] || '').toLowerCase() === visitorId) out[String(row[1])] = true;
  });
  return out;
}

function getPublicData(clientHash, visitor) {
  const state = srPublicState_();
  const playlist = srPlaylistFor_(clientHash);
  const visitorId = srVisitorOrNull_(visitor);
  const voted = srVotesForVisitor_(state, visitorId);
  return {
    session: state.session,
    requests: state.requests.map(row => ({
      id: row.id,
      song: row.song,
      artist: row.artist,
      status: row.status,
      order: row.order,
      votes: row.votes,
      createdAt: row.createdAt,
      playedAt: row.playedAt,
      voted: visitorId ? !!voted[row.id] : false
    })),
    playlistHash: playlist.hash,
    playlist: playlist.songs,
    serverTime: Date.now()
  };
}

// First request for a song creates one item; subsequent visitors vote on it.
function requestSong(song, artist, requester, visitor) {
  return srLock_(() => {
    srValidVisitor_(visitor);
    const session = srCurrent_();
    if (!session || session.state !== 'OPEN') throw new Error('Requests are closed.');
    const match = srPlaylist_().find(item => srKey_(item.song, item.artist) === srKey_(song, artist));
    if (!match) throw new Error('Song is not in the repertoire.');
    const existing = srRequests_(session.id).find(row =>
      srKey_(row[2], row[3]) === srKey_(song, artist) && String(row[5]) !== 'CANC');
    if (existing) {
      return { alreadyExists: true, id: String(existing[0]), status: String(existing[5]) };
    }
    const id = Utilities.getUuid();
    const ownerToken = Utilities.getUuid() + Utilities.getUuid();
    srSheet_(SR.requests).appendRow([
      id, session.id, match.song, match.artist, String(requester || '').slice(0, 50),
      'REQU', '', ownerToken, new Date(), ''
    ]);
    srInvalidatePublic_();
    return { id: id, ownerToken: ownerToken, alreadyExists: false };
  });
}

function voteSong(id, visitor) {
  return srLock_(() => {
    const who = srValidVisitor_(visitor);
    const session = srCurrent_();
    if (!session || session.state !== 'OPEN') throw new Error('Voting is closed.');
    const request = srFindRequest_(id, session.id).row;
    if (['REQU', 'QUED'].indexOf(String(request[5])) < 0) throw new Error('Voting is closed for this song.');
    const sheet = srSheet_(SR.votes);
    const rows = srRows_(SR.votes);
    let index = -1;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === session.id && String(rows[i][1]) === String(id) && String(rows[i][2]).toLowerCase() === who) {
        index = i;
        break;
      }
    }
    if (index >= 0) {
      const sheetRow = index + 2;
      if (sheetRow < 2) throw new Error('Could not update that vote.');
      sheet.deleteRow(sheetRow);
      srInvalidatePublic_();
      return { voted: false };
    }
    sheet.appendRow([session.id, id, who, new Date()]);
    srInvalidatePublic_();
    return { voted: true };
  });
}

function cancelOwnRequest(id, ownerToken) {
  return srLock_(() => {
    const session = srCurrent_();
    if (!session || session.state !== 'OPEN') throw new Error('Requests are closed.');
    if (!/^[-0-9a-f]{70,80}$/i.test(String(ownerToken || ''))) throw new Error('Cannot cancel this request.');
    const found = srFindRequest_(id, session.id);
    if (String(found.row[5]) !== 'REQU' || String(found.row[7]) !== String(ownerToken)) {
      throw new Error('Cannot cancel this request.');
    }
    found.sh.getRange(found.sheetRow, 6).setValue('CANC');
    srInvalidatePublic_();
    return true;
  });
}

function srLoginGate_() {
  let raw = '';
  try { raw = Session.getTemporaryActiveUserKey() || ''; } catch (ignore) { raw = ''; }
  if (!raw) {
    return {
      blocked: function () { return false; },
      fail: function () {},
      clear: function () {}
    };
  }
  const cache = CacheService.getScriptCache();
  const key = 'SR_LOGIN_' + srHash_(raw).slice(0, 32);
  return {
    blocked: function () { return Number(cache.get(key) || '0') >= 8; },
    fail: function () {
      const next = Number(cache.get(key) || '0') + 1;
      cache.put(key, String(next), 600);
    },
    clear: function () { cache.remove(key); }
  };
}

// Admin identity is checked on every privileged operation.
// ?admin=1 only chooses the page. It does not authorize anyone.
function adminLogin(password) {
  const gate = srLoginGate_();
  if (gate.blocked()) throw new Error('Too many sign-in attempts. Try again in a few minutes.');
  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty('SR_SALT');
  const hash = props.getProperty('SR_HASH');
  if (!salt || !hash) throw new Error('Song Requests setup is incomplete. Run setupSongRequests in the Apps Script editor.');
  const actual = srHash_(salt + String(password || ''));
  if (!srHashEquals_(actual, hash)) {
    gate.fail();
    throw new Error('Incorrect admin password.');
  }
  gate.clear();
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('SR_ADMIN_' + srHash_(token), 'ok', 21600);
  return token;
}

function adminLogout(token) {
  const value = String(token || '');
  if (/^[0-9a-f-]{70,80}$/i.test(value)) {
    CacheService.getScriptCache().remove('SR_ADMIN_' + srHash_(value));
  }
  return true;
}

function srAuth_(token) {
  const value = String(token || '');
  if (!/^[0-9a-f-]{70,80}$/i.test(value)) throw new Error('Admin login required or expired.');
  const cache = CacheService.getScriptCache();
  const key = 'SR_ADMIN_' + srHash_(value);
  if (cache.get(key) !== 'ok') throw new Error('Admin login required or expired.');
  cache.put(key, 'ok', 21600);
}

function adminData(token) {
  srAuth_(token);
  const state = srBuildPublicState_();
  return { session: state.session, requests: state.requests, serverTime: Date.now() };
}

function adminStartSession(token, venue) {
  srAuth_(token);
  return srLock_(() => {
    if (srCurrent_()) throw new Error('Close the current session first.');
    const name = String(venue || '').trim().slice(0, 100) || 'Live performance';
    const id = 'S-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 6);
    srSheet_(SR.sessions).appendRow([id, new Date(), name, 'OPEN', new Date(), '', '']);
    srInvalidatePublic_();
    return id;
  });
}

function srSessionSheetRow_(sid) {
  const index = srRows_(SR.sessions).findIndex(row => String(row[0]) === String(sid));
  if (index < 0) throw new Error('Session missing.');
  return index + 2;
}

function adminCloseRequests(token) {
  srAuth_(token);
  return srLock_(() => {
    const session = srCurrent_();
    if (!session || session.state !== 'OPEN') throw new Error('No session accepting requests.');
    srSheet_(SR.sessions).getRange(srSessionSheetRow_(session.id), 4).setValue('NO_REQUESTS');
    srInvalidatePublic_();
    return true;
  });
}

function adminReopenRequests(token) {
  srAuth_(token);
  return srLock_(() => {
    const session = srCurrent_();
    if (!session || session.state !== 'NO_REQUESTS') throw new Error('Requests are already open, or the session is closed.');
    srSheet_(SR.sessions).getRange(srSessionSheetRow_(session.id), 4).setValue('OPEN');
    srInvalidatePublic_();
    return true;
  });
}

function adminSetRequest(token, id, status, order) {
  srAuth_(token);
  return srLock_(() => {
    const session = srCurrent_();
    if (!session) throw new Error('No active session.');
    const nextStatus = String(status || '').toUpperCase();
    if (['REQU', 'PARK', 'QUED', 'PLYD'].indexOf(nextStatus) < 0) throw new Error('Invalid status.');
    const found = srFindRequest_(id, session.id);
    if (String(found.row[5]) === 'CANC') throw new Error('Cancelled requests cannot be changed.');
    if (nextStatus === 'QUED') {
      const parsed = srParseOrder_(order);
      const n = parsed == null ? srNextOrder_(session.id) : parsed;
      found.sh.getRange(found.sheetRow, 6, 1, 2).setValues([['QUED', n]]);
      found.sh.getRange(found.sheetRow, 10).clearContent();
    } else {
      found.sh.getRange(found.sheetRow, 6, 1, 2).setValues([[nextStatus, '']]);
      found.sh.getRange(found.sheetRow, 10).setValue(nextStatus === 'PLYD' ? new Date() : '');
    }
    srInvalidatePublic_();
    return true;
  });
}

function srFormatWhen_(value) {
  if (value == null || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

function srCsvCell_(value) {
  let text = srFormatWhen_(value);
  if (/^[\t\r ]*[=+\-@]/.test(text)) text = "'" + text;
  return '"' + text.replace(/"/g, '""') + '"';
}

function srCsv_(rows) {
  return '\uFEFF' + rows.map(row => row.map(srCsvCell_).join(',')).join('\r\n');
}

function adminCloseSession(token) {
  srAuth_(token);
  return srLock_(() => {
    const session = srCurrent_();
    if (!session) throw new Error('No active session.');
    const requests = srRequests_(session.id);
    const votes = srVotes_(session.id);
    const counts = {};
    votes.forEach(row => { counts[String(row[1])] = (counts[String(row[1])] || 0) + 1; });
    const rows = [['Session', 'Song', 'Artist', 'Requester', 'Status', 'Order', 'Votes', 'RequestedAt', 'PlayedAt']];
    requests.forEach(row => rows.push([
      session.id, row[2], row[3], row[4], row[5], row[6], counts[String(row[0])] || 0, row[8], row[9]
    ]));
    const folder = DriveApp.createFolder('Song Requests ' + session.id);
    folder.createFile('Requests.csv', srCsv_(rows), MimeType.CSV);
    folder.createFile('Votes.csv', srCsv_([['SessionID', 'RequestID', 'VisitorID', 'CreatedAt']].concat(votes)), MimeType.CSV);
    const sh = srSheet_(SR.sessions);
    const n = srSessionSheetRow_(session.id);
    sh.getRange(n, 4).setValue('CLOSED');
    sh.getRange(n, 6).setValue(new Date());
    sh.getRange(n, 7).setValue(folder.getUrl());
    srInvalidatePublic_();
    return folder.getUrl();
  });
}
