// ============================================================
// BLAZE BOT – Full Script with Diagnostic Status Command
// ============================================================

const SID = '';
const BID = '';
// ⚠️ CRITICAL: This MUST be your Personal Access Token (from dev.groupme.com/session), NOT the Bot Token.
const TKN = ''; 
const GID = '';

let _db = null, _idx = null, _cfg = null, _groupData = null;

// ------------------- Refresh Group Data -------------------
function refreshMembers() {
  CacheService.getScriptCache().remove('group_data_cache');
  _groupData = null;
  return fetchGroupData();
}

// ------------------- Sheet Init -------------------
function initSheet() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('si') === new Date().toDateString()) return;
  const ss = SpreadsheetApp.openById(SID);
  ['Config', 'Users', 'Warnings', 'Notes', 'Activity_Log', 'Welcome'].forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  const c = ss.getSheetByName('Config');
  if (c.getLastRow() === 0) c.getRange(1, 1, 8, 2).setValues([['Setting', 'Value'], ['RaidMode', 'false'], ['AllowAdd', 'true'], ['AllowChat', 'true'], ['AutoReadd', 'true'], ['WelcomeEnabled', 'true'], ['MaxStrikes', '3'], ['KickStrikeEnabled', 'true']]);
  const u = ss.getSheetByName('Users');
  if (u.getLastRow() === 0) u.appendRow(['UserID', 'Name', 'Trusted', 'Muted', 'Banned', 'MuteExpiry', 'MuteStrikes', 'Phone', 'ProfStrikes', 'Kicked']);
  if (ss.getSheetByName('Activity_Log').getLastRow() === 0) ss.getSheetByName('Activity_Log').appendRow(['Timestamp', 'UserID', 'Username', 'Action', 'Details']);
  const w = ss.getSheetByName('Welcome');
  if (w.getLastRow() === 0) { w.appendRow(['Enabled', 'Message']); w.appendRow(['true', 'Hello $u, Please add to this group.']); }
  p.setProperty('si', new Date().toDateString());
}

function logError(action, details) {
  try { SpreadsheetApp.openById(SID).getSheetByName('Activity_Log').appendRow([new Date(), 'SYSTEM', 'ERROR', action, String(details).substring(0, 500)]); } catch(e) {}
}

// ------------------- Group Data Fetch with Caching -------------------
function fetchGroupData() {
  if (_groupData) return _groupData;
  const c = CacheService.getScriptCache();
  const cached = c.get('group_data_cache');
  if (cached) { try { _groupData = JSON.parse(cached); if(_groupData) return _groupData; } catch(e) {} }
  try {
    const url = `https://api.groupme.com/v3/groups/${GID}?token=${TKN}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      _groupData = JSON.parse(response.getContentText()).response;
      if (_groupData) c.put('group_data_cache', JSON.stringify(_groupData), 21600);
      return _groupData;
    } else {
      logError('API_FAIL', `HTTP ${response.getResponseCode()}: ${response.getContentText().substring(0, 200)}`);
    }
  } catch (e) { logError('API_EXCEPTION', e.toString()); }
  return null;
}

function gmMembers() {
  const data = fetchGroupData();
  if (data && data.members && data.members.length > 0) return data.members;
  return [];
}

function gmInfo() { return fetchGroupData(); }

// ------------------- Membership ID Mapping -------------------
function getMembershipId(userId) {
  refreshMembers();
  const members = gmMembers();
  const member = members.find(m => String(m.user_id).trim() === String(userId).trim());
  return member ? member.id : null;
}

// ------------------- KICK -------------------
function rmMember(userId) {
  refreshMembers();
  const membershipId = getMembershipId(userId);
  if (!membershipId) {
    return { success: false, error: 'User not found in group (membership_id missing).' };
  }
  try {
    const url = `https://api.groupme.com/v3/groups/${GID}/members/${membershipId}/remove?token=${TKN}`;
    let r = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true });
    let code = r.getResponseCode();
    if ([200, 201, 202, 204].includes(code)) {
      return { success: true };
    }
    let error = `POST ${code}: ${r.getContentText().substring(0, 100)}`;
    r = UrlFetchApp.fetch(url, { method: 'delete', muteHttpExceptions: true });
    code = r.getResponseCode();
    if ([200, 201, 202, 204].includes(code)) {
      return { success: true };
    }
    error += ` | DELETE ${code}: ${r.getContentText().substring(0, 100)}`;
    return { success: false, error };
  } catch (e) {
    return { success: false, error: 'Exception: ' + e.toString() };
  }
}

// ------------------- ADD / READD (With Explicit Error Reporting) -------------------
function addMem(phone, uid) {
  try {
    let payload = {};
    if (phone) {
      let cleanPhone = String(phone).replace(/[^0-9+]/g, '');
      if (!cleanPhone.startsWith('+')) cleanPhone = '+' + cleanPhone;
      payload.phone_number = cleanPhone;
      payload.nickname = "Invited User";
    }
    if (uid) {
      payload.user_id = String(uid);
    }
    if (!phone && !uid) return { success: false, error: 'No phone or user_id provided.' };

    const url = `https://api.groupme.com/v3/groups/${GID}/members/add?token=${TKN}`;
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ members: [payload] }),
      muteHttpExceptions: true
    });
    
    const code = res.getResponseCode();
    if ([200, 201, 202].includes(code)) {
      return { success: true };
    }
    
    let errorBody = 'Unknown API error';
    try { 
      const errJson = JSON.parse(res.getContentText()); 
      if (errJson.meta && errJson.meta.errors && Array.isArray(errJson.meta.errors)) {
        errorBody = errJson.meta.errors.join(', ');
      } else if (errJson.errors && Array.isArray(errJson.errors)) {
        errorBody = errJson.errors.join(', ');
      } else {
        errorBody = res.getContentText().substring(0, 200);
      }
    } catch(e) { 
      errorBody = res.getContentText().substring(0, 200); 
    }
    
    return { success: false, error: `HTTP ${code}: ${errorBody}` };
  } catch (e) {
    return { success: false, error: 'Exception: ' + e.toString() };
  }
}

// ------------------- Messaging -------------------
function send(t) { 
  try { 
    UrlFetchApp.fetch('https://api.groupme.com/v3/bots/post', { 
      method: 'post', 
      contentType: 'application/json', 
      payload: JSON.stringify({ bot_id: BID, text: t }), 
      muteHttpExceptions: true 
    }); 
  } catch (e) {} 
}

function isT(v) { 
  if (v === true || v === 1 || v === '1') return true; 
  const s = String(v).toLowerCase().trim(); 
  return s === 'true' || s === 'yes'; 
}

// ------------------- Config & DB Loaders -------------------
function loadCfg() {
  if (_cfg) return; _cfg = {};
  const d = SpreadsheetApp.openById(SID).getSheetByName('Config').getDataRange().getValues();
  for (let i = 1; i < d.length; i++) _cfg[String(d[i][0]).toLowerCase().replace(/[\s_]/g, '')] = String(d[i][1]).toLowerCase();
}

function loadDB() {
  if (_db) return;
  const s = SpreadsheetApp.openById(SID).getSheetByName('Users');
  _db = s.getDataRange().getValues();
  if (_db.length === 0) return;
  const h = _db[0].map(x => String(x).toLowerCase().replace(/[\s_]/g, ''));
  const f = n => h.findIndex(x => x === n.toLowerCase().replace(/[\s_]/g, ''));
  let nc = h.length + 1;
  const ensure = (name) => { let i = f(name); if (i === -1) { h.push(name.toLowerCase().replace(/[\s_]/g, '')); i = h.length - 1; s.getRange(1, nc++).setValue(name); } return i; };
  _idx = { 
    id: ensure('UserID'), nm: ensure('Name'), tr: ensure('Trusted'), mu: ensure('Muted'), 
    ba: ensure('Banned'), ex: ensure('MuteExpiry'), st: ensure('MuteStrikes'), 
    ph: ensure('Phone'), pr: ensure('ProfStrikes'), ki: ensure('Kicked') 
  };
}

function getCfg(k) { loadCfg(); return _cfg[String(k).toLowerCase().replace(/[\s_]/g, '')] || null; }
function setCfg(k, v) {
  const cleanK = String(k).toLowerCase().replace(/[\s_]/g, '');
  _cfg[cleanK] = v; const s = SpreadsheetApp.openById(SID).getSheetByName('Config'); const d = s.getDataRange().getValues();
  for (let i = 0; i < d.length; i++) { if (String(d[i][0]).toLowerCase().replace(/[\s_]/g, '') === cleanK) { s.getRange(i + 1, 2).setValue(v); return; } } 
  s.appendRow([k, v]);
}

// ------------------- User DB Operations -------------------
function getUser(uid) {
  loadDB(); let lastMatch = null;
  for (let r = 1; r < _db.length; r++) {
    if (!_db[r][_idx.id]) continue;
    if (String(_db[r][_idx.id]).trim() === String(uid).trim()) {
      lastMatch = { 
        row: r, userId: _db[r][_idx.id], name: _db[r][_idx.nm] || 'Unknown', 
        trusted: isT(_db[r][_idx.tr]), muted: isT(_db[r][_idx.mu]), 
        banned: isT(_db[r][_idx.ba]), muteExpiry: parseInt(_db[r][_idx.ex]) || 0, 
        muteStrikes: parseInt(_db[r][_idx.st]) || 0, phone: _db[r][_idx.ph] || '', 
        profStrikes: parseInt(_db[r][_idx.pr]) || 0, kicked: isT(_db[r][_idx.ki]) 
      };
    }
  }
  return lastMatch;
}

function setUser(uid, name, u) {
  loadDB(); const s = SpreadsheetApp.openById(SID).getSheetByName('Users'); const i = _idx; let found = false;
  for (let r = 1; r < _db.length; r++) {
    if (!_db[r][i.id]) continue;
    if (String(_db[r][i.id]).trim() === String(uid).trim()) {
      found = true; const row = r + 1;
      const set = (col, val) => { s.getRange(row, col + 1).setValue(val); _db[r][col] = val; };
      if (u.name !== undefined) set(i.nm, u.name);
      if (u.trusted !== undefined) set(i.tr, u.trusted ? 'true' : 'false');
      if (u.muted !== undefined) set(i.mu, u.muted ? 'true' : 'false');
      if (u.banned !== undefined) set(i.ba, u.banned ? 'true' : 'false');
      if (u.muteExpiry !== undefined) set(i.ex, u.muteExpiry);
      if (u.muteStrikes !== undefined) set(i.st, u.muteStrikes);
      if (u.phone !== undefined) set(i.ph, u.phone);
      if (u.profStrikes !== undefined) set(i.pr, u.profStrikes);
      if (u.kicked !== undefined) set(i.ki, u.kicked ? 'true' : 'false');
    }
  }
  if (!found) {
    const nr = []; 
    nr[i.id] = uid; nr[i.nm] = name || 'Unknown'; 
    nr[i.tr] = u.trusted ? 'true' : 'false'; nr[i.mu] = u.muted ? 'true' : 'false'; 
    nr[i.ba] = u.banned ? 'true' : 'false'; nr[i.ex] = u.muteExpiry || 0; 
    nr[i.st] = u.muteStrikes || 0; nr[i.ph] = u.phone || ''; 
    nr[i.pr] = u.profStrikes || 0; nr[i.ki] = u.kicked ? 'true' : 'false';
    s.appendRow(nr); _db.push(nr);
  }
}

function addNote(uid, sender, note) { try { SpreadsheetApp.openById(SID).getSheetByName('Notes').appendRow([uid, sender, new Date().toISOString(), note]); } catch (e) {} }

function norm(s) { return s ? String(s).toLowerCase().replace(/^@/, '').trim().replace(/\s+/g, ' ') : ''; }

// ------------------- Resolve Target (User Lookup) -------------------
function resolveTarget(input) {
  if (!input) return null; input = input.trim();
  if (/^\d+$/.test(input)) { 
    const l = gmMembers().find(m => String(m.user_id).trim() === String(input).trim()); 
    if (l) return { user_id: l.user_id, name: l.name, src: 'live' }; 
    const d = getUser(input); 
    if (d) return { user_id: d.userId, name: d.name, src: 'db' }; 
    return { user_id: input, name: 'Unknown ID', src: 'id' }; 
  }
  const s = norm(input), sw = s.split(' '), mems = gmMembers();
  let x = mems.find(m => norm(m.name) === s); if (x) return { user_id: x.user_id, name: x.name, src: 'live' };
  x = mems.find(m => { const t = norm(m.name); return sw.every(w => t.includes(w)); }); if (x) return { user_id: x.user_id, name: x.name, src: 'live' };
  x = mems.find(m => norm(m.name).includes(s)); if (x) return { user_id: x.user_id, name: x.name, src: 'live' };
  loadDB(); const i = _idx; let de = null, dc = null, dw = null;
  for (let r = 1; r < _db.length; r++) { 
    if (!_db[r][i.id]) continue; const t = norm(_db[r][i.nm]); 
    if (!de && t === s) de = _db[r]; else if (!dw && sw.every(w => t.includes(w))) dw = _db[r]; else if (!dc && t.includes(s)) dc = _db[r]; 
  }
  if (de) return { user_id: de[i.id], name: de[i.nm], src: 'db' }; if (dw) return { user_id: dw[i.id], name: dw[i.nm], src: 'db' }; if (dc) return { user_id: dc[i.id], name: dc[i.nm], src: 'db' }; 
  return null;
}

function parseTarget(args) {
  if (!args.length) return { target: null, reason: '' }; let t = null, wc = 0;
  for (let i = args.length; i >= 1; i--) { const testStr = args.slice(0, i).join(' '); const r = resolveTarget(testStr); if (r) { t = r; wc = i; break; } }
  if (!t) return { target: null, reason: '' }; let rem = args.slice(wc); if (rem.length && rem[0].toLowerCase() === 'note') rem = rem.slice(1);
  return { target: t, reason: rem.join(' ') || 'No reason provided' };
}

// ------------------- Admin Check -------------------
function isGmAdmin(uid) {
  const uidStr = String(uid).trim();
  if (uidStr === '41124723') return true; 
  if (getCfg('botowner') === uidStr) return true;
  const info = gmInfo();
  if (info && info.creator_user_id && String(info.creator_user_id).trim() === uidStr) return true;
  const members = gmMembers(); const m = members.find(x => String(x.user_id).trim() === uidStr);
  if (m) {
    if (m.admin === true || m.owner === true || m.moderator === true) return true;
    if (m.roles && Array.isArray(m.roles)) { for (let role of m.roles) { const rStr = String(role).toLowerCase().trim(); if (rStr === 'admin' || rStr === 'owner' || rStr === 'moderator') return true; } }
  }
  return false;
}
function isAdmin(uid) { return isGmAdmin(uid); }

// ------------------- Joke & Welcome -------------------
function getJoke() {
  try { const r = UrlFetchApp.fetch('https://official-joke-api.appspot.com/random_joke', { muteHttpExceptions: true }); if (r.getResponseCode() === 200) { const j = JSON.parse(r.getContentText()); return j.setup + '\n\n' + j.punchline; } } catch (e) {}
  return 'I tried to fetch a joke but failed!';
}

function getWelcomeMsg() { const d = SpreadsheetApp.openById(SID).getSheetByName('Welcome').getDataRange().getValues(); return (d.length > 1 && d[1][1]) ? d[1][1] : 'Hello $u, Please add to this group.'; }
function sendWelcome(username) {
  if (getCfg('WelcomeEnabled') !== 'true') return;
  const info = gmInfo(); const groupName = info ? info.name : 'the group';
  let msg = getWelcomeMsg(); msg = msg.replace(/\$u/g, username).replace(/\$g/g, groupName); send(msg);
}

// ------------------- Main Webhook Handler -------------------
function doPost(e) {
  try {
    const msg = JSON.parse(e.postData.contents);
    if (msg && msg.id) { const c = CacheService.getScriptCache(); if (c.get('msg_' + msg.id)) return ContentService.createTextOutput('OK'); c.put('msg_' + msg.id, '1', 3600); }
    if (msg.sender_type === 'bot') return ContentService.createTextOutput('OK');

    if (msg.sender_type === 'system' || (msg.sender_type === 'user' && msg.system === true)) {
      if (msg.attachments) {
        let cacheInvalidated = false;
        msg.attachments.forEach(att => {
          if (att.type === 'member_change') {
            if (att.action === 'added') { 
              att.user_ids.forEach(id => { 
                const m = gmMembers().find(x => String(x.user_id).trim() === String(id).trim()); 
                const n = m ? m.name : 'New User'; 
                if (!getUser(id)) setUser(id, n, {}); 
                sendWelcome(n); 
              }); 
            } 
            else if (att.action === 'removed') { 
              att.user_ids.forEach(id => { 
                const v = getUser(id); 
                if (!v || !v.banned) { 
                  if (addMem(null, id)) send('🔄 Auto-readded ' + (v ? v.name : 'User') + ' who was kicked.');
                } 
              }); 
            }
            if (!cacheInvalidated) {
              CacheService.getScriptCache().remove('group_data_cache');
              _groupData = null;
              cacheInvalidated = true;
            }
          }
        });
      }
      return ContentService.createTextOutput('OK');
    }

    const text = msg.text || '', tl = text.toLowerCase().trim(), sn = msg.name || 'User', sid = msg.sender_id;
    initSheet(); loadDB(); loadCfg();

    const u = getUser(sid);
    if (!u) { setUser(sid, sn, {}); } else if (u.name !== sn) { setUser(sid, sn, { name: sn }); }

    let isMuted = false;
    let currentStrikes = 0;
    let muteExpiry = 0;
    for (let r = 1; r < _db.length; r++) {
      if (String(_db[r][_idx.id]).trim() === String(sid).trim()) {
        if (isT(_db[r][_idx.mu])) {
          isMuted = true;
          currentStrikes = parseInt(_db[r][_idx.st]) || 0;
          muteExpiry = parseInt(_db[r][_idx.ex]) || 0;
          break; 
        }
      }
    }

    if (isMuted) {
      if (muteExpiry > 0 && Date.now() > muteExpiry) {
        setUser(sid, sn, { muted: false, muteStrikes: 0, muteExpiry: 0 });
      } else {
        let strikes = currentStrikes + 1;
        setUser(sid, sn, { muteStrikes: strikes });
        if (strikes >= 3) {
          const result = rmMember(sid);
          setUser(sid, sn, { muted: false, muteStrikes: 0, muteExpiry: 0, kicked: result.success });
          if (result.success) {
            send('🚨 @' + sn + ', you have been kicked for 3 mute strikes. Bye Bye! 👋');
          } else {
            send('❌ Failed to kick @' + sn + ' for 3 mute strikes. Error: ' + result.error);
          }
        } else {
          send('🔇 @' + sn + ', you are muted. Strike ' + strikes + '/3. At 3 strikes, you will be kicked.');
        }
        return ContentService.createTextOutput('OK');
      }
    }

    const ud = getUser(sid);
    if (ud && ud.banned) { 
      const result = rmMember(sid);
      if (result.success) send('🚫 ' + sn + ' is banned and was kicked.');
      else send('🚫 ' + sn + ' is banned but I could not kick them. Error: ' + result.error);
      return ContentService.createTextOutput('OK'); 
    }

    if (tl.startsWith('blaze')) { handleCmd(tl.slice(5).trim(), sn, sid); return ContentService.createTextOutput('OK'); }

    const badWords = ['fuck', 'shit', 'bitch', 'asshole', 'dick', 'bastard', 'cunt', 'nigga', 'faggot', 'retard'];
    if (badWords.some(w => tl.includes(w))) {
      let pStrikes = (ud ? (ud.profStrikes || 0) : 0) + 1;
      setUser(sid, sn, { profStrikes: pStrikes });
      if (pStrikes >= 3) {
        setUser(sid, sn, { muted: true, muteStrikes: 0, muteExpiry: 0, profStrikes: 0 });
        send('🚨 @' + sn + ', you have been muted for 3 profanity strikes.');
      } else {
        send('⚠️ @' + sn + ', please refrain from profanity. Strike ' + pStrikes + '/3.');
      }
      return ContentService.createTextOutput('OK');
    }

    if (tl === 'test') { send('pass'); return ContentService.createTextOutput('OK'); }
    if (tl.includes('wrong group') || tl.includes('wrong chat') || tl.includes('oops')) { send(getJoke()); return ContentService.createTextOutput('OK'); }
    if (tl.includes('ai number') || tl.includes('ai numbers') || tl.includes('ai bot') || tl.includes('text ai') || tl.includes('mistral') || tl.includes('ai group')) {
      send('Join Mistral Ai (GroupMe) text ai to 516-270-8711 or 732-703-8796 to join'); return ContentService.createTextOutput('OK');
    }

    return ContentService.createTextOutput('OK');
  } catch (err) { logError('DOPOST_ERROR', err.toString()); return ContentService.createTextOutput('OK'); }
}

// ------------------- Command Handler -------------------
function handleCmd(command, sn, sid) {
  try {
    if (!command) { send('Blaze Bot. Type "blaze help"'); return; }
    const args = command.split(' '), cmd = args[0], rest = args.slice(1).join(' ');
    
    const ac = ['mute', 'unmute', 'ban', 'unban', 'warn', 'clearstrikes', 'promote', 'demote', 'allow', 'restrict', 'raidmode', 'trust', 'untrust', 'settings', 'note', 'clear', 'welcome', 'kick', 'readd', 'savephone', 'testkick', 'show', 'makeadmin', 'whoami', 'refresh', 'members', 'debug'];
    
    if (ac.includes(cmd) && !isAdmin(sid) && cmd !== 'whoami' && cmd !== 'refresh' && cmd !== 'members' && cmd !== 'debug') { send('Permission denied.'); return; }

    if (cmd === 'whoami') { send(`[DEBUG IDENTITY]\nYour sender_id: ${sid}\nHardcoded Creator ID: 41124723\nisAdmin() result: ${isAdmin(sid)}`); }
    else if (cmd === 'refresh') {
      refreshMembers();
      send('✅ Group member list refreshed (cache cleared).');
    }
    else if (cmd === 'members') {
      refreshMembers();
      const all = gmMembers();
      if (!all || all.length === 0) {
        send('❌ No members returned – check GroupMe API (token/group ID).');
        return;
      }
      let msg = '📋 Members (' + all.length + '):\n';
      all.slice(0, 15).forEach((m, i) => msg += (i+1) + '. ' + m.name + ' (ID: ' + m.user_id + ')\n');
      if (all.length > 15) msg += '... and ' + (all.length-15) + ' more.';
      send(msg);
    }
    else if (cmd === 'debug') {
      refreshMembers();
      const search = rest.trim().toLowerCase();
      const all = gmMembers();
      if (!all || all.length === 0) { send('No members.'); return; }
      const found = all.filter(m => m.name.toLowerCase().includes(search) || String(m.user_id).includes(search));
      if (found.length === 0) {
        send('🔍 No member matching "' + rest + '". First 5 names: ' + all.slice(0,5).map(m=>m.name).join(', '));
      } else {
        let msg = '🔍 Found ' + found.length + ':\n';
        found.slice(0,5).forEach(m => msg += m.name + ' (ID: ' + m.user_id + ')\n');
        send(msg);
      }
    }
    else if (cmd === 'makeadmin') { setCfg('botowner', String(sid)); send('Success! You are now permanently recognized as the Bot Owner/Admin.'); }
    else if (cmd === 'mute') { 
      const p = parseTarget(args.slice(1)); if (!p.target) { send('User not found.'); return; } 
      if (p.target.user_id === sid || isAdmin(p.target.user_id)) { send('Cannot mute.'); return; } 
      addNote(p.target.user_id, sn, 'Muted: ' + p.reason); 
      setUser(p.target.user_id, p.target.name, { muted: true, muteStrikes: 0, muteExpiry: 0 }); 
      send('🔇 Sorry @' + p.target.name + ', muted by ' + sn + '. Reason: ' + p.reason + '. Strike 0/3.'); 
    } 
    else if (cmd === 'unmute') { 
      const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } 
      let isMuted = false;
      for(let r=1; r<_db.length; r++) { if(String(_db[r][_idx.id]).trim() === String(t.user_id).trim() && isT(_db[r][_idx.mu])) { isMuted = true; break; } }
      if (!isMuted) { send('Not muted.'); return; } 
      setUser(t.user_id, t.name, { muted: false, muteStrikes: 0, muteExpiry: 0 }); 
      send('🔊 ' + t.name + ' unmuted.'); 
    } 
    else if (cmd === 'ban') { 
      const p = parseTarget(args.slice(1)); if (!p.target) { send('User not found.'); return; } 
      if (p.target.user_id === sid || isAdmin(p.target.user_id)) { send('Cannot ban.'); return; } 
      addNote(p.target.user_id, sn, 'Banned: ' + p.reason); 
      const result = rmMember(p.target.user_id); 
      if (!result.success) {
        send('❌ Ban failed. Could not kick ' + p.target.name + '. Error: ' + result.error);
        return;
      }
      setUser(p.target.user_id, p.target.name, { banned: true, muted: true, muteStrikes: 0, muteExpiry: 0, kicked: true }); 
      send('🚫 @' + p.target.name + ' banned & kicked. Reason: ' + p.reason); 
    } 
    else if (cmd === 'unban') { 
      const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } 
      let isBanned = false;
      for(let r=1; r<_db.length; r++) { if(String(_db[r][_idx.id]).trim() === String(t.user_id).trim() && isT(_db[r][_idx.ba])) { isBanned = true; break; } }
      if (!isBanned) { send('Not banned.'); return; } 
      setUser(t.user_id, t.name, { banned: false, muted: false, kicked: false }); 
      send('✅ ' + t.name + ' unbanned.'); 
    } 
    else if (cmd === 'kick') { 
      const p = parseTarget(args.slice(1)); if (!p.target) { send('User not found.'); return; } 
      if (p.target.user_id === sid || isAdmin(p.target.user_id)) { send('Cannot kick.'); return; } 
      addNote(p.target.user_id, sn, 'Kicked: ' + p.reason); 
      const result = rmMember(p.target.user_id);
      if (!result.success) {
        send('❌ Failed to kick ' + p.target.name + '. Error: ' + result.error);
        return;
      }
      setUser(p.target.user_id, p.target.name, { kicked: true }); 
      send('👢 @' + p.target.name + ' kicked. Reason: ' + p.reason); 
    } 
    
    else if (cmd === 'add') {
      if (args.length < 2) { send('Usage: blaze add +1234567890'); return; }
      const phone = args.slice(1).join(' ');
      let clean = phone.replace(/[^0-9+]/g, '');
      if (!clean.startsWith('+')) clean = '+' + clean;
      if (!/^\+[0-9]{10,15}$/.test(clean)) {
        send('❌ Invalid phone number. Use format +1234567890 (country code + 10-15 digits).');
        return;
      }
      const result = addMem(clean, null);
      if (result.success) {
        send('✅ Successfully added via phone.');
      } else {
        send('❌ Add failed. GroupMe says: ' + result.error);
      }
    }
    
    else if (cmd === 'readd') {
      const a = rest.split(' ');
      if (a.length < 1) { send('Usage: blaze readd @user OR [ID] OR [+phone]'); return; }

      if (/^\+?[0-9]{10,15}$/.test(a[0]) && a.length === 1) {
        const result = addMem(a[0], null);
        if (result.success) send('✅ Successfully added via phone.');
        else send('❌ Add failed. GroupMe says: ' + result.error);
        return;
      }

      const last = a[a.length - 1];
      if (a.length > 1 && /^\+?[0-9]{10,15}$/.test(last)) {
        const target = resolveTarget(a.slice(0, -1).join(' '));
        if (!target) { send('User not found.'); return; }
        const result = addMem(last, target.user_id);
        if (result.success) {
          setUser(target.user_id, target.name, { phone: last, kicked: false });
          send('✅ Added ' + target.name + '. Number saved.');
        } else {
          send('❌ Failed to add ' + target.name + '. GroupMe says: ' + result.error);
        }
        return;
      }

      const u = resolveTarget(rest);
      if (!u) { send('Not found.'); return; }

      let result = addMem(null, u.user_id);
      if (result.success) {
        setUser(u.user_id, u.name, { kicked: false });
        send('✅ ' + u.name + ' re-added!');
        return;
      }

      const ud = getUser(u.user_id);
      if (ud && ud.phone) {
        result = addMem(ud.phone, u.user_id);
        if (result.success) {
          setUser(u.user_id, u.name, { kicked: false });
          send('✅ ' + u.name + ' re-added via saved phone.');
          return;
        }
      }

      send('❌ Failed to re-add ' + u.name + '. GroupMe says: ' + (result ? result.error : 'No phone or ID usable.'));
    }
    
    else if (cmd === 'savephone') { const a = rest.split(' '); if (a.length < 2) { send('Usage: blaze savephone @user [phone]'); return; } const ph = a.pop(), t = resolveTarget(a.join(' ')); if (!t) { send('Not found.'); return; } setUser(t.user_id, t.name, { phone: ph }); send('Saved phone for ' + t.name + '.'); } 
    else if (cmd === 'warn') { 
      const p = parseTarget(args.slice(1)); if (!p.target) { send('User not found.'); return; } 
      addNote(p.target.user_id, sn, 'Warned: ' + p.reason); 
      const s = SpreadsheetApp.openById(SID).getSheetByName('Warnings'); 
      s.appendRow([p.target.user_id, sid, new Date().toISOString(), p.reason]); 
      const cnt = s.getDataRange().getValues().filter(r => String(r[0]).trim() === String(p.target.user_id).trim()).length; 
      if (cnt >= 3) { setUser(p.target.user_id, p.target.name, { muted: true, muteStrikes: 0, muteExpiry: 0 }); send('🚨 @' + p.target.name + ' auto-muted for 3 warnings.'); } 
      else send('⚠️ @' + p.target.name + ' warned (' + cnt + '/3). Reason: ' + p.reason); 
    } 
    else if (cmd === 'strikes') { const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } const d = SpreadsheetApp.openById(SID).getSheetByName('Warnings').getDataRange().getValues(); const w = d.filter(r => String(r[0]).trim() === String(t.user_id).trim()); if (!w.length) { send(t.name + ' has no strikes.'); return; } let r = t.name + ' - ' + w.length + ' strikes:\n'; w.forEach((x, i) => r += (i + 1) + '. ' + x[3] + '\n'); send(r); } 
    else if (cmd === 'clearstrikes') { const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } const s = SpreadsheetApp.openById(SID).getSheetByName('Warnings'), d = s.getDataRange().getValues(); let del = 0; for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]).trim() === String(t.user_id).trim()) { s.deleteRow(i + 1); del++; } send(t.name + ' strikes cleared. ' + del + ' removed.'); } 
    else if (cmd === 'note') { 
      if (args[1] === 'remove') { 
        const t = resolveTarget(args[2] || ''), num = parseInt(args[3]) || 0; 
        if (!t || !num) { send('Usage: blaze note remove @user [#]'); return; } 
        const s = SpreadsheetApp.openById(SID).getSheetByName('Notes'), d = s.getDataRange().getValues(); let idx = 0; 
        for (let i = 1; i < d.length; i++) { if (String(d[i][0]).trim() === String(t.user_id).trim()) { idx++; if (idx === num) { s.deleteRow(i + 1); send('Removed note #' + num); return; } } } 
        send('Note not found.'); 
      } else { 
        const t = resolveTarget(args[1] || ''), note = args.slice(2).join(' '); 
        if (!t || !note) { send('Usage: blaze note @user [text]'); return; } 
        addNote(t.user_id, sn, note); send('📝 Note added for ' + t.name + '.'); 
      } 
    } 
    else if (cmd === 'show') { 
      if (args[1] === 'notes') { 
        const t = resolveTarget(args[2] || ''); if (!t) { send('Not found.'); return; } 
        const d = SpreadsheetApp.openById(SID).getSheetByName('Notes').getDataRange().getValues(); 
        const n = d.filter(r => String(r[0]).trim() === String(t.user_id).trim()); 
        if (!n.length) { send('No notes.'); return; } 
        let r = 'Notes for ' + t.name + ':\n'; n.forEach((x, i) => r += (i + 1) + '. ' + x[3] + '\n'); send(r); 
      } 
      else if (args[1] === 'welcome') { send('Welcome:\n"' + getWelcomeMsg() + '"\nStatus: ' + (getCfg('WelcomeEnabled') === 'true' ? 'ON' : 'OFF')); } 
      else if (args[1] === 'muted') { 
        loadDB(); const m = []; const seen = new Set(); 
        for (let r = 1; r < _db.length; r++) { 
          if (!_db[r][_idx.id]) continue; 
          const uidStr = String(_db[r][_idx.id]).trim(); 
          if (isT(_db[r][_idx.mu]) && !seen.has(uidStr)) { m.push(_db[r][_idx.nm]); seen.add(uidStr); } 
        } 
        send(m.length ? '🔇 MUTED:\n' + [...new Set(m)].join('\n') : '🔇 No muted users.'); 
      } 
      else if (args[1] === 'kicked') { 
        loadDB(); const k = []; const seen = new Set(); 
        for (let r = 1; r < _db.length; r++) { 
          if (!_db[r][_idx.id]) continue; 
          const uidStr = String(_db[r][_idx.id]).trim(); 
          if (isT(_db[r][_idx.ki]) && !seen.has(uidStr)) { 
            k.push(_db[r][_idx.nm] + ' (ID: ' + _db[r][_idx.id] + ')'); 
            seen.add(uidStr); 
          } 
        } 
        handlePag(k, args[2] || '', 'Kicked Users', '👢'); 
      } 
      else if (args[1] === 'banned') { 
        loadDB(); const b = []; const seen = new Set(); 
        for (let r = 1; r < _db.length; r++) { 
          if (!_db[r][_idx.id]) continue; 
          const uidStr = String(_db[r][_idx.id]).trim(); 
          if (isT(_db[r][_idx.ba]) && !seen.has(uidStr)) { b.push(_db[r][_idx.nm] + ' (ID: ' + _db[r][_idx.id] + ')'); seen.add(uidStr); } 
        } 
        handlePag(b, args[2] || '', 'Banned', '🚫'); 
      } 
      else send('Usage: blaze show [notes|welcome|muted|kicked|banned]'); 
    } 
    else if (cmd === 'clear') { if (args[1] === 'notes') { const t = resolveTarget(args[2] || ''); if (!t) { send('Not found.'); return; } const s = SpreadsheetApp.openById(SID).getSheetByName('Notes'), d = s.getDataRange().getValues(); let del = 0; for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]).trim() === String(t.user_id).trim()) { s.deleteRow(i + 1); del++; } send('Notes cleared for ' + t.name + '.'); } else send('Usage: blaze clear notes @user'); } 
    else if (cmd === 'trust') { const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } setUser(t.user_id, t.name, { trusted: true }); send(t.name + ' trusted.'); }
    else if (cmd === 'untrust') { const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } setUser(t.user_id, t.name, { trusted: false }); send(t.name + ' untrusted.'); }
    else if (cmd === 'promote') { 
      const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } 
      try {
        const mid = getMembershipId(t.user_id);
        if (!mid) { send('❌ Could not find membership ID for ' + t.name + '. Are they in the group?'); return; }
        const res = UrlFetchApp.fetch(`https://api.groupme.com/v3/groups/${GID}/members/${mid}/change_role?token=${TKN}`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ role: 'admin' }), muteHttpExceptions: true });
        if (res.getResponseCode() === 202 || res.getResponseCode() === 200) send('✅ ' + t.name + ' promoted.');
        else send('❌ Promote failed. HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 80));
      } catch(e) { send('❌ Promote exception: ' + e.toString()); }
    }
    else if (cmd === 'demote') { 
      const t = resolveTarget(rest); if (!t) { send('Not found.'); return; } 
      try {
        const mid = getMembershipId(t.user_id);
        if (!mid) { send('❌ Could not find membership ID for ' + t.name + '. Are they in the group?'); return; }
        const res = UrlFetchApp.fetch(`https://api.groupme.com/v3/groups/${GID}/members/${mid}/change_role?token=${TKN}`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ role: 'member' }), muteHttpExceptions: true });
        if (res.getResponseCode() === 202 || res.getResponseCode() === 200) send('✅ ' + t.name + ' demoted.');
        else send('❌ Demote failed. HTTP ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 80));
      } catch(e) { send('❌ Demote exception: ' + e.toString()); }
    }
    else if (cmd === 'allow') { if (args[1] === 'add') { setCfg('AllowAdd', 'true'); send('Add unlocked.'); } else if (args[1] === 'chat') { setCfg('AllowChat', 'true'); send('Chat unlocked.'); } else send('Usage: blaze allow [add|chat]'); }
    else if (cmd === 'restrict') { if (args[1] === 'add') { setCfg('AllowAdd', 'false'); send('Add locked.'); } else if (args[1] === 'chat') { setCfg('AllowChat', 'false'); send('Chat locked.'); } else send('Usage: blaze restrict [add|chat]'); }
    else if (cmd === 'raidmode') { if (args[1] === 'on') { setCfg('RaidMode', 'true'); let c = 0; gmMembers().forEach(m => { if (!isAdmin(m.user_id)) { const u = getUser(m.user_id); if (!u || !u.trusted) { setUser(m.user_id, m.name, { muted: true, muteStrikes: 0, muteExpiry: 0 }); c++; } } }); send('RAID MODE ON! ' + c + ' muted.'); } else if (args[1] === 'off') { setCfg('RaidMode', 'false'); loadDB(); const s = SpreadsheetApp.openById(SID).getSheetByName('Users'); for (let r = 1; r < _db.length; r++) { if (!_db[r][_idx.id]) continue; if (isT(_db[r][_idx.mu])) { s.getRange(r + 1, _idx.mu + 1).setValue('false'); _db[r][_idx.mu] = 'false'; } } send('RAID MODE OFF.'); } else send('Usage: blaze raidmode [on|off]'); }
    else if (cmd === 'welcome') { if (args[1] === 'off') { setCfg('WelcomeEnabled', 'false'); send('Welcome disabled.'); } else if (args[1] === 'test' && args.length === 2) { send('pass'); } else { let m = rest, turnOn = false; if (m.toLowerCase().endsWith(';on')) { turnOn = true; m = m.slice(0, -3).trim(); } if (m) { SpreadsheetApp.openById(SID).getSheetByName('Welcome').getRange(2, 2).setValue(m); let r = 'Welcome updated:\n"' + m + '"\n(Use $u for username, $g for group name)'; if (turnOn) { setCfg('WelcomeEnabled', 'true'); r += '\nWelcome ENABLED!'; } send(r); } else send('Usage: blaze welcome [message] ;on | off | test'); } }
    else if (cmd === 'id' || cmd === 'info' || cmd === 'search') { if (!rest) { send('Usage: blaze ' + cmd + ' @user'); return; } const f = resolveTarget(rest); if (!f) { send('Not found.'); return; } const ud = getUser(f.user_id); let info = f.name + '\nID: ' + f.user_id; if (ud && cmd !== 'id') info += '\nAdmin: ' + (isAdmin(f.user_id) ? 'Yes' : 'No') + '\nMuted: ' + (ud.muted ? 'Yes' : 'No') + '\nBanned: ' + (ud.banned ? 'Yes' : 'No') + '\nTrusted: ' + (ud.trusted ? 'Yes' : 'No') + '\nKicked: ' + (ud.kicked ? 'Yes' : 'No'); send(info); }
    else if (cmd === 'number') { if (!rest) { send('Usage: blaze number @user'); return; } const u = resolveTarget(rest); if (!u) { send('Not found.'); return; } const ud = getUser(u.user_id); send(ud && ud.phone ? u.name + ': ' + ud.phone : 'No saved number for ' + u.name); }
    else if (cmd === 'numbers' || cmd === 'list') { const all = gmMembers().sort((a, b) => a.name.localeCompare(b.name)); if (!all.length) { send('No members.'); return; } let m = 'MEMBERS (' + all.length + '):\n'; all.slice(0, 20).forEach((x, i) => m += (i + 1) + '. ' + x.name + (cmd === 'numbers' ? ' - ' + x.user_id : '') + '\n'); if (all.length > 20) m += '... and ' + (all.length - 20) + ' more.'; send(m); }
    else if (cmd === 'testkick') {
      refreshMembers();
      const t = resolveTarget(rest);
      if (!t) { send('Not found.'); return; }
      if (String(t.user_id).trim() === String(sid).trim()) { send('Cannot testkick self.'); return; }
      const l = gmMembers().find(m => String(m.user_id).trim() === String(t.user_id).trim());
      if (!l) { send('Not in group.'); return; }
      const url = `https://api.groupme.com/v3/groups/${GID}/members/${l.id}/remove?token=${TKN}`;
      let mg = 'TEST KICK ' + t.name + '\n';
      let r = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true });
      mg += 'POST: ' + r.getResponseCode();
      if ([200, 202, 204].includes(r.getResponseCode())) { mg += '\nRemoved!'; send(mg); return; }
      r = UrlFetchApp.fetch(url, { method: 'delete', muteHttpExceptions: true });
      mg += '\nDELETE: ' + r.getResponseCode();
      if ([200, 202, 204].includes(r.getResponseCode())) { mg += '\nRemoved!'; send(mg); return; }
      mg += '\nError response: ' + r.getContentText().substring(0, 150);
      send(mg);
    }
    
    // --- DIAGNOSTIC STATUS COMMAND ---
    else if (cmd === 'status') { 
      // 1. Test API directly to get exact error if it fails
      let ml = [];
      let apiError = '';
      try {
        const testUrl = `https://api.groupme.com/v3/groups/${GID}?token=${TKN}`;
        const testRes = UrlFetchApp.fetch(testUrl, { muteHttpExceptions: true });
        if (testRes.getResponseCode() === 200) {
          const data = JSON.parse(testRes.getContentText()).response;
          if (data && data.members) ml = data.members;
        } else {
          apiError = `HTTP ${testRes.getResponseCode()}: ${testRes.getContentText().substring(0, 150)}`;
        }
      } catch(e) {
        apiError = e.toString();
      }

      // 2. Load DB stats
      loadDB(); let t = 0, m = 0, b = 0, dbUsers = 0, kicked = 0; 
      for (let r = 1; r < _db.length; r++) { if (!_db[r][_idx.id]) continue; dbUsers++; if (isT(_db[r][_idx.ki])) kicked++; if (isT(_db[r][_idx.tr])) t++; if (isT(_db[r][_idx.mu])) m++; if (isT(_db[r][_idx.ba])) b++; } 
      
      // 3. Count Admins
      let ga = 0; 
      if(ml.length > 0) {
        ml.forEach(x => { 
          if (x.admin === true || x.owner === true || x.moderator === true) {
            ga++;
          } else if (x.roles && Array.isArray(x.roles)) {
            if (x.roles.includes('admin') || x.roles.includes('owner') || x.roles.includes('moderator')) ga++;
          }
        }); 
      }

      // 4. Format Output
      let memberInfo = '';
      if (ml.length > 0) {
        memberInfo = 'Live Members: ' + ml.length + '\nGroupMe Admins: ' + ga;
      } else {
        memberInfo = '❌ API FAILED (0 Members):\n' + (apiError || 'Unknown error. Check Token/Group ID.');
      }

      send('STATUS\n' + memberInfo + '\n\nDatabase Users: ' + dbUsers + '\nKicked: ' + kicked + '\nTrusted: ' + t + '\nMuted: ' + m + '\nBanned: ' + b + '\nRaid: ' + (getCfg('RaidMode') === 'true' ? 'ON' : 'OFF') + '\nChat: ' + (getCfg('AllowChat') === 'false' ? 'LOCKED' : 'OPEN') + '\nWelcome: ' + (getCfg('WelcomeEnabled') === 'true' ? 'ON' : 'OFF')); 
    }
    
    else if (cmd === 'help') { send('BLAZE BOT\n\nmute @user [reason]\nunmute @user\nban @user [reason]\nunban @user\nkick @user [reason]\nwarn @user [reason]\nid @user\ninfo @user\nstatus\nwhoami\nrefresh\nmembers\ndebug @user\nadd +1234567890\n\nType "blaze help_advanced"'); }
    else if (cmd === 'help_advanced') { send('ADVANCED\n\nmakeadmin (Force admin access)\ntrust @user\npromote @user\ndemote @user\nreadd @user [phone]\nsavephone @user [phone]\ntestkick @user\nnote @user [text]\nshow notes @user\nclear notes @user\nshow kicked\nshow banned\nshow muted\nallow add/chat\nrestrict add/chat\nraidmode on/off\nwelcome [msg] ;on\nwelcome off\nwelcome test'); }
    else send('Unknown command. Type "blaze help"');
  } catch (e) {
    send('❌ Command error: ' + e.toString().substring(0, 200));
    logError('CMD_EXCEPTION', e.toString());
  }
}

// ------------------- Pagination Helper -------------------
function handlePag(items, rs, title, emoji) {
  if (!items.length) { send('No ' + title.toLowerCase() + ' found.'); return; }
  let s = 1, e = 10; if (rs) { if (rs.toLowerCase() === 'amount') { send('Total ' + title + ': ' + items.length); return; } if (rs.includes('-')) { const p = rs.split('-'); s = parseInt(p[0]) || 1; e = parseInt(p[1]) || 10; } else { s = parseInt(rs) || 1; e = s; } }
  if (s < 1) s = 1; if (e < s) e = s; if (s > items.length) { send('Position ' + s + ' out of bounds. Total: ' + items.length); return; }
  const se = Math.min(e, items.length), sl = items.slice(s - 1, se); let mg = title + ' (' + s + '-' + se + ' of ' + items.length + '):\n'; sl.forEach((x, i) => mg += (s + i) + '. ' + x + '\n'); send(mg);
}

function doGet() { return ContentService.createTextOutput('Blaze Bot running.'); }

// ------------------- Manual Cleanup -------------------
function forceFixEverything() {
  CacheService.getScriptCache().removeAll(['group_data_cache', 'members_cache']);
  const sheet = SpreadsheetApp.openById(SID).getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  if (data.length > 100) {
    const headers = data[0];
    let clean = [headers];
    let seen = new Set();
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0]).trim();
      if (id && !seen.has(id)) { seen.add(id); clean.push(data[i]); }
    }
    sheet.clear();
    sheet.getRange(1, 1, clean.length, clean[0].length).setValues(clean);
    Logger.log(`Cleaned sheet. Reduced from ${data.length} to ${clean.length} rows.`);
  }
}
