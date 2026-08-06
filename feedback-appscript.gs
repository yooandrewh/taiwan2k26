/* Taiwan 2026 — Feedback backend (Google Apps Script)
 *
 * WHERE THIS LIVES
 *   Open the Feedback Google Sheet → Extensions → Apps Script.
 *   (It has to be created from inside the Sheet — a standalone project can't
 *   see it, SpreadsheetApp.getActiveSpreadsheet() would come back empty.)
 *   Paste this file over everything that's there, Save, then
 *   Deploy → Manage deployments → edit the existing one → Version: New version
 *   → Deploy.  Keep the SAME deployment so the /exec URL doesn't change.
 *   Execute as: Me · Who has access: Anyone.
 *
 * WHAT IT DOES
 *   POST  {person,pin,edited,g,ev,vbs,tmf,hs,n_g..n_hs,gcl,ds,lp,eric,last}
 *         → upserts that person's row.  The first PIN seen for a name claims
 *           the row; a different one afterwards is rejected.
 *   GET   ?person=NAME&pin=HASH  → that one row, only if the hash matches.
 *   GET   ?status=1              → progress for everyone: how many characters
 *           are in each section, whether each was marked "nothing to add",
 *           which sign-ups are ticked, when it was last edited.  No answer
 *           text ever leaves the sheet through this call — that's what lets
 *           the app show a team status board without breaking the privacy
 *           promise.
 */

var SHEET = 'Feedback';
var HEAD = ['Timestamp','Person','PinHash','Edited','General','EV Week','VBS Week',
            'TMF Week','Hot Springs','GCL Meeting','D&S Letter','LP Letter',
            'Eric Letter','Steve&Sehee Letter',
            'Q General','Q EV','Q VBS','Q TMF','Q Hot Springs'];
var SEC = {g:4, ev:5, vbs:6, tmf:7, hs:8};              // 0-based column index
var SU  = {gcl:9, ds:10, lp:11, eric:12, last:13};
// "Do you have notes for this one?" — 'y', 'n', or blank for not-asked-yet.
// This is what lets the status board tell "done, nothing to add" apart from
// "hasn't touched it".
var NOT = {g:14, ev:15, vbs:16, tmf:17, hs:18};

function getSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if(!sh){ sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); return sh; }
  // Self-healing header: pasting a newer version of this file adds any new
  // columns on its own, so nobody has to edit the sheet by hand.
  var hdr = sh.getRange(1, 1, 1, HEAD.length).getValues()[0];
  for(var c = 0; c < HEAD.length; c++){
    if(String(hdr[c]) !== HEAD[c]){ sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]); break; }
  }
  return sh;
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function isTrue(v){
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function rowToObj(r){
  var o = {person: String(r[1]), edited: Number(r[3]) || 0};
  for(var k in SEC) o[k] = r[SEC[k]] === null || r[SEC[k]] === undefined ? '' : String(r[SEC[k]]);
  for(var s in SU)  o[s] = isTrue(r[SU[s]]);
  for(var nk in NOT) o['n_' + nk] = String(r[NOT[nk]] || '').trim().toLowerCase();
  return o;
}

function doGet(e){
  // e is undefined when you press ▶ Run in the editor — don't blow up on it.
  e = e || {};
  var q = e.parameter || {};
  var data = getSheet().getDataRange().getValues();

  // ---- Team status: counts only, never any of the writing itself. ----
  if(q.status){
    var out = [];
    for(var i = 1; i < data.length; i++){
      var r = data[i];
      if(!r[1]) continue;
      var chars = {}, notes = {};
      for(var k in SEC) chars[k] = String(r[SEC[k]] == null ? '' : r[SEC[k]]).trim().length;
      for(var nk in NOT) notes[nk] = String(r[NOT[nk]] || '').trim().toLowerCase();
      var sign = [];
      for(var s in SU) if(isTrue(r[SU[s]])) sign.push(s);
      out.push({
        person:  String(r[1]),
        chars:   chars,
        notes:   notes,
        signups: sign,
        edited:  Number(r[3]) || 0,
        locked:  !!String(r[2] || '').trim()
      });
    }
    return json({status: out});
  }

  // ---- One person's own entry.  The PIN gate is enforced HERE, on the
  //      server, so a teammate can't read someone else's answers just by
  //      hitting the URL by hand. ----
  var who = String(q.person || '');
  var pin = String(q.pin || '');
  if(!who) return json({rows: []});
  for(var j = 1; j < data.length; j++){
    if(String(data[j][1]) !== who) continue;
    var stored = String(data[j][2] || '').trim();
    if(stored){ if(stored !== pin) return json({error: 'pin'}); }
    else if(!pin){ return json({error: 'pin'}); }   // unclaimed row still needs a PIN
    return json({rows: [rowToObj(data[j])]});
  }
  return json({rows: []});
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(err){ return json({error: 'busy'}); }
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var who = String(b.person || '').trim();
    if(!who) return json({error: 'person'});
    var pin = String(b.pin || '').trim();

    var sh = getSheet();
    var data = sh.getDataRange().getValues();
    var row = 0, stored = '';
    for(var i = 1; i < data.length; i++){
      if(String(data[i][1]) === who){ row = i + 1; stored = String(data[i][2] || '').trim(); break; }
    }
    if(row && stored && pin && stored !== pin) return json({error: 'pin'});
    if(!row){ sh.appendRow([new Date(), who, pin, 0, '','','','','', false,false,false,false,false, '','','','','']); row = sh.getLastRow(); }

    var out = new Array(HEAD.length);
    out[0] = new Date();
    out[1] = who;
    out[2] = stored || pin;                 // first PIN claims the row
    out[3] = Number(b.edited) || 0;
    for(var k in SEC) out[SEC[k]] = b[k] == null ? '' : String(b[k]);
    for(var s in SU)  out[SU[s]]  = !!b[s];
    for(var nk in NOT) out[NOT[nk]] = String(b['n_' + nk] || '').trim().toLowerCase();
    sh.getRange(row, 1, 1, HEAD.length).setValues([out]);
    return json({ok: true});
  } finally {
    lock.releaseLock();
  }
}
