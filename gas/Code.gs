// ═══════════════════════════════════════════════════════════════
//  INVESTOR TRIP BOOK 2026 — Google Apps Script Backend
//  Deploy: Extensions → Apps Script → Deploy → Web App
//         Execute as: Me  |  Access: Anyone (or your org)
// ═══════════════════════════════════════════════════════════════

var SHEET_ID_KEY = 'TRIPBOOK_SHEET_ID';

// ── Entry point ──────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Investor Trip Book · HK & Canton Fair 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ── Spreadsheet bootstrap ────────────────────────────────────
function getOrCreateSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty(SHEET_ID_KEY);

  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* deleted — make a new one */ }
  }

  var ss = SpreadsheetApp.create('Investor Trip Book · HK & Canton Fair 2026');
  props.setProperty(SHEET_ID_KEY, ss.getId());

  var sheetNames = ['Overview','HK_Events','CF_Events','Actions','Contacts','Logistics','Investor_Notes'];
  ss.getSheets()[0].setName('Overview');
  sheetNames.slice(1).forEach(function(n) { ss.insertSheet(n); });
  initHeaders_(ss);

  return ss;
}

function initHeaders_(ss) {
  var EV = ['ID','Date','Time','Duration','Color','Status','Title','Company',
            'Venue','Attendees','Notes','Tags','Investor Tier','Budget','Outcome','Commitment'];
  var widths = [55,105,65,65,115,100,280,200,240,220,360,200,140,120,340,340];

  ['HK_Events','CF_Events'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(1,1,1,EV.length).setValues([EV]);
    formatHeader_(sh, EV.length);
    widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
  });

  var actHdr = ['ID','Task','Owner','Due Date','Phase','Status','Priority'];
  writeHeader_(ss, 'Actions', actHdr, [55,320,120,110,130,90,90]);

  var conHdr = ['ID','Name','Role','Company','Phase','Tags','Email','Phone',
                'Investor Type','Deal Value','Meeting Objective','Notes'];
  writeHeader_(ss, 'Contacts', conHdr, [55,150,160,200,130,160,200,140,140,150,240,340]);

  var logHdr = ['ID','Type','Title','Details','Status'];
  writeHeader_(ss, 'Logistics', logHdr, [55,90,280,400,110]);

  var invHdr = ['Phase','Investor Tier','Title','Company','Budget','Outcome','Commitment','Date'];
  writeHeader_(ss, 'Investor_Notes', invHdr, [130,140,280,200,120,360,360,110]);
}

function writeHeader_(ss, name, hdr, widths) {
  var sh = ss.getSheetByName(name);
  if (!sh) return;
  sh.getRange(1,1,1,hdr.length).setValues([hdr]);
  formatHeader_(sh, hdr.length);
  if (widths) widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
}

function formatHeader_(sh, cols) {
  sh.getRange(1,1,1,cols)
    .setBackground('#0B2740')
    .setFontColor('#D4A843')
    .setFontWeight('bold')
    .setFontSize(10);
  sh.setFrozenRows(1);
}

function getOrMake_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── Public API (called from client via google.script.run) ────

/** Returns spreadsheet metadata so the client can show a link. */
function getSpreadsheetInfo() {
  try {
    var ss = getOrCreateSpreadsheet_();
    return { ok: true, id: ss.getId(), url: ss.getUrl(), name: ss.getName() };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

/** Writes all client data to the spreadsheet. */
function saveData(jsonStr) {
  try {
    var d  = JSON.parse(jsonStr);
    var ss = getOrCreateSpreadsheet_();

    writeEvents_(ss, 'HK_Events', d.HK_EVENTS || []);
    writeEvents_(ss, 'CF_Events', d.CF_EVENTS || []);
    writeActions_(ss,   d.ACTIONS   || []);
    writeContacts_(ss,  d.CONTACTS  || []);
    writeLogistics_(ss, d.LOGISTICS || {});
    writeInvestorNotes_(ss, d.HK_EVENTS || [], d.CF_EVENTS || []);
    writeOverview_(ss, d);

    return { ok: true, ts: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') };
  } catch(e) {
    Logger.log(e);
    return { ok: false, error: e.toString() };
  }
}

/** Reads all data from the spreadsheet and returns it as an object. */
function loadData() {
  try {
    var ss = getOrCreateSpreadsheet_();
    return {
      ok:         true,
      HK_EVENTS:  readEvents_(ss, 'HK_Events'),
      CF_EVENTS:  readEvents_(ss, 'CF_Events'),
      ACTIONS:    readActions_(ss),
      CONTACTS:   readContacts_(ss),
      LOGISTICS:  readLogistics_(ss),
      ts: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
    };
  } catch(e) {
    Logger.log(e);
    return { ok: false, error: e.toString() };
  }
}

// ── Sheet writers ─────────────────────────────────────────────

var EV_HDR_ = ['ID','Date','Time','Duration','Color','Status','Title','Company',
               'Venue','Attendees','Notes','Tags','Investor Tier','Budget','Outcome','Commitment'];

function writeEvents_(ss, sheetName, events) {
  var sh = getOrMake_(ss, sheetName);
  sh.clearContents();
  var rows = [EV_HDR_];
  events.forEach(function(e) {
    rows.push([
      e.id||'', e.date||'', e.time||'', e.dur||'',
      e.color||'', e.status||'', e.title||'', e.co||'',
      e.venue||'', e.att||'', e.notes||'',
      (e.tags||[]).join(', '),
      e.investorTier||'', e.budget||'', e.outcome||'', e.commitment||''
    ]);
  });
  if (rows.length > 1) {
    sh.getRange(1, 1, rows.length, EV_HDR_.length).setValues(rows);
    formatHeader_(sh, EV_HDR_.length);
    // Alternate row banding
    if (rows.length > 2) {
      sh.getRange(2, 1, rows.length-1, EV_HDR_.length)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    }
  }
}

function writeActions_(ss, actions) {
  var sh = getOrMake_(ss, 'Actions');
  sh.clearContents();
  var hdr  = ['ID','Task','Owner','Due Date','Phase','Status','Priority'];
  var rows = [hdr];
  actions.forEach(function(a) {
    rows.push([
      a.id||'', a.text||'', a.owner||'Delegate', a.due||'TBD',
      a.phase==='hk'?'Hong Kong':'Canton Fair',
      a.done?'Done':'Open',
      a.priority||'medium'
    ]);
  });
  sh.getRange(1,1,rows.length,hdr.length).setValues(rows);
  formatHeader_(sh, hdr.length);
  // Colour Done rows green
  for (var i=2; i<=rows.length; i++) {
    if (rows[i-1][5]==='Done') {
      sh.getRange(i,1,1,hdr.length).setBackground('#e8f5e9');
    }
  }
}

function writeContacts_(ss, contacts) {
  var sh  = getOrMake_(ss, 'Contacts');
  sh.clearContents();
  var hdr = ['ID','Name','Role','Company','Phase','Tags','Email','Phone',
             'Investor Type','Deal Value','Meeting Objective','Notes'];
  var rows = [hdr];
  contacts.forEach(function(c) {
    rows.push([
      c.id||'', c.name||'', c.role||'', c.co||'',
      c.phase==='hk'?'Hong Kong':'Canton Fair',
      (c.tags||[]).join(', '),
      c.email||'', c.phone||'',
      c.investorType||'', c.dealValue||'', c.meetingObjective||'', c.notes||''
    ]);
  });
  sh.getRange(1,1,rows.length,hdr.length).setValues(rows);
  formatHeader_(sh, hdr.length);
}

function writeLogistics_(ss, logistics) {
  var sh   = getOrMake_(ss, 'Logistics');
  sh.clearContents();
  var hdr  = ['ID','Type','Title','Details','Status'];
  var rows = [hdr];
  var push = function(arr, type) {
    (arr||[]).forEach(function(l) {
      rows.push([l.id||'', type, l.title||'', l.detail||'',
                 l.stat==='conf'?'Confirmed':'Pending']);
    });
  };
  push(logistics.flights, 'Flight');
  push(logistics.hotels,  'Hotel');
  push(logistics.ground,  'Transfer');
  sh.getRange(1,1,rows.length,hdr.length).setValues(rows);
  formatHeader_(sh, hdr.length);
}

function writeInvestorNotes_(ss, hkEvs, cfEvs) {
  var sh   = getOrMake_(ss, 'Investor_Notes');
  sh.clearContents();
  var hdr  = ['Phase','Investor Tier','Title','Company','Budget','Outcome','Commitment','Date'];
  var rows = [hdr];
  var addEvs = function(arr, phase) {
    arr.forEach(function(e) {
      if (e.investorTier || e.outcome || e.commitment || e.budget) {
        rows.push([phase, e.investorTier||'', e.title||'', e.co||'',
                   e.budget||'', e.outcome||'', e.commitment||'', e.date||'']);
      }
    });
  };
  addEvs(hkEvs, 'Hong Kong');
  addEvs(cfEvs, 'Canton Fair');
  sh.getRange(1,1,rows.length,hdr.length).setValues(rows);
  formatHeader_(sh, hdr.length);
  // Highlight tier1 rows gold
  for (var i=2; i<=rows.length; i++) {
    if (rows[i-1][1]==='tier1') sh.getRange(i,1,1,hdr.length).setBackground('#FBF3E0');
  }
}

function writeOverview_(ss, d) {
  var sh    = getOrMake_(ss, 'Overview');
  sh.clearContents();
  var hkEvs    = d.HK_EVENTS  || [];
  var cfEvs    = d.CF_EVENTS  || [];
  var actions  = d.ACTIONS    || [];
  var contacts = d.CONTACTS   || [];
  var allEvs   = hkEvs.concat(cfEvs);
  var tz       = Session.getScriptTimeZone();
  var now      = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm');

  var rows = [
    ['INVESTOR TRIP BOOK 2026','',''],
    ['Wellness & Recovery Lab Investiture · Canton Fair Phase 3','',''],
    ['','',''],
    ['METRIC','VALUE','LAST UPDATED'],
    ['HK Meetings',         hkEvs.length,   now],
    ['CF Events',           cfEvs.length,   ''],
    ['Confirmed',           allEvs.filter(function(e){return e.status==='confirmed';}).length, ''],
    ['Pending / TBC',       allEvs.filter(function(e){return e.status!=='confirmed';}).length, ''],
    ['Open Actions',        actions.filter(function(a){return !a.done;}).length,  ''],
    ['Completed Actions',   actions.filter(function(a){return a.done;}).length,   ''],
    ['Total Contacts',      contacts.length, ''],
    ['Lead Investors',      contacts.filter(function(c){return c.investorType==='lead_investor';}).length, ''],
    ['Strategic Partners',  contacts.filter(function(c){return c.investorType==='strategic';}).length,    ''],
    ['Meetings w/ Budget',  allEvs.filter(function(e){return e.budget;}).length,     ''],
    ['Outcomes Recorded',   allEvs.filter(function(e){return e.outcome;}).length,    ''],
    ['Commitments Logged',  allEvs.filter(function(e){return e.commitment;}).length, ''],
    ['','',''],
    ['Sheet generated by Investor Trip Book App','',''],
  ];

  sh.getRange(1,1,rows.length,3).setValues(rows);

  // Title styling
  sh.getRange(1,1,2,3)
    .setBackground('#0B2740')
    .setFontColor('#D4A843')
    .setFontWeight('bold')
    .setFontSize(12);

  // Header row
  sh.getRange(4,1,1,3)
    .setBackground('#163d5e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Data rows alternate
  sh.getRange(5,1,12,3).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 180);
}

// ── Sheet readers ─────────────────────────────────────────────

function readEvents_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .filter(function(r){ return r[0] !== '' && r[0] !== null; })
    .map(function(r) {
      return {
        id:           Number(r[0]) || 0,
        date:         String(r[1]  || ''),
        time:         String(r[2]  || ''),
        dur:          String(r[3]  || ''),
        color:        String(r[4]  || 'gcal-blue'),
        status:       String(r[5]  || 'pending'),
        title:        String(r[6]  || ''),
        co:           String(r[7]  || ''),
        venue:        String(r[8]  || ''),
        att:          String(r[9]  || ''),
        notes:        String(r[10] || ''),
        tags:         String(r[11] || '').split(',').map(function(t){return t.trim();}).filter(Boolean),
        investorTier: String(r[12] || ''),
        budget:       String(r[13] || ''),
        outcome:      String(r[14] || ''),
        commitment:   String(r[15] || '')
      };
    });
}

function readActions_(ss) {
  var sh = ss.getSheetByName('Actions');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .filter(function(r){ return r[0] !== ''; })
    .map(function(r) {
      return {
        id:       Number(r[0]) || Date.now(),
        text:     String(r[1] || ''),
        owner:    String(r[2] || 'Delegate'),
        due:      String(r[3] || 'TBD'),
        phase:    String(r[4] || '') === 'Hong Kong' ? 'hk' : 'cf',
        done:     String(r[5] || '') === 'Done',
        priority: String(r[6] || 'medium')
      };
    });
}

function readContacts_(ss) {
  var sh = ss.getSheetByName('Contacts');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1)
    .filter(function(r){ return r[1] !== ''; })
    .map(function(r) {
      return {
        id:               Number(r[0]) || 0,
        name:             String(r[1]  || ''),
        role:             String(r[2]  || ''),
        co:               String(r[3]  || ''),
        phase:            String(r[4]  || '') === 'Hong Kong' ? 'hk' : 'cf',
        tags:             String(r[5]  || '').split(',').map(function(t){return t.trim();}).filter(Boolean),
        email:            String(r[6]  || ''),
        phone:            String(r[7]  || ''),
        investorType:     String(r[8]  || ''),
        dealValue:        String(r[9]  || ''),
        meetingObjective: String(r[10] || ''),
        notes:            String(r[11] || ''),
        bg:               '#0B2740'
      };
    });
}

function readLogistics_(ss) {
  var sh = ss.getSheetByName('Logistics');
  if (!sh) return { flights:[], hotels:[], ground:[] };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { flights:[], hotels:[], ground:[] };

  var result = { flights:[], hotels:[], ground:[] };
  data.slice(1).filter(function(r){ return r[2] !== ''; }).forEach(function(r) {
    var t    = String(r[1]||'');
    var item = {
      id:     Number(r[0]) || 0,
      type:   t==='Flight'?'fl' : t==='Hotel'?'ht' : 'tr',
      ico:    t==='Flight'?'✈️' : t==='Hotel'?'🏨' : '🚗',
      title:  String(r[2]||''),
      detail: String(r[3]||''),
      stat:   String(r[4]||'') === 'Confirmed' ? 'conf' : 'pend'
    };
    if (item.type==='fl')      result.flights.push(item);
    else if (item.type==='ht') result.hotels.push(item);
    else                       result.ground.push(item);
  });
  return result;
}
