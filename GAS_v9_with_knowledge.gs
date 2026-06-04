// ════════════════════════════════════════════════════════════════════
// Google Apps Script — 輻射事件應變資訊平台 v9
// 新增：getKnowledge action（讀 Drive 資料夾，GAS 中介，無 CORS 問題）
// ════════════════════════════════════════════════════════════════════

// ── 觸發 Drive 授權（首次部署後在 GAS 介面手動執行一次）──
function authDrive() {
  var folder = DriveApp.getRootFolder();
  var test = folder.createFile('_auth_test.txt', 'test', MimeType.PLAIN_TEXT);
  test.setTrashed(true);
}

// ── SHEET_ID 改用 function 避免高並發失敗 ──
function getSheet() {
  return SpreadsheetApp.openById('YOUR_SPREADSHEET_ID_HERE');
}

var GROUP_SHEETS = {
  'tech-rdd':       '技術組-輻射彈',
  'tech-material':  '技術組-放射性物料',
  'tech-substance': '技術組-放射性物質',
  'staff':          '參謀組',
  'press':          '新聞組'
};

var LOGIN_SHEETS = [
  '技術組-輻射彈',
  '技術組-放射性物料',
  '技術組-放射性物質',
  '參謀組',
  '新聞組'
];

var TECH_GROUPS = ['tech-rdd', 'tech-material', 'tech-substance'];

// ════════════════════════════════════════════════════════════════════
// doGet 主路由
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  var action   = e.parameter.action   || 'maxSeq';
  var callback = e.parameter.callback || null;

  // ── fetchCwaWind ──────────────────────────────────────────────────
  if (action === 'fetchCwaWind') {
    var evLat = parseFloat(e.parameter.lat);
    var evLng = parseFloat(e.parameter.lng);
    if (isNaN(evLat) || isNaN(evLng)) {
      return respond(e, { error: '缺少座標參數' });
    }
    try {
      var apiUrl = 'https://opendata.cwa.gov.tw/api/v1/rest/datastore/O-A0003-001' +
        '?Authorization=CWA-YOUR-KEY&format=JSON&limit=200';
      var resp    = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
      var raw     = JSON.parse(resp.getContentText());
      var stations = raw.records && raw.records.Station ? raw.records.Station : [];

      var best = null, bestDist = Infinity;
      stations.forEach(function(st) {
        var lat = parseFloat(st.GeoInfo.Coordinates[0].StationLatitude);
        var lng = parseFloat(st.GeoInfo.Coordinates[0].StationLongitude);
        var d = Math.hypot(lat - evLat, lng - evLng);
        if (d < bestDist) { bestDist = d; best = st; }
      });

      if (!best) return respond(e, { error: '找不到最近測站' });

      var we = best.WeatherElement;
      var windDir   = null, windSpeed = null;
      we.forEach(function(w) {
        if (w.ElementName === 'WindDirection') windDir   = w.ElementValue[0].value;
        if (w.ElementName === 'WindSpeed')     windSpeed = w.ElementValue[0].value;
      });
      return respond(e, {
        ok:         true,
        stationName: best.StationName,
        windDir:    windDir,
        windSpeed:  windSpeed
      });
    } catch(err) {
      return respond(e, { error: err.message });
    }
  }

  // ── login ─────────────────────────────────────────────────────────
  if (action === 'login') {
    var acc = String(e.parameter.acc || '').trim();
    var pwd = String(e.parameter.pwd || '').trim();
    if (!acc || !pwd) return respond(e, { ok: false, msg: '請輸入帳號和密碼' });

    var ss = getSheet();
    for (var si = 0; si < LOGIN_SHEETS.length; si++) {
      var sh = ss.getSheetByName(LOGIN_SHEETS[si]);
      if (!sh) continue;
      var data = sh.getDataRange().getValues();
      if (data.length < 2) continue;

      var headers = data[0];
      var colAcc=-1, colPwd=-1, colAdmin=-1, colName=-1, colTeam=-1;
      for (var h = 0; h < headers.length; h++) {
        var hs = String(headers[h]).trim();
        if (hs==='帳號')   colAcc   = h;
        if (hs==='密碼')   colPwd   = h;
        if (hs==='管理員') colAdmin = h;
        if (hs==='姓名')   colName  = h;
        if (hs==='隊別')   colTeam  = h;
      }
      if (colAcc < 0 || colPwd < 0) continue;

      for (var r = 1; r < data.length; r++) {
        var ra = String(data[r][colAcc]||'').trim();
        var rp = String(data[r][colPwd]||'').trim();
        if (!ra && !rp) continue;
        if (ra === acc && rp === pwd) {
          var uName   = colName  >= 0 ? String(data[r][colName] ||'').trim() : acc;
          var uTeam   = colTeam  >= 0 ? String(data[r][colTeam] ||'').trim() : '';
          var isAdmin = colAdmin >= 0
            ? String(data[r][colAdmin]||'').trim().toUpperCase() === 'V'
            : false;
          return respond(e, { ok:true, name:uName, team:uTeam, isAdmin:isAdmin });
        }
      }
    }
    return respond(e, { ok:false, msg:'帳號或密碼錯誤，請重試' });
  }

  // ── getKnowledge：讀 Drive 資料夾所有 Google Doc 回傳文字 ──────────
  if (action === 'getKnowledge') {
    var folderId = e.parameter.folderId || '1XRM7zIkRotFaqX_afKw7ij4PlG-6z5Fq';
    try {
      var folder = DriveApp.getFolderById(folderId);
      var files  = folder.getFiles();
      var docs   = [];
      while (files.hasNext()) {
        var file = files.next();
        try {
          var content = '';
          var mt = file.getMimeType();
          if (mt === 'application/vnd.google-apps.document') {
            // Google Doc → 直接取純文字
            content = DocumentApp.openById(file.getId()).getBody().getText();
          } else if (mt === 'text/plain') {
            // 純文字檔
            content = file.getBlob().getDataAsString('utf-8');
          }
          // PDF、圖片等跳過
          if (content && content.length > 50) {
            docs.push({ name: file.getName(), content: content });
          }
        } catch(fe) {
          // 單一檔案失敗不影響其他
          docs.push({ name: file.getName(), content: '', error: fe.message });
        }
      }
      return respond(e, { ok: true, docs: docs });
    } catch(err) {
      return respond(e, { ok: false, docs: [], error: err.message });
    }
  }

  // ── maxSeq ────────────────────────────────────────────────────────
  if (action === 'maxSeq') {
    var ss2  = getSheet();
    var sh2  = ss2.getSheetByName('事件紀錄');
    var vals = sh2 ? sh2.getRange('A:A').getValues().flat().filter(function(v){
      return typeof v === 'number' && v > 0;
    }) : [];
    return respond(e, { maxSeq: vals.length ? Math.max.apply(null,vals) : 0 });
  }

  // ── events ────────────────────────────────────────────────────────
  if (action === 'events') {
    var ss3 = getSheet();
    var sh3 = ss3.getSheetByName('事件紀錄');
    if (!sh3) return respond(e, { events: [] });
    var rows = sh3.getDataRange().getValues();
    var evs  = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      evs.push({
        id:         String(r[0]),
        name:       String(r[1]||''),
        type:       String(r[2]||''),
        date:       String(r[3]||''),
        time:       String(r[4]||''),
        address:    String(r[5]||''),
        lat:        r[6]||'',
        lng:        r[7]||'',
        note:       String(r[8]||''),
        created:    String(r[9]||''),
        status:     String(r[10]||''),
        personnel:  String(r[11]||''),
        instrument: String(r[12]||''),
        windDir:    r[13]!==''?r[13]:'',
        doseMax:    r[15]!==''?r[15]:'',
        windSpeed:  r[14]!==''?r[14]:'',
        windLocked: String(r[18]||'').trim().toUpperCase() === 'V'
      });
    }
    return respond(e, { events: evs });
  }

  // ── measurements ─────────────────────────────────────────────────
  if (action === 'measurements') {
    var eventId = String(e.parameter.eventId || '').trim();
    if (!eventId) return respond(e, { measurements: [] });
    var ss4 = getSheet();
    var sh4 = ss4.getSheetByName('量測回報');
    if (!sh4) return respond(e, { measurements: [] });
    var rows4 = sh4.getDataRange().getValues();
    var msrs  = [];
    var hdr4  = rows4[0] || [];
    // 動態找欄位
    var ci = {}; ['事件編號','回報者','數值','污染值','緯度','經度','時間','待量測','備註'].forEach(function(n){
      var idx = hdr4.indexOf(n); if (idx >= 0) ci[n] = idx;
    });
    for (var j = 1; j < rows4.length; j++) {
      var row = rows4[j];
      if (!row[ci['事件編號']] || String(row[ci['事件編號']]).trim() !== eventId) continue;
      msrs.push({
        reporter:      String(row[ci['回報者']]    ||''),
        value:         row[ci['數值']]      !== undefined ? row[ci['數值']]      : '',
        contamination: row[ci['污染值']]    !== undefined ? row[ci['污染值']]    : '',
        lat:           row[ci['緯度']]      !== undefined ? row[ci['緯度']]      : '',
        lng:           row[ci['經度']]      !== undefined ? row[ci['經度']]      : '',
        time:          String(row[ci['時間']]      ||''),
        pending:       String(row[ci['待量測']]    ||''),
        note:          String(row[ci['備註']]      ||'')
      });
    }
    return respond(e, { measurements: msrs });
  }

  // ── tasks ─────────────────────────────────────────────────────────
  if (action === 'tasks') {
    var ss5 = getSheet();
    var sh5 = ss5.getSheetByName('任務派送');
    if (!sh5) return respond(e, { tasks: [] });
    var rows5 = sh5.getDataRange().getValues();
    var tasks = [];
    var hdr5  = rows5[0] || [];
    var tc = {}; ['任務編號','任務內容','派送對象','狀態','建立時間','截止時間','回報內容'].forEach(function(n){
      var idx = hdr5.indexOf(n); if (idx>=0) tc[n]=idx;
    });
    for (var k = 1; k < rows5.length; k++) {
      var tr = rows5[k];
      if (!tr[tc['任務編號']]) continue;
      // 多目標：派送對象可能有逗號分隔
      var targets = String(tr[tc['派送對象']]||'').split(/[,，]/).map(function(t){return t.trim();}).filter(Boolean);
      tasks.push({
        id:       String(tr[tc['任務編號']]),
        content:  String(tr[tc['任務內容']]  ||''),
        target:   String(tr[tc['派送對象']]  ||''),
        targets:  targets,
        status:   String(tr[tc['狀態']]       ||''),
        created:  String(tr[tc['建立時間']]   ||''),
        deadline: String(tr[tc['截止時間']]   ||''),
        report:   String(tr[tc['回報內容']]   ||'')
      });
    }
    return respond(e, { tasks: tasks });
  }

  // ── instruments ──────────────────────────────────────────────────
  if (action === 'instruments') {
    var ss6 = getSheet();
    var sh6 = ss6.getSheetByName('儀器管理');
    if (!sh6) return respond(e, { instruments: [] });
    var rows6 = sh6.getDataRange().getValues();
    var insts = [];
    var hdr6  = rows6[0] || [];
    var ic = {}; ['儀器名稱','型號','序號','狀態','備註'].forEach(function(n){
      var idx = hdr6.indexOf(n); if(idx>=0) ic[n]=idx;
    });
    for (var m = 1; m < rows6.length; m++) {
      var ir = rows6[m];
      if (!ir[ic['儀器名稱']]) continue;
      insts.push({
        name:   String(ir[ic['儀器名稱']]||''),
        model:  String(ir[ic['型號']]    ||''),
        serial: String(ir[ic['序號']]    ||''),
        status: String(ir[ic['狀態']]    ||''),
        note:   String(ir[ic['備註']]    ||'')
      });
    }
    return respond(e, { instruments: insts });
  }

  // ── groupData（各組表單資料）────────────────────────────────────
  if (action === 'groupData') {
    var group  = e.parameter.group  || '';
    var evtId2 = e.parameter.eventId || '';
    var shName = GROUP_SHEETS[group];
    if (!shName) return respond(e, { rows: [] });
    var ss7 = getSheet();
    var sh7 = ss7.getSheetByName(shName);
    if (!sh7) return respond(e, { rows: [] });
    var rows7 = sh7.getDataRange().getValues();
    var hdrs7 = rows7[0] || [];
    var out7  = [];
    var colEvId = hdrs7.indexOf('事件編號');
    for (var n = 1; n < rows7.length; n++) {
      var row7 = rows7[n];
      if (evtId2 && colEvId >= 0 && String(row7[colEvId]).trim() !== evtId2) continue;
      var obj = {};
      hdrs7.forEach(function(h, idx){ obj[h] = row7[idx]; });
      out7.push(obj);
    }
    return respond(e, { rows: out7 });
  }

  // ── envMonitor：抓官方環境輻射監測站即時資料 ──────────────────────
  if (action === 'envMonitor') {
    try {
      // 原能會環境輻射即時資料（CSV）
      var csvUrl = 'https://www.trb.org.tw/OpenData/CSV/all.csv';
      var resp = UrlFetchApp.fetch(csvUrl, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      var csv = resp.getContentText('UTF-8');
      var lines = csv.split(/\r?\n/).filter(function(l){ return l.trim(); });
      if (!lines.length) return respond(e, { stations: [] });

      // 第一列是標頭，找欄位位置
      var headers = lines[0].split(',').map(function(h){ return h.trim().replace(/^"|"$/g,''); });
      var iName   = headers.indexOf('測站名稱');
      var iNameEn = headers.indexOf('測站英文名稱');
      var iLat    = headers.indexOf('緯度');
      var iLng    = headers.indexOf('經度');
      var iVal    = headers.indexOf('劑量率(nSv/h)');
      var iTime   = headers.indexOf('更新時間');
      // 若找不到欄位名稱，用位置推定（不同版本 CSV 欄位順序可能不同）
      if (iName  < 0) iName   = 0;
      if (iNameEn< 0) iNameEn = 1;
      if (iLat   < 0) iLat    = 4;
      if (iLng   < 0) iLng    = 5;
      if (iVal   < 0) iVal    = 6;
      if (iTime  < 0) iTime   = 7;

      var stations = [];
      for (var si = 1; si < lines.length; si++) {
        var cols = lines[si].split(',').map(function(c){ return c.trim().replace(/^"|"$/g,''); });
        if (cols.length < 4) continue;
        var lat = parseFloat(cols[iLat]);
        var lng = parseFloat(cols[iLng]);
        var val = parseFloat(cols[iVal]);
        if (isNaN(lat)||isNaN(lng)||isNaN(val)) continue;
        // 將 nSv/h 轉成 μSv/h（前端 _fmtDose 會自動換算單位顯示）
        stations.push({
          name:   cols[iName]   || '',
          nameEn: cols[iNameEn] || '',
          lat:    lat,
          lng:    lng,
          value:  val / 1000,   // nSv/h → μSv/h
          time:   cols[iTime]   || ''
        });
      }
      return respond(e, { stations: stations });
    } catch(err) {
      return respond(e, { stations: [], error: err.message });
    }
  }

  return respond(e, { error: '未知 action: ' + action });
}

// ════════════════════════════════════════════════════════════════════
// doPost：寫入資料
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  var action = '';
  var payload = {};
  try {
    payload = JSON.parse(e.postData.contents);
    action  = payload.action || '';
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:'JSON parse error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var ss = getSheet();

  // ── saveEvent ─────────────────────────────────────────────────────
  if (action === 'saveEvent') {
    var sh = ss.getSheetByName('事件紀錄');
    if (!sh) return jsonResp({ ok:false, error:'工作表不存在' });
    var ev = payload.event || {};
    // 找是否已存在此 id
    var data = sh.getDataRange().getValues();
    var found = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(ev.id||'').trim()) { found = i+1; break; }
    }
    var row = [
      ev.id, ev.name, ev.type, ev.date, ev.time, ev.address,
      ev.lat, ev.lng, ev.note, ev.created, ev.status,
      ev.personnel, ev.instrument, ev.windDir, ev.windSpeed,
      ev.doseMax, ev.impactRange, ev.attachments, ev.windLocked?'V':''
    ];
    if (found > 0) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonResp({ ok: true });
  }

  // ── saveMeasurement ───────────────────────────────────────────────
  if (action === 'saveMeasurement') {
    var sh2 = ss.getSheetByName('量測回報');
    if (!sh2) return jsonResp({ ok:false, error:'量測回報工作表不存在' });
    var m = payload.measurement || {};
    sh2.appendRow([
      m.eventId, m.reporter, m.value, m.contamination,
      m.lat, m.lng, m.time, m.pending, m.note
    ]);
    return jsonResp({ ok: true });
  }

  // ── saveTask ──────────────────────────────────────────────────────
  if (action === 'saveTask') {
    var sh3 = ss.getSheetByName('任務派送');
    if (!sh3) return jsonResp({ ok:false, error:'任務派送工作表不存在' });
    var t = payload.task || {};
    // 找是否已存在
    var data3 = sh3.getDataRange().getValues();
    var hdr3  = data3[0] || [];
    var colId = hdr3.indexOf('任務編號');
    var fnd3  = -1;
    for (var j = 1; j < data3.length; j++) {
      if (colId>=0 && String(data3[j][colId]).trim() === String(t.id||'').trim()) { fnd3=j+1; break; }
    }
    var taskRow = [t.id, t.content, t.target, t.status, t.created, t.deadline, t.report];
    if (fnd3 > 0) {
      sh3.getRange(fnd3, 1, 1, taskRow.length).setValues([taskRow]);
    } else {
      sh3.appendRow(taskRow);
    }
    return jsonResp({ ok: true });
  }

  // ── saveGroupData ─────────────────────────────────────────────────
  if (action === 'saveGroupData') {
    var group   = payload.group   || '';
    var rowData = payload.rowData || {};
    var shName2 = GROUP_SHEETS[group];
    if (!shName2) return jsonResp({ ok:false, error:'未知群組' });
    var sh4 = ss.getSheetByName(shName2);
    if (!sh4) return jsonResp({ ok:false, error:'工作表不存在：'+shName2 });
    var hdrs4 = sh4.getRange(1, 1, 1, sh4.getLastColumn()).getValues()[0];
    var newRow = hdrs4.map(function(h){ return rowData[h] !== undefined ? rowData[h] : ''; });
    sh4.appendRow(newRow);
    return jsonResp({ ok: true });
  }

  // ── windLock ──────────────────────────────────────────────────────
  if (action === 'windLock') {
    var sh5  = ss.getSheetByName('事件紀錄');
    var evId = String(payload.eventId||'').trim();
    var lock = payload.lock ? 'V' : '';
    var wd   = payload.windDir   !== undefined ? payload.windDir   : null;
    var ws   = payload.windSpeed !== undefined ? payload.windSpeed : null;
    if (!sh5 || !evId) return jsonResp({ ok:false });
    var rows5 = sh5.getDataRange().getValues();
    for (var k = 1; k < rows5.length; k++) {
      if (String(rows5[k][0]).trim() === evId) {
        sh5.getRange(k+1, 19).setValue(lock); // col S = windLocked
        if (wd   !== null) sh5.getRange(k+1, 14).setValue(wd);
        if (ws   !== null) sh5.getRange(k+1, 15).setValue(ws);
        return jsonResp({ ok: true });
      }
    }
    return jsonResp({ ok:false, error:'找不到事件' });
  }

  return jsonResp({ ok: false, error: '未知 action: ' + action });
}

// ════════════════════════════════════════════════════════════════════
// 工具函式
// ════════════════════════════════════════════════════════════════════
function respond(e, obj) {
  var cb  = e.parameter ? e.parameter.callback : null;
  var out = JSON.stringify(obj);
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
