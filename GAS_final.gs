// ════════════════════════════════════════════════════════════════════
// Google Apps Script — 輻射事件應變資訊平台
// ════════════════════════════════════════════════════════════════════

// ── 觸發 Drive 授權（首次部署後在 GAS 介面手動執行一次）──
function authDrive() {
  var folder = DriveApp.getRootFolder();
  var test = folder.createFile('_auth_test.txt', 'test', MimeType.PLAIN_TEXT);
  test.setTrashed(true);
}

// ── SHEET_ID 改用 function 避免高並發失敗 ──
function getSheet() {
  return SpreadsheetApp.openById('1WS8x91EL_fzH1AJSfB_9t076u37xmR9Nq-G2NvZYlZY');
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

// ════════════════════════════════════════════════════════════════════
// doGet 主路由
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  var action   = e.parameter.action   || 'maxSeq';

  // ── fetchCwaWind ──────────────────────────────────────────────────
  if (action === 'fetchCwaWind') {
    var evLat = parseFloat(e.parameter.lat);
    var evLng = parseFloat(e.parameter.lng);
    if (isNaN(evLat) || isNaN(evLng)) return respond(e, { error: '缺少座標參數' });
    try {
      var apiUrl = 'https://opendata.cwa.gov.tw/api/v1/rest/datastore/O-A0003-001' +
        '?Authorization=CWA-YOUR-KEY&format=JSON&limit=200';
      var resp = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
      var raw  = JSON.parse(resp.getContentText());
      var stations = raw.records && raw.records.Station ? raw.records.Station : [];
      var best = null, bestDist = Infinity;
      stations.forEach(function(st) {
        var lat = parseFloat(st.GeoInfo.Coordinates[0].StationLatitude);
        var lng = parseFloat(st.GeoInfo.Coordinates[0].StationLongitude);
        var d = Math.hypot(lat - evLat, lng - evLng);
        if (d < bestDist) { bestDist = d; best = st; }
      });
      if (!best) return respond(e, { error: '找不到最近測站' });
      var windDir = null, windSpeed = null;
      best.WeatherElement.forEach(function(w) {
        if (w.ElementName === 'WindDirection') windDir   = w.ElementValue[0].value;
        if (w.ElementName === 'WindSpeed')     windSpeed = w.ElementValue[0].value;
      });
      return respond(e, { ok: true, stationName: best.StationName, windDir: windDir, windSpeed: windSpeed });
    } catch(err) { return respond(e, { error: err.message }); }
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
          var isAdmin = colAdmin >= 0 ? String(data[r][colAdmin]||'').trim().toUpperCase() === 'V' : false;
          return respond(e, { ok: true, name: uName, team: uTeam, isAdmin: isAdmin });
        }
      }
    }
    return respond(e, { ok: false, msg: '帳號或密碼錯誤，請重試' });
  }

  // ── getKnowledge：透過 GAS 讀 Drive 資料夾（繞過 CORS）─────────────
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
            content = DocumentApp.openById(file.getId()).getBody().getText();
          } else if (mt === 'text/plain') {
            content = file.getBlob().getDataAsString('utf-8');
          }
          if (content && content.length > 50) {
            docs.push({ name: file.getName(), content: content });
          }
        } catch(fe) {
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
    var ss2 = getSheet();
    var sh2 = ss2.getSheetByName('事件紀錄');
    var vals = sh2 ? sh2.getRange('A:A').getValues().flat().filter(function(v){
      return typeof v === 'number' && v > 0;
    }) : [];
    return respond(e, { maxSeq: vals.length ? Math.max.apply(null, vals) : 0 });
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
      // 欄位對應（A~S）：
      // A=事件編號 B=事件名稱 C=事件類型 D=日期 E=時間 F=地址
      // G=緯度 H=經度 I=備註 J=建立時間 K=狀態
      // L=出勤人員 M=出動儀器 N=風向 O=風速 P=最大劑量率
      // Q=預估影響範圍 R=附件連結 S=風向鎖定
      evs.push({
        id:           String(r[0]),
        name:         String(r[1]||''),
        type:         String(r[2]||''),
        date:         String(r[3]||''),
        time:         String(r[4]||''),
        address:      String(r[5]||''),
        lat:          r[6]||'',
        lng:          r[7]||'',
        note:         String(r[8]||''),
        created:      String(r[9]||''),
        status:       String(r[10]||''),
        personnel:    String(r[11]||''),
        personnelStr: String(r[11]||''),   // 前端相容欄位
        instrument:   String(r[12]||''),
        instrumentStr:String(r[12]||''),   // 前端相容欄位
        windDir:      r[13] !== '' ? r[13] : '',
        windSpeed:    r[14] !== '' ? r[14] : '',
        doseMax:      r[15] !== '' ? r[15] : '',
        impactRange:  String(r[16]||''),
        attachments:  String(r[17]||''),
        windLocked:   String(r[18]||'').trim().toUpperCase() === 'V'
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
    // 動態找欄位（支援多種可能的標頭名稱）
    var ci = {};
    var colMap = {
      '事件編號': ['事件編號', '事件Id', 'eventId', 'event_id'],
      '回報者':   ['回報者', '回報人員', '人員', 'reporter'],
      '數值':     ['數值', '數值(μSv/h)', '數值(uSv/h)', '劑量率', '劑量率(μSv/h)', '劑量率(uSv/h)', 'value', 'dose'],
      '污染值':   ['污染值', '污染', '污染值(cps)', 'contamination'],
      '緯度':     ['緯度', 'lat', 'latitude', 'Lat'],
      '經度':     ['經度', 'lng', 'longitude', 'Lng'],
      '時間':     ['時間', '回報時間', '時間戳記', 'time', 'timestamp'],
      '待量測':   ['待量測', 'pending'],
      '備註':     ['備註', '使用儀器', '儀器', 'note', 'instrument']
    };
    Object.keys(colMap).forEach(function(key) {
      colMap[key].forEach(function(alias) {
        if (ci[key] === undefined) {
          var idx = hdr4.indexOf(alias);
          if (idx >= 0) ci[key] = idx;
        }
      });
    });
    for (var j = 1; j < rows4.length; j++) {
      var row = rows4[j];
      var evCol = ci['事件編號'];
      if (evCol === undefined || !row[evCol]) continue;
      if (String(row[evCol]).trim() !== eventId) continue;
      msrs.push({
        reporter:      ci['回報者']  !== undefined ? String(row[ci['回報者']] ||'') : '',
        value:         ci['數值']    !== undefined ? row[ci['數值']]                : '',
        contamination: ci['污染值']  !== undefined ? row[ci['污染值']]              : '',
        lat:           ci['緯度']    !== undefined ? row[ci['緯度']]                : '',
        lng:           ci['經度']    !== undefined ? row[ci['經度']]                : '',
        time:          ci['時間']    !== undefined ? String(row[ci['時間']]   ||'') : '',
        pending:       ci['待量測']  !== undefined ? String(row[ci['待量測']] ||'') : '',
        note:          ci['備註']    !== undefined ? String(row[ci['備註']]   ||'') : ''
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
    var tc = {};
    ['任務編號','任務內容','派送對象','狀態','建立時間','截止時間','回報內容'].forEach(function(n){
      var idx = hdr5.indexOf(n); if (idx >= 0) tc[n] = idx;
    });
    for (var k = 1; k < rows5.length; k++) {
      var tr = rows5[k];
      if (!tr[tc['任務編號']]) continue;
      var targets = String(tr[tc['派送對象']]||'').split(/[,，]/).map(function(t){return t.trim();}).filter(Boolean);
      tasks.push({
        id:       String(tr[tc['任務編號']]),
        content:  String(tr[tc['任務內容']] ||''),
        target:   String(tr[tc['派送對象']] ||''),
        targets:  targets,
        status:   String(tr[tc['狀態']]      ||''),
        created:  String(tr[tc['建立時間']]  ||''),
        deadline: String(tr[tc['截止時間']]  ||''),
        report:   String(tr[tc['回報內容']]  ||'')
      });
    }
    return respond(e, { tasks: tasks });
  }

  // ── instruments ──────────────────────────────────────────────────
  if (action === 'instruments') {
    var ss6 = getSheet();
    var sh6 = ss6.getSheetByName('儀器管理');
    if (!sh6) return respond(e, { columns: ['儀器名稱','型號','序號','狀態','備註','使用者'], rows: [] });
    var rows6 = sh6.getDataRange().getValues();
    if (rows6.length < 1) return respond(e, { columns: ['儀器名稱','型號','序號','狀態','備註','使用者'], rows: [] });
    var hdr6 = rows6[0].map(function(h){ return String(h||''); });
    // 確保有「使用者」欄
    if (hdr6.indexOf('使用者') < 0) hdr6.push('使用者');
    var outRows = [];
    for (var m = 1; m < rows6.length; m++) {
      var ir = rows6[m];
      if (!ir[0]) continue;
      var rowArr = hdr6.map(function(_, ci){ return String(ir[ci]||''); });
      outRows.push(rowArr);
    }
    return respond(e, { columns: hdr6, rows: outRows });
  }

  // ── group：依 id 取人員名單（前端 create/edit 用）────────────────────
  if (action === 'group') {
    var gId   = e.parameter.id || '';
    var shNameG = GROUP_SHEETS[gId];
    if (!shNameG) return respond(e, { members: [] });
    var ssG = getSheet();
    var shG = ssG.getSheetByName(shNameG);
    if (!shG) return respond(e, { members: [] });
    var rowsG = shG.getDataRange().getValues();
    var hdrsG = rowsG[0] || [];
    var members = [];
    for (var gi = 1; gi < rowsG.length; gi++) {
      var rg = rowsG[gi];
      if (!rg[0]) continue;
      var mObj = {};
      hdrsG.forEach(function(h, ci){ mObj[h] = String(rg[ci]||''); });
      // 統一欄位名稱
      mObj.name  = mObj['姓名']  || mObj['name']  || String(rg[0]||'');
      mObj.team  = mObj['小組']  || mObj['隊別'] || mObj['team']  || '';
      mObj.phone = mObj['電話']  || mObj['phone'] || '';
      mObj.rank  = mObj['職稱']  || mObj['rank']  || '';
      members.push(mObj);
    }
    return respond(e, { members: members });
  }

  // ── getSquads：取小組清單 ────────────────────────────────────────────
  if (action === 'getSquads') {
    return respond(e, { squads: Object.keys(GROUP_SHEETS) });
  }

  // ── groupData ────────────────────────────────────────────────────
  if (action === 'groupData') {
    var group  = e.parameter.group   || '';
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
      hdrs7.forEach(function(hh, idx){ obj[hh] = row7[idx]; });
      out7.push(obj);
    }
    return respond(e, { rows: out7 });
  }

  // ── envDebug：偵錯用，確認 CSV 標頭（部署後瀏覽器直接呼叫）──
  if (action === 'envDebug') {
    try {
      var dr = UrlFetchApp.fetch('https://www.nusc.gov.tw/open/gammamonitor_u.csv',
        { muteHttpExceptions: true, followRedirects: true });
      var dc = dr.getResponseCode();
      var preview = dr.getContentText('UTF-8').substring(0, 800);
      return respond(e, { code: dc, preview: preview });
    } catch(de) {
      return respond(e, { error: de.message });
    }
  }

  // ── envMonitor：抓官方環境輻射監測站即時資料 ─────────────────────
  if (action === 'envMonitor') {
    var csvUrls = [
      'https://www.nusc.gov.tw/open/gammamonitor_u.csv'
    ];
    var csv = '';
    var lastErr = '';
    for (var ui = 0; ui < csvUrls.length; ui++) {
      try {
        var resp = UrlFetchApp.fetch(csvUrls[ui], {
          muteHttpExceptions: true,
          followRedirects: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS/1.0)' }
        });
        var code = resp.getResponseCode();
        if (code === 200) {
          csv = resp.getContentText('UTF-8');
          if (csv && csv.length > 100) break;
        }
        lastErr = 'HTTP ' + code + ' from ' + csvUrls[ui];
      } catch(fetchErr) {
        lastErr = fetchErr.message;
      }
    }
    if (!csv || csv.length < 100) {
      return respond(e, { stations: [], error: '無法取得資料：' + lastErr });
    }
    try {
      var lines = csv.split(/\r?\n/).filter(function(l){ return l.trim(); });
      if (!lines.length) return respond(e, { stations: [] });

      // 動態找欄位（支援多種可能的標頭名稱）
      var headers = lines[0].split(',').map(function(h){ return h.trim().replace(/^"|"$/g,''); });
      function findCol(names) {
        for (var ni = 0; ni < names.length; ni++) {
          var idx = headers.indexOf(names[ni]);
          if (idx >= 0) return idx;
        }
        return -1;
      }
      var iName   = findCol(['監測站','測站名稱','站名','name']);
      var iNameEn = findCol(['監測站(英文)','測站英文名稱','英文名稱','nameEn','name_en']);
      var iVal    = findCol(['監測值(微西弗/時)','劑量率(nSv/h)','劑量率','dose_rate','value','nSv/h']);
      var iTime   = findCol(['時間','更新時間','time','datetime']);
      var iLng    = findCol(['GPS經度','經度','Lng','lng','longitude']);
      var iLat    = findCol(['GPS緯度','緯度','Lat','lat','latitude']);
      // 找不到就用位置推定（對應 nusc.gov.tw 格式）
      if (iName   < 0) iName   = 0;
      if (iNameEn < 0) iNameEn = 1;
      if (iVal    < 0) iVal    = 2;
      if (iTime   < 0) iTime   = 3;
      if (iLng    < 0) iLng    = 4;
      if (iLat    < 0) iLat    = 5;

      var stations = [];
      for (var si = 1; si < lines.length; si++) {
        // 處理 CSV 引號包圍的欄位
        var cols = [];
        var line = lines[si];
        var inQ = false, cur = '';
        for (var ci2 = 0; ci2 < line.length; ci2++) {
          var ch = line[ci2];
          if (ch === '"') { inQ = !inQ; }
          else if (ch === ',' && !inQ) { cols.push(cur.trim()); cur = ''; }
          else { cur += ch; }
        }
        cols.push(cur.trim());
        if (cols.length < 4) continue;
        var lat = parseFloat(cols[iLat]);
        var lng = parseFloat(cols[iLng]);
        var val = parseFloat(cols[iVal]);
        if (isNaN(lat) || isNaN(lng) || isNaN(val)) continue;
        // 判斷單位：nSv/h → μSv/h；若值 > 1000 可能已是 nSv/h
        var valueInUsv = (iVal >= 0 && headers[iVal] && headers[iVal].indexOf('nSv') >= 0)
          ? val / 1000
          : val;
        stations.push({
          name:   cols[iName]   || '',
          nameEn: cols[iNameEn] || '',
          lat:    lat,
          lng:    lng,
          value:  valueInUsv,
          time:   cols[iTime]   || ''
        });
      }
      return respond(e, { stations: stations });
    } catch(parseErr) {
      return respond(e, { stations: [], error: '解析失敗：' + parseErr.message });
    }
  }

  // ── getBirdHistory：讀取指定事件的 AI 對話紀錄 ──────────────────
  if (action === 'getBirdHistory') {
    var eventId = String(e.parameter.eventId || '').trim();
    var limit   = parseInt(e.parameter.limit  || '20') || 20;
    var ss = getSheet();
    var sh = ss.getSheetByName('AI對話紀錄');
    if (!sh) return respond(e, { history: [] });
    var rows = sh.getDataRange().getValues();
    if (rows.length < 2) return respond(e, { history: [] });
    var hdr = rows[0];
    var ciEv   = hdr.indexOf('事件ID');
    var ciRole = hdr.indexOf('角色');
    var ciCont = hdr.indexOf('內容');
    var ciTime = hdr.indexOf('時間');
    if (ciEv < 0 || ciRole < 0 || ciCont < 0) return respond(e, { history: [] });
    var history = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (eventId && String(r[ciEv]).trim() !== eventId) continue;
      history.push({
        time:    ciTime >= 0 ? String(r[ciTime]||'') : '',
        role:    String(r[ciRole]||'user'),
        content: String(r[ciCont]||'')
      });
    }
    return respond(e, { history: history.slice(-limit) });
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

  // ── uploadFile：上傳附件到 Drive ────────────────────────────────────
  if (action === 'uploadFile') {
    try {
      var fileName = payload.fileName || ('attachment_' + Date.now());
      var mimeType = payload.mimeType || 'application/octet-stream';
      var base64   = payload.data || '';
      var decoded  = Utilities.base64Decode(base64);
      var blob     = Utilities.newBlob(decoded, mimeType, fileName);
      var folder   = DriveApp.getRootFolder();
      var file     = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return jsonResp({ ok: true, url: file.getUrl(), id: file.getId() });
    } catch(ue) {
      return jsonResp({ ok: false, error: ue.message });
    }
  }

  // ── saveEvent ─────────────────────────────────────────────────────
  if (action === 'saveEvent') {
    var sh = ss.getSheetByName('事件紀錄');
    if (!sh) return jsonResp({ ok:false, error:'工作表不存在' });
    var ev = payload.event || payload;
    var data = sh.getDataRange().getValues();
    var found = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(ev.id||'').trim()) { found = i+1; break; }
    }
    var row = [
      ev.id, ev.name, ev.type, ev.date, ev.time, ev.address,
      ev.lat, ev.lng, ev.note, ev.created, ev.status,
      ev.personnelStr || ev.personnel || '',
      ev.instrumentStr || ev.instrument || '',
      ev.windDir, ev.windSpeed,
      ev.doseMax, ev.impactRange, ev.attachments,
      ev.windLocked ? 'V' : ''
    ];
    if (found > 0) {
      sh.getRange(found, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return jsonResp({ ok: true });
  }

  // ── updateEvent：部分欄位更新（不覆寫整筆）─────────────────────────
  if (action === 'updateEvent') {
    var sh = ss.getSheetByName('事件紀錄');
    if (!sh) return jsonResp({ ok:false, error:'工作表不存在' });
    var evId = String(payload.eventId || '').trim();
    if (!evId) return jsonResp({ ok:false, error:'缺少 eventId' });
    var rows = sh.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === evId) { rowIdx = i + 1; break; }
    }
    if (rowIdx < 0) return jsonResp({ ok:false, error:'找不到事件：' + evId });

    // 欄位對應：只更新有傳入的欄位
    var colMap = {
      name:         2,   // B
      type:         3,   // C
      date:         4,   // D
      time:         5,   // E
      address:      6,   // F
      lat:          7,   // G
      lng:          8,   // H
      note:         9,   // I
      status:       11,  // K
      personnelStr: 12,  // L
      personnel:    12,  // L（別名）
      instrumentStr:13,  // M
      instrument:   13,  // M（別名）
      windDir:      14,  // N
      windSpeed:    15,  // O
      doseMax:      16,  // P
      impactRange:  17,  // Q
      attachments:  18   // R
    };

    Object.keys(colMap).forEach(function(key) {
      if (payload[key] !== undefined) {
        sh.getRange(rowIdx, colMap[key]).setValue(payload[key]);
      }
    });

    // windManual：true → 鎖定（S欄設 V），false → 解鎖（S欄清空）
    if (payload.windManual !== undefined) {
      sh.getRange(rowIdx, 19).setValue(payload.windManual ? 'V' : '');
    }

    return jsonResp({ ok: true });
  }

  // ── saveMeasurement ───────────────────────────────────────────────
  if (action === 'saveMeasurement') {
    var sh2 = ss.getSheetByName('量測回報');
    if (!sh2) return jsonResp({ ok:false, error:'量測回報工作表不存在' });
    var m = payload.measurement || {};
    // 支援舊格式（flat）和新格式（nested measurement）
    var eventId       = m.eventId       || payload.eventId       || '';
    var reporter      = m.reporter      || payload.reporter      || '';
    var value         = m.value         !== undefined ? m.value         : (payload.value !== undefined ? payload.value : '');
    var contamination = m.contamination || payload.contamination || '';
    var lat           = m.lat           || payload.lat           || '';
    var lng           = m.lng           || payload.lng           || '';
    var time          = m.time          || payload.time          || new Date().toLocaleString('zh-TW');
    var pending       = m.pending       || payload.pending       || '';
    var note          = m.note          || payload.note          || payload.instrument || '';
    sh2.appendRow([eventId, reporter, value, contamination, lat, lng, time, pending, note]);
    return jsonResp({ ok: true });
  }

  // ── addMeasurement（前端舊格式相容）──────────────────────────────
  if (action === 'addMeasurement') {
    var sh2b = ss.getSheetByName('量測回報');
    if (!sh2b) return jsonResp({ ok:false, error:'量測回報工作表不存在' });
    sh2b.appendRow([
      payload.eventId       || '',
      payload.reporter      || '',
      payload.value         !== undefined ? payload.value : '',
      payload.contamination || '',
      payload.lat           || '',
      payload.lng           || '',
      new Date().toLocaleString('zh-TW'),
      '',
      payload.instrument    || ''
    ]);
    return jsonResp({ ok: true });
  }

  // ── saveTask ──────────────────────────────────────────────────────
  if (action === 'saveTask') {
    var sh3 = ss.getSheetByName('任務派送');
    if (!sh3) return jsonResp({ ok:false, error:'任務派送工作表不存在' });
    var t = payload.task || {};
    var data3 = sh3.getDataRange().getValues();
    var hdr3  = data3[0] || [];
    var colId = hdr3.indexOf('任務編號');
    var fnd3  = -1;
    for (var j = 1; j < data3.length; j++) {
      if (colId >= 0 && String(data3[j][colId]).trim() === String(t.id||'').trim()) { fnd3 = j+1; break; }
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
    if (!sh5 || !evId) return jsonResp({ ok:false });
    var rows5 = sh5.getDataRange().getValues();
    for (var k = 1; k < rows5.length; k++) {
      if (String(rows5[k][0]).trim() === evId) {
        sh5.getRange(k+1, 19).setValue(payload.lock ? 'V' : '');  // S = windLocked
        if (payload.windDir   !== undefined) sh5.getRange(k+1, 14).setValue(payload.windDir);
        if (payload.windSpeed !== undefined) sh5.getRange(k+1, 15).setValue(payload.windSpeed);
        return jsonResp({ ok: true });
      }
    }
    return jsonResp({ ok:false, error:'找不到事件' });
  }

  // ── saveBirdMsg：儲存 AI 對話紀錄 ────────────────────────────────
  if (action === 'saveBirdMsg') {
    var ss2 = getSheet();
    var sh2 = ss2.getSheetByName('AI對話紀錄');
    if (!sh2) {
      sh2 = ss2.insertSheet('AI對話紀錄');
      sh2.appendRow(['事件ID','時間','角色','內容']);
    }
    sh2.appendRow([
      String(payload.eventId || ''),
      new Date().toLocaleString('zh-TW'),
      String(payload.role    || 'user'),
      String(payload.content || '').substring(0, 3000)
    ]);
    return jsonResp({ ok: true });
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
