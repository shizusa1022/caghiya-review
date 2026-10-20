// カギヤ 口コミのカギ — アンケートの保存と読み出し（v2・2026-10-20）
// v2：お店への意見（4欄）・区分（店舗／醸造部イベント）・段階を足した。見出しの名前で列を合わせるので、前の列の並びのままでも動く。
// 読み出し（doGet）：スクリプトのプロパティ READ_KEY と同じ key を付けたときだけ、回答をJSONで返す（社員ポータルのダッシュボード用）。
var COLS = ['日時','区分','店舗・イベント','よかった点','ビール','同伴','テイスト','ひとこと','生成文',
            '意見_ビール','意見_接客','意見_お店','意見_そのほか','段階','流入(QR)','端末'];
var OLD = {'店舗':'店舗・イベント'}; // v1の見出しの読み替え

function sheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('回答') || ss.insertSheet('回答');
  if (sh.getLastRow() === 0) { sh.appendRow(COLS); return sh; }
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return OLD[h] || h; });
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  COLS.forEach(function(c){ if (head.indexOf(c) < 0) { sh.getRange(1, head.length + 1).setValue(c); head.push(c); } });
  return sh;
}

function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var d = JSON.parse(e.postData.contents);
    var o = d.opinions || {};
    var sh = sheet_();
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var v = {
      '日時': d.ts ? new Date(d.ts) : new Date(),
      '区分': d.kind === 'brewery' ? '醸造部' : '店舗',
      '店舗・イベント': d.kind === 'brewery' ? (d.event || 'イベント') : (d.store || ''),
      'よかった点': (d.points || []).join(' / '),
      'ビール': (d.beers || []).join(' / '),
      '同伴': d.withwho || '', 'テイスト': d.tone || '', 'ひとこと': d.personal || '', '生成文': d.draft || '',
      '意見_ビール': o.beer || '', '意見_接客': o.staff || '', '意見_お店': o.shop || '', '意見_そのほか': o.other || '',
      '段階': d.stage || '', '流入(QR)': d.src || '', '端末': d.device || ''
    };
    sh.appendRow(head.map(function(h){ return v.hasOwnProperty(h) ? v[h] : ''; }));
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }finally{
    lock.releaseLock();
  }
}

function doGet(e){
  var key = PropertiesService.getScriptProperties().getProperty('READ_KEY');
  if (!key || !e || !e.parameter || e.parameter.key !== key) return ContentService.createTextOutput('OK');
  var sh = sheet_();
  var vals = sh.getDataRange().getValues();
  var head = vals.shift();
  var rows = vals.map(function(r){
    var x = {}; head.forEach(function(h, i){ if (h !== '端末') x[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; }); return x;
  });
  return ContentService.createTextOutput(JSON.stringify({ok:true, rows:rows})).setMimeType(ContentService.MimeType.JSON);
}
