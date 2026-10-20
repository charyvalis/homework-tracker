/**
 * 作業缺交通知看板（多班級版・免生日）- Google Apps Script 後端
 * 學生／家長只需選「班級」並輸入「座號」即可查詢。
 *
 * 使用前：專案設定 → 指令碼屬性，新增一個：
 *   TEACHER_PW = 老師密碼
 * 部署：部署 → 新增部署 → 網頁應用程式 → 執行身分「我」、存取權「所有人」
 * 注意：改完程式碼後，要「管理部署作業 → 編輯 → 新版本」才會生效。
 */
const P = PropertiesService.getScriptProperties();
const SS = SpreadsheetApp.getActiveSpreadsheet();
const H = {
  Assignments: ['id', 'cls', 'title', 'published', 'created'],
  Records: ['aid', 'seat', 'status', 'date'],
  Roster: ['cls', 'seat']
};

function sheet(n){ let s = SS.getSheetByName(n); if(!s){ s = SS.insertSheet(n); s.appendRow(H[n]); } return s; }
function rows(n){ const v = sheet(n).getDataRange().getValues(); v.shift(); return v; }
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function classList(){
  const set = {};
  rows('Roster').forEach(x => { set[String(x[0])] = 1; });
  return Object.keys(set).sort((a,b)=> a.localeCompare(b, 'zh-Hant', {numeric:true}));
}

function doPost(e){
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return out(handle(JSON.parse(e.postData.contents))); }
  catch(err){ return out({ok:false, error:String(err)}); }
  finally { lock.releaseLock(); }
}
function doGet(){ return out({ok:true, msg:'homework tracker api'}); }

function handle(r){
  // ---- 公開（學生／家長）----
  if(r.action === 'classes') return {ok:true, classes: classList()};
  if(r.action === 'query') return query(r.cls, r.seat);

  // ---- 老師（需密碼）----
  if(!r.pw || r.pw !== P.getProperty('TEACHER_PW')) return {ok:false, error:'密碼錯誤'};
  switch(r.action){
    case 'load': {
      const classes = classList();
      const cls = classes.indexOf(String(r.cls)) >= 0 ? String(r.cls) : (classes[0] || '');
      const as = rows('Assignments').filter(x=>String(x[1])===cls)
        .map(x=>({id:String(x[0]), title:x[2], published:Number(x[3])===1}));
      const ids = {}; as.forEach(a => ids[a.id] = 1);
      const rs = rows('Records').filter(x=>ids[String(x[0])])
        .map(x=>({aid:String(x[0]), seat:Number(x[1]), status:x[2], date:String(x[3]).replace(/^'/,'')}));
      const seats = rows('Roster').filter(x=>String(x[0])===cls).map(x=>Number(x[1])).sort((a,b)=>a-b);
      return {ok:true, classes, cls, assignments:as.reverse(), records:rs, seats};
    }
    case 'add': {
      const id = Date.now().toString(36);
      sheet('Assignments').appendRow([id, "'" + r.cls, r.title, 0, new Date()]);
      return {ok:true, id};
    }
    case 'publish': {
      const s = sheet('Assignments'), v = s.getDataRange().getValues();
      for(let i=1;i<v.length;i++) if(String(v[i][0])===r.aid){ s.getRange(i+1,4).setValue(r.value?1:0); break; }
      return {ok:true};
    }
    case 'del': {
      del('Assignments', x=>String(x[0])===r.aid);
      del('Records', x=>String(x[0])===r.aid);
      return {ok:true};
    }
    case 'set': {
      del('Records', x=>String(x[0])===r.aid && Number(x[1])===Number(r.seat));
      if(r.status !== 'done') sheet('Records').appendRow([r.aid, r.seat, r.status, "'" + (r.date||'')]);
      return {ok:true};
    }
    case 'roster': {
      // 設定某班的座號，例如："1-32" 或 "1-10, 12-30"
      const cls = String(r.cls||'').trim();
      if(!cls) return {ok:false, error:'請輸入班級名稱'};
      const seats = parseSeats(r.text);
      if(!seats.length) return {ok:false, error:'座號格式不正確，例如 1-32'};
      del('Roster', x=>String(x[0])===cls);
      const s = sheet('Roster');
      const data = seats.map(n => ["'" + cls, n]);
      s.getRange(s.getLastRow()+1, 1, data.length, 2).setValues(data);
      return {ok:true, count:seats.length, cls};
    }
    case 'delClass': {
      const cls = String(r.cls);
      const ids = {};
      rows('Assignments').forEach(x => { if(String(x[1])===cls) ids[String(x[0])] = 1; });
      del('Records', x=>ids[String(x[0])]);
      del('Assignments', x=>String(x[1])===cls);
      del('Roster', x=>String(x[0])===cls);
      return {ok:true};
    }
  }
  return {ok:false, error:'unknown action'};
}

function parseSeats(text){
  const set = {};
  String(text||'').split(/[,，、\s]+/).forEach(t => {
    const m = t.match(/^(\d{1,2})(?:[-~～－](\d{1,2}))?$/);
    if(!m) return;
    const a = Number(m[1]), b = m[2] ? Number(m[2]) : a;
    for(let i=Math.min(a,b); i<=Math.max(a,b); i++) if(i>=1 && i<=99) set[i] = 1;
  });
  return Object.keys(set).map(Number).sort((a,b)=>a-b);
}

function del(name, test){
  const s = sheet(name), v = s.getDataRange().getValues();
  for(let i=v.length-1;i>=1;i--) if(test(v[i])) s.deleteRow(i+1);
}

function query(cls, seat){
  cls = String(cls||'').trim(); seat = parseInt(seat, 10);
  if(!cls || !seat) return {ok:false, error:'請選擇班級並輸入座號'};
  const ok = rows('Roster').some(x=>String(x[0])===cls && Number(x[1])===seat);
  if(!ok) return {ok:false, error:'找不到這個班級或座號，請再確認'};
  const recs = rows('Records').filter(x=>Number(x[1])===seat);
  const items = rows('Assignments').filter(x=>String(x[1])===cls && Number(x[3])===1).map(a=>{
    const rec = recs.find(x=>String(x[0])===String(a[0]));
    return {title:a[2], status: rec?rec[2]:'done', date: rec?String(rec[3]).replace(/^'/,''):''};
  }).reverse();
  return {ok:true, cls, seat, items};
}
