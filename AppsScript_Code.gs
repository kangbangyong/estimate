/**
 * 그리다 견적 — Google Sheets API
 * ------------------------------------------------------------------
 * 설치
 *   1) 구글 드라이브에서 빈 스프레드시트를 하나 만듭니다. 이름: "그리다 견적 DB"
 *   2) 확장 프로그램 → Apps Script → 기본 코드를 지우고 이 파일 전체를 붙여넣기
 *   3) 아래 TOKEN 을 아무도 모르는 긴 문자열로 바꿉니다
 *   4) 상단 함수 목록에서 초기화 를 골라 ▶ 실행 (권한 요청이 뜨면 허용)
 *   5) 배포 → 새 배포 → 유형: 웹 앱
 *        · 실행 사용자: 나
 *        · 액세스 권한: 모든 사용자
 *      → 배포 후 나오는 /exec 로 끝나는 주소를 복사해 둡니다
 *
 * 이 주소는 비밀번호와 같습니다. 직원에게만 알려주세요.
 * ------------------------------------------------------------------
 */

/** ★ 반드시 바꾸세요. 길고 아무 뜻 없는 문자열이면 됩니다. */
const TOKEN = 'grida-CHANGE-ME-1234abcd';

const TZ = 'Asia/Seoul';
const SS = () => SpreadsheetApp.getActiveSpreadsheet();

/* ================================================================
   1. 최초 1회 — 시트 만들기
   ================================================================ */
const SCHEMA = {
  '단가DB':   ['공종','품목명','규격','단위','재료비','노무비','경비','원가','단가합계','사용'],
  '견적이력': ['견적번호','작성일시','현장명','발주처','담당','공급가액','부가세','합계','원가합계','마진','품목수','비고'],
  '견적상세': ['견적번호','순번','공종','품목명','규격','단위','수량','재료비','노무비','경비','단가합계','금액','원가'],
  '설정':     ['항목','값']
};

const 설정기본값 = [
  ['회사명',''], ['사업자등록번호',''], ['주소',''], ['연락처',''], ['담당자',''],
  ['부가세율', 10], ['견적번호 접두사','G'], ['견적 유효기간','견적일로부터 30일']
];

function 초기화() {
  const ss = SS();
  Object.keys(SCHEMA).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const head = SCHEMA[name];
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#22282b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });

  // 설정 기본값은 비어 있을 때만 채웁니다
  const cfg = ss.getSheetByName('설정');
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 설정기본값.length, 2).setValues(설정기본값);
  }

  // 단가합계 = 재료비 + 노무비 + 경비 (500행까지)
  const db = ss.getSheetByName('단가DB');
  db.getRange('I2:I500').setFormulaR1C1('=IF(RC2="","",RC5+RC6+RC7)');   // 행마다 자기 행 참조
  db.getRange('E2:I500').setNumberFormat('#,##0');
  db.setColumnWidth(1, 150); db.setColumnWidth(2, 240); db.setColumnWidth(3, 180);

  const h = ss.getSheetByName('견적이력');
  h.getRange('F2:J1000').setNumberFormat('#,##0');
  h.getRange('B2:B1000').setNumberFormat('yyyy-mm-dd hh:mm');

  const d = ss.getSheetByName('견적상세');
  d.getRange('G2:M5000').setNumberFormat('#,##0');

  // 기본 Sheet1 이 남아 있으면 정리
  const junk = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (junk && ss.getSheets().length > 1) ss.deleteSheet(junk);

  Logger.log('초기화 완료 — 시트 4개 준비됨. 다음: 배포 → 새 배포 → 웹 앱');
  try { ss.toast('시트 4개가 준비됐습니다. 다음: 배포 → 새 배포 → 웹 앱', '초기화 완료', 8); } catch (e) {}
}

/* ================================================================
   2. 읽기 API
   ================================================================ */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.token !== TOKEN) return json({ ok: false, error: '토큰이 맞지 않습니다.' });

    switch (p.action) {
      case 'catalog':  return json({ ok: true, items: readCatalog(), settings: readSettings() });
      case 'settings': return json({ ok: true, settings: readSettings() });
      case 'history':  return json({ ok: true, rows: readHistory() });
      case 'estimate': return json({ ok: true, estimate: readEstimate(p.no) });
      default:         return json({ ok: true, pong: true, time: now() });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function readCatalog() {
  const sh = SS().getSheetByName('단가DB');
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const head = v[0].map(String);
  const at = name => head.indexOf(name);
  const items = [];
  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    const name = String(row[at('품목명')] || '').trim();
    if (!name) continue;
    if (String(row[at('사용')] || '').trim().toUpperCase() === 'N') continue;   // 단종·보류 품목 제외
    items.push({
      id:   'r' + (r + 1),                       // 시트 행 번호 = 품목 ID
      gong: String(row[at('공종')] || '').trim(),
      name: name,
      spec: String(row[at('규격')] || '').trim(),
      unit: String(row[at('단위')] || '식').trim(),
      mat:  num(row[at('재료비')]),
      lab:  num(row[at('노무비')]),
      exp:  num(row[at('경비')]),
      cost: num(row[at('원가')])
    });
  }
  return items;
}

function readSettings() {
  const sh = SS().getSheetByName('설정');
  if (!sh || sh.getLastRow() < 2) return {};
  const v = sh.getDataRange().getValues();
  const o = {};
  for (let r = 1; r < v.length; r++) {
    const k = String(v[r][0] || '').trim();
    if (k) o[k] = v[r][1];
  }
  return o;
}

function readHistory() {
  const sh = SS().getSheetByName('견적이력');
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(1, 1, last, SCHEMA['견적이력'].length).getValues()
           .map(r => r.map(c => (c instanceof Date) ? Utilities.formatDate(c, TZ, 'yyyy-MM-dd HH:mm') : c));
}

function readEstimate(no) {
  if (!no) return null;
  const det = SS().getSheetByName('견적상세');
  const last = det.getLastRow();
  if (last < 2) return null;
  const v = det.getRange(2, 1, last - 1, SCHEMA['견적상세'].length).getValues();
  const rows = v.filter(r => String(r[0]) === String(no)).map(r => ({
    gong: r[2], name: r[3], spec: r[4], unit: r[5], qty: r[6],
    mat: r[7], lab: r[8], exp: r[9], cost: r[12] / (r[6] || 1)
  }));

  const hist = SS().getSheetByName('견적이력');
  const hv = hist.getLastRow() > 1
    ? hist.getRange(2, 1, hist.getLastRow() - 1, SCHEMA['견적이력'].length).getValues()
    : [];
  const h = hv.find(r => String(r[0]) === String(no));

  return { no: no, client: h ? h[2] : '', owner: h ? h[3] : '', staff: h ? h[4] : '', rows: rows };
}

/* ================================================================
   3. 저장 API
   ================================================================ */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.token !== TOKEN) return json({ ok: false, error: '토큰이 맞지 않습니다.' });

    switch (body.action) {
      case 'save':      return json(saveEstimate(body.estimate || {}));
      case 'addItems':  return json(addItems(body.items || []));
      default:          return json({ ok: false, error: '알 수 없는 요청입니다: ' + body.action });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** 견적서 저장. 같은 견적번호가 이미 있으면 덮어씁니다. */
function saveEstimate(est) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss   = SS();
    const hist = ss.getSheetByName('견적이력');
    const det  = ss.getSheetByName('견적상세');
    const no   = est.no || nextNo();

    removeRows(hist, no);
    removeRows(det,  no);

    const rows = est.rows || [];
    let sup = 0, costSum = 0;

    const detRows = rows.map((r, i) => {
      const u   = num(r.mat) + num(r.lab) + num(r.exp);
      const qty = num(r.qty);
      const amt = u * qty;
      const c   = num(r.cost) * qty;
      sup += amt; costSum += c;
      return [no, i + 1, r.gong || '', r.name || '', r.spec || '', r.unit || '',
              qty, num(r.mat), num(r.lab), num(r.exp), u, amt, c];
    });

    if (detRows.length) {
      det.getRange(det.getLastRow() + 1, 1, detRows.length, detRows[0].length).setValues(detRows);
    }

    const rate  = num(readSettings()['부가세율'] || 10) / 100;
    const vat   = Math.round(sup * rate);
    const total = sup + vat;

    hist.appendRow([no, new Date(), est.client || '', est.owner || '', est.staff || '',
                    sup, vat, total, costSum, sup - costSum, rows.length, est.memo || '']);

    return { ok: true, no: no, sup: sup, vat: vat, total: total, cost: costSum, margin: sup - costSum };
  } finally {
    lock.releaseLock();
  }
}

/** 화면에서 등록한 신규 품목을 단가DB 끝에 덧붙입니다. */
/** 중복 기준: 공종 + 품목명 + 규격 (공백·대소문자 무시). 이미 있는 건 건너뜁니다. */
function itemKey(gong, name, spec) {
  return [gong, name, spec].map(function (s) {
    return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  }).join('|');
}

function addItems(items) {
  if (!items.length) return { ok: true, added: 0, skipped: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const db = SS().getSheetByName('단가DB');
    const have = {};
    if (db.getLastRow() > 1) {
      db.getRange(2, 1, db.getLastRow() - 1, 3).getValues().forEach(function (r) {
        if (String(r[1] || '').trim()) have[itemKey(r[0], r[1], r[2])] = true;
      });
    }
    const rows = [];
    let skipped = 0;
    items.forEach(function (i) {
      const k = itemKey(i.gong, i.name, i.spec);
      if (have[k]) { skipped++; return; }
      have[k] = true;
      rows.push([i.gong || '', i.name || '', i.spec || '', i.unit || '식',
                 num(i.mat), num(i.lab), num(i.exp), num(i.cost), '', '']);
    });
    if (!rows.length) return { ok: true, added: 0, skipped: skipped };
    const start = db.getLastRow() + 1;
    db.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
    // R1C1: 각 행이 자기 행을 참조 (setFormula 는 같은 수식을 그대로 복사해서 전부 첫 행을 가리킴)
    db.getRange(start, 9, rows.length, 1).setFormulaR1C1('=IF(RC2="","",RC5+RC6+RC7)');
    return { ok: true, added: rows.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/* ================================================================
   4. 보조
   ================================================================ */
/** G-260916-01 형태로 오늘 날짜 기준 다음 번호를 뽑습니다. */
function nextNo() {
  const prefix = String(readSettings()['견적번호 접두사'] || 'G').trim();
  const base   = prefix + '-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-';
  const hist   = SS().getSheetByName('견적이력');
  const last   = hist.getLastRow();
  let max = 0;
  if (last > 1) {
    hist.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      const s = String(r[0] || '');
      if (s.indexOf(base) === 0) {
        const n = parseInt(s.slice(base.length), 10);
        if (n > max) max = n;
      }
    });
  }
  return base + ('0' + (max + 1)).slice(-2);
}

function removeRows(sh, no) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let r = v.length - 1; r >= 0; r--) {
    if (String(v[r][0]) === String(no)) sh.deleteRow(r + 2);
  }
}

function num(v) {
  const n = Number(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function now() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

/* ================================================================
   5. 시트에서 직접 확인용 (배포 전 점검)
   ================================================================ */
function 점검() {
  const items = readCatalog();
  const cfg   = readSettings();
  Logger.log('단가DB 품목: %s개', items.length);
  Logger.log('설정: %s', JSON.stringify(cfg));
  Logger.log('다음 견적번호: %s', nextNo());
  if (items.length) Logger.log('첫 품목: %s', JSON.stringify(items[0]));
}
