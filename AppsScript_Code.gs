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
  '단가DB':   ['대공종','공종','품목명','규격','단위','자재비','인건비','경비','단가합계','원가','사용'],
  '견적이력': ['견적번호','작성일시','현장명','발주처','담당','공급가액','부가세','합계','원가합계','마진','품목수','비고','시공사','발주처사업자번호','메타'],
  '시공사':   ['사업자등록번호','상호','성명','사업장주소','업태','종목','이메일','담당자성명','담당자연락처','로고이미지','도장이미지','FAX'],
  '발주처':   ['사업자등록번호','상호','성명','사업장주소','업태','종목','이메일','담당자성명','담당자연락처','로고이미지'],
  '견적상세': ['견적번호','순번','공종','품목명','규격','단위','수량','자재비','인건비','경비','단가합계','금액','원가'],
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

  // 단가DB 정리: 품목명 없는 빈 행을 걷어내고 데이터를 2행부터 붙여 올립니다
  // (예전 초기화가 I열에 수식을 500행까지 깔아둔 탓에 새 품목이 501행 아래로 들어가던 문제 복구)
  const db = ss.getSheetByName('단가DB');
  const last = db.getLastRow();
  if (last > 1) {
    const all  = db.getRange(2, 1, last - 1, 11).getValues();
    const keep = all.filter(function (r) { return String(r[2] || '').trim(); });   // C열 품목명 있는 행만
    db.getRange(2, 1, last - 1, 11).clearContent();
    if (keep.length) {
      keep.forEach(function (r) { r[8] = ''; });          // I열은 아래 배열수식이 계산
      db.getRange(2, 1, keep.length, 11).setValues(keep);
      db.getRange(2, 9, keep.length, 1).clearContent();  // I열은 완전히 비워야 배열수식이 퍼짐
    }
  }
  // 단가합계 = 자재비 + 인건비 + 경비 — 배열수식 한 칸으로 (빈 행을 만들지 않음)
  db.getRange('I2').setFormula('=ARRAYFORMULA(IF(C2:C="","",F2:F+G2:G+H2:H))');
  db.getRange('F2:J2000').setNumberFormat('#,##0');
  db.setColumnWidth(1, 70); db.setColumnWidth(2, 150); db.setColumnWidth(3, 240); db.setColumnWidth(4, 180);

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
      case 'parties':  return json({ ok: true, contractors: readParties('시공사'), clients: readParties('발주처') });
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
  // 열 이름은 자재비·인건비. 예전 시트(재료비·노무비)도 그대로 읽습니다
  const ALIAS = { '자재비': '재료비', '인건비': '노무비' };
  const at = name => { const i = head.indexOf(name); return i >= 0 ? i : (ALIAS[name] ? head.indexOf(ALIAS[name]) : -1); };
  const items = [];
  for (let r = 1; r < v.length; r++) {
    const row = v[r];
    const name = String(row[at('품목명')] || '').trim();
    if (!name) continue;
    items.push({
      id:   'r' + (r + 1),                       // 시트 행 번호 = 품목 ID
      use:  String(row[at('사용')] || '').trim().toUpperCase() !== 'N',   // N = 숨김 (견적 화면에서 안 보임)
      dae:  at('대공종') >= 0 ? String(row[at('대공종')] || '').trim() : '',
      gong: String(row[at('공종')] || '').trim(),
      name: name,
      spec: String(row[at('규격')] || '').trim(),
      unit: String(row[at('단위')] || '식').trim(),
      mat:  num(row[at('자재비')]),
      lab:  num(row[at('인건비')]),
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
  // 목록에는 '메타'(마지막 열, JSON)를 빼고 보냅니다 — 목록이 가벼워야 해서
  return sh.getRange(1, 1, last, SCHEMA['견적이력'].length - 1).getValues()
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

  let meta = {};
  try { meta = h && h[14] ? JSON.parse(h[14]) : {}; } catch (e) { meta = {}; }
  return { no: no, client: h ? h[2] : '', owner: h ? h[3] : '', staff: h ? h[4] : '',
           contractor: h ? h[12] : '', clientBiz: h ? h[13] : '', meta: meta, rows: rows };
}

/* ---------------- 거래처 (시공사 · 발주처) ---------------- */
const PARTY_KEYS = { '사업자등록번호':'bizno','상호':'name','성명':'ceo','사업장주소':'addr','업태':'uptae','종목':'jongmok',
                     '이메일':'email','담당자성명':'mgr','담당자연락처':'mgrTel','로고이미지':'logo','도장이미지':'stamp','FAX':'fax' };

function bizDigits(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }
function bizFormat(v) { const d = bizDigits(v); return d.length === 10 ? d.slice(0,3) + '-' + d.slice(3,5) + '-' + d.slice(5) : String(v || '').trim(); }

function readParties(sheetName) {
  const sh = SS().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const head = v[0].map(String);
  const out = [];
  for (let r = 1; r < v.length; r++) {
    const o = { id: 'r' + (r + 1) };
    head.forEach(function (hname, c) { const k = PARTY_KEYS[hname]; if (k) o[k] = String(v[r][c] == null ? '' : v[r][c]).trim(); });
    if (o.name || o.bizno) out.push(o);
  }
  return out;
}

/** 등록·수정. 같은 사업자등록번호가 있으면 그 행을 고치고, 없으면 새 행. (번호가 없으면 상호로 찾음) */
function saveParty(type, p) {
  const sheetName = type === 'contractor' ? '시공사' : '발주처';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SS().getSheetByName(sheetName);
    if (!sh) return { ok: false, error: sheetName + ' 시트가 없습니다. 초기화를 먼저 실행하세요.' };
    const head = SCHEMA[sheetName];
    if (!String(p.name || '').trim()) return { ok: false, error: '상호는 필수입니다.' };
    const biz = bizDigits(p.bizno);
    let row = 0;
    const last = sh.getLastRow();
    if (last > 1) {
      const v = sh.getRange(2, 1, last - 1, 2).getValues();
      for (let i = 0; i < v.length; i++) {
        const sameBiz  = biz && bizDigits(v[i][0]) === biz;
        const sameName = !biz && !bizDigits(v[i][0]) && String(v[i][1]).trim() === String(p.name).trim();
        if (sameBiz || sameName) { row = i + 2; break; }
      }
    }
    const inv = {}; Object.keys(PARTY_KEYS).forEach(function (k) { inv[PARTY_KEYS[k]] = k; });
    const vals = head.map(function (hname) {
      const key = PARTY_KEYS[hname];
      let val = p[key] == null ? '' : String(p[key]);
      if (key === 'bizno') val = bizFormat(val);
      if ((key === 'logo' || key === 'stamp') && val.length > 49000) val = '';   // 셀 한도(5만 자) 보호
      return val;
    });
    if (!row) row = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(row, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);   // 번호 앞자리 0 보존 위해 텍스트로
    return { ok: true, id: 'r' + row, created: row > last };
  } finally {
    lock.releaseLock();
  }
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
      case 'updateItem': return json(updateItem(body.item || {}));
      case 'saveParty':  return json(saveParty(body.type, body.party || {}));
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

    // 갑지 총계(경비·이윤·일반관리비·별도공사·절사 반영, 부가세 별도)가 오면 그걸 공급가액으로 기록합니다.
    // 없으면 예전처럼 품목 합계를 공급가액으로.
    const rate   = num(readSettings()['부가세율'] || 10) / 100;
    const supply = est.grandTotal != null && num(est.grandTotal) > 0 ? num(est.grandTotal) : sup;
    const vat    = Math.round(supply * rate);
    const total  = supply + vat;
    const net    = sup;          // 순공사비 (품목 합계) — 마진 계산은 이 기준
    sup = supply;

    // 메타: 시공사·발주처 선택, 표지 문구, 집계 비율, 제출 양식 등 견적서 머리 전체 (이미지는 넣지 않음)
    const metaJson = JSON.stringify(est.meta || {});
    hist.appendRow([no, new Date(), est.client || '', est.owner || '', est.staff || '',
                    sup, vat, total, costSum, sup - costSum, rows.length, est.memo || '',
                    est.contractor || '', est.clientBiz || '', metaJson.length < 45000 ? metaJson : '{}']);

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

/** 품목명(B열) 기준 마지막 데이터 행. getLastRow() 는 수식만 있는 빈 행도 세므로 쓰지 않는다. */
function lastDataRow(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const v = sh.getRange(2, 3, last - 1, 1).getValues();   // C열 품목명
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0] || '').trim()) return i + 2;
  return 1;
}

function addItems(items) {
  if (!items.length) return { ok: true, added: 0, skipped: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const db = SS().getSheetByName('단가DB');
    const have = {};
    const ld = lastDataRow(db);
    if (ld > 1) {
      db.getRange(2, 2, ld - 1, 3).getValues().forEach(function (r) {   // B 공종 · C 품목명 · D 규격
        if (String(r[1] || '').trim()) have[itemKey(r[0], r[1], r[2])] = true;
      });
    }
    const rows = [];
    let skipped = 0;
    items.forEach(function (i) {
      const k = itemKey(i.gong, i.name, i.spec);
      if (have[k]) { skipped++; return; }
      have[k] = true;
      rows.push([String(i.dae || ''), i.gong || '', i.name || '', i.spec || '', i.unit || '식',
                 num(i.mat), num(i.lab), num(i.exp), num(i.cost)]);
    });
    if (!rows.length) return { ok: true, added: 0, skipped: skipped };
    const start = lastDataRow(db) + 1;                 // 품목명 있는 마지막 행 바로 아래
    db.getRange(start, 1,  rows.length, 8).setValues(rows.map(function (r) { return r.slice(0, 8); }));  // A~H. I열은 배열수식
    db.getRange(start, 10, rows.length, 1).setValues(rows.map(function (r) { return [r[8]]; }));         // J 원가. K 사용은 비움
    return { ok: true, added: rows.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/** 화면 '단가 수정' 탭에서 한 행을 고칩니다. id = 'r<시트 행 번호>'. */
function updateItem(it) {
  const row = parseInt(String(it.id || '').replace(/^r/, ''), 10);
  if (!row || row < 2) return { ok: false, error: '잘못된 행 번호입니다.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const db = SS().getSheetByName('단가DB');
    // 안전장치: 화면에서 보던 품목명과 시트의 그 행이 같은지 확인 (행이 밀렸으면 엉뚱한 행을 덮지 않도록)
    const curName = String(db.getRange(row, 3).getValue() || '').trim();   // C열 품목명
    if (it.prevName != null && curName !== String(it.prevName).trim()) {
      return { ok: false, error: '시트가 그새 바뀌었습니다. ↻ 새로고침 후 다시 시도하세요.' };
    }
    db.getRange(row, 1, 1, 8).setValues([[String(it.dae || ''), it.gong || '', it.name || '', it.spec || '',
                                          it.unit || '식', num(it.mat), num(it.lab), num(it.exp)]]);   // A~H
    db.getRange(row, 10).setValue(num(it.cost));                  // J 원가
    db.getRange(row, 11).setValue(it.use === false ? 'N' : '');   // K 사용
    return { ok: true, id: it.id };
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
