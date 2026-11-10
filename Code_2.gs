/**
 * 기록 보관함 - Google 시트 저장용 Apps Script
 * 사용법: 시트에서 [확장 프로그램 > Apps Script] 에 붙여넣고 웹 앱으로 배포하세요.
 * (이전 버전을 쓰셨다면 'records' 탭을 삭제한 뒤 다시 배포하세요. 열 구성이 바뀌었습니다.)
 */
const NOTIFY_EMAIL = '';                                // 시작/종료 때 알림 메일 받을 주소 (비우면 안 보냄)
const SHEET_NAME = 'records';      // 원본 기록 (수정 금지)
const SUMMARY_NAME = '기록목록';   // 사람이 보기 쉬운 요약 (자동 갱신)
const FIELDS = ['seq','type','sid','at','by','who','topic','planned','text','prev'];
const HEADER = FIELDS.concat(['hash','srv','srvKST']);

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold').setBackground('#e8f0ff');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, HEADER.length).setNumberFormat('@');
  }
  return sh;
}
function canon_(e) { return JSON.stringify(FIELDS.map(k => e[k] === undefined || e[k] === null ? '' : e[k])); }
function sha_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function kst_(ms, f) { return ms ? Utilities.formatDate(new Date(Number(ms)), 'Asia/Seoul', f || 'yyyy-MM-dd HH:mm:ss') : ''; }

function rows_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEADER.length).getValues().map(r => ({
    seq: Number(r[0]), type: String(r[1]), sid: String(r[2]), at: Number(r[3]), by: String(r[4]),
    who: String(r[5]), topic: String(r[6]), planned: String(r[7]), text: String(r[8]),
    prev: String(r[9]), hash: String(r[10]), srv: Number(r[11])
  }));
}

/** 사람이 보기 쉬운 '기록목록' 시트를 다시 만든다 (원본 records는 건드리지 않음) */
function rebuildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = rows_(sheet_());
  const m = {};
  list.forEach(e => {
    const s = m[e.sid] || (m[e.sid] = { sid: e.sid, status: '', memos: [], by: {} });
    if (e.by) s.by[e.by] = 1;
    if (e.type === 'plan') { s.who = e.who; s.topic = e.topic; s.planned = e.planned; s.status = '예정'; }
    else if (e.type === 'start') { s.start = e.at; s.startSrv = e.srv; if (e.who) s.who = e.who; if (e.topic) s.topic = e.topic; if (e.planned) s.planned = e.planned; s.status = '진행 중'; }
    else if (e.type === 'end') { s.end = e.at; s.endSrv = e.srv; s.status = '완료'; }
    else if (e.type === 'cancel') { s.status = '예정 취소'; }
    else if (e.type === 'memo') { s.memos.push(kst_(e.at, 'HH:mm') + ' ' + e.by + ': ' + e.text); }
    else if (e.type === 'retitle') { s.topic = e.text; }
  });
  const rows = Object.keys(m).map(k => m[k]).sort((a, b) => (b.start || 0) - (a.start || 0))
    .map(s => [
      s.status, s.topic || '', s.who || '', Object.keys(s.by).join(', '),
      (s.planned || '').replace('T', ' '), kst_(s.start), kst_(s.startSrv), kst_(s.end), kst_(s.endSrv),
      s.start && s.end ? Math.round((s.end - s.start) / 60000) : '', s.memos.join('\n')
    ]);
  let sh = ss.getSheetByName(SUMMARY_NAME) || ss.insertSheet(SUMMARY_NAME);
  sh.clear();
  const head = ['상태', '제목', '대상', '입력자', '예정 시각', '시작(기기)', '시작(서버)', '종료(기기)', '종료(서버)', '소요(분)', '메모'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8f0ff');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, head.length - 1);
  sh.setColumnWidth(head.length, 320);
}

/** 원본 기록의 해시 연결을 검사한다 */
function verifyChain() {
  const list = rows_(sheet_());
  let prev = 'GENESIS', bad = [];
  list.forEach((e, i) => {
    if (e.seq !== i + 1) bad.push('#' + (i + 1) + ' 번호 불연속');
    if (e.prev !== prev) bad.push('#' + e.seq + ' 연결 불일치');
    if (sha_(canon_(e)) !== e.hash) bad.push('#' + e.seq + ' 내용 변조 의심');
    prev = e.hash;
  });
  SpreadsheetApp.getUi().alert(bad.length ? '이상 발견:\n' + bad.join('\n') : '정상: 기록 ' + list.length + '건 모두 일치');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('기록')
    .addItem('기록목록 새로고침', 'rebuildSummary')
    .addItem('무결성 검사', 'verifyChain')
    .addToUi();
}

function doGet() { return out_({ ok: true, msg: 'counsel-log alive' }); }

function doPost(ev) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(ev.postData.contents);
    const sh = sheet_();
    if (req.action === 'list') return out_({ ok: true, events: rows_(sh) });

    if (req.action === 'append') {
      const e = req.event || {};
      const list = rows_(sh);
      const last = list[list.length - 1];
      if (Number(e.seq) !== (last ? last.seq + 1 : 1) || e.prev !== (last ? last.hash : 'GENESIS'))
        return out_({ ok: false, code: 'SEQ', error: '다른 입력자가 먼저 기록했습니다. 동기화 후 다시 시도하세요.' });
      if (sha_(canon_(e)) !== e.hash) return out_({ ok: false, error: '해시가 맞지 않습니다' });

      const now = new Date();
      const row = FIELDS.map(k => e[k] === undefined || e[k] === null ? '' : e[k]);
      row.push(e.hash, now.getTime(), kst_(now.getTime()));
      sh.getRange(sh.getLastRow() + 1, 1, 1, HEADER.length).setNumberFormat('@').setValues([row.map(String)]);

      try { rebuildSummary(); } catch (x) {}
      if (NOTIFY_EMAIL && (e.type === 'start' || e.type === 'end')) {
        try {
          MailApp.sendEmail(NOTIFY_EMAIL,
            '[기록] ' + (e.type === 'start' ? '시작' : '종료') + ' ' + kst_(now.getTime(), 'MM-dd HH:mm:ss'),
            '입력자: ' + e.by + '\n기록 #' + e.seq + '\n서버 시각: ' + kst_(now.getTime()) + '\n해시: ' + e.hash);
        } catch (x) {}
      }
      return out_({ ok: true, srv: now.getTime() });
    }
    return out_({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
