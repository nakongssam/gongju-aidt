/**
 * 공주시 AIDT — 구글 시트 제출 수신 스크립트
 *
 * 설치 방법은 '사용안내.md'의 [구글 시트 제출 설정]을 참고하세요.
 *  1) 새 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 *  2) 아래 TEACHER_KEY를 선생님만 아는 비밀번호로 바꾸기
 *  3) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포
 *  4) 받은 웹 앱 URL을 aidt-config.js의 submitUrl에 붙여넣기
 */

// 교사용 대시보드에서 시트 데이터를 불러올 때 쓰는 비밀번호 (꼭 바꾸세요)
const TEACHER_KEY = '여기에-비밀번호-입력';

// 제출 마감: 시트 아래에 '제출마감'이라는 이름의 탭을 추가하면 제출이 막히고, 탭을 지우면 다시 열립니다.
const CLOSE_TAB = '제출마감';

const SHEET_SUMMARY = '제출기록';
const SHEET_RESP = '서술형응답';
const SUMMARY_HEAD = ['제출시각', '키', '학교', '학년', '반', '번호', '이름', '과목', '차시', '주제',
  '진행률(%)', '성취도(점)', '수준', '자기진단(%)', '학습시간(분)', '영상시청', '문항결과', '원본데이터'];
const RESP_HEAD = ['제출시각', '키', '학번', '이름', '과목', '차시', '페이지', '질문', '답변'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSheetByName(CLOSE_TAB)) return out_({ ok: false, error: 'closed' });
    const r = JSON.parse(e.postData.contents);
    if (r.format !== 'gongju-aidt-v1' || !r.lesson || !r.student) return out_({ ok: false, error: 'format' });
    const st = r.student, s = r.summary || {};
    const key = [r.lesson.id, st.school, st.grade, st.cls, st.num, st.name].join('|');
    const now = new Date();

    // 1) 제출기록: 학생·차시당 1행(다시 제출하면 덮어쓰기)
    const sh = sheet_(ss, SHEET_SUMMARY, SUMMARY_HEAD);
    const row = [now, key, st.school, st.grade, st.cls, st.num, st.name,
      r.lesson.subject, r.lesson.code, r.lesson.title,
      s.progress, s.score, s.level, s.self, Math.round((s.timeSec || 0) / 60),
      (s.videosWatched || 0) + '/' + (s.videosTotal || 0),
      (r.quiz || []).map(function (q) { return q.num + ':' + (!q.done ? '-' : q.firstCorrect ? 'O' : q.correct ? '△' : 'X'); }).join(' '),
      raw_(r)];
    const idx = findRow_(sh, key);
    if (idx > 0) sh.getRange(idx, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    // 2) 서술형응답: 해당 학생·차시의 이전 응답을 지우고 새로 기록
    const rs = sheet_(ss, SHEET_RESP, RESP_HEAD);
    const last = rs.getLastRow();
    if (last > 1) {
      const ks = rs.getRange(2, 2, last - 1, 1).getValues();
      for (let i = ks.length - 1; i >= 0; i--) if (ks[i][0] === key) rs.deleteRow(i + 2);
    }
    const sid = [st.grade, String(st.cls || '').padStart(2, '0'), String(st.num || '').padStart(2, '0')].join('');
    const rows = (r.responses || []).filter(function (x) { return x.answer; })
      .map(function (x) { return [now, key, sid, st.name, r.lesson.subject, r.lesson.code, x.page, x.prompt, x.answer]; });
    if (rows.length) rs.getRange(rs.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 교사용 대시보드가 ?key=비밀번호 로 호출하면 전체 제출 기록을 돌려줍니다.
function doGet(e) {
  if (!e.parameter.key || e.parameter.key !== TEACHER_KEY) return out_({ ok: false, error: 'key' });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUMMARY);
  if (!sh || sh.getLastRow() < 2) return out_({ ok: true, records: [] });
  const col = SUMMARY_HEAD.indexOf('원본데이터') + 1;
  const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  const records = [];
  vals.forEach(function (v) { if (v[0]) { try { records.push(JSON.parse(v[0])); } catch (x) { } } });
  return out_({ ok: true, records: records });
}

// 셀 한 칸 최대 50,000자 제한 대비
function raw_(r) {
  let t = JSON.stringify(r);
  if (t.length > 49000) { (r.responses || []).forEach(function (x) { x.answer = String(x.answer || '').slice(0, 400); }); r.timePerPage = []; t = JSON.stringify(r); }
  return t.slice(0, 49900);
}
function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#d1fae5');
    sh.setFrozenRows(1);
    if (name === SHEET_SUMMARY) { sh.hideColumns(head.indexOf('키') + 1); sh.hideColumns(head.indexOf('원본데이터') + 1); }
    if (name === SHEET_RESP) sh.hideColumns(head.indexOf('키') + 1);
  }
  return sh;
}
function findRow_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const f = sh.getRange(2, 2, n - 1, 1).createTextFinder(key).matchEntireCell(true).findNext();
  return f ? f.getRow() : -1;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
