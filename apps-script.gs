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
    let r = JSON.parse(e.postData.contents);
    if (r.format !== 'gongju-aidt-v1' || !r.lesson || !r.student) return out_({ ok: false, error: 'format' });
    const st = r.student;
    const key = [r.lesson.id, st.school, st.grade, st.cls, st.num, st.name].join('|');
    const now = new Date();

    // 1) 제출기록: 학생·차시당 1행. 다시 제출하거나 다른 기기에서 제출하면 이전 기록과 합칩니다.
    const sh = sheet_(ss, SHEET_SUMMARY, SUMMARY_HEAD);
    const idx = findRow_(sh, key);
    if (idx > 0) {
      let old = null;
      try { old = JSON.parse(sh.getRange(idx, SUMMARY_HEAD.indexOf('원본데이터') + 1).getValue()); } catch (x) { }
      if (old) r = merge_(old, r);
    }
    const s = r.summary || {};
    const row = [now, key, st.school, st.grade, st.cls, st.num, st.name,
      r.lesson.subject, r.lesson.code, r.lesson.title,
      s.progress, s.score, s.level, s.self, Math.round((s.timeSec || 0) / 60),
      (s.videosWatched || 0) + '/' + (s.videosTotal || 0),
      (r.quiz || []).map(function (q) { return q.num + ':' + (!q.done ? '-' : q.firstCorrect ? 'O' : q.correct ? '△' : 'X'); }).join(' '),
      raw_(r)];
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

    return out_({ ok: true, merged: idx > 0, summary: s });
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

// 이전 제출(a)과 새 제출(b)을 합치기
//  - 퀴즈: 먼저 완료한 결과를 유지(다른 기기에서 다시 풀어도 처음 결과가 기록됨)
//  - 서술형·자기진단: 새 제출이 비어 있으면 이전 내용 유지
//  - 영상: 어느 기기에서든 본 적 있으면 시청
//  - 학습 시간: 페이지별로 더 긴 값
function merge_(a, b) {
  const r = JSON.parse(JSON.stringify(b));
  const byKey = function (arr, k) { const m = {}; (arr || []).forEach(function (x) { m[x[k]] = x; }); return m; };
  const aq = byKey(a.quiz, 'num');
  r.quiz = (b.quiz || []).map(function (q) { const o = aq[q.num]; return o && o.done ? o : q; });
  const ar = byKey(a.responses, 'key');
  r.responses = (b.responses || []).map(function (x) { const o = ar[x.key]; return !x.answer && o && o.answer ? o : x; });
  const as = byKey(a.selfAssessment, 'item');
  r.selfAssessment = (b.selfAssessment || []).map(function (x) { const o = as[x.item]; return x.score == null && o && o.score != null ? o : x; });
  const av = byKey(a.videos, 'id');
  r.videos = (b.videos || []).map(function (v) { return { n: v.n, id: v.id, watched: !!(v.watched || (av[v.id] && av[v.id].watched)) }; });
  const at = byKey(a.timePerPage, 'page');
  r.timePerPage = (b.timePerPage || []).map(function (t) { return { page: t.page, sec: Math.max(t.sec || 0, at[t.page] ? at[t.page].sec || 0 : 0) }; });
  r.summary = summary_(r);
  return r;
}
function summary_(r) {
  const Q = r.quiz || [], R = r.responses || [], A = r.selfAssessment || [], V = r.videos || [];
  const done = Q.filter(function (q) { return q.done; }).length + R.filter(function (x) { return String(x.answer || '').trim(); }).length +
    A.filter(function (x) { return x.score != null; }).length + V.filter(function (v) { return v.watched; }).length;
  const total = Q.length + R.length + A.length + V.length;
  let answered = 0, first = 0, second = 0, wrong = 0;
  Q.forEach(function (q) { if (!q.done) return; answered++; if (q.firstCorrect) first++; else if (q.correct) second++; else wrong++; });
  const score = answered ? Math.round((first + second * 0.5) / answered * 100) : null;
  const prov = score == null ? '-' : score >= 80 ? '상' : score >= 50 ? '중' : '하';
  const rated = A.filter(function (x) { return x.score != null; });
  return {
    progress: total ? Math.round(done / total * 100) : 0, done: done, total: total, quizTotal: Q.length, answered: answered, prov: prov,
    first: first, second: second, wrong: wrong, score: score, level: answered === 0 ? '-' : answered < Q.length ? '미완료' : prov,
    self: rated.length ? Math.round(rated.reduce(function (s, x) { return s + x.score; }, 0) / (rated.length * 2) * 100) : null,
    timeSec: (r.timePerPage || []).reduce(function (s, t) { return s + (t.sec || 0); }, 0),
    videosWatched: V.filter(function (v) { return v.watched; }).length, videosTotal: V.length
  };
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
