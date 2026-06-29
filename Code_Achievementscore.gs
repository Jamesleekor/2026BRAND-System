/*************************************************************************
 * [업적 점수 집계 모듈]  Code_AchievementScore.gs
 * 목적: "직전 MVP 선정 직후 ~ 다음 측정 전날" 기간에 새로 달성한 업적만
 *       점수로 합산 (누적 개수가 아니라 등급별 점수 기준)
 *
 * 설계 원칙
 *  - 로그인(doGet)은 절대 건드리지 않음. 무거운 집계는 교사가 버튼/메뉴로 실행.
 *  - 점수(G열)는 교사 전용. 학생 화면(도감)에는 절대 내려가지 않게 할 것.
 *  - 결과는 교사 전용 시트 '업적점수집계'에만 기록.
 *
 * 사용 순서 (처음 1번)
 *  1) importAchievementScores()        → 업적마스터 G열에 점수 채우기
 *  2) setupScoreSheet()                → 집계 시트 + 날짜 입력칸 만들기
 * 매달 측정할 때
 *  3) 업적점수집계 시트 B1=시작일, B2=종료일 입력 후
 *     calcAchievementScoresFromCells() 실행
 *************************************************************************/

// ===== 시트/컬럼 설정 (실제 구조에 맞춤) =====
var ASC_MASTER_SHEET   = '업적마스터';
var ASC_MASTER_ID_COL  = 1;   // A열: 업적ID
var ASC_MASTER_PT_COL  = 7;   // G열: 점수 (신설)

var ASC_ACHIEVE_SHEET  = '학생업적달성';
var ASC_ACH_NAME_COL   = 1;   // A열: 학생성명
var ASC_ACH_ID_COL     = 2;   // B열: 업적ID
var ASC_ACH_DATE_COL   = 5;   // E열: 달성날짜  ← 과거 디버깅으로 확인된 위치

var ASC_RESULT_SHEET   = '업적점수집계';   // 교사 전용 출력

// ===== v5 점수표 (업적ID: 점수) — 처음 G열 채울 때만 사용 =====
var ACHIEVEMENT_SCORE_MAP = {
// 총 128 개
    'START-001':4, 'START-002':4, 'START-003':5, 'ECO-001':5, 'ECO-002':26, 'ECO-003':13,
    'ECO-004':11, 'ECO-005':9, 'ECO-006':7, 'ECO-007':4, 'ECO-008':5, 'ECO-009':9,
    'ECO-010':13, 'ECO-011':9, 'ECO-012':11, 'ECO-013':5, 'ECO-014':18, 'ECO-015':110,
    'ECO-016':11, 'ECO-017':34, 'ECO-018':9, 'ECO-020':5, 'ECO-021':4, 'ECO-022':7,
    'ECO-023':11, 'ECO-024':7, 'ECO-025':4, 'ECO-026':7, 'ECO-027':22, 'ECO-028':18,
    'ECO-029':13, 'ECO-030':34, 'ECO-031':11, 'ECO-032':9, 'ECO-033':5, 'ECO-034':26,
    'ECO-035':22, 'ECO-036':26, 'ECO-037':26, 'STORY-001':18, 'LIFE-002':7, 'LIFE-003':9,
    'LIFE-005':26, 'LIFE-006':7, 'LIFE-007':26, 'LIFE-008':5, 'LIFE-009':5, 'LIFE-010':9,
    'LIFE-011':7, 'LIFE-012':11, 'LIFE-013':5, 'LIFE-014':13, 'LIFE-015':5, 'LIFE-016':7,
    'LIFE-017':145, 'LIFE-018':5, 'LIFE-019':155, 'LIFE-020':26, 'LIFE-021':18, 'ART-001':22,
    'ART-002':9, 'MVP-001':7, 'MVP-002':34, 'RANK-001':60, 'RANK-002':60, 'RANK-003':60,
    'RANK-004':60, 'RANK-005':72, 'RANK-006':90, 'RANK-007':100, 'GUILD-001':4, 'GUILD-002':4,
    'GUILD-003':5, 'GUILD-004':7, 'GUILD-005':7, 'GUILD-006':7, 'GUILD-007':11, 'GUILD-008':22,
    'GUILD-009':22, 'GUILD-010':11, 'GUILD-011':13, 'GUILD-012':34, 'GUILD-013':72, 'GUILD-014':160,
    'STU-001':11, 'STU-002':11, 'STU-003':11, 'STU-005':11, 'STU-006':4, 'STU-007':11,
    'STU-008':34, 'STU-009':7, 'STU-010':5, 'STU-011':9, 'STU-012':11, 'STU-013':9,
    'STU-014':22, 'TEAM-001':13, 'TEAM-002':11, 'TEAM-003':11, 'TEAM-004':11, 'CONS-001':13,
    'CONS-002':5, 'ACH-001':4, 'ACH-002':4, 'CHAL-001':13, 'CHAL-002':22, 'CHAL-003':26,
    'CHAL-004':34, 'CHAL-005':11, 'CHAL-006':4, 'CHAL-007':34, 'CHAL-008':100, 'CHAL-009':165,
    'CHAL-010':300, 'CHAL-011':18, 'CHAL-012':72, 'CHAL-013':105, 'CHAL-014':105, 'HID-001':11,
    'HID-002':5, 'HID-003':5, 'HID-004':11, 'HID-005':200, 'HID-006':138, 'HID-007':5,
    'HID-008':5, 'CHAL-015':100,
};

/**
 * [1단계] 업적마스터 G열에 점수를 채워넣음 (업적ID로 매칭 → 순서 바뀌어도 안전)
 * 한 번만 실행. 이후 점수 조정은 G열에서 직접 수정하면 됨.
 */
function importAchievementScores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASC_MASTER_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('업적마스터 시트를 찾을 수 없습니다.'); return; }

  var data = sheet.getDataRange().getValues();
  // 헤더(G1)에 '점수' 표시
  sheet.getRange(1, ASC_MASTER_PT_COL).setValue('점수');

  var updated = 0, missing = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][ASC_MASTER_ID_COL - 1]).trim();
    if (!id) continue;
    if (ACHIEVEMENT_SCORE_MAP.hasOwnProperty(id)) {
      sheet.getRange(i + 1, ASC_MASTER_PT_COL).setValue(ACHIEVEMENT_SCORE_MAP[id]);
      updated++;
    } else {
      missing.push(id);   // 점수표에 없는 업적
    }
  }
  var msg = '점수 입력 완료: ' + updated + '개';
  if (missing.length) msg += '\n\n점수표에 없어 비워둔 업적: ' + missing.join(', ');
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * [핵심] 기간을 받아 학생별 점수를 집계해서 반환
 * @param {Date} startDate 시작일 (그날 00:00부터 포함)
 * @param {Date} endDate   종료일 (그날 23:59까지 포함)
 * @return {Object} { period:{이름:점수}, count:{이름:개수}, cumulative:{이름:누적점수} }
 */
function calcAchievementScores(startDate, endDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) 업적마스터 G열에서 {업적ID: 점수} 맵 만들기
  var masterData = ss.getSheetByName(ASC_MASTER_SHEET).getDataRange().getValues();
  var scoreMap = {};
  for (var i = 1; i < masterData.length; i++) {
    var id = String(masterData[i][ASC_MASTER_ID_COL - 1]).trim();
    var pt = Number(masterData[i][ASC_MASTER_PT_COL - 1]) || 0;
    if (id) scoreMap[id] = pt;
  }

  // 2) 기간 경계 정리
  var start = new Date(startDate); start.setHours(0, 0, 0, 0);
  var end   = new Date(endDate);   end.setHours(23, 59, 59, 999);

  // 3) 학생업적달성 훑으며 합산
  var achieveData = ss.getSheetByName(ASC_ACHIEVE_SHEET).getDataRange().getValues();
  var period = {}, count = {}, cumulative = {};

  for (var j = 1; j < achieveData.length; j++) {
    var name  = String(achieveData[j][ASC_ACH_NAME_COL - 1]).trim();
    var achId = String(achieveData[j][ASC_ACH_ID_COL   - 1]).trim();
    var raw   = achieveData[j][ASC_ACH_DATE_COL - 1];
    if (!name || !achId) continue;

    var pt = scoreMap[achId] || 0;          // 점수 없는 업적은 0점
    cumulative[name] = (cumulative[name] || 0) + pt;   // 누적(명예의전당용)

    if (raw) {
      var d = new Date(raw);
      if (!isNaN(d.getTime()) && d >= start && d <= end) {   // 기간 내 신규만
        period[name] = (period[name] || 0) + pt;
        count[name]  = (count[name]  || 0) + 1;
      }
    }
  }
  return { period: period, count: count, cumulative: cumulative };
}

/**
 * [준비] 집계 시트 + 날짜 입력칸 생성 (한 번만)
 */
function setupScoreSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ASC_RESULT_SHEET) || ss.insertSheet(ASC_RESULT_SHEET);
  sh.getRange('A1').setValue('측정 시작일').setFontWeight('bold');
  sh.getRange('A2').setValue('측정 종료일').setFontWeight('bold');
  sh.getRange('B1:B2').setNumberFormat('yyyy-mm-dd');
  sh.getRange('A4').setValue('▼ calcAchievementScoresFromCells() 실행 시 아래에 결과가 채워집니다');
  SpreadsheetApp.getUi().alert("'업적점수집계' 시트 준비 완료.\nB1에 시작일, B2에 종료일을 입력한 뒤 집계를 실행하세요.");
}

/**
 * [실행] 입력칸(B1,B2)의 날짜를 읽어 집계 → 시트에 표로 출력
 */
function calcAchievementScoresFromCells() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ASC_RESULT_SHEET);
  if (!sh) { setupScoreSheet(); return; }

  var s = sh.getRange('B1').getValue();
  var e = sh.getRange('B2').getValue();
  if (!(s instanceof Date) || !(e instanceof Date)) {
    SpreadsheetApp.getUi().alert('B1(시작일), B2(종료일)에 날짜를 정확히 입력해주세요.');
    return;
  }

  var r = calcAchievementScores(s, e);

  // 기간내 점수 내림차순 정렬
  var names = Object.keys(r.cumulative);
  names.sort(function(a, b) { return (r.period[b] || 0) - (r.period[a] || 0); });

  var out = [['순위', '학생명', '기간내 점수', '기간내 달성수', '누적 점수(명예의전당)']];
  names.forEach(function(name, idx) {
    out.push([idx + 1, name, r.period[name] || 0, r.count[name] || 0, r.cumulative[name] || 0]);
  });

  // 기존 결과 영역 비우고 다시 쓰기 (A5부터)
  sh.getRange(5, 1, Math.max(sh.getMaxRows() - 4, 1), 5).clearContent();
  sh.getRange(5, 1, out.length, 5).setValues(out);
  sh.getRange(5, 1, 1, 5).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('집계 완료: ' + names.length + '명');
}

/* ───────────────────────────────────────────────
 * (선택) 교사 메뉴에 버튼 추가하고 싶으면, 기존 onOpen() 안에
 * 메뉴 .addItem 줄 옆에 아래 한 줄을 추가하세요:
 *
 *   .addItem('업적 점수 집계', 'calcAchievementScoresFromCells')
 *
 * onOpen 전체를 새로 만들지 말고, 있는 메뉴에 한 줄만 끼워넣으면 됩니다.
 * ─────────────────────────────────────────────── */
