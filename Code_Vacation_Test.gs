// ================================================================
// Code_Vacation_Test.gs — 방학 이벤트 사전 점검 도구
// ================================================================
// 방학이 시작되기 전에 실제로 작동하는지 확인하기 위한 파일입니다.
// 개학 후에는 이 파일을 지워도 되고, 그냥 둬도 아무 영향 없습니다.
//
// ★ 핵심 안전장치 ★
//   테스트 모드는 "지정한 학생 1명"에게만 적용됩니다.
//   다른 학생이 그 사이에 로그인해도 진짜 날짜로 처리되므로,
//   방학 기간 전이면 그 학생들에게는 여전히 아무 일도 일어나지 않습니다.
// ================================================================

// ▼▼▼ 테스트에 사용할 학생 이름을 여기에 적으세요 ▼▼▼
const VAC_TEST_STUDENT = '테스트요원';
// 실제 학생 이름을 쓰면 그 학생 자산이 늘어납니다.
// vacTestReset() 이 전부 되돌려주지만, 아래 vacTestCreateDummy() 로
// 가짜 학생을 만들어 쓰는 쪽이 가장 안전합니다.

// ════════════════════════════════════════════════════════════════
// 【1단계】 가짜 학생 만들기 (권장)
// ════════════════════════════════════════════════════════════════
function vacTestCreateDummy() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(SHEET_MAIN);
  const data = main.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NAME - 1]).trim() === VAC_TEST_STUDENT) {
      return 'ℹ️ 이미 있습니다: ' + VAC_TEST_STUDENT + ' (' + (i + 1) + '행)';
    }
  }
  const row = new Array(main.getLastColumn()).fill('');
  row[COL_BRAND - 1]    = '테스트';
  row[COL_NAME - 1]     = VAC_TEST_STUDENT;
  row[COL_VALUE - 1]    = 30000;   // 브랜드가치 (변하지 않는지 확인용)
  row[COL_ASSET - 1]    = 0;       // 자산 (골드가 쌓이는지 확인용)
  row[COL_PASSWORD - 1] = '1234';  // 로그인용 비밀번호
  main.appendRow(row);

  // 방학_출석 시트에도 등록
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  if (att) {
    const aData = att.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < aData.length; i++) {
      if (String(aData[i][VC.NAME - 1]).trim() === VAC_TEST_STUDENT) { found = true; break; }
    }
    if (!found) {
      att.appendRow([VAC_TEST_STUDENT, '', 0, 0, VAC_CFG.RESTORE_TICKETS, 0, 0, '', '', '', 0, 0, '', '']);
    }
  }
  return '✅ 테스트 학생 생성 완료\n  이름: ' + VAC_TEST_STUDENT + '\n  비밀번호: 1234\n' +
         '  브랜드가치 30,000 / 자산 0 에서 시작합니다.';
}

// ════════════════════════════════════════════════════════════════
// 【2단계】 테스트 모드 켜기 — 첫날(7/27)로 설정
// ════════════════════════════════════════════════════════════════
function vacTestOn() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('VAC_TEST_NAME', VAC_TEST_STUDENT);
  props.setProperty('VAC_TEST_DATE', VAC_CFG.START);
  return '🧪 테스트 모드 ON\n' +
         '  대상: ' + VAC_TEST_STUDENT + ' (이 학생에게만 적용)\n' +
         '  가짜 날짜: ' + VAC_CFG.START + '\n\n' +
         '이제 웹앱에서 ' + VAC_TEST_STUDENT + ' 으로 로그인해 보세요.';
}

// 날짜를 하루 넘기기 — 브라우저 새로고침 전에 이걸 누르면 다음 날이 됩니다
function vacTestNextDay() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty('VAC_TEST_DATE') || VAC_CFG.START;
  const next = _vacAddDays_(cur, 1);
  props.setProperty('VAC_TEST_DATE', next);
  return '📅 ' + cur + ' → ' + next + '\n웹앱을 새로고침하면 다음 별이 켜집니다.';
}

// 하루 건너뛰기 (연속 끊김 + 복구권 테스트용)
function vacTestSkipDay() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty('VAC_TEST_DATE') || VAC_CFG.START;
  const next = _vacAddDays_(cur, 2);
  props.setProperty('VAC_TEST_DATE', next);
  return '⏭️ ' + cur + ' → ' + next + ' (하루 건너뜀)\n' +
         '새로고침하면 연속이 끊기고, 복구 버튼이 나타나야 정상입니다.';
}

// 특정 날짜로 점프 (예: 광복절 확인)
function vacTestGoToLiberationDay() {
  PropertiesService.getScriptProperties().setProperty('VAC_TEST_DATE', VAC_CFG.LIBERATION_DATE);
  return '🕊️ ' + VAC_CFG.LIBERATION_DATE + ' 로 이동\n새로고침하면 한 줄 소감 입력창이 떠야 정상입니다.';
}
function vacTestGoToMakeup() {
  PropertiesService.getScriptProperties().setProperty('VAC_TEST_DATE', VAC_CFG.MAKEUP_START);
  return '🔁 보충 기간(' + VAC_CFG.MAKEUP_START + ')으로 이동\n' +
         '새로고침을 3번 하면 3칸까지만 열리고 멈춰야 정상입니다.';
}

// 현재 테스트 상태 확인
function vacTestStatus() {
  const props = PropertiesService.getScriptProperties();
  const name = props.getProperty('VAC_TEST_NAME');
  if (!name) return '⚪ 테스트 모드 OFF (모든 학생이 진짜 날짜로 동작)';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const att = ss.getSheetByName(SHEET_VAC_ATT).getDataRange().getValues();
  let line = null;
  for (let i = 1; i < att.length; i++) {
    if (String(att[i][VC.NAME - 1]).trim() === name) { line = att[i]; break; }
  }
  let classTotal = 0;
  for (let i = 1; i < att.length; i++) classTotal += Number(att[i][VC.CELLS - 1]) || 0;

  return '🧪 테스트 모드 ON\n' +
    '  대상    : ' + name + '\n' +
    '  가짜날짜: ' + props.getProperty('VAC_TEST_DATE') + '\n' +
    '  진짜날짜: ' + _todayStr() + '\n' +
    (line ? ('  ─────────────\n' +
      '  누적칸수: ' + line[VC.CELLS - 1] + ' / ' + VAC_CFG.TOTAL_CELLS + '\n' +
      '  연속일수: ' + line[VC.STREAK - 1] + '  (최고 ' + line[VC.BEST - 1] + ')\n' +
      '  토큰    : ' + line[VC.TOKENS - 1] + '\n' +
      '  복구권  : ' + line[VC.TICKETS - 1] + '\n' +
      '  보물기록: ' + (line[VC.TREASURE - 1] || '-') + '\n' +
      '  보너스  : ' + (line[VC.BONUS_LOG - 1] || '-') + '\n') : '') +
    '  학급합계: ' + classTotal + '회';
}

// ════════════════════════════════════════════════════════════════
// 【3단계】 자동 시뮬레이션 — 브라우저 없이 28일을 한 번에 돌려봄
//   실행 후 [실행 기록] 또는 아래 반환값에서 결과를 확인하세요.
// ════════════════════════════════════════════════════════════════
function vacTestSimulate28Days() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('VAC_TEST_NAME', VAC_TEST_STUDENT);

  const log = [];
  let date = VAC_CFG.START;

  for (let day = 1; day <= 30 && date <= VAC_CFG.END; day++) {
    // 12일째와 13일째는 일부러 건너뛰어 연속 끊김·복구를 확인
    if (day === 12) {
      log.push(date + ' : (일부러 결석)');
      date = _vacAddDays_(date, 1);
      continue;
    }
    props.setProperty('VAC_TEST_DATE', date);
    const r = vacationOnLogin(VAC_TEST_STUDENT);

    if (!r) { log.push(date + ' : (기간 밖)'); }
    else if (r.error) { log.push(date + ' : ⚠️ 오류 발생'); }
    else {
      let line = date + ' : ' + String(r.cells).padStart(2) + '칸 · 연속' +
                 String(r.streak).padStart(2) + '일 · 토큰' + String(r.tokens).padStart(3);
      if (r.restoreAvailable) {
        const rr = useVacationRestore(VAC_TEST_STUDENT);
        line += rr.success ? '  → 🌠복구 성공(연속 ' + rr.streak + '일)' : '  → 복구실패:' + rr.msg;
      }
      if (r.messages && r.messages.length) line += '  ' + r.messages.join(' ');
      log.push(line);
    }
    date = _vacAddDays_(date, 1);
  }

  // 공동 목표 확인까지 실행
  checkVacationClassMilestones();
  log.push('─────────────────────────');
  log.push(vacTestStatus());

  const out = log.join('\n');
  Logger.log(out);
  return out;
}

// ════════════════════════════════════════════════════════════════
// 【4단계】 원복 — 테스트 흔적을 전부 지웁니다 (반드시 실행!)
// ════════════════════════════════════════════════════════════════
function vacTestReset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const name = VAC_TEST_STUDENT;
  const report = [];

  // ① 테스트 모드 끄기
  props.deleteProperty('VAC_TEST_NAME');
  props.deleteProperty('VAC_TEST_DATE');
  props.deleteProperty('VAC_MILESTONE_GRANTED');
  report.push('✅ 테스트 모드 OFF · 공동목표 지급기록 초기화');

  // ② 방학_출석 행 초기화
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  if (att) {
    const d = att.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][VC.NAME - 1]).trim() === name) {
        att.getRange(i + 1, 1, 1, 14).setValues([[
          name, '', 0, 0, VAC_CFG.RESTORE_TICKETS, 0, 0, '', '', '', 0, 0, '', ''
        ]]);
        report.push('✅ 방학_출석 초기화 (' + (i + 1) + '행)');
        break;
      }
    }
  }

  // ③ 방학 히스토리에서 테스트 기록 삭제 (아래에서 위로 지워야 행이 안 밀림)
  const vh = ss.getSheetByName(SHEET_VAC_HIST);
  if (vh && vh.getLastRow() >= 2) {
    const d = vh.getDataRange().getValues();
    let cnt = 0;
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1]).trim() === name) { vh.deleteRow(i + 1); cnt++; }
    }
    report.push('✅ 방학_출석_히스토리 ' + cnt + '행 삭제');
  }

  // ④ 메인 히스토리에서 [방학...] 기록 삭제 + 지급된 골드 회수
  const hist = ss.getSheetByName(SHEET_HISTORY);
  let goldBack = 0;
  if (hist && hist.getLastRow() >= 2) {
    const d = hist.getDataRange().getValues();
    let cnt = 0;
    for (let i = d.length - 1; i >= 1; i--) {
      const note = String(d[i][7] || '');
      if (String(d[i][1]).trim() === name &&
          (note.indexOf('[방학') === 0 || note.indexOf('[공동목표]') === 0)) {
        goldBack += Number(d[i][4]) || 0;
        hist.deleteRow(i + 1);
        cnt++;
      }
    }
    report.push('✅ 메인 히스토리 ' + cnt + '행 삭제 (회수 골드 ' + goldBack + ')');
  }

  // ⑤ 메인 시트 자산 원복
  const main = ss.getSheetByName(SHEET_MAIN);
  const md = main.getDataRange().getValues();
  for (let i = 1; i < md.length; i++) {
    if (String(md[i][COL_NAME - 1]).trim() === name) {
      const cur = Number(md[i][COL_ASSET - 1]) || 0;
      main.getRange(i + 1, COL_ASSET).setValue(Math.max(0, cur - goldBack));
      report.push('✅ 자산 원복: ' + cur + ' → ' + Math.max(0, cur - goldBack));
      break;
    }
  }

  // ⑥ 제안함에서 테스트 제안 삭제
  const pr = ss.getSheetByName(SHEET_VAC_PROP);
  if (pr && pr.getLastRow() >= 2) {
    const d = pr.getDataRange().getValues();
    let cnt = 0;
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1]).trim() === name) { pr.deleteRow(i + 1); cnt++; }
    }
    if (cnt) report.push('✅ 제안함 ' + cnt + '행 삭제');
  }

  updateRankings();
  report.push('\n🎉 원복 완료. 이제 실제 방학 시작을 기다리면 됩니다.');
  return report.join('\n');
}

// ════════════════════════════════════════════════════════════════
// 【복구】 이미 만든 시트에 날짜 자동변환 버그가 있었다면 — 1회 실행
// ════════════════════════════════════════════════════════════════
// 증상: 연속일수가 1에서 안 올라감 (구글시트가 "2026-07-27" 문자열을
// 자기 맘대로 날짜 타입으로 바꿔버려서 다음 날 비교가 항상 어긋남)
// 이 함수는: ① 날짜 열을 텍스트 서식으로 고정 ② 이미 날짜 타입으로
// 바뀐 값들을 전부 원래 문자열로 되돌립니다. 딱 1번만 실행하면 됩니다.
function vacFixDateFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = [];

  const att = ss.getSheetByName(SHEET_VAC_ATT);
  if (att) {
    const lastRow = Math.max(att.getLastRow(), 2);
    [VC.LAST_DATE, VC.MAKEUP_DATE, VC.RESTORE_DATE].forEach(function(col) {
      att.getRange(2, col, Math.max(lastRow - 1, 500), 1).setNumberFormat('@');
    });
    const data = att.getDataRange().getValues();
    let fixed = 0;
    for (let i = 1; i < data.length; i++) {
      [VC.LAST_DATE, VC.MAKEUP_DATE, VC.RESTORE_DATE].forEach(function(col) {
        const v = data[i][col - 1];
        if (v instanceof Date) {
          att.getRange(i + 1, col).setValue(_vacDateKey_(v));
          fixed++;
        }
      });
    }
    report.push('✅ 방학_출석: 날짜열 텍스트 서식 고정, 오염된 값 ' + fixed + '개 복구');
  } else {
    report.push('⚠️ 방학_출석 시트가 없습니다. setupVacationSheets 를 먼저 실행하세요.');
  }

  const hist = ss.getSheetByName(SHEET_VAC_HIST);
  if (hist && hist.getLastRow() >= 1) {
    hist.getRange(2, 1, Math.max(hist.getLastRow(), 500), 1).setNumberFormat('@');
    report.push('✅ 방학_출석_히스토리: 날짜열 텍스트 서식 고정');
  }

  const prop = ss.getSheetByName(SHEET_VAC_PROP);
  if (prop && prop.getLastRow() >= 1) {
    prop.getRange(2, 1, Math.max(prop.getLastRow(), 500), 1).setNumberFormat('@');
    report.push('✅ 방학_제안함: 날짜열 텍스트 서식 고정');
  }

  return report.join('\n');
}


function vacTestDeleteDummy() {
  const msg = vacTestReset();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  [SHEET_MAIN, SHEET_VAC_ATT].forEach(function(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const col = (sheetName === SHEET_MAIN) ? COL_NAME : VC.NAME;
    const d = sh.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) {
      if (String(d[i][col - 1]).trim() === VAC_TEST_STUDENT) sh.deleteRow(i + 1);
    }
  });
  updateRankings();
  return msg + '\n🗑️ 테스트 학생(' + VAC_TEST_STUDENT + ') 행 삭제 완료';
}
