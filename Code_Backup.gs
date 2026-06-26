// ════════════════════════════════════════════════════════════════
// ██ 자동 백업 시스템
// ════════════════════════════════════════════════════════════════

// ── 자동 백업 트리거 설정 (매일 자정 실행) ───────────────────────
// ※ 이 함수는 최초 1회만 실행하면 됩니다 (메뉴 → [백업] 자동 백업 스케줄 설정)
function setupDailyBackupTrigger() {
  const ui = SpreadsheetApp.getUi();

  // 기존 백업 트리거 중복 방지 (이미 설정된 경우 재설정 안 함)
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyBackup') {
      ui.alert(
        '✅ 자동 백업 이미 설정됨',
        '매일 자정 자동 백업이 이미 설정되어 있습니다.\n\n' +
        '중복 설정은 불필요합니다.',
        ui.ButtonSet.OK
      );
      return;
    }
  }

  // 매일 자정(0시~1시 사이) 실행 트리거 등록
  ScriptApp.newTrigger('runDailyBackup')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  ui.alert(
    '✅ 자동 백업 설정 완료',
    '매일 자정에 자동으로 백업이 실행됩니다.\n\n' +
    '백업 위치: 구글 드라이브 → [' + BACKUP_FOLDER_NAME + '] 폴더\n\n' +
    '※ 이 설정은 1회만 하면 됩니다.',
    ui.ButtonSet.OK
  );
}

// ── 브랜드가치 자동 기록 트리거 설정 ─────────────────────────────
function setupDailyTrackerTrigger() {
  // 기존 트리거 중복 방지
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'finalizeDailyTracker') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 매일 오후 11시에 자동 기록
  ScriptApp.newTrigger('finalizeDailyTracker')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
  SpreadsheetApp.getUi().alert('✅ 브랜드가치 자동 기록이 설정되었습니다.\n매일 오후 11시에 자동으로 기록됩니다.');
}

// ── 자동 백업 트리거 해제 (필요 시) ────────────────────────────
function removeDailyBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyBackup') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  SpreadsheetApp.getUi().alert(
    removed > 0
      ? '✅ 자동 백업 트리거가 해제되었습니다.'
      : '⚠️ 설정된 자동 백업 트리거가 없습니다.'
  );
}

// ── 트리거에 의해 매일 자동 실행되는 백업 함수 ──────────────────
function runDailyBackup() {
  _executeBackup('자동');
}

// ── 메뉴에서 수동으로 실행하는 즉시 백업 ────────────────────────
function runManualBackup() {
  _executeBackup('수동');
  SpreadsheetApp.getUi().alert(
    '✅ 백업 완료',
    '구글 드라이브 → [' + BACKUP_FOLDER_NAME + '] 폴더에 백업이 저장되었습니다.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ── 백업 실행 핵심 로직 ─────────────────────────────────────────
function _executeBackup(type) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const ssId     = ss.getId();
    const ssName   = ss.getName();
    const tz       = Session.getScriptTimeZone();
    const dateStr  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const timeStr  = Utilities.formatDate(new Date(), tz, 'HH:mm');
    const copyName = '[' + type + '백업] ' + ssName + ' (' + dateStr + ' ' + timeStr + ')';

    // 백업 폴더 찾기 또는 생성
    const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    const folder  = folders.hasNext()
      ? folders.next()
      : DriveApp.createFolder(BACKUP_FOLDER_NAME);

    // 스프레드시트 복사
    const copy = DriveApp.getFileById(ssId).makeCopy(copyName, folder);

    // 백업 로그 시트에 기록
    _recordBackupLog(ss, type, dateStr, timeStr, copyName, copy.getId());

    // 오래된 백업 자동 정리 (30개 초과 시 가장 오래된 것부터 삭제)
    _cleanOldBackups(folder, 30);

    Logger.log('[B.R.A.N.D 백업] ' + type + ' 백업 완료: ' + copyName);
  } catch (e) {
    Logger.log('[B.R.A.N.D 백업] 오류: ' + e.message);
    // 백업 실패해도 시스템 동작에는 영향 없음
  }
}

// ── 백업 로그 시트 기록 헬퍼 ────────────────────────────────────
function _recordBackupLog(ss, type, dateStr, timeStr, copyName, fileId) {
  try {
    let logSheet = ss.getSheetByName('백업로그');
    if (!logSheet) {
      logSheet = ss.insertSheet('백업로그');
      logSheet.appendRow(['날짜', '시각', '유형', '백업파일명', '파일ID']);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      logSheet.setColumnWidth(4, 300);
    }
    logSheet.appendRow([dateStr, timeStr, type, copyName, fileId]);
  } catch (e) {
    Logger.log('[B.R.A.N.D 백업] 로그 기록 오류: ' + e.message);
  }
}

// ── 오래된 백업 정리 헬퍼 (maxCount 초과분 삭제) ───────────────
function _cleanOldBackups(folder, maxCount) {
  try {
    const files = [];
    const iter  = folder.getFiles();
    while (iter.hasNext()) {
      const f = iter.next();
      files.push({ file: f, date: f.getDateCreated() });
    }
    // 오래된 순서로 정렬
    files.sort(function(a, b) { return a.date - b.date; });
    // maxCount 초과분 삭제
    if (files.length > maxCount) {
      const toDelete = files.slice(0, files.length - maxCount);
      for (let i = 0; i < toDelete.length; i++) {
        toDelete[i].file.setTrashed(true);
        Logger.log('[B.R.A.N.D 백업] 오래된 백업 삭제: ' + toDelete[i].file.getName());
      }
    }
  } catch (e) {
    Logger.log('[B.R.A.N.D 백업] 오래된 백업 정리 오류: ' + e.message);
  }
}

// ── 백업 목록 확인 (메뉴에서 호출) ─────────────────────────────
function showBackupList() {
  const ui = SpreadsheetApp.getUi();
  try {
    const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    if (!folders.hasNext()) {
      ui.alert('📋 백업 없음', '아직 백업이 없습니다.\n[백업] 지금 즉시 백업을 먼저 실행해주세요.', ui.ButtonSet.OK);
      return;
    }
    const folder = folders.next();
    const files  = [];
    const iter   = folder.getFiles();
    while (iter.hasNext()) {
      const f = iter.next();
      files.push(f.getName());
    }
    files.sort().reverse(); // 최신순 정렬
    const preview = files.slice(0, 10); // 최근 10개만 표시
    const msg = '총 ' + files.length + '개 백업 보관 중 (최근 10개)\n\n' + preview.join('\n');
    ui.alert('📋 백업 목록', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ 오류', '백업 목록 확인 중 오류: ' + e.message, ui.ButtonSet.OK);
  }
}


// ════════════════════════════════════════════════════════════════
// ██ 배포 안내 (초기 설정 가이드)
// ════════════════════════════════════════════════════════════════
// ※ [웹앱 URL 업데이트 방법]
//   1. GAS 편집기 상단 메뉴 → [배포] → [배포 관리]
//   2. 현재 배포 항목 옆 ✏️ 수정 아이콘 클릭
//   3. 버전을 "새 버전"으로 선택 후 [배포]
//   4. 표시되는 "웹 앱 URL"을 복사
//   5. 아래 url 변수의 따옴표 안에 붙여넣기 후 저장
//   → 이후 메뉴에서 [배포 URL 확인]을 실행하면 학생들에게 공유할 URL이 표시됩니다.

function showDeployInfo() {
  const ui  = SpreadsheetApp.getUi();
  const url = '여기에_새_URL_붙여넣기'; // ← GAS 배포 후 실제 URL로 교체하세요
  const msg = '✅ 현재 배포된 웹앱 URL\n\n' +
    url + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '학생 대시보드:\n' + url + '\n\n' +
    '선생님 경매 패널:\n' + url + '?page=admin\n\n' +
    '경매 실시간 중계:\n' + url + '?page=display\n\n' +
    '경제 수호대:\n' + url + '?page=guard\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '※ URL이 변경되지 않도록 "새 배포" 대신\n   "배포 관리 → 수정"으로 업데이트하세요.';
  ui.alert('🚀 B.R.A.N.D 웹앱 배포 정보', msg, ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// 업적검증 데이터 보강용 — 미확인 시트 구조 덤프
//
// 사용법
//   1) 이 함수를 아무 .gs 파일 맨 아래에 붙여넣기
//   2) 함수목록에서 dumpSheetsForVerify 선택 후 ▶ 실행
//   3) [실행 기록] 또는 [보기 > 로그]의 출력을 그대로 복사해서 전달
//
// 목적: 신용/경매/예금이자/벌점/상점구매 시트의 진짜 컬럼 구조를 확인하여,
//       추측 없이 정확한 정량 업적 자동집계를 구현하기 위함.
// ════════════════════════════════════════════════════════════════
function dumpSheetsForVerify() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];
  function log(s){ out.push(s); Logger.log(s); }

  var letters = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"];

  // 확인이 필요한 시트들
  var targets = [
    '학생별가입예금',   // 예금 이자/수익 컬럼 확인 (ECO-013/014/015)
    '예금상품',         // 이자율 등
    '신용점수이력',     // ECO-010
    '경매관리',         // ECO-007/016/017/023/025/027 (낙찰 기록 위치)
    '수호대적발로그',   // 벌점 (LIFE-006, CHAL-008)
    '상점_구매로그',    // ECO-032
    '대출현황',         // 대출 잔액/상태 (ECO-024/029/030)
    '히스토리'          // 비고(태그) 종류 확인 (ECO-009/003)
  ];

  targets.forEach(function(nm){
    log("\n══════════ [" + nm + "] ══════════");
    var sh = ss.getSheetByName(nm);
    if (!sh) { log("  ❌ 시트 없음"); return; }
    var data = sh.getDataRange().getValues();
    log("  데이터 행수: " + (data.length - 1) + " / 열수: " + (data.length ? data[0].length : 0));

    // 헤더
    if (data.length > 0) {
      log("  ── 헤더 ──");
      for (var c = 0; c < Math.min(data[0].length, 15); c++) {
        log("     " + letters[c] + "열: \"" + data[0][c] + "\"");
      }
    }
    // 샘플 3행
    log("  ── 샘플(최대 3행) ──");
    for (var r = 1; r < Math.min(data.length, 4); r++) {
      var parts = [];
      for (var cc = 0; cc < Math.min(data[r].length, 15); cc++) {
        var v = data[r][cc];
        if (v instanceof Date) v = "[Date]" + Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd HH:mm");
        v = String(v);
        if (v.length > 25) v = v.substring(0, 25) + "…";
        parts.push(letters[cc] + "=" + v);
      }
      log("     행" + (r+1) + ": " + parts.join(" | "));
    }
  });

  // 히스토리 비고(태그) 고유값 — ECO-009(하루 3경로) 판정 기준 파악용
  log("\n══════════ [히스토리 비고 태그 분석] ══════════");
  var hist = ss.getSheetByName('히스토리');
  if (hist) {
    var hData = hist.getDataRange().getValues();
    var tagCount = {};
    for (var i = 1; i < hData.length; i++) {
      var note = String(hData[i][7] || '').trim();   // H열: 비고
      // [xxx] 형태의 태그 추출
      var m = note.match(/\[([^\]]+)\]/);
      var key = m ? ('[' + m[1] + ']') : (note ? note.substring(0, 10) : '(빈칸)');
      tagCount[key] = (tagCount[key] || 0) + 1;
    }
    var keys = Object.keys(tagCount).sort(function(a,b){ return tagCount[b]-tagCount[a]; });
    keys.slice(0, 25).forEach(function(k){ log("   " + k + " : " + tagCount[k] + "건"); });
  }

  // 자산사용 사용항목(D열) 고유값 — 간식/상점/거래소/기부 구분 파악용
  log("\n══════════ [자산사용 D열(사용항목) 종류] ══════════");
  var spend = ss.getSheetByName('자산사용');
  if (spend) {
    var sData = spend.getDataRange().getValues();
    var catCount = {};
    for (var j = 1; j < sData.length; j++) {
      var cat = String(sData[j][3] || '').trim();   // D열
      catCount[cat] = (catCount[cat] || 0) + 1;
    }
    Object.keys(catCount).sort(function(a,b){ return catCount[b]-catCount[a]; })
      .slice(0, 25).forEach(function(k){ log("   \"" + k + "\" : " + catCount[k] + "건"); });
  }

  log("\n════════ 덤프 끝 ════════");
  return out.join("\n");
}

