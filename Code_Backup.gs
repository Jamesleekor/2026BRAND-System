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
function showDeployInfo() {
  const ui  = SpreadsheetApp.getUi();
  const url = '여기에_새_URL_붙여넣기';
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

