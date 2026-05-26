// ════════════════════════════════════════════════════════════════
// 과제 제출 시스템  (Code_Assignment.gs)
// ════════════════════════════════════════════════════════════════
// 상수: SHEET_ASSIGN_LIST, SHEET_ASSIGN_SUBMIT, ASSIGN_DRIVE_FOLDER → Code.gs
// ────────────────────────────────────────────────────────────────

// ── 시트 자동 생성 ────────────────────────────────────────────────
function _ensureAssignmentSheets(ss) {
  let listSheet = ss.getSheetByName(SHEET_ASSIGN_LIST);
  if (!listSheet) {
    listSheet = ss.insertSheet(SHEET_ASSIGN_LIST);
    listSheet.appendRow(['과제ID','과제명','설명','마감일시','허용타입','상태']);
    listSheet.getRange(1,1,1,6).setFontWeight('bold');
  }
  let submitSheet = ss.getSheetByName(SHEET_ASSIGN_SUBMIT);
  if (!submitSheet) {
    submitSheet = ss.insertSheet(SHEET_ASSIGN_SUBMIT);
    submitSheet.appendRow(['제출일시','학생명','브랜드명','과제ID','과제명',
                           '제출타입','제출내용','원본파일명','메모','확인여부','교사피드백']);
    submitSheet.getRange(1,1,1,11).setFontWeight('bold');
  }
  return { listSheet, submitSheet };
}

// ── Google Drive 과제 폴더/하위폴더 확보 ─────────────────────────
function _getAssignFolder(assignId) {
  let root;
  const it = DriveApp.getFoldersByName(ASSIGN_DRIVE_FOLDER);
  root = it.hasNext() ? it.next() : DriveApp.createFolder(ASSIGN_DRIVE_FOLDER);

  const subName = assignId;
  const subIt   = root.getFoldersByName(subName);
  return subIt.hasNext() ? subIt.next() : root.createFolder(subName);
}

// ════════════════════════════════════════════════════════════════
// 1. 과제 목록 조회 (학생별 제출 현황 포함)
// ════════════════════════════════════════════════════════════════
function getAssignmentList(studentName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { listSheet, submitSheet } = _ensureAssignmentSheets(ss);
    const listData   = listSheet.getDataRange().getValues();
    const submitData = submitSheet.getDataRange().getValues();
    const tz         = Session.getScriptTimeZone();
    const now        = new Date();

    // 이 학생의 제출 기록 맵 (assignId → 최신 제출 정보)
    // ── [FIX 2026-05] 잘못된 날짜 값(빈문자열·Invalid Date) 방어 처리 ──
    const submitMap = {};
    for (let i = 1; i < submitData.length; i++) {
      if (String(submitData[i][1]).trim() !== String(studentName).trim()) continue;
      const aId = String(submitData[i][3]).trim();
      let ts;
      try {
        ts = submitData[i][0] ? new Date(submitData[i][0]) : new Date(0);
        if (isNaN(ts.getTime())) ts = new Date(0);
      } catch(e) { ts = new Date(0); }
      let submittedAtStr = '';
      try {
        if (ts.getTime() > 0) submittedAtStr = Utilities.formatDate(ts, tz, 'MM/dd HH:mm');
      } catch(e) { submittedAtStr = ''; }

      if (!submitMap[aId] || ts > submitMap[aId]._ts) {
        submitMap[aId] = {
          _ts:        ts,
          type:       submitData[i][5],
          content:    submitData[i][6],
          fileName:   submitData[i][7],
          memo:       submitData[i][8],
          checked:    submitData[i][9] === true || submitData[i][9] === 'TRUE',
          feedback:   submitData[i][10],
          submittedAt: submittedAtStr
        };
      }
    }

    const assignments = [];
    for (let i = 1; i < listData.length; i++) {
      const assignId = String(listData[i][0] || '').trim();
      if (!assignId) continue;
      const status = String(listData[i][5] || '').trim();
      if (status === '숨김') continue;

      // ── 마감일 변환 (방어 처리) ──
      const deadlineRaw = listData[i][3];
      let deadline = '';
      let dDay     = null;
      if (deadlineRaw) {
        try {
          const d = new Date(deadlineRaw);
          if (!isNaN(d.getTime())) {
            deadline = Utilities.formatDate(d, tz, 'MM/dd HH:mm');
            dDay     = Math.ceil((d - now) / (1000 * 60 * 60 * 24));
          }
        } catch(e) { /* 잘못된 날짜는 마감일 미정으로 표시 */ }
      }

      assignments.push({
        id:          assignId,
        name:        String(listData[i][1] || ''),
        description: String(listData[i][2] || ''),
        deadline:    deadline,
        allowType:   String(listData[i][4] || '둘다'),
        status:      status,
        dDay:        dDay,
        submitted:   submitMap[assignId] || null
      });
    }

    return { success: true, assignments: assignments };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// 2. 링크 제출
// ════════════════════════════════════════════════════════════════
function submitAssignmentLink(studentName, assignId, assignName, link, memo) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }

  try {
    const trimmedLink = String(link || '').trim();
    if (!trimmedLink)                          return { success: false, msg: '링크를 입력해주세요.' };
    if (!trimmedLink.startsWith('http'))       return { success: false, msg: '올바른 링크를 입력해주세요. (http://  또는 https:// 로 시작해야 합니다)' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { submitSheet } = _ensureAssignmentSheets(ss);
    const brand = _getStudentBrand(ss, studentName);

    // 기존 제출 행 찾기 (재제출 시 덮어쓰기)
    const submitData   = submitSheet.getDataRange().getValues();
    let   existingRow  = -1;
    for (let i = 1; i < submitData.length; i++) {
      if (String(submitData[i][1]).trim() === String(studentName).trim() &&
          String(submitData[i][3]).trim() === String(assignId).trim()) {
        existingRow = i + 1;
        break;
      }
    }

    const row = [new Date(), studentName, brand, assignId, assignName,
                 '링크', trimmedLink, '', String(memo || '').trim(), false, ''];
    if (existingRow > 0) {
      submitSheet.getRange(existingRow, 1, 1, 11).setValues([row]);
    } else {
      submitSheet.appendRow(row);
    }

    return { success: true, msg: '제출이 완료되었습니다! 선생님이 확인 후 피드백을 남겨드릴게요.' };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// 3. 파일 제출 (Base64 → Google Drive 저장)
//    클라이언트에서 FileReader로 base64 변환 후 호출
//    최대 권장 크기: 5MB (GAS 파라미터 한도)
// ════════════════════════════════════════════════════════════════
function submitAssignmentFile(studentName, assignId, assignName, base64Data, fileName, mimeType, memo) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }

  try {
    if (!base64Data) return { success: false, msg: '파일 데이터가 없습니다.' };
    if (!fileName)   return { success: false, msg: '파일명이 없습니다.' };

    // Drive에 파일 저장
    const tz       = Session.getScriptTimeZone();
    const dateStr  = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
    const safeName = studentName.replace(/[\\/:*?"<>|]/g, '_');
    const ext      = fileName.split('.').pop();
    const driveName = safeName + '_' + dateStr + '.' + ext;

    const bytes  = Utilities.base64Decode(base64Data);
    const blob   = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', driveName);
    const folder = _getAssignFolder(assignId);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    // 시트에 기록
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { submitSheet } = _ensureAssignmentSheets(ss);
    const brand = _getStudentBrand(ss, studentName);

    const submitData  = submitSheet.getDataRange().getValues();
    let   existingRow = -1;
    for (let i = 1; i < submitData.length; i++) {
      if (String(submitData[i][1]).trim() === String(studentName).trim() &&
          String(submitData[i][3]).trim() === String(assignId).trim()) {
        existingRow = i + 1;
        break;
      }
    }

    const row = [new Date(), studentName, brand, assignId, assignName,
                 '파일', fileUrl, fileName, String(memo || '').trim(), false, ''];
    if (existingRow > 0) {
      submitSheet.getRange(existingRow, 1, 1, 11).setValues([row]);
    } else {
      submitSheet.appendRow(row);
    }

    return { success: true, msg: '파일 제출이 완료되었습니다! 선생님이 확인 후 피드백을 남겨드릴게요.' };
  } catch(e) {
    return { success: false, msg: '파일 저장 중 오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// 내부 헬퍼
// ════════════════════════════════════════════════════════════════
function _getStudentBrand(ss, studentName) {
  const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      return mainData[i][COL_BRAND - 1];
    }
  }
  return '';
}
