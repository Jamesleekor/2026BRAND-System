// ════════════════════════════════════════════════════════════════
// ██ 기능 1: 우편함(Mailbox) ██
// 시트 컬럼: A=메시지ID, B=수신자, C=제목, D=내용, E=타입(승인/반려),
//            F=읽음여부(TRUE/FALSE), G=발송일시
// ════════════════════════════════════════════════════════════════

// 우편함 메시지 전송 (내부 헬퍼, approveAchievement에서 호출)
function _sendMail(recipientName, subject, body, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MAILBOX);
    sheet.appendRow(['메시지ID','수신자','제목','내용','타입','읽음','발송일시']);
  }
  const msgId = 'MSG_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,5);
  const ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([msgId, recipientName, subject, body, type, false, ts]);
}

// 학생의 읽지 않은 메시지 수 반환
function getUnreadMailCount(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(studentName).trim() &&
        String(data[i][5]).toUpperCase() !== 'TRUE') {
      count++;
    }
  }
  return count;
}

// 학생의 전체 메시지 목록 반환 + 읽음 처리
function getMailboxMessages(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const msgs = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== String(studentName).trim()) continue;
    const isRead = String(data[i][5]).toUpperCase() === 'TRUE';
    msgs.push({
      msgId:   String(data[i][0]),
      subject: String(data[i][2]),
      body:    String(data[i][3]),
      type:    String(data[i][4]),   // '승인' | '반려'
      isRead:  isRead,
      sentAt:  String(data[i][6]),
      rowNum:  i + 1
    });
    // 읽음 처리
    if (!isRead) sheet.getRange(i + 1, 6).setValue(true);
  }
  return msgs.reverse(); // 최신순
}

// ════════════════════════════════════════════════════════════════
// approveAchievement 확장: 승인/반려 시 우편함 메시지 자동 발송
// 기존 함수를 덮어쓰지 않고, 래퍼 함수로 우편함 호출을 삽입합니다.
// → AuctionAdmin.html에서 approveAchievement 대신
//   approveAchievementWithMail 을 호출하도록 변경하세요.
// ════════════════════════════════════════════════════════════════
function approveAchievementWithMail(rowNumber, isApproved, finalAchievementId, rejectReason) {
  const result = approveAchievement(rowNumber, isApproved, finalAchievementId);
  if (!result.success) return result;

  // 수신자 이름 다시 조회
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  if (!logSheet) return result;
  const row         = logSheet.getRange(rowNumber, 1, 1, 5).getValues()[0];
  const studentName = String(row[1]).trim();
  const achId       = String(row[2]).trim();

  // 업적명 조회
  let achName = achId;
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (String(mData[m][0]).trim() === (finalAchievementId || achId)) {
        achName = String(mData[m][1]).trim();
        break;
      }
    }
  }

  if (isApproved) {
    _sendMail(
      studentName,
      `✅ 업적 승인: ${achName}`,
      `🎉 축하합니다! [${achName}] 업적 신청이 승인되었습니다. 나의 업적 창에서 확인해보세요!`,
      '승인'
    );
  } else {
    const reason = rejectReason ? rejectReason : '조건 미충족';
    _sendMail(
      studentName,
      `❌ 업적 반려: ${achName}`,
      `[${achName}] 업적 신청이 반려되었습니다.\n\n반려 사유: ${reason}\n\n조건을 다시 확인하고 재신청해주세요.`,
      '반려'
    );
  }
  return result;
}

// ════════════════════════════════════════════════════════════════
// ██ 기능 2: 업적 전광판 (Global Alert)
// 전광판 조건: ① 업적 10개 단위 ② 유일/초월 등급 획득
// 시트: 전역알림 (기존 SHEET_GLOBAL_NOTIFY) 재활용
// ════════════════════════════════════════════════════════════════

// 학생 업적 개수 & 등급 기반 전광판 메시지 생성 (approveAchievement 후 호출)
function _checkAndPostGlobalAlert(studentName, achName, achGrade) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const achSheet = ss.getSheetByName(SHEET_ACH_STUDENT);
  const notify   = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
  if (!achSheet || !notify) return;

  // 해당 학생 총 업적 수
  const achData = achSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < achData.length; i++) {
    if (String(achData[i][0]).trim() === String(studentName).trim()) count++;
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  let msg = null;
  let noticeId = null;

  // ① 10개 단위 달성
  if (count > 0 && count % 10 === 0) {
    noticeId = `MILESTONE_${studentName}_${count}_${new Date().getTime()}`;
    msg = `🏆 [${studentName}] 학생이 업적 ${count}개 달성! 대단한 업적 수집가가 탄생했습니다!`;
  }

  // ② 유일/초월 등급 획득 (더 우선순위 높음)
  if (achGrade === '유일' || achGrade === '초월') {
    const gradeLabel = achGrade === '유일' ? '🌌 유일' : '✨ 초월';
    noticeId = `GRADE_${achGrade}_${studentName}_${new Date().getTime()}`;
    msg = `${gradeLabel} 등급 업적 [${achName}] 발현! [${studentName}] 학생의 이름이 아카식 레코드의 최상단에 기록됩니다!`;
  }

  if (msg && noticeId) {
    notify.appendRow([noticeId, msg, ts, 'ALERT']); // D열 = 'ALERT' 타입 표시
  }
}

// ── 티어 최초 진입 전역 알림 (checkAndGrantAchievements에서 호출) ──
function _postTierFirstAlert(studentName, tierName) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const notify = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
  if (!notify) return;

  const tierEmojis = {
    '금 광석': '🥇',
    '루비 원석': '💎',
    '다이아 원석': '💠',
    '마스터': '👑',
    '천상의 마스터': '👑',
    '그랜드마스터': '🏆'
  };
  const emoji = tierEmojis[tierName] || '🎉';
  const noticeId = 'TIER_' + tierName.replace(/\s/g, '') + '_' + new Date().getTime();
  const msg = emoji + ' [경고] 세계선이 재편됩니다.\n[' + studentName + '] 학생이 최초로 경계를 넘었습니다.\n이 기록은 영원히 소멸되지 않습니다.';
  const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  notify.appendRow([noticeId, msg, ts, 'ALERT']);
}

// 전광판 최신 메시지 조회 (프론트에서 폴링)
function getLatestGlobalAlert(lastSeenId, loginTimeStr) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();

  // 로그인 시각 파싱 (프론트에서 ISO 문자열로 전달)
  const loginTime = loginTimeStr ? new Date(loginTimeStr) : null;

  // 최신 행부터 탐색 — ALERT 타입이고, 로그인 이후 발생했고, 이미 본 것이 아닌 것만 반환
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][3]) !== 'ALERT') continue;
    if (String(data[i][0]) === String(lastSeenId)) continue;
    // 로그인 시각 이후 발생한 알림만 허용
    if (loginTime) {
      const alertTime = new Date(data[i][2]);
      if (!isNaN(alertTime.getTime()) && alertTime < loginTime) continue;
    }
    return { noticeId: String(data[i][0]), msg: String(data[i][1]), ts: String(data[i][2]) };
  }
  return null;
}

// ════════════════════════════════════════════════════════════════
// 전체 메시지 발송 (Admin 패널용)
// ════════════════════════════════════════════════════════════════

// 전체 학생 목록 반환 (Admin 패널에서 수신자 선택용)
function getStudentListForMail() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };
    const data = mainSheet.getDataRange().getValues();
    const students = [];
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][COL_NAME - 1]).trim();
      if (name) students.push(name);
    }
    return { success: true, students: students };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}

// 전체(또는 선택) 학생에게 메시지 발송
function sendBroadcastMail(subject, body, targetNames) {
  // targetNames: 배열. 비어있으면 전체 발송
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };

    const data = mainSheet.getDataRange().getValues();
    const allStudents = [];
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][COL_NAME - 1]).trim();
      if (name) allStudents.push(name);
    }

    const recipients = (targetNames && targetNames.length > 0)
                       ? targetNames
                       : allStudents;

    if (recipients.length === 0) return { success: false, msg: '발송 대상 학생이 없습니다.' };

    for (let i = 0; i < recipients.length; i++) {
      _sendMail(recipients[i], subject, body, '공지');
    }
    return { success: true, count: recipients.length };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}
