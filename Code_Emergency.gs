// ════════════════════════════════════════════════════════════════
// 국가 비상사태 시스템
// ════════════════════════════════════════════════════════════════

// ── 비상사태 시트 자동 생성 ───────────────────────────────────────
function _ensureEmergencySheet(ss) {
  let sheet = ss.getSheetByName(SHEET_EMERGENCY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EMERGENCY);
    sheet.appendRow(['시나리오유형', '상태', '선포시각', '종료예정시각', '교사메모', '동결비율', '트리거ID']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

// ── 현재 진행 중인 비상사태 반환 ─────────────────────────────────
function _getActiveEmergency(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EMERGENCY);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === '진행중') {
      return {
        rowNum    : i + 1,
        type      : String(data[i][0]).trim(),
        status    : '진행중',
        startedAt : data[i][2],
        endsAt    : data[i][3],
        memo      : String(data[i][4]),
        freezeRate: Number(data[i][5]) || 30,
        triggerId : String(data[i][6])
      };
    }
  }
  return null;
}

// ── 비상사태 활성 여부 확인 (지출 함수에서 호출) ──────────────────
function _isEmergencyActive(type) {
  const e = _getActiveEmergency();
  if (!e) return false;
  return type ? e.type === type : true;
}

// ── 현재 비상사태 상태 조회 (Admin/학생 대시보드용) ───────────────
function getEmergencyStatus() {
  try {
    const e = _getActiveEmergency();
    if (!e) return { active: false };
    return {
      active    : true,
      type      : e.type,
      startedAt : e.startedAt ? Utilities.formatDate(new Date(e.startedAt), Session.getScriptTimeZone(), 'MM/dd HH:mm') : '',
      endsAt    : e.endsAt    ? Utilities.formatDate(new Date(e.endsAt),    Session.getScriptTimeZone(), 'MM/dd HH:mm') : '',
      memo      : e.memo,
      freezeRate: e.freezeRate
    };
  } catch(err) {
    return { active: false };
  }
}

// ── 비상사태 선포 ────────────────────────────────────────────────
function declareEmergency(type, endDatetimeStr, memo, freezeRate) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _ensureEmergencySheet(ss);

    // 이미 진행 중인 비상사태가 있으면 차단
    if (_getActiveEmergency(ss)) {
      return { success: false, msg: '이미 진행 중인 비상사태가 있습니다. 먼저 해제해 주세요.' };
    }

    const now       = new Date();
    const endDate   = endDatetimeStr ? new Date(endDatetimeStr) : null;
    const ts        = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const fRate     = Number(freezeRate) || 30;

    // 시트 기록
    sheet.appendRow([type, '진행중', now, endDate || '', memo || '', fRate, '']);
    const newRowNum = sheet.getLastRow();

    // 시간 기반 자동 종료 트리거 등록
    let triggerId = '';
    if (endDate && endDate > now) {
      const trigger = ScriptApp.newTrigger('autoEndEmergency')
        .timeBased().at(endDate).create();
      triggerId = trigger.getUniqueId();
      sheet.getRange(newRowNum, 7).setValue(triggerId);
    }

    // ── 시나리오별 즉시 효과 ──────────────────────────────────
    if (type === '고용 한파') {
      _executeLayoffs(ss);
    }

    // ── 전역 알림 발송 ────────────────────────────────────────
    const notifySheet = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
    if (notifySheet) {
      const msgMap = {
        '하이퍼인플레이션': '🚨 [국가 경제 적색경보] 화폐 발행량 증가로 가치가 폭락합니다! 모든 물가가 급등합니다.',
        '고용 한파'       : '📉 [고용 한파 발령] 국가 긴급 구조조정이 시작됩니다. 일부 인원이 구조조정됩니다. 직업을 잃은 사람은 실업급여 50p를 받습니다.',
        '자산 동결'       : '🔒 [경제 위기로 인한 긴급 통제] 현재 경제 위기로 인해 보유 자산의 일부만 사용 가능합니다.'
      };
      const noticeId = 'EMERGENCY_' + type.replace(/\s/g,'') + '_' + now.getTime();
      notifySheet.appendRow([noticeId, msgMap[type] || '🚨 국가 비상사태가 선포되었습니다.', ts, 'ALERT']);
    }

    // ── 전체 우편 발송 ────────────────────────────────────────
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    const mainData  = mainSheet.getDataRange().getValues();
    const mailMap = {
      '하이퍼인플레이션': {
        subject: '🚨 하이퍼인플레이션 발령',
        body   : `국가 경제 위기가 발생했습니다!\n\n모든 간식·상품 가격이 200%로 상승합니다.\n현명한 소비 전략을 세워 위기를 극복하세요.\n\n해제 예정: ${endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'MM/dd HH:mm') : '미정'}`
      },
      '고용 한파': {
        subject: '📉 고용 한파 발령',
        body   : `국가 긴급 구조조정이 시작되었습니다.\n\n일부 학생의 직업이 변경될 수 있습니다.\n자세한 내용은 별도 통보를 확인하세요.\n\n해제 예정: ${endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'MM/dd HH:mm') : '미정'}`
      },
      '자산 동결': {
        subject: '🔒 자산 동결 발령',
        body   : `금융 긴급 통제가 시작되었습니다.\n\n현재 보유 자산의 ${fRate}%만 사용 가능합니다.\n나머지 자산은 동결 해제 시까지 사용할 수 없습니다.\n\n해제 예정: ${endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'MM/dd HH:mm') : '미정'}`
      }
    };
    const mail = mailMap[type];
    if (mail) {
      for (let i = 1; i < mainData.length; i++) {
        const name = String(mainData[i][COL_NAME - 1]).trim();
        if (name) _sendMail(name, mail.subject, mail.body, '비상사태');
      }
    }

    return { success: true, msg: `[${type}] 비상사태가 선포되었습니다.` };
  } catch(err) {
    return { success: false, msg: '오류: ' + err.message };
  }
}

// ── 고용 한파: 해고 처리 ──────────────────────────────────────────
function _executeLayoffs(ss) {
  const jobSheet  = ss.getSheetByName(SHEET_JOB);
  if (!jobSheet) return;
  const jobData   = jobSheet.getDataRange().getValues();

  // 일급 합산
  let totalSalary = 0;
  const employees = [];
  for (let i = 1; i < jobData.length; i++) {
    const name   = String(jobData[i][0]).trim();
    const salary = Number(jobData[i][2]) || 0;
    if (!name || salary === 0) continue;
    totalSalary += salary;
    employees.push({ rowNum: i + 1, name: name, salary: salary });
  }

  const BUDGET      = 3000; // 국가 예산 기준
  const TARGET      = 2800; // 목표 총임금
  if (totalSalary <= BUDGET) return; // 예산 초과 아니면 해고 없음

  // 일급 낮은 순으로 정렬
  employees.sort(function(a, b) { return a.salary - b.salary; });

  const now = new Date();
  for (let e = 0; e < employees.length && totalSalary > TARGET; e++) {
    const emp = employees[e];
    totalSalary -= emp.salary;
    // 직업 → 무직, 일급 → 50 (실업급여)
    jobSheet.getRange(emp.rowNum, 2).setValue('무직');
    jobSheet.getRange(emp.rowNum, 3).setValue(50);
    // 해고 통보 우편
    _sendMail(emp.name, '📋 [구조조정 통보서]',
      `안타깝게도 국가 예산 초과로 인해 귀하의 직책이 무직으로 변경되었습니다.\n실업급여 $50이 지급됩니다.\n비상사태 해제 후 선생님께 복직을 문의하세요.`, '비상사태');
  }

  // 잔류 학생 경고 우편
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();
  const firedNames = employees.filter(function(e2) {
    const row = jobSheet.getRange(e2.rowNum, 2).getValue();
    return row === '무직';
  }).map(function(e2) { return e2.name; });

  for (let i = 1; i < mainData.length; i++) {
    const name = String(mainData[i][COL_NAME - 1]).trim();
    if (name && firedNames.indexOf(name) === -1) {
      _sendMail(name, '⚠️ [고용 한파 생존 통보]',
        '이번 구조조정에서 직위가 유지되었습니다.\n하지만 경제 위기는 계속됩니다. 방심하지 마세요.', '비상사태');
    }
  }
}

// ── 비상사태 해제 ────────────────────────────────────────────────
function endEmergency() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const e  = _getActiveEmergency(ss);
    if (!e) return { success: false, msg: '현재 진행 중인 비상사태가 없습니다.' };

    // 시트 상태 업데이트
    const sheet = ss.getSheetByName(SHEET_EMERGENCY);
    sheet.getRange(e.rowNum, 2).setValue('종료');

    // 시간 트리거 삭제
    if (e.triggerId) {
      const triggers = ScriptApp.getProjectTriggers();
      for (let t = 0; t < triggers.length; t++) {
        if (triggers[t].getUniqueId() === e.triggerId) {
          ScriptApp.deleteTrigger(triggers[t]);
          break;
        }
      }
    }

    // 전역 알림
    const notifySheet = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
    if (notifySheet) {
      const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const nid = 'EMERGENCY_END_' + new Date().getTime();
      notifySheet.appendRow([nid, `✅ [비상사태 해제] ${e.type} 비상사태가 종료되었습니다. 경제가 정상화됩니다.`, ts, 'ALERT']);
    }

    // 전체 우편
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    const mainData  = mainSheet.getDataRange().getValues();
    const endMsg = e.type === '고용 한파'
      ? `${e.type} 비상사태가 해제되었습니다.\n복직을 원하는 학생은 선생님께 문의하세요.`
      : `${e.type} 비상사태가 해제되었습니다. 모든 제한이 풀렸습니다.`;
    for (let i = 1; i < mainData.length; i++) {
      const name = String(mainData[i][COL_NAME - 1]).trim();
      if (name) _sendMail(name, '✅ [비상사태 해제]', endMsg, '비상사태');
    }

    return { success: true, msg: `[${e.type}] 비상사태가 해제되었습니다.` };
  } catch(err) {
    return { success: false, msg: '오류: ' + err.message };
  }
}

// ── 자동 종료 트리거 함수 ────────────────────────────────────────
function autoEndEmergency() {
  endEmergency();
}


// ── 예금 만기 자동 처리 트리거 설정 ──────────────────────────────
function setupDepositTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyDepositCheck') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runDailyDepositCheck')
    .timeBased().everyDays(1).atHour(12).nearMinute(30).create();
  SpreadsheetApp.getUi().alert('✅ 매일 12:30 예금 만기 자동 처리 트리거가 설정되었습니다.');
}

function runDailyDepositCheck() {
  checkAndPayDeposits(null); // null = 전체 학생 처리
}