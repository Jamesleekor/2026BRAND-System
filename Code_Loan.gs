// ════════════════════════════════════════════════════════════════
// ██ 대출 시스템
// 시트: 대출현황
//   A=대출ID, B=학생명, C=대출액, D=금리(주%), E=신용등급,
//   F=용도, G=실행일, H=상환기한(yyyy-MM-dd), I=상환상태(정상/연체/완료),
//   J=연체주수, K=잔여원금
// 시트: 대출신청로그
//   A=타임스탬프, B=학생명, C=신청금액, D=용도, E=신용점수,
//   F=신용등급, G=최대대출한도, H=상태(대기/승인/반려)
// ════════════════════════════════════════════════════════════════

// ── 시트 자동 생성 헬퍼 ──────────────────────────────────────────
function _ensureLoanStatusSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_LOAN_STATUS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOAN_STATUS);
    sheet.appendRow(['대출ID','학생명','대출액','금리(주%)','신용등급','용도','실행일','상환기한','상환상태','연체주수','잔여원금']);
  }
  return sheet;
}

function _ensureLoanLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_LOAN_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOAN_LOG);
    sheet.appendRow(['타임스탬프','학생명','신청금액','용도','신용점수','신용등급','최대대출한도','상태']);
  }
  return sheet;
}

// ── 대출 신청 ────────────────────────────────────────────────────
// amount: 신청 금액, purpose: 용도 문자열
function applyLoan(studentName, amount, purpose) {
  amount = Number(amount);

  // ── 기본 유효성 검사 ────────────────────────────────────────
  if (!studentName) return { success: false, msg: '학생 정보가 없습니다.' };
  if (!amount || amount <= 0) return { success: false, msg: '대출 금액을 입력해주세요.' };
  if (amount % 100 !== 0) return { success: false, msg: '금액은 100 단위로 입력해주세요.' };
  if (!purpose || !String(purpose).trim()) return { success: false, msg: '대출 용도를 선택해주세요.' };

  const allowedPurposes = ['경매 참여 자금', '2차 직업 사업 자금', '정기 예금 가입', '기타 생산적 용도'];
  if (!allowedPurposes.includes(String(purpose).trim())) {
    return { success: false, msg: '허용되지 않는 대출 용도입니다.' };
  }

  // ── 중복 실행 방지 Lock ──────────────────────────────────────
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, msg: '잠시 후 다시 시도해주세요. (동시 요청 충돌)' };
  }

  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    const mainData  = mainSheet.getDataRange().getValues();

    // ── 학생 행 찾기 ────────────────────────────────────────
    let studentIdx = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
        studentIdx = i; break;
      }
    }
    if (studentIdx === -1) return { success: false, msg: '학생 정보를 찾을 수 없습니다.' };

    const curAsset = Number(mainData[studentIdx][COL_ASSET - 1]) || 0;

    // ── 신용점수 계산 및 등급 확인 ───────────────────────────
    const creditResult = calcCreditScore(studentName);
    const gradeInfo    = getCreditGrade(creditResult.total);

    if (gradeInfo.grade === '대출불가') {
      return { success: false, msg: '현재 신용등급(대출불가)으로는 대출 신청이 불가합니다.' };
    }

    // ── 신청 시점 자산 기준 최대 한도 계산 ──────────────────
    const maxLoan = Math.floor(curAsset * gradeInfo.maxLoanRate);
    if (amount > maxLoan) {
      return { success: false, msg: `최대 대출 한도는 $${maxLoan.toLocaleString()}입니다. (현재 자산 $${curAsset.toLocaleString()} × ${Math.round(gradeInfo.maxLoanRate * 100)}%)` };
    }

    // ── 동시 1건 제한 확인 ──────────────────────────────────
    const loanSheet = ss.getSheetByName(SHEET_LOAN_STATUS);
    if (loanSheet) {
      const loanData = loanSheet.getDataRange().getValues();
      for (let i = 1; i < loanData.length; i++) {
        if (String(loanData[i][1]).trim() !== String(studentName).trim()) continue;
        const st = String(loanData[i][8]).trim(); // I: 상환상태
        if (st === '정상' || st === '연체') {
          return { success: false, msg: '이미 진행 중인 대출이 있습니다. 상환 후 재신청해주세요.' };
        }
      }
    }

    // ── 대출신청로그에 기록 ──────────────────────────────────
    const logSheet = _ensureLoanLogSheet(ss);
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([
      ts,                       // A: 타임스탬프
      studentName,              // B: 학생명
      amount,                   // C: 신청금액
      String(purpose).trim(),   // D: 용도
      creditResult.total,       // E: 신용점수
      gradeInfo.grade,          // F: 신용등급
      maxLoan,                  // G: 최대대출한도
      '대기'                    // H: 상태
    ]);

    // ── 학생에게 접수 확인 우편 발송 ────────────────────────
    _sendMail(
      studentName,
      '💳 대출 신청 접수',
      `$${amount.toLocaleString()} 대출 신청이 접수되었습니다.\n용도: ${purpose}\n선생님 검토 후 승인/반려 결과를 알려드립니다.`,
      'loan'
    );

    return { success: true, msg: `✅ 대출 신청이 접수되었습니다. 선생님 승인을 기다려주세요.` };

  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── 대기 중인 대출 신청 목록 반환 (교사 패널용) ──────────────────
function getPendingLoans() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOAN_LOG);
  if (!sheet) return [];

  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]).trim() !== '대기') continue; // H: 상태
    result.push({
      rowNum:    i + 1,
      ts:        String(data[i][0]),
      name:      String(data[i][1]).trim(),
      amount:    Number(data[i][2]) || 0,
      purpose:   String(data[i][3]).trim(),
      score:     Number(data[i][4]) || 0,
      grade:     String(data[i][5]).trim(),
      maxLoan:   Number(data[i][6]) || 0
    });
  }
  return result.reverse(); // 최신순
}

// ── 대출 승인 / 반려 (교사 패널에서 호출) ───────────────────────
// rowNum: 대출신청로그 행 번호, isApproved: true/false, rejectReason: 반려 사유
function approveLoan(rowNum, isApproved, rejectReason) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '잠시 후 다시 시도해주세요.' }; }

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_LOAN_LOG);
    if (!logSheet) return { success: false, msg: '대출신청로그 시트를 찾을 수 없습니다.' };

    const logData = logSheet.getDataRange().getValues();
    if (rowNum < 2 || rowNum > logData.length) return { success: false, msg: '유효하지 않은 행 번호입니다.' };

    const row     = logData[rowNum - 1];
    const name    = String(row[1]).trim();
    const amount  = Number(row[2]) || 0;
    const purpose = String(row[3]).trim();
    const grade   = String(row[5]).trim();

    // 이미 처리된 신청인지 확인
    if (String(row[7]).trim() !== '대기') {
      return { success: false, msg: '이미 처리된 신청입니다.' };
    }

    // ── 반려 처리 ────────────────────────────────────────────
    if (!isApproved) {
      logSheet.getRange(rowNum, 8).setValue('반려');
      _sendMail(name, '💳 대출 신청 반려',
        `대출 신청이 반려되었습니다.\n사유: ${rejectReason || '선생님 검토 결과 반려'}`, 'loan');
      return { success: true, msg: `${name} 학생의 대출 신청이 반려되었습니다.` };
    }

    // ── 승인 처리 ────────────────────────────────────────────
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    const mainData  = mainSheet.getDataRange().getValues();
    let studentIdx  = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][COL_NAME - 1]).trim() === name) { studentIdx = i; break; }
    }
    if (studentIdx === -1) return { success: false, msg: '학생 정보를 찾을 수 없습니다.' };

    const curAsset = Number(mainData[studentIdx][COL_ASSET - 1]) || 0;
    const newAsset = curAsset + amount;
    const curHonor = Number(mainData[studentIdx][COL_VALUE - 1]) || 0;
    const today    = _todayStr();

    // 상환 기한: 실행일 + 28일
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 28);
    const dueDateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // 금리 조회
    const gradeInfo  = getCreditGrade(calcCreditScore(name).total);
    const weeklyRate = gradeInfo.weeklyRate;

    // 자산 증가
    mainSheet.getRange(studentIdx + 1, COL_ASSET).setValue(newAsset);

    // 히스토리 기록 (8개 값 필수)
    ss.getSheetByName(SHEET_HISTORY).appendRow([
      today, name, mainData[studentIdx][COL_BRAND - 1],
      0, amount, curHonor, newAsset,
      `[대출 실행] ${purpose} / ${grade}등급 / 주${weeklyRate}%`
    ]);

    // 대출현황 시트에 기록
    const statusSheet = _ensureLoanStatusSheet(ss);
    const loanId = 'LOAN_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 4);
    statusSheet.appendRow([
      loanId,       // A: 대출ID
      name,         // B: 학생명
      amount,       // C: 대출액
      weeklyRate,   // D: 금리(주%)
      grade,        // E: 신용등급
      purpose,      // F: 용도
      today,        // G: 실행일
      dueDateStr,   // H: 상환기한
      '정상',        // I: 상환상태
      0,            // J: 연체주수
      amount        // K: 잔여원금
    ]);

    // 대출신청로그 상태 → 승인
    logSheet.getRange(rowNum, 8).setValue('승인');

    // 랭킹 갱신 + 캐시 무효화
    updateRankings();
    CacheService.getScriptCache().remove('student_' + name);

    // 학생에게 승인 우편 발송
    _sendMail(name, '✅ 대출 승인',
      `$${amount.toLocaleString()} 대출이 승인되었습니다!\n용도: ${purpose}\n상환 기한: ${dueDateStr} (28일 이내 일시 상환)\n주간 금리: ${weeklyRate}%`,
      'loan');

    return { success: true, msg: `✅ ${name} 학생에게 $${amount.toLocaleString()} 대출이 실행되었습니다.` };

  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── 대출 상환 (학생이 호출) ──────────────────────────────────────
function repayLoan(studentName, loanId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '잠시 후 다시 시도해주세요.' }; }

  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const statusSheet = ss.getSheetByName(SHEET_LOAN_STATUS);
    if (!statusSheet) return { success: false, msg: '대출현황 시트를 찾을 수 없습니다.' };

    const loanData = statusSheet.getDataRange().getValues();
    let loanRowIdx = -1;
    let loanRow    = null;
    for (let i = 1; i < loanData.length; i++) {
      if (String(loanData[i][0]).trim() === String(loanId).trim() &&
          String(loanData[i][1]).trim() === String(studentName).trim()) {
        loanRowIdx = i; loanRow = loanData[i]; break;
      }
    }
    if (loanRowIdx === -1) return { success: false, msg: '해당 대출을 찾을 수 없습니다.' };

    const repayAmount = Number(loanRow[10]) || 0; // K: 잔여원금
    const status      = String(loanRow[8]).trim(); // I: 상환상태
    if (status === '완료') return { success: false, msg: '이미 상환 완료된 대출입니다.' };

    // 자산 확인
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    const mainData  = mainSheet.getDataRange().getValues();
    let studentIdx  = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
        studentIdx = i; break;
      }
    }
    if (studentIdx === -1) return { success: false, msg: '학생 정보를 찾을 수 없습니다.' };

    const curAsset = Number(mainData[studentIdx][COL_ASSET - 1]) || 0;
    if (curAsset < repayAmount) {
      return { success: false, msg: `잔액이 부족합니다. 상환 필요액: $${repayAmount.toLocaleString()} / 현재 자산: $${curAsset.toLocaleString()}` };
    }

    const newAsset = curAsset - repayAmount;
    const curHonor = Number(mainData[studentIdx][COL_VALUE - 1]) || 0;
    const today    = _todayStr();

    // 자산 차감
    mainSheet.getRange(studentIdx + 1, COL_ASSET).setValue(newAsset);

    // 히스토리 기록 (8개 값 필수)
    ss.getSheetByName(SHEET_HISTORY).appendRow([
      today, studentName, mainData[studentIdx][COL_BRAND - 1],
      0, -repayAmount, curHonor, newAsset,
      `[대출 상환] 대출ID: ${loanId}`
    ]);

    // 대출현황 상태 → 완료
    statusSheet.getRange(loanRowIdx + 1, 9).setValue('완료');   // I: 상환상태
    statusSheet.getRange(loanRowIdx + 1, 11).setValue(0);       // K: 잔여원금

    // 랭킹 갱신 + 캐시 무효화
    updateRankings();
    CacheService.getScriptCache().remove('student_' + studentName);

    // 우편 발송
    _sendMail(studentName, '✅ 대출 상환 완료',
      `$${repayAmount.toLocaleString()} 대출이 상환 완료되었습니다. 수고하셨습니다!`, 'loan');

    return { success: true, msg: `✅ $${repayAmount.toLocaleString()} 상환이 완료되었습니다!` };

  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── 대시보드용: 현재 대출 상태 반환 ─────────────────────────────
// 진행중(정상/연체) 대출 1건 반환. 없으면 null.
function getLoanStatus(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOAN_STATUS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== String(studentName).trim()) continue;
    const st = String(data[i][8]).trim(); // I: 상환상태
    if (st !== '정상' && st !== '연체') continue;
    return {
      loanId:     String(data[i][0]).trim(),
      amount:     Number(data[i][2]) || 0,
      rate:       Number(data[i][3]) || 0,
      grade:      String(data[i][4]).trim(),
      purpose:    String(data[i][5]).trim(),
      startDate:  String(data[i][6]),
      dueDate:    String(data[i][7]),
      status:     st,
      overdueWeeks: Number(data[i][9]) || 0,
      remaining:  Number(data[i][10]) || 0
    };
  }
  return null;
}

// ── 연체 처리 (매주 월요일 오전 트리거) ─────────────────────────
function checkLoanOverdue() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const statusSheet = ss.getSheetByName(SHEET_LOAN_STATUS);
  if (!statusSheet) return;

  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();
  const loanData  = statusSheet.getDataRange().getValues();
  const today     = _todayStr();

  for (let i = 1; i < loanData.length; i++) {
    const st      = String(loanData[i][8]).trim(); // I: 상환상태
    // 정상 또는 연체 상태만 처리 (완료 제외)
    if (st !== '정상' && st !== '연체') continue;

    const dueDate = String(loanData[i][7]).trim().substring(0, 10); // H: 상환기한
    // 상환기한이 오늘보다 이전인 경우만 연체 처리
    if (dueDate >= today) continue;

    const name         = String(loanData[i][1]).trim();
    const baseRate     = Number(loanData[i][3]) || 0;  // D: 기본 금리
    const overdueWeeks = Number(loanData[i][9]) || 0;  // J: 현재 연체주수
    const remaining    = Number(loanData[i][10]) || 0; // K: 잔여원금

    // 연체 금리 = 기본금리 + 2 × (연체주수 + 1)
    const penaltyRate = baseRate + 2 * (overdueWeeks + 1);
    const newRemaining = Math.floor(remaining * (1 + penaltyRate / 100));
    const newOverdueWeeks = overdueWeeks + 1;

    // 대출현황 업데이트
    statusSheet.getRange(i + 1, 9).setValue('연체');           // I: 상환상태
    statusSheet.getRange(i + 1, 10).setValue(newOverdueWeeks); // J: 연체주수
    statusSheet.getRange(i + 1, 11).setValue(newRemaining);    // K: 잔여원금

    // 브랜드가치 -50
    let studentIdx = -1;
    for (let j = 1; j < mainData.length; j++) {
      if (String(mainData[j][COL_NAME - 1]).trim() === name) { studentIdx = j; break; }
    }
    if (studentIdx !== -1) {
      const curHonor  = Number(mainData[studentIdx][COL_VALUE - 1]) || 0;
      const newHonor  = Math.max(0, curHonor - 50);
      mainSheet.getRange(studentIdx + 1, COL_VALUE).setValue(newHonor);
      // mainData 인메모리도 갱신 (같은 루프 내 중복 처리 방지)
      mainData[studentIdx][COL_VALUE - 1] = newHonor;
    }

    // 학생에게 연체 알림 우편
    _sendMail(name, '⚠️ 대출 연체 알림',
      `대출이 연체되었습니다! (${newOverdueWeeks}주차)\n현재 잔여 원금: $${newRemaining.toLocaleString()}\n즉시 상환하지 않으면 매주 원금이 불어납니다.`,
      'loan');

    CacheService.getScriptCache().remove('student_' + name);
  }

  // 전체 처리 후 랭킹 갱신
  updateRankings();
}
