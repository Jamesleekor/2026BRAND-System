// ════════════════════════════════════════════════════════════════
// ██ 정기 예금 시스템
// 시트: 예금상품 (A=상품ID, B=상품명, C=1주이자율, D=2주이자율,
//                E=3주이자율, F=4주이자율, G=최소금액, H=최대금액,
//                I=패널티율, J=상태, K=론칭일)
//       학생별가입예금 (A=예금ID, B=학생명, C=원금, D=이자율,
//                      E=거치기간(주), F=시작일, G=만기일,
//                      H=상태, I=지급이자액, J=처리일, K=상품ID)
// ════════════════════════════════════════════════════════════════

// ── 현재 활성 예금 상품 반환 ──────────────────────────────────────
function getActiveDepositProduct() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9]).trim() === '활성') {  // J열: 상태
      return {
        prodId:    String(data[i][0]).trim(),   // A
        prodName:  String(data[i][1]).trim(),   // B
        rate1:     Number(data[i][2]) || 0,     // C
        rate2:     Number(data[i][3]) || 0,     // D
        rate3:     Number(data[i][4]) || 0,     // E
        rate4:     Number(data[i][5]) || 0,     // F
        minAmount: Number(data[i][6]) || 500,   // G
        maxAmount: Number(data[i][7]) || 5000,  // H
        penalty:   Number(data[i][8]) || 5,     // I
        launchDate: String(data[i][10]).trim()  // K
      };
    }
  }
  return null;
}

// ── 신규 예금 상품 론칭 (AuctionAdmin에서 호출) ───────────────────
// rates: { r1, r2, r3, r4 } — 각 주별 이자율(%)
function launchDepositProduct(prodName, rates, penalty, minAmount, maxAmount) {
  // 유효성 검사
  if (!prodName || !prodName.trim())
    return { success: false, msg: '상품명을 입력해주세요.' };
  if (!rates.r1 || !rates.r2 || !rates.r3 || !rates.r4)
    return { success: false, msg: '이자율 4개를 모두 입력해주세요.' };
  penalty   = Number(penalty)   || 5;
  minAmount = Number(minAmount) || 500;
  maxAmount = Number(maxAmount) || 5000;

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!prodSheet) return { success: false, msg: '예금상품 시트를 찾을 수 없습니다.' };

  // 기존 활성 상품 → 종료 처리
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (String(prodData[i][9]).trim() === '활성') {
      prodSheet.getRange(i + 1, 10).setValue('종료');
    }
  }

  // 신규 상품 행 추가
  const today  = _todayStr();
  const prodId = 'PROD_' + today.replace(/-/g, '');
  prodSheet.appendRow([
    prodId,
    prodName.trim(),
    Number(rates.r1),
    Number(rates.r2),
    Number(rates.r3),
    Number(rates.r4),
    minAmount,
    maxAmount,
    penalty,
    '활성',
    today
  ]);

  // 전체 학생에게 우편 발송
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (mainSheet) {
    const mainData = mainSheet.getDataRange().getValues();
    for (let i = 1; i < mainData.length; i++) {
      const name = String(mainData[i][COL_NAME - 1]).trim();
      if (!name) continue;
      _sendMail(
        name,
        `🏦 새 예금 상품 출시: ${prodName}`,
        `새로운 정기예금 상품이 출시되었습니다!\n\n` +
        `📌 상품명: ${prodName}\n` +
        `💰 이자율: 1주 ${rates.r1}% / 2주 ${rates.r2}% / 3주 ${rates.r3}% / 4주 ${rates.r4}%\n` +
        `💵 가입 한도: 최소 $${minAmount} ~ 최대 $${maxAmount} (100단위)\n` +
        `⚠️ 중도해지 패널티: 원금의 ${penalty}%\n\n` +
        `지금 대시보드에서 가입하세요!`,
        '공지'
      );
    }
  }

  return { success: true, msg: `✅ [${prodName}] 상품이 론칭되었습니다. 전체 학생에게 우편이 발송되었습니다.` };
}

// ── 현재 활성 상품 패널티율 수정 (AuctionAdmin에서 호출) ──────────
function setPenaltyRate(rate) {
  rate = Number(rate);
  if (isNaN(rate) || rate < 0 || rate > 100)
    return { success: false, msg: '0~100 사이의 숫자를 입력해주세요.' };

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!prodSheet) return { success: false, msg: '예금상품 시트를 찾을 수 없습니다.' };

  const data = prodSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9]).trim() === '활성') {
      prodSheet.getRange(i + 1, 9).setValue(rate);  // I열
      return { success: true, msg: `✅ 패널티율이 ${rate}%로 변경되었습니다.` };
    }
  }
  return { success: false, msg: '현재 활성 상품이 없습니다.' };
}

// ── 예금 가입 (학생 → Index.html에서 호출) ───────────────────────
function createDeposit(studentName, amount, weeks) {
  amount = Number(amount);
  weeks  = Number(weeks);

  // ── 유효성 검사 ──────────────────────────────────────────────
  if (!studentName) return { success: false, msg: '학생 정보가 없습니다.' };
  if (![1, 2, 3, 4].includes(weeks))
    return { success: false, msg: '거치 기간은 1~4주 중 선택해주세요.' };
  if (!amount || amount <= 0)
    return { success: false, msg: '금액을 입력해주세요.' };
  if (amount % 100 !== 0)
    return { success: false, msg: '금액은 100 단위로 입력해주세요. (예: 500, 1000)' };

  // 활성 상품 조회
  const prod = getActiveDepositProduct();
  if (!prod) return { success: false, msg: '현재 가입 가능한 예금 상품이 없습니다.' };

  if (amount < prod.minAmount)
    return { success: false, msg: `최소 가입 금액은 $${prod.minAmount.toLocaleString()}입니다.` };
  if (amount > prod.maxAmount)
    return { success: false, msg: `1회 최대 가입 금액은 $${prod.maxAmount.toLocaleString()}입니다.` };

  // ── 상품당 누적 한도 체크 ─────────────────────────────────
  // 현재 활성 상품에 진행중인 예금 합산액이 maxAmount를 초과하지 않도록
  const logSheetCheck = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DEPOSIT_LOG);
  if (logSheetCheck) {
    const logCheck = logSheetCheck.getDataRange().getValues();
    let alreadyDeposited = 0;
    for (let i = 1; i < logCheck.length; i++) {
      if (String(logCheck[i][1]).trim() === String(studentName).trim() &&
          String(logCheck[i][10]).trim() === prod.prodId &&
          String(logCheck[i][7]).trim() === '진행중') {
        alreadyDeposited += Number(logCheck[i][2]) || 0;
      }
    }
    const remaining = prod.maxAmount - alreadyDeposited;
    if (remaining <= 0)
      return { success: false, msg: `이번 상품의 최대 한도 $${prod.maxAmount.toLocaleString()}에 이미 도달했습니다.` };
    if (amount > remaining)
      return { success: false, msg: `이번 상품에 추가 가입 가능한 금액은 $${remaining.toLocaleString()}입니다. (한도: $${prod.maxAmount.toLocaleString()})` };
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
  // 학생 행 찾기
  const mainData = mainSheet.getDataRange().getValues();
  let studentIdx = -1;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      studentIdx = i; break;
    }
  }
  if (studentIdx === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

  const curAsset = Number(mainData[studentIdx][COL_ASSET - 1]) || 0;
  if (curAsset < amount)
    return { success: false, msg: `잔액이 부족합니다. (현재: $${curAsset.toLocaleString()})` };

  // 이자율 결정
  const rateMap = { 1: prod.rate1, 2: prod.rate2, 3: prod.rate3, 4: prod.rate4 };
  const rate    = rateMap[weeks];

  // 만기일 계산 (시작일 + weeks*7일)
  const today     = _todayStr();
  const dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + weeks * 7);
  dueDate.setHours(12, 0, 0, 0);  // 정오 12:00 고정
  const dueDateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 자산 차감
  const newAsset = curAsset - amount;
  mainSheet.getRange(studentIdx + 1, COL_ASSET).setValue(newAsset);

  // 학생별가입예금 시트에 기록
  const logSheet = ss.getSheetByName(SHEET_DEPOSIT_LOG);
  if (!logSheet) {
    // 차감 롤백
    mainSheet.getRange(studentIdx + 1, COL_ASSET).setValue(curAsset);
    return { success: false, msg: '학생별가입예금 시트를 찾을 수 없습니다.' };
  }
  const depId = 'DEP_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 4);
  logSheet.appendRow([
    depId,          // A: 예금ID
    studentName,    // B: 학생명
    amount,         // C: 원금
    rate,           // D: 이자율
    weeks,          // E: 거치기간
    today,          // F: 시작일
    dueDateStr,     // G: 만기일
    '진행중',        // H: 상태
    0,              // I: 지급이자액 (미정)
    '',             // J: 처리일
    prod.prodId     // K: 상품ID
  ]);

  // 히스토리 기록
  const histSheet = ss.getSheetByName(SHEET_HISTORY);
  if (histSheet) {
    histSheet.appendRow([
      today, studentName, mainData[studentIdx][COL_BRAND - 1],
      0, -amount,
      mainData[studentIdx][COL_VALUE - 1], newAsset,
      `[예금가입] ${prod.prodName} ${weeks}주 $${amount.toLocaleString()} (만기: ${dueDateStr})`
    ]);
  }

  /// 캐시 무효화
  CacheService.getScriptCache().remove('student_' + studentName);
  updateRankings();

  return {
    success:    true,
    msg:        `✅ 예금 가입 완료! $${amount.toLocaleString()} 예치 (만기일: ${dueDateStr})`,
    newBalance: newAsset,
    dueDate:    dueDateStr
  };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ── 나의 예금 목록 반환 (Index.html에서 호출) ────────────────────
function getMyDeposits(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEPOSIT_LOG);
  if (!sheet) return [];
  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== String(studentName).trim()) continue;
    let startVal = data[i][5];
    let dueVal   = data[i][6];
    let procVal  = data[i][9];
    if (startVal instanceof Date)
      startVal = Utilities.formatDate(startVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (dueVal instanceof Date)
      dueVal = Utilities.formatDate(dueVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (procVal instanceof Date)
      procVal = Utilities.formatDate(procVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    result.push({
      rowNum:    i + 1,
      depId:     String(data[i][0]).trim(),
      amount:    Number(data[i][2]) || 0,
      rate:      Number(data[i][3]) || 0,
      weeks:     Number(data[i][4]) || 0,
      startDate: String(startVal),
      dueDate:   String(dueVal),
      status:    String(data[i][7]).trim(),
      paidInt:   Number(data[i][8]) || 0,
      procDate:  String(procVal),
      prodId:    String(data[i][10]).trim()
    });
  }
  // 진행중 먼저, 나머지는 최신순
  result.sort(function(a, b) {
    if (a.status === '진행중' && b.status !== '진행중') return -1;
    if (a.status !== '진행중' && b.status === '진행중') return 1;
    return b.rowNum - a.rowNum;
  });
  return result;
}

// ── 중도 해지 (학생 → Index.html에서 호출) ───────────────────────
function cancelDeposit(studentName, depId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_DEPOSIT_LOG);
  if (!logSheet) return { success: false, msg: '학생별가입예금 시트를 찾을 수 없습니다.' };

  const data = logSheet.getDataRange().getValues();
  let targetIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(depId).trim() &&
        String(data[i][1]).trim() === String(studentName).trim() &&
        String(data[i][7]).trim() === '진행중') {
      targetIdx = i; break;
    }
  }
  if (targetIdx === -1)
    return { success: false, msg: '해당 예금을 찾을 수 없거나 이미 처리된 예금입니다.' };

  const amount = Number(data[targetIdx][2]) || 0;

  // 패널티율 — 가입 당시 상품에서 조회, 없으면 현재 활성 상품, 없으면 5%
  let penaltyRate = 5;
  const prodId    = String(data[targetIdx][10]).trim();
  const prodSheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (prodSheet && prodId) {
    const prodData = prodSheet.getDataRange().getValues();
    for (let p = 1; p < prodData.length; p++) {
      if (String(prodData[p][0]).trim() === prodId) {
        penaltyRate = Number(prodData[p][8]) || 5; break;
      }
    }
  }

  const penalty   = Math.floor(amount * penaltyRate / 100);
  const refund    = amount - penalty;
  const today     = _todayStr();

  // 메인 시트에서 학생 자산 증가
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };
  const mainData = mainSheet.getDataRange().getValues();
  let studentIdx = -1;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      studentIdx = i; break;
    }
  }
  if (studentIdx === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

  const curAsset = Number(mainData[studentIdx][COL_ASSET - 1]) || 0;
  const newAsset = curAsset + refund;
  mainSheet.getRange(studentIdx + 1, COL_ASSET).setValue(newAsset);

  // 예금 상태 업데이트
  logSheet.getRange(targetIdx + 1, 8).setValue('중도해지');   // H: 상태
  logSheet.getRange(targetIdx + 1, 9).setValue(0);            // I: 지급이자액
  logSheet.getRange(targetIdx + 1, 10).setValue(today);       // J: 처리일

  // 히스토리 기록
  const histSheet = ss.getSheetByName(SHEET_HISTORY);
  if (histSheet) {
    histSheet.appendRow([
      today, studentName, mainData[studentIdx][COL_BRAND - 1],
      0, refund,
      mainData[studentIdx][COL_VALUE - 1], newAsset,
      `[예금중도해지] 원금 $${amount.toLocaleString()} → 패널티 $${penalty.toLocaleString()} 차감 → 반환 $${refund.toLocaleString()}`
    ]);
  }

  // 우편함 알림
  _sendMail(
    studentName,
    '❌ 예금 중도 해지 처리',
    `예금이 중도 해지되었습니다.\n\n` +
    `원금: $${amount.toLocaleString()}\n` +
    `패널티 (${penaltyRate}%): -$${penalty.toLocaleString()}\n` +
    `반환액: $${refund.toLocaleString()}\n\n` +
    `반환금이 자산에 추가되었습니다.`,
    '알림'
  );

  CacheService.getScriptCache().remove('student_' + studentName);
  updateRankings();

  return {
    success:    true,
    msg:        `중도 해지 완료. 패널티 $${penalty.toLocaleString()} 차감 후 $${refund.toLocaleString()} 반환되었습니다.`,
    newBalance: newAsset,
    penalty,
    refund
  };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ── 만기 체크 및 이자 지급 (트리거 + getStudentData에서 호출) ─────
// studentName 지정 시 해당 학생만, null 이면 전체 처리
function checkAndPayDeposits(studentName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_DEPOSIT_LOG);
  if (!logSheet) return;

  const data      = logSheet.getDataRange().getValues();
  const today     = _todayStr();
  let   processed = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]).trim() !== '진행중') continue;
    if (studentName && String(data[i][1]).trim() !== String(studentName).trim()) continue;

    // 만기일 문자열로 비교
    let dueVal = data[i][6];
    if (dueVal instanceof Date)
      dueVal = Utilities.formatDate(dueVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const dueStr  = String(dueVal).substring(0, 10); // "yyyy-MM-dd"
    const dueTs   = new Date(dueStr + 'T12:00:00+09:00').getTime(); // KST 정오
    const nowTs   = new Date().getTime();
    if (dueTs <= nowTs) {
      _payOneDeposit(ss, logSheet, i, data[i]);
      processed++;
    }
  }
  return processed;
}

// ── 내부 헬퍼: 만기 이자 지급 ────────────────────────────────────
function _payOneDeposit(ss, logSheet, rowIdx, row) {
  // 이중 처리 방지: 다시 확인
  const freshStatus = String(logSheet.getRange(rowIdx + 1, 8).getValue()).trim();
  if (freshStatus !== '진행중') return;

  const studentName = String(row[1]).trim();
  const amount      = Number(row[2]) || 0;
  const rate        = Number(row[3]) || 0;
  const today       = _todayStr();

  // 이자 계산
  const grossInt  = Math.floor(amount * rate / 100);   // 세전 이자
  const taxAmount = Math.floor(grossInt * 0.1);        // 소득세 10%
  const netInt    = grossInt - taxAmount;               // 세후 이자
  const totalBack = amount + netInt;                    // 원금 + 세후 이자

  // 메인 시트 업데이트
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;
  const mainData = mainSheet.getDataRange().getValues();
  let studentIdx = -1;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === studentName) {
      studentIdx = i; break;
    }
  }
  if (studentIdx === -1) return;

  const curAsset  = Number(mainData[studentIdx][COL_ASSET - 1]) || 0;
  const curTax    = Number(mainData[studentIdx][COL_TAX - 1])   || 0;
  const newAsset  = curAsset + totalBack;
  const newTax    = curTax + taxAmount;

  mainSheet.getRange(studentIdx + 1, COL_ASSET).setValue(newAsset);
  mainSheet.getRange(studentIdx + 1, COL_TAX  ).setValue(newTax);

  // 예금 상태 업데이트
  logSheet.getRange(rowIdx + 1, 8).setValue('만기');        // H
  logSheet.getRange(rowIdx + 1, 9).setValue(netInt);        // I: 지급이자액(세후)
  logSheet.getRange(rowIdx + 1, 10).setValue(today);        // J: 처리일

  // 히스토리 기록
  const histSheet = ss.getSheetByName(SHEET_HISTORY);
  if (histSheet) {
    histSheet.appendRow([
      today, studentName, mainData[studentIdx][COL_BRAND - 1],
      0, totalBack,
      mainData[studentIdx][COL_VALUE - 1], newAsset,
      `[예금만기] 원금 $${amount.toLocaleString()} + 세후이자 $${netInt.toLocaleString()} (세금 $${taxAmount.toLocaleString()} 복지기금 납부)`
    ]);
  }

  // 예금 만기 우편함 알림
  _sendMail(
    studentName,
    '🎉 예금 만기 지급 완료!',
    `예금이 만기되어 원금과 이자가 지급되었습니다.\n\n` +
    `원금:         $${amount.toLocaleString()}\n` +
    `세전 이자 (${rate}%): $${grossInt.toLocaleString()}\n` +
    `소득세 (10%): -$${taxAmount.toLocaleString()}\n` +
    `─────────────────\n` +
    `실수령액:      $${totalBack.toLocaleString()}\n\n` +
    `수고하셨습니다! 💰`,
    '보상'
  );

  CacheService.getScriptCache().remove('student_' + studentName);
}
