// ════════════════════════════════════════════════════════════════
// ██ 정기 예금 시스템
// ════════════════════════════════════════════════════════════════
// ※ [트리거 함수 위치 안내]
//   예금 관련 트리거 설정 함수들은 이 파일 맨 아래에 있습니다:
//
//     • setupDepositTrigger()   — 매일 12:30 만기 처리 트리거 등록 (한 번만 실행)
//     • runDailyDepositCheck()  — 트리거가 매일 자동 호출하는 실제 만기 처리 함수
//
//   예금 관련 코드를 수정·추가할 때는 이 파일만 보면 됩니다.
// ════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════
// ██ 정기 예금 시스템
// 시트: 예금상품 (A=상품ID, B=상품명, C=1주이자율, D=2주이자율,
//                E=3주이자율, F=4주이자율, G=최소금액, H=최대금액,
//                I=패널티율, J=상태, K=론칭일)
//       학생별가입예금 (A=예금ID, B=학생명, C=원금, D=이자율,
//                      E=거치기간(주), F=시작일, G=만기일,
//                      H=상태, I=지급이자액, J=처리일, K=상품ID)
// ════════════════════════════════════════════════════════════════

// ── 시트 행 → 상품 객체 변환 (내부 헬퍼) ──────────────────────────
function _parseDepositProdRow(row) {
  return {
    prodId:    String(row[0]).trim(),   // A
    prodName:  String(row[1]).trim(),   // B
    rate1:     Number(row[2]) || 0,     // C
    rate2:     Number(row[3]) || 0,     // D
    rate3:     Number(row[4]) || 0,     // E
    rate4:     Number(row[5]) || 0,     // F
    minAmount: Number(row[6]) || 500,   // G
    maxAmount: Number(row[7]) || 5000,  // H
    penalty:   Number(row[8]) || 5,     // I
    status:    String(row[9]).trim(),   // J
    launchDate: String(row[10]).trim()  // K
  };
}

// ── 현재 활성 예금 상품 1개 반환 (하위호환용 — 첫 번째 활성 상품) ──
function getActiveDepositProduct() {
  const list = getActiveDepositProducts();
  return list.length > 0 ? list[0] : null;
}

// ── 현재 활성 예금 상품 전체 반환 (복수 상품 지원) ────────────────
function getActiveDepositProducts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9]).trim() === '활성') {  // J열: 상태
      out.push(_parseDepositProdRow(data[i]));
    }
  }
  return out;
}

// ── 특정 상품ID로 상품 1개 조회 (내부/검증용) ─────────────────────
function _getDepositProductById(prodId) {
  if (!prodId) return null;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) {
      return _parseDepositProdRow(data[i]);
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

  // [복수 상품 지원] 기존 활성 상품을 종료하지 않고 그대로 유지 → 동시 운영
  // 상품ID 중복 방지: 같은 날 여러 상품 론칭 시 뒤에 일련번호 부여
  const today  = _todayStr();
  const baseId = 'PROD_' + today.replace(/-/g, '');
  let   prodId = baseId;
  const existIds = {};
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    existIds[String(prodData[i][0]).trim()] = true;
  }
  if (existIds[prodId]) {
    let n = 2;
    while (existIds[baseId + '_' + n]) n++;
    prodId = baseId + '_' + n;
  }

  // 신규 상품 행 추가
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

  return { success: true, msg: `✅ [${prodName}] 상품이 론칭되었습니다. (기존 상품과 함께 동시 운영) 전체 학생에게 우편이 발송되었습니다.` };
}

// ── 관리자: 활성 예금 상품 목록 반환 (AuctionAdmin 관리 패널용) ────
function getAdminDepositProducts() {
  const list = getActiveDepositProducts();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_DEPOSIT_LOG);
  // 상품별 진행중 예금 건수/합계 집계
  const stats = {};
  if (logSheet) {
    const log = logSheet.getDataRange().getValues();
    for (let i = 1; i < log.length; i++) {
      if (String(log[i][7]).trim() !== '진행중') continue;
      const pid = String(log[i][10]).trim();
      if (!stats[pid]) stats[pid] = { count: 0, total: 0 };
      stats[pid].count++;
      stats[pid].total += Number(log[i][2]) || 0;
    }
  }
  return list.map(function(p) {
    const s = stats[p.prodId] || { count: 0, total: 0 };
    p.activeCount = s.count;   // 진행중 예금 건수
    p.activeTotal = s.total;   // 진행중 예금 원금 합계
    return p;
  });
}

// ── 관리자: 특정 예금 상품 종료 (신규 가입만 차단, 기존 예금은 만기까지 유지) ──
function endDepositProduct(prodId) {
  if (!prodId) return { success: false, msg: '상품 정보가 없습니다.' };
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!prodSheet) return { success: false, msg: '예금상품 시트를 찾을 수 없습니다.' };

  const data = prodSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) {
      if (String(data[i][9]).trim() !== '활성')
        return { success: false, msg: '이미 종료된 상품입니다.' };
      prodSheet.getRange(i + 1, 10).setValue('종료');  // J열
      return { success: true, msg: `✅ [${String(data[i][1]).trim()}] 상품을 종료했습니다. (신규 가입만 차단되며, 진행중 예금은 만기까지 정상 유지됩니다.)` };
    }
  }
  return { success: false, msg: '해당 상품을 찾을 수 없습니다.' };
}

// ── 특정 활성 상품 패널티율 수정 (prodId 지정) ────────────────────
// prodId 미지정 시 첫 번째 활성 상품 (하위호환)
function setPenaltyRate(rate, prodId) {
  rate = Number(rate);
  if (isNaN(rate) || rate < 0 || rate > 100)
    return { success: false, msg: '0~100 사이의 숫자를 입력해주세요.' };

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName(SHEET_DEPOSIT_PROD);
  if (!prodSheet) return { success: false, msg: '예금상품 시트를 찾을 수 없습니다.' };

  const data = prodSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const isActive = String(data[i][9]).trim() === '활성';
    const idMatch  = prodId ? (String(data[i][0]).trim() === String(prodId).trim()) : isActive;
    if (idMatch && isActive) {
      prodSheet.getRange(i + 1, 9).setValue(rate);  // I열
      return { success: true, msg: `✅ [${String(data[i][1]).trim()}] 패널티율이 ${rate}%로 변경되었습니다.` };
    }
  }
  return { success: false, msg: '대상 활성 상품이 없습니다.' };
}

// ── 예금 가입 (학생 → Index.html에서 호출) ───────────────────────
function createDeposit(studentName, amount, weeks, prodId) {
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

  // 상품 조회 — prodId 지정 시 해당 상품, 미지정 시 첫 번째 활성 상품(하위호환)
  let prod;
  if (prodId) {
    prod = _getDepositProductById(prodId);
    if (!prod)               return { success: false, msg: '선택한 예금 상품을 찾을 수 없습니다.' };
    if (prod.status !== '활성') return { success: false, msg: '이미 종료된 상품입니다. 다른 상품을 선택해주세요.' };
  } else {
    prod = getActiveDepositProduct();
  }
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
      _nowStr(), studentName, mainData[studentIdx][COL_BRAND - 1],
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
      _nowStr(), studentName, mainData[studentIdx][COL_BRAND - 1],
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
      _nowStr(), studentName, mainData[studentIdx][COL_BRAND - 1],
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


// ════════════════════════════════════════════════════════════════
// ██ 예금 트리거 설정 / 자동 실행 함수
//
// ※ [이동 이력]
//   원래 Code_Emergency.gs에 있던 함수들을 예금 코드 일관성을 위해
//   이 파일(Code_Deposit.gs) 하단으로 옮겼습니다.
//
// ※ [GAS 트리거 안전성]
//   GAS 트리거는 '파일명'이 아닌 '함수명'으로 등록됩니다.
//   함수 이름(runDailyDepositCheck)이 그대로이므로 기존 트리거는
//   파일 이동 후에도 정상 작동합니다.
//
// ※ [트리거 재등록 필요 여부]
//   이미 트리거가 설정되어 있다면 재등록 불필요합니다.
//   트리거가 없는 경우에만 GAS 편집기 메뉴에서
//   setupDepositTrigger() 를 한 번 실행하세요.
// ════════════════════════════════════════════════════════════════

// ── 예금 만기 자동 처리 트리거 설정 (한 번만 실행하면 됩니다) ──────
function setupDepositTrigger() {
  // 기존 동일 트리거가 있으면 먼저 삭제 (중복 방지)
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyDepositCheck') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 매일 12:30에 runDailyDepositCheck() 자동 실행되도록 등록
  ScriptApp.newTrigger('runDailyDepositCheck')
    .timeBased().everyDays(1).atHour(12).nearMinute(30).create();
  SpreadsheetApp.getUi().alert('✅ 매일 12:30 예금 만기 자동 처리 트리거가 설정되었습니다.');
}

// ── 트리거가 매일 자동 호출하는 실제 만기 처리 실행 함수 ────────────
function runDailyDepositCheck() {
  checkAndPayDeposits(null); // null = 전체 학생 처리(정기예금 만기)
  // 적금: 매주 자동 납입 + 이자 누적 + 만기 처리 (다음납입일 도래분만 처리)
  try {
    if (typeof checkAndProcessSavings === 'function') checkAndProcessSavings(null);
  } catch(e) {
    // 적금 시트 미구성 등 예외 시 정기예금 처리에는 영향 없도록 무시
  }
}
