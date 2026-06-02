// ════════════════════════════════════════════════════════════════
// ██ 적금(積金) 시스템  — 매주 자동 납입 / 관대형 미납 처리
// ════════════════════════════════════════════════════════════════
// ※ 이자 방식: B방식(회차별 이자)
//    매주 그 시점까지 적립된 누적 원금에 '주당이자율'을 적용해 이자를 누적.
//    → 먼저·꾸준히 넣을수록 이자가 커짐.
//    전부 정상 납입 시 세전이자 = 회차당금액 × 주당이자율% × N(N+1)/2
//
// ※ 미납 처리: 관대형
//    납입일에 잔액 부족 시 그 회차는 '건너뜀'(미납횟수+1).
//    강제 해지 없이 진행되며, 만기에는 실제 적립된 원금·이자만 지급.
//
// ※ 트리거: 기존 '매일 12:30' runDailyDepositCheck() 안에서 함께 호출됨
//    (Code_Deposit.gs 하단 참고). 별도 트리거 등록 불필요.
//
// 시트: 적금상품
//   A=상품ID, B=상품명, C=주당이자율(%), D=회차당최소, E=회차당최대,
//   F=최대회차, G=중도해지패널티율(%), H=상태(활성/종료), I=론칭일
//
// 시트: 학생별적금
//   A=적금ID, B=학생명, C=회차당금액, D=주당이자율, E=총회차(목표),
//   F=납입회차(성공), G=경과회차(틱), H=누적원금, I=누적이자(세전),
//   J=미납횟수, K=시작일, L=다음납입일, M=만기일,
//   N=상태(진행중/만기/중도해지), O=처리일, P=상품ID
// ════════════════════════════════════════════════════════════════

const SHEET_SAVING_PROD = '적금상품';
const SHEET_SAVING_LOG  = '학생별적금';

// ── 시트 보장(없으면 헤더와 함께 생성) ────────────────────────────
function _ensureSavingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prod = ss.getSheetByName(SHEET_SAVING_PROD);
  if (!prod) {
    prod = ss.insertSheet(SHEET_SAVING_PROD);
    prod.appendRow(['상품ID','상품명','주당이자율','회차당최소','회차당최대',
                    '최대회차','패널티율','상태','론칭일']);
  }
  let log = ss.getSheetByName(SHEET_SAVING_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_SAVING_LOG);
    log.appendRow(['적금ID','학생명','회차당금액','주당이자율','총회차',
                   '납입회차','경과회차','누적원금','누적이자','미납횟수',
                   '시작일','다음납입일','만기일','상태','처리일','상품ID']);
  }
  return { prod, log };
}

// ── 시트 행 → 상품 객체 ───────────────────────────────────────────
function _parseSavingProdRow(row) {
  return {
    prodId:     String(row[0]).trim(),   // A
    prodName:   String(row[1]).trim(),   // B
    weeklyRate: Number(row[2]) || 0,     // C 주당이자율(%)
    minPer:     Number(row[3]) || 100,   // D 회차당 최소
    maxPer:     Number(row[4]) || 2000,  // E 회차당 최대
    maxRounds:  Number(row[5]) || 8,     // F 최대회차
    penalty:    Number(row[6]) || 5,     // G 패널티율
    status:     String(row[7]).trim(),   // H
    launchDate: String(row[8]).trim()    // I
  };
}

// ── 활성 적금 상품 전체 반환 (학생 UI 드롭다운용) ─────────────────
function getActiveSavingProducts() {
  const { prod } = _ensureSavingSheets();
  const data = prod.getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]).trim() === '활성') out.push(_parseSavingProdRow(data[i]));
  }
  return out;
}

// ── 상품ID로 단일 상품 조회 ───────────────────────────────────────
function _getSavingProductById(prodId) {
  if (!prodId) return null;
  const { prod } = _ensureSavingSheets();
  const data = prod.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) return _parseSavingProdRow(data[i]);
  }
  return null;
}

// ── 관리자: 신규 적금 상품 론칭 (복수 동시 운영) ──────────────────
// opts: { prodName, weeklyRate, minPer, maxPer, maxRounds, penalty }
function launchSavingProduct(prodName, weeklyRate, minPer, maxPer, maxRounds, penalty) {
  if (!prodName || !prodName.trim())
    return { success: false, msg: '상품명을 입력해주세요.' };
  weeklyRate = Number(weeklyRate);
  if (!weeklyRate || weeklyRate <= 0)
    return { success: false, msg: '주당 이자율을 입력해주세요.' };
  minPer    = Number(minPer)    || 100;
  maxPer    = Number(maxPer)    || 2000;
  maxRounds = Number(maxRounds) || 8;
  penalty   = Number(penalty)   || 5;
  if (minPer > maxPer)
    return { success: false, msg: '회차당 최소금액이 최대금액보다 큽니다.' };
  if (maxRounds < 2 || maxRounds > 52)
    return { success: false, msg: '최대 회차는 2~52 사이로 설정해주세요.' };

  const { prod } = _ensureSavingSheets();
  const today  = _todayStr();
  const baseId = 'SAV_' + today.replace(/-/g, '');
  let   prodId = baseId;
  const exist  = {};
  const pData  = prod.getDataRange().getValues();
  for (let i = 1; i < pData.length; i++) exist[String(pData[i][0]).trim()] = true;
  if (exist[prodId]) { let n = 2; while (exist[baseId + '_' + n]) n++; prodId = baseId + '_' + n; }

  prod.appendRow([prodId, prodName.trim(), weeklyRate, minPer, maxPer,
                  maxRounds, penalty, '활성', today]);

  // 전체 학생 우편 공지
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (mainSheet) {
    const md = mainSheet.getDataRange().getValues();
    for (let i = 1; i < md.length; i++) {
      const name = String(md[i][COL_NAME - 1]).trim();
      if (!name) continue;
      _sendMail(
        name,
        `🐷 새 적금 상품 출시: ${prodName}`,
        `매주 조금씩 모으는 적금 상품이 출시되었습니다!\n\n` +
        `📌 상품명: ${prodName}\n` +
        `📈 주당 이자율: ${weeklyRate}% (먼저·꾸준히 넣을수록 이자가 커져요!)\n` +
        `💵 회차당 납입: $${minPer.toLocaleString()} ~ $${maxPer.toLocaleString()} (100단위)\n` +
        `🗓️ 최대 ${maxRounds}회 (매주 1회 자동 납입)\n` +
        `⚠️ 잔액이 부족하면 그 주는 건너뜁니다(미납). 중도해지 패널티: 원금의 ${penalty}%\n\n` +
        `지금 대시보드에서 가입하세요!`,
        '공지'
      );
    }
  }
  return { success: true, msg: `✅ [${prodName}] 적금 상품이 출시되었습니다. (기존 상품과 동시 운영) 전체 학생에게 우편 발송 완료.` };
}

// ── 관리자: 활성 적금 상품 목록 + 통계 ────────────────────────────
function getAdminSavingProducts() {
  const list = getActiveSavingProducts();
  const { log } = _ensureSavingSheets();
  const stats = {};
  const data = log.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][13]).trim() !== '진행중') continue;  // N=상태
    const pid = String(data[i][15]).trim();                  // P=상품ID
    if (!stats[pid]) stats[pid] = { count: 0, total: 0 };
    stats[pid].count++;
    stats[pid].total += Number(data[i][7]) || 0;             // H=누적원금
  }
  return list.map(function(p) {
    const s = stats[p.prodId] || { count: 0, total: 0 };
    p.activeCount = s.count;
    p.activeTotal = s.total;
    return p;
  });
}

// ── 관리자: 적금 상품 종료 (신규 가입만 차단, 기존 적금은 유지) ────
function endSavingProduct(prodId) {
  if (!prodId) return { success: false, msg: '상품 정보가 없습니다.' };
  const { prod } = _ensureSavingSheets();
  const data = prod.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) {
      if (String(data[i][7]).trim() !== '활성') return { success: false, msg: '이미 종료된 상품입니다.' };
      prod.getRange(i + 1, 8).setValue('종료');  // H열
      return { success: true, msg: `✅ [${String(data[i][1]).trim()}] 적금 상품을 종료했습니다. (진행중 적금은 만기까지 유지)` };
    }
  }
  return { success: false, msg: '해당 상품을 찾을 수 없습니다.' };
}

// ── 관리자: 적금 패널티율 수정 ────────────────────────────────────
function setSavingPenaltyRate(rate, prodId) {
  rate = Number(rate);
  if (isNaN(rate) || rate < 0 || rate > 100)
    return { success: false, msg: '0~100 사이의 숫자를 입력해주세요.' };
  const { prod } = _ensureSavingSheets();
  const data = prod.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const isActive = String(data[i][7]).trim() === '활성';
    const idMatch  = prodId ? (String(data[i][0]).trim() === String(prodId).trim()) : isActive;
    if (idMatch && isActive) {
      prod.getRange(i + 1, 7).setValue(rate);  // G열
      return { success: true, msg: `✅ [${String(data[i][1]).trim()}] 패널티율이 ${rate}%로 변경되었습니다.` };
    }
  }
  return { success: false, msg: '대상 활성 상품이 없습니다.' };
}

// ── 미리보기 계산 (프론트 검증용·공용) ────────────────────────────
// 전부 정상 납입 가정. 세전이자 = perAmount × rate% × N(N+1)/2
function calcSavingPreview(perAmount, rounds, weeklyRate) {
  perAmount  = Number(perAmount) || 0;
  rounds     = Number(rounds) || 0;
  weeklyRate = Number(weeklyRate) || 0;
  const principal = perAmount * rounds;
  const grossInt  = Math.floor(perAmount * (weeklyRate / 100) * rounds * (rounds + 1) / 2);
  const tax       = Math.floor(grossInt * 0.1);
  const netInt    = grossInt - tax;
  return {
    principal: principal,
    grossInt:  grossInt,
    tax:       tax,
    netInt:    netInt,
    total:     principal + netInt
  };
}

// ── 학생: 적금 가입 (첫 회차는 가입 즉시 납입) ────────────────────
function createSaving(studentName, perAmount, rounds, prodId) {
  perAmount = Number(perAmount);
  rounds    = Number(rounds);
  if (!studentName) return { success: false, msg: '학생 정보가 없습니다.' };
  if (!perAmount || perAmount <= 0) return { success: false, msg: '회차당 금액을 입력해주세요.' };
  if (perAmount % 100 !== 0) return { success: false, msg: '금액은 100 단위로 입력해주세요.' };

  const prod = _getSavingProductById(prodId) ||
               (getActiveSavingProducts().length ? getActiveSavingProducts()[0] : null);
  if (!prod) return { success: false, msg: '현재 가입 가능한 적금 상품이 없습니다.' };
  if (prod.status !== '활성') return { success: false, msg: '이미 종료된 상품입니다. 다른 상품을 선택해주세요.' };
  if (rounds < 2 || rounds > prod.maxRounds)
    return { success: false, msg: `회차는 2 ~ ${prod.maxRounds}회 중 선택해주세요.` };
  if (perAmount < prod.minPer) return { success: false, msg: `회차당 최소 금액은 $${prod.minPer.toLocaleString()}입니다.` };
  if (perAmount > prod.maxPer) return { success: false, msg: `회차당 최대 금액은 $${prod.maxPer.toLocaleString()}입니다.` };

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
    const { log } = _ensureSavingSheets();
    const mainData = mainSheet.getDataRange().getValues();
    let sIdx = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) { sIdx = i; break; }
    }
    if (sIdx === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

    const curAsset = Number(mainData[sIdx][COL_ASSET - 1]) || 0;
    if (curAsset < perAmount)
      return { success: false, msg: `잔액이 부족합니다. 첫 회차 $${perAmount.toLocaleString()} 이상 필요합니다. (현재: $${curAsset.toLocaleString()})` };

    // 첫 회차 납입 + 첫 주 이자 누적
    const newAsset = curAsset - perAmount;
    mainSheet.getRange(sIdx + 1, COL_ASSET).setValue(newAsset);

    const firstInt = perAmount * (prod.weeklyRate / 100);   // 1회차 분 이자(세전, 소수 허용 → 만기시 floor)
    const today    = _todayStr();
    const next = new Date(); next.setDate(next.getDate() + 7); next.setHours(12,0,0,0);
    const nextStr = Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const due  = new Date(); due.setDate(due.getDate() + (rounds - 1) * 7); due.setHours(12,0,0,0);
    const dueStr  = Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const savId = 'SAVL_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,4);
    log.appendRow([
      savId, studentName, perAmount, prod.weeklyRate, rounds,
      1, 1, perAmount, firstInt, 0,
      today, (rounds > 1 ? nextStr : ''), dueStr, '진행중', '', prod.prodId
    ]);

    // 히스토리
    const histSheet = ss.getSheetByName(SHEET_HISTORY);
    if (histSheet) {
      histSheet.appendRow([
        _nowStr(), studentName, mainData[sIdx][COL_BRAND - 1],
        0, -perAmount, mainData[sIdx][COL_VALUE - 1], newAsset,
        `[적금가입] ${prod.prodName} 회차당 $${perAmount.toLocaleString()} × ${rounds}회 (1회차 납입)`
      ]);
    }

    CacheService.getScriptCache().remove('student_' + studentName);
    updateRankings();

    const pv = calcSavingPreview(perAmount, rounds, prod.weeklyRate);
    return {
      success: true,
      msg: `✅ 적금 가입 완료! 1회차 $${perAmount.toLocaleString()} 납입. 매주 자동 납입되며 만기 예상 실수령액 약 $${pv.total.toLocaleString()}입니다.`,
      newBalance: newAsset
    };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ── 학생: 나의 적금 목록 ──────────────────────────────────────────
function getMySavings(studentName) {
  const { log } = _ensureSavingSheets();
  const data = log.getDataRange().getValues();
  const out  = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== String(studentName).trim()) continue;
    const fmt = function(v) {
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return String(v);
    };
    out.push({
      rowNum:    i + 1,
      savId:     String(data[i][0]).trim(),
      perAmount: Number(data[i][2]) || 0,
      rate:      Number(data[i][3]) || 0,
      rounds:    Number(data[i][4]) || 0,
      paidCount: Number(data[i][5]) || 0,
      elapsed:   Number(data[i][6]) || 0,
      principal: Number(data[i][7]) || 0,
      accruedInt:Number(data[i][8]) || 0,
      missed:    Number(data[i][9]) || 0,
      startDate: fmt(data[i][10]),
      nextDate:  fmt(data[i][11]),
      dueDate:   fmt(data[i][12]),
      status:    String(data[i][13]).trim(),
      procDate:  fmt(data[i][14]),
      prodId:    String(data[i][15]).trim()
    });
  }
  out.sort(function(a, b) {
    if (a.status === '진행중' && b.status !== '진행중') return -1;
    if (a.status !== '진행중' && b.status === '진행중') return 1;
    return b.rowNum - a.rowNum;
  });
  return out;
}

// ── 학생: 적금 중도 해지 (누적원금에서 패널티 차감 후 반환, 이자 소멸) ──
function cancelSaving(studentName, savId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { log } = _ensureSavingSheets();
    const data = log.getDataRange().getValues();
    let idx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(savId).trim() &&
          String(data[i][1]).trim() === String(studentName).trim() &&
          String(data[i][13]).trim() === '진행중') { idx = i; break; }
    }
    if (idx === -1) return { success: false, msg: '해당 적금을 찾을 수 없거나 이미 처리되었습니다.' };

    const principal = Number(data[idx][7]) || 0;
    const prodId    = String(data[idx][15]).trim();
    let penaltyRate = 5;
    const prodObj = _getSavingProductById(prodId);
    if (prodObj) penaltyRate = prodObj.penalty;

    const penalty = Math.floor(principal * penaltyRate / 100);
    const refund  = principal - penalty;
    const today   = _todayStr();

    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    const mainData  = mainSheet.getDataRange().getValues();
    let sIdx = -1;
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) { sIdx = i; break; }
    }
    if (sIdx === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

    const curAsset = Number(mainData[sIdx][COL_ASSET - 1]) || 0;
    const newAsset = curAsset + refund;
    mainSheet.getRange(sIdx + 1, COL_ASSET).setValue(newAsset);

    log.getRange(idx + 1, 14).setValue('중도해지');  // N
    log.getRange(idx + 1, 15).setValue(today);        // O

    const histSheet = ss.getSheetByName(SHEET_HISTORY);
    if (histSheet) {
      histSheet.appendRow([
        _nowStr(), studentName, mainData[sIdx][COL_BRAND - 1],
        0, refund, mainData[sIdx][COL_VALUE - 1], newAsset,
        `[적금중도해지] 누적원금 $${principal.toLocaleString()} → 패널티 $${penalty.toLocaleString()} 차감 → 반환 $${refund.toLocaleString()} (이자 소멸)`
      ]);
    }
    _sendMail(
      studentName, '❌ 적금 중도 해지 처리',
      `적금이 중도 해지되었습니다.\n\n` +
      `누적 원금: $${principal.toLocaleString()}\n` +
      `패널티 (${penaltyRate}%): -$${penalty.toLocaleString()}\n` +
      `반환액: $${refund.toLocaleString()}\n` +
      `※ 적립 이자는 중도해지 시 지급되지 않습니다.\n\n` +
      `반환금이 자산에 추가되었습니다.`,
      '알림'
    );

    CacheService.getScriptCache().remove('student_' + studentName);
    updateRankings();
    return { success: true, msg: `중도 해지 완료. 패널티 $${penalty.toLocaleString()} 차감 후 $${refund.toLocaleString()} 반환.`, newBalance: newAsset };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ── 주간 자동 납입 + 이자 누적 + 만기 처리 (트리거에서 호출) ───────
// studentName 지정 시 해당 학생만, null이면 전체
function checkAndProcessSavings(studentName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { log } = _ensureSavingSheets();
  const data = log.getDataRange().getValues();
  const todayStr = _todayStr();
  let processed = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][13]).trim() !== '진행중') continue;
    if (studentName && String(data[i][1]).trim() !== String(studentName).trim()) continue;

    // 다음 납입일이 도래한 만큼 회차를 따라잡으며 처리(미실행 누락 대비)
    let guard = 0;
    while (guard++ < 60) {
      const fresh = log.getRange(i + 1, 1, 1, 16).getValues()[0];
      if (String(fresh[13]).trim() !== '진행중') break;
      const rounds  = Number(fresh[4]) || 0;
      const elapsed = Number(fresh[6]) || 0;
      let nextVal = fresh[11];
      if (nextVal instanceof Date) nextVal = Utilities.formatDate(nextVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const nextStr = String(nextVal).substring(0, 10);
      if (!nextStr || elapsed >= rounds) break;
      if (nextStr > todayStr) break;  // 아직 납입일 전
      _processSavingTick(ss, log, i + 1, fresh);
      processed++;
    }
  }
  return processed;
}

// ── 내부: 한 회차(주) 처리 ────────────────────────────────────────
function _processSavingTick(ss, log, rowNum, row) {
  const studentName = String(row[1]).trim();
  const perAmount   = Number(row[2]) || 0;
  const rate        = Number(row[3]) || 0;
  const rounds      = Number(row[4]) || 0;
  let   paidCount   = Number(row[5]) || 0;
  let   elapsed     = Number(row[6]) || 0;
  let   principal   = Number(row[7]) || 0;
  let   accruedInt  = Number(row[8]) || 0;
  let   missed      = Number(row[9]) || 0;

  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();
  let sIdx = -1;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === studentName) { sIdx = i; break; }
  }

  // 1) 납입 시도 (관대형: 잔액 부족 시 건너뜀)
  if (sIdx !== -1) {
    const curAsset = Number(mainData[sIdx][COL_ASSET - 1]) || 0;
    if (curAsset >= perAmount) {
      const newAsset = curAsset - perAmount;
      mainSheet.getRange(sIdx + 1, COL_ASSET).setValue(newAsset);
      principal += perAmount;
      paidCount += 1;
      const histSheet = ss.getSheetByName(SHEET_HISTORY);
      if (histSheet) {
        histSheet.appendRow([
          _nowStr(), studentName, mainData[sIdx][COL_BRAND - 1],
          0, -perAmount, mainData[sIdx][COL_VALUE - 1], newAsset,
          `[적금납입] ${paidCount}회차 $${perAmount.toLocaleString()} (누적 $${principal.toLocaleString()})`
        ]);
      }
    } else {
      missed += 1;
      _sendMail(
        studentName, '⚠️ 적금 납입 실패 (미납)',
        `이번 주 적금 회차 납입에 실패했습니다.\n\n` +
        `필요 금액: $${perAmount.toLocaleString()}\n` +
        `사유: 잔액 부족\n\n` +
        `이번 회차는 건너뜁니다(미납). 적금은 계속 진행되며, 만기에는 실제 적립된 만큼만 지급됩니다.`,
        '알림'
      );
    }
  }

  // 2) 이번 주 이자 누적 (현재 누적 원금 기준)
  accruedInt += principal * (rate / 100);
  elapsed    += 1;

  // 3) 시트 반영
  log.getRange(rowNum, 6).setValue(paidCount);
  log.getRange(rowNum, 7).setValue(elapsed);
  log.getRange(rowNum, 8).setValue(principal);
  log.getRange(rowNum, 9).setValue(accruedInt);
  log.getRange(rowNum, 10).setValue(missed);

  // 다음 납입일 갱신(남은 회차 있을 때만)
  if (elapsed < rounds) {
    const next = new Date(); next.setDate(next.getDate() + 7); next.setHours(12,0,0,0);
    log.getRange(rowNum, 12).setValue(Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  } else {
    log.getRange(rowNum, 12).setValue('');
  }

  // 4) 만기 도달 시 지급
  if (elapsed >= rounds) {
    _paySaving(ss, log, rowNum, studentName, principal, accruedInt, paidCount, missed, rate);
  }

  CacheService.getScriptCache().remove('student_' + studentName);
}

// ── 내부: 적금 만기 지급 ──────────────────────────────────────────
function _paySaving(ss, log, rowNum, studentName, principal, accruedInt, paidCount, missed, rate) {
  const grossInt  = Math.floor(accruedInt);
  const taxAmount = Math.floor(grossInt * 0.1);
  const netInt    = grossInt - taxAmount;
  const totalBack = principal + netInt;

  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();
  let sIdx = -1;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === studentName) { sIdx = i; break; }
  }
  if (sIdx === -1) return;

  const curAsset = Number(mainData[sIdx][COL_ASSET - 1]) || 0;
  const curTax   = Number(mainData[sIdx][COL_TAX - 1])   || 0;
  const newAsset = curAsset + totalBack;
  mainSheet.getRange(sIdx + 1, COL_ASSET).setValue(newAsset);
  mainSheet.getRange(sIdx + 1, COL_TAX).setValue(curTax + taxAmount);

  log.getRange(rowNum, 14).setValue('만기');          // N
  log.getRange(rowNum, 15).setValue(_todayStr());      // O

  const histSheet = ss.getSheetByName(SHEET_HISTORY);
  if (histSheet) {
    histSheet.appendRow([
      _nowStr(), studentName, mainData[sIdx][COL_BRAND - 1],
      0, totalBack, mainData[sIdx][COL_VALUE - 1], newAsset,
      `[적금만기] 원금 $${principal.toLocaleString()} + 세후이자 $${netInt.toLocaleString()} (세금 $${taxAmount.toLocaleString()} 복지기금, 납입 ${paidCount}회/미납 ${missed}회)`
    ]);
  }
  _sendMail(
    studentName, '🎉 적금 만기 지급 완료!',
    `적금이 만기되어 원금과 이자가 지급되었습니다.\n\n` +
    `누적 원금:      $${principal.toLocaleString()} (납입 ${paidCount}회 / 미납 ${missed}회)\n` +
    `세전 이자:      $${grossInt.toLocaleString()}\n` +
    `소득세 (10%):  -$${taxAmount.toLocaleString()}\n` +
    `─────────────────\n` +
    `실수령액:       $${totalBack.toLocaleString()}\n\n` +
    `꾸준히 모으느라 수고하셨습니다! 🐷💰`,
    '보상'
  );
  CacheService.getScriptCache().remove('student_' + studentName);
}
