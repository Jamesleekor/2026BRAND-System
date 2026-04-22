// ════════════════════════════════════════════════════════════════
// 물품 거래소 - 인벤토리 시스템
// ════════════════════════════════════════════════════════════════

// ── 인벤토리 시트 자동 생성 ───────────────────────────────────────
function _ensureInventorySheet(ss) {
  let sheet = ss.getSheetByName(SHEET_INVENTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_INVENTORY);
    sheet.appendRow(['구매시기', '구매자', '품목명', '구매가격', '구매수량', '사용수량', '사용여부', '사용시기', '교사확인여부']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  return sheet;
}

// ── 아이템 사용 로그 시트 자동 생성 ──────────────────────────────

// ── 주간 구매 수량 확인 ───────────────────────────────────────────
function _getWeeklyBuyCount(invSheet, studentName, itemName) {
  const now = new Date();
  // 이번 주 월요일 00:00:00 구하기
  const day = now.getDay(); // 0=일, 1=월 ...
  const diff = (day === 0) ? 6 : day - 1;
  const monday = new Date(now);
  monday.setDate(now.getDate() - diff);
  monday.setHours(0, 0, 0, 0);

  const data = invSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    const rowName = String(data[i][1]).trim();
    const rowItem = String(data[i][2]).trim();
    const rowQty  = Number(data[i][4]) || 1;
    if (rowName === studentName && rowItem === itemName && rowDate >= monday) {
      count += rowQty;
    }
  }
  return count;
}

// ── 구매하기 ─────────────────────────────────────────────────────
function buyItem(studentName, itemName, quantity) {
  quantity = quantity || 1;

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }

  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet  = ss.getSheetByName(SHEET_MAIN);
    const snackSheet = ss.getSheetByName(SHEET_SNACK);
    const invSheet   = _ensureInventorySheet(ss);
    const mainData   = mainSheet.getDataRange().getValues();
    const dateStr    = _todayStr();

    // ── 학생 찾기
    let studentRowNum = -1;
    let brand = '';
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
        studentRowNum = i + 1;
        brand = mainData[i][COL_BRAND - 1];
        break;
      }
    }
    if (studentRowNum === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };



    // ── 간식 시트에서 품목 찾기 + 재고/가격 확인
    const snackData = snackSheet.getDataRange().getValues();
    let snackRowNum = -1;
    let currentStock = 0;
    let currentPrice = 0;
    for (let n = 1; n < snackData.length; n++) {
      if (String(snackData[n][0]).trim() === String(itemName).trim()) {
        snackRowNum  = n + 1;
        currentStock = Number(snackData[n][3]) || 0;
        // 현재 시세는 getSnackData()와 동일한 비선형 가격 함수로 계산
        const basePrice  = Number(snackData[n][1]) || 0;
        const baseStock  = Number(snackData[n][2]) || 1;
        // 재고 100%→1배, 20% 이하→2.5배 상한 (getSnackData와 동일 공식)
        const ratio      = (currentStock > 0) ? Math.min(1, currentStock / baseStock) : 0;
        const multiplier = Math.max(1, Math.min(2.5, 1 + 1.5 * (1 - ratio) / 0.8));
        // ── 하이퍼인플레이션: getSnackData()와 동일하게 ×2 적용
        const _emgItem = _getActiveEmergency();
        const _infOn   = _emgItem && _emgItem.type === '하이퍼인플레이션';
        currentPrice = Math.round(basePrice * multiplier * (_infOn ? 2 : 1));
        break;
      }
    }
    if (snackRowNum === -1) return { success: false, msg: '해당 품목을 찾을 수 없습니다.' };
    if (currentStock < quantity) return { success: false, msg: `재고가 부족합니다! (현재 재고: ${currentStock}개)` };

    // ── 주간 구매 제한 확인 (품목별 한도, null이면 제한 없음)
    const weeklyLimit = (snackRowNum !== -1 && snackData[snackRowNum - 1][4] !== '' && snackData[snackRowNum - 1][4] !== null && snackData[snackRowNum - 1][4] !== undefined)
                        ? Number(snackData[snackRowNum - 1][4]) : null;
    if (weeklyLimit !== null) {
      const weeklyCount = _getWeeklyBuyCount(invSheet, studentName, itemName);
      if (weeklyCount + quantity > weeklyLimit) {
        const remaining = Math.max(0, weeklyLimit - weeklyCount);
        return { success: false, msg: `이번 주 "${itemName}" 구매 가능 수량이 ${remaining}개 남았습니다. (주간 한도: ${weeklyLimit}개)` };
      }
    }

    // ── 잔액 확인
    const curAsset  = Number(mainSheet.getRange(studentRowNum, COL_ASSET).getValue()) || 0;
    const totalCost = currentPrice * quantity;
    // ── 자산 동결 체크
    const _emgB = _getActiveEmergency();
    if (_emgB && _emgB.type === '자산 동결') {
      const _usable = Math.floor(curAsset * (_emgB.freezeRate / 100));
      if (totalCost > _usable) return { success: false, msg: `🔒 자산 동결 중! 사용 가능 금액: $${_usable.toLocaleString()} (보유액의 ${_emgB.freezeRate}%)` };
    }
    if (curAsset < totalCost) return { success: false, msg: `잔액이 부족합니다! (필요: $${totalCost.toLocaleString()}, 현재: $${curAsset.toLocaleString()})` };

    // ── 자산 차감
    const newAsset = curAsset - totalCost;
    mainSheet.getRange(studentRowNum, COL_ASSET).setValue(newAsset);

    // ── 재고 차감
    snackSheet.getRange(snackRowNum, 4).setValue(currentStock - quantity);

    // ── 인벤토리 기록 (quantity만큼 행 추가)
    // ── 인벤토리 기록 (1행에 수량 통합, 사용수량 0으로 초기화)
    const timestamp = new Date();
    invSheet.appendRow([timestamp, studentName, itemName, currentPrice, quantity, 0, false, '', false]);

    // ── 자산사용 시트 기록
    const curValue = Number(mainSheet.getRange(studentRowNum, COL_VALUE).getValue()) || 0;
    ss.getSheetByName(SHEET_SPEND).appendRow([
      dateStr, studentName, brand,
      `[물품구매] ${itemName}`, totalCost, newAsset, `수량 ${quantity}개`
    ]);

    // ── 히스토리 시트 기록
    ss.getSheetByName(SHEET_HISTORY).appendRow([
      dateStr, studentName, brand,
      0, -totalCost, curValue, newAsset, `[물품구매] ${itemName} x${quantity}`
    ]);

    // ── 캐시 무효화
    CacheService.getScriptCache().remove('student_' + studentName);

    updateRankings();
    return { success: true, newBalance: newAsset, price: currentPrice, quantity: quantity };

  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── 인벤토리 조회 ─────────────────────────────────────────────────
function getMyInventory(studentName) {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _ensureInventorySheet(ss);
    const data     = invSheet.getDataRange().getValues();

    // 열 인덱스 (0-based)
    // A=0 구매시기, B=1 구매자, C=2 품목명, D=3 구매가격
    // E=4 구매수량, F=5 사용수량, G=6 사용여부, H=7 사용시기, I=8 교사확인여부

    const unusedMap = {}; // key: itemName → 미사용(잔여) 합산
    const usedMap   = {}; // key: itemName → 사용 완료 합산

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(studentName).trim()) continue;

      const itemName  = String(data[i][2]).trim();
      const price     = Number(data[i][3]) || 0;
      const buyQty    = Number(data[i][4]) || 0;
      const useQty    = Number(data[i][5]) || 0;
      const isAllUsed = data[i][6] === true || data[i][6] === 'TRUE';
      const usedAt    = data[i][7] ? Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone(), 'MM/dd HH:mm') : '';
      const leftQty   = buyQty - useQty; // 잔여 수량

      // 잔여 수량이 있으면 미사용 합산
      if (!isAllUsed && leftQty > 0) {
        if (!unusedMap[itemName]) {
          unusedMap[itemName] = { itemName: itemName, price: price, totalQty: 0 };
        }
        unusedMap[itemName].totalQty += leftQty;
      }

      // 사용 완료분 합산 (사용수량 기준)
      if (useQty > 0) {
        if (!usedMap[itemName]) {
          usedMap[itemName] = { itemName: itemName, price: price, totalQty: 0, usedAt: usedAt };
        }
        usedMap[itemName].totalQty += useQty;
        if (usedAt > usedMap[itemName].usedAt) usedMap[itemName].usedAt = usedAt;
      }
    }

    return {
      success: true,
      unused: Object.values(unusedMap),
      used:   Object.values(usedMap)
    };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}

// ── 사용하기 (구매수량 불변, 사용수량 누적 방식) ───────────────────
function useItem(studentName, itemName, useQty) {
  useQty = Number(useQty) || 1;

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _ensureInventorySheet(ss);
    const data     = invSheet.getDataRange().getValues();
    const now      = new Date();

    // 열 인덱스 (0-based)
    // E=4 구매수량, F=5 사용수량, G=6 사용여부, H=7 사용시기

    // 해당 학생 + 품목의 미사용(잔여) 행 수집
    const targetRows = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() !== String(studentName).trim()) continue;
      if (String(data[i][2]).trim() !== String(itemName).trim()) continue;
      if (data[i][6] === true || data[i][6] === 'TRUE') continue; // 전량 사용된 행 제외
      const buyQty  = Number(data[i][4]) || 0;
      const usedQty = Number(data[i][5]) || 0;
      const leftQty = buyQty - usedQty;
      if (leftQty <= 0) continue;
      targetRows.push({ rowNum: i + 1, buyQty: buyQty, usedQty: usedQty, leftQty: leftQty });
    }

    // 잔여 수량 합산
    const totalLeft = targetRows.reduce(function(s, r) { return s + r.leftQty; }, 0);
    if (totalLeft === 0) return { success: false, msg: '사용 가능한 아이템이 없습니다.' };
    if (useQty > totalLeft) return { success: false, msg: `보유 수량(${totalLeft}개)보다 많이 사용할 수 없습니다.` };

    // 앞 행부터 순서대로 사용수량 누적
    let remaining = useQty;
    for (let r = 0; r < targetRows.length && remaining > 0; r++) {
      const row        = targetRows[r];
      const consume    = Math.min(row.leftQty, remaining);
      remaining       -= consume;
      const newUsedQty = row.usedQty + consume;

      invSheet.getRange(row.rowNum, 6).setValue(newUsedQty); // F열: 사용수량 누적

      // 전량 소진 시 사용여부 TRUE + 사용시기 기록
      if (newUsedQty >= row.buyQty) {
        invSheet.getRange(row.rowNum, 7).setValue(true); // G열: 사용여부
        invSheet.getRange(row.rowNum, 8).setValue(now);  // H열: 사용시기
      }
    }

    // ── 아이템 사용 로그 시트에 기록 ─────────────────────────────
    const useLogSheet = ss.getSheetByName('아이템사용로그');
    if (useLogSheet) {
      const logRow = useLogSheet.getLastRow() + 1;
      useLogSheet.getRange(logRow, 1).setValue(now);
      useLogSheet.getRange(logRow, 2).setValue(studentName);
      useLogSheet.getRange(logRow, 3).setValue(itemName);
      useLogSheet.getRange(logRow, 4).setValue(useQty);
      useLogSheet.getRange(logRow, 5).setValue(false);
    }
    // ─────────────────────────────────────────────────────────────

    return {
      success : true,
      itemName: itemName,
      useQty  : useQty,
      leftQty : totalLeft - useQty,
      msg     : `${itemName} ${useQty}개 사용이 완료되었습니다.\n간식은 제과점 담당자에게 확인받고 간식을 받아가세요.\n이외 상품은 선생님에게 알려주고 해당 상품을 이용하세요.`
    };

  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
