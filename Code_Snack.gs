// 간식 시세 계산 (재고 100%→1배, 20% 이하→2.5배 상한, 선형 상승)
function getSnackData() {
  const snackSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SNACK);
  if (!snackSheet) return [];
  const sData  = snackSheet.getDataRange().getValues();
  const result = [];
  const _emg         = _getActiveEmergency();
  const _inflationOn = _emg && _emg.type === '하이퍼인플레이션';
  for (let n = 1; n < sData.length; n++) {
    if (!sData[n][0]) continue;
    const basePrice    = Number(sData[n][1]) || 0;
    const baseStock    = Number(sData[n][2]) || 1;
    const currentStock = Number(sData[n][3]);
    const weeklyLimit  = (sData[n][4] !== '' && sData[n][4] !== null && sData[n][4] !== undefined)
                         ? Number(sData[n][4]) : null;
    // 재고 100%→1배, 20% 이하→2.5배 상한 (선형)
    // currentStock/baseStock = 1.0(100%)일 때 1배, 0.2(20%)일 때 2.5배
    const ratio      = (currentStock > 0) ? Math.min(1, currentStock / baseStock) : 0;
    const multiplier = Math.max(1, Math.min(2.5, 1 + 1.5 * (1 - ratio) / 0.8));
    // ── 하이퍼인플레이션: 가격 ×2
    const finalPrice = _inflationOn
      ? Math.round(basePrice * multiplier * 2)
      : Math.round(basePrice * multiplier);
    result.push({
      name:        sData[n][0],
      price:       finalPrice,
      stock:       currentStock,
      weeklyLimit: weeklyLimit,
      inflated:    _inflationOn
    });
  }
  return result;
}

// ════════════════════════════════════════════════════════════════
// 7. 간식 판매 처리 (Snackdialog.html 에서 호출)
// ════════════════════════════════════════════════════════════════

// 간식 판매 다이얼로그 초기 데이터 (학생 목록 + 간식 목록)
function getSnackInitData() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
  const students = [];
  for (let i = 1; i < mainData.length; i++) {
    const name = mainData[i][COL_NAME - 1];
    if (name) {
      students.push({
        name:    name,
        brand:   mainData[i][COL_BRAND - 1],
        balance: Number(mainData[i][COL_ASSET - 1]) || 0
      });
    }
  }
  return { students, snacks:        getSnackData(),
    achievements:  getStudentAchievements(studentName)
  };
}

// 간식 구매 실행 (잔액 차감 + 재고 감소 + 시트 기록)
function executeSnackPurchase(studentName, itemName, price) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet  = ss.getSheetByName(SHEET_MAIN);
  const snackSheet = ss.getSheetByName(SHEET_SNACK);
  const mainData   = mainSheet.getDataRange().getValues();
  const dateStr    = _todayStr();

  // 학생 찾기
  let studentRowNum = -1;
  let brand = '';
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      studentRowNum = i + 1; // 시트 실제 행 번호 (헤더 포함, 1-indexed)
      brand = mainData[i][COL_BRAND - 1];
      break;
    }
  }
  if (studentRowNum === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

  // 잔액 확인
  const curAsset = Number(mainSheet.getRange(studentRowNum, COL_ASSET).getValue()) || 0;
  if (curAsset < price) return { success: false, msg: `잔액이 부족합니다! (현재 잔액: $${curAsset})` };

  // 간식 재고 감소
  if (snackSheet) {
    const snackData = snackSheet.getDataRange().getValues();
    let found = false;
    for (let n = 1; n < snackData.length; n++) {
      if (String(snackData[n][0]).trim() === String(itemName).trim()) {
        const currentStock = Number(snackData[n][3]) || 0;
        if (currentStock <= 0) return { success: false, msg: '재고가 없습니다!' };
        snackSheet.getRange(n + 1, 4).setValue(currentStock - 1);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, msg: '해당 간식을 찾을 수 없습니다.' };
  }

  // 자산 차감
  const newAsset = curAsset - price;
  mainSheet.getRange(studentRowNum, COL_ASSET).setValue(newAsset);

  // 자산사용 시트 기록
  ss.getSheetByName(SHEET_SPEND).appendRow([
    dateStr, studentName, brand, `[간식구매] ${itemName}`, price, newAsset, '간식 구매'
  ]);
  // 히스토리 시트 기록
  const curValue = Number(mainSheet.getRange(studentRowNum, COL_VALUE).getValue()) || 0;
  ss.getSheetByName(SHEET_HISTORY).appendRow([
    dateStr, studentName, brand, 0, -price, curValue, newAsset, `[간식구매] ${itemName}`
  ]);

  updateRankings();
  return { success: true, newBalance: newAsset };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}
