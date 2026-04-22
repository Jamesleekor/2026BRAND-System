// ════════════════════════════════════════════════════════════════
// 4. 관리자용 경매 초기 데이터 (AuctionAdmin.html 에서 호출)
// ════════════════════════════════════════════════════════════════
function getAuctionInitData() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const mainData    = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
  const auctionData = ss.getSheetByName(SHEET_AUCTION).getDataRange().getValues();

  const students = [];
  for (let i = 1; i < mainData.length; i++) {
    if (mainData[i][COL_NAME - 1]) {
      students.push({
        rowIdx:  i,
        name:    mainData[i][COL_NAME - 1],
        brand:   mainData[i][COL_BRAND - 1],
        balance: Number(mainData[i][COL_ASSET - 1]) || 0
      });
    }
  }

  const items = [];
  for (let j = 1; j < auctionData.length; j++) {
    if (auctionData[j][0] && auctionData[j][1]) {
      let avgPrice = Number(auctionData[j][11]); // L열 = 다음달 시작가
      if (!avgPrice || avgPrice <= 0) avgPrice = 100;
      else avgPrice = Math.round(avgPrice);

      // ── 하이퍼인플레이션 비상사태 시 시작가 2배 적용
      const _emgAuction = _getActiveEmergency();
      if (_emgAuction && _emgAuction.type === '하이퍼인플레이션') {
        avgPrice = avgPrice * 2;
      }

      items.push({
        // 이름 형식: "카테고리 - 상세명" (executeAuctionSold에서 split(' - ')로 파싱)
        name:       `${auctionData[j][0]} - ${auctionData[j][1]}`,
        startPrice: avgPrice
      });
    }
  }
  return { students, items };
}


// ════════════════════════════════════════════════════════════════
// 5. 경매 상태 관리 (캐시 우선 → Properties 백업)
// ════════════════════════════════════════════════════════════════
function setAuctionState(stateObj) {
  const stateStr = JSON.stringify(stateObj);
  CacheService.getScriptCache().put('AUCTION_STATE', stateStr, 21600);
  PropertiesService.getScriptProperties().setProperty('AUCTION_STATE', stateStr);
  return stateObj;
}

function getAuctionState() {
  let stateStr = CacheService.getScriptCache().get('AUCTION_STATE');
  if (!stateStr) {
    stateStr = PropertiesService.getScriptProperties().getProperty('AUCTION_STATE');
    if (stateStr) CacheService.getScriptCache().put('AUCTION_STATE', stateStr, 21600);
  }
  return stateStr ? JSON.parse(stateStr) : { status: 'idle' };
}

function addAuctionTime(ms) {
  let stateStr = CacheService.getScriptCache().get('AUCTION_STATE')
               || PropertiesService.getScriptProperties().getProperty('AUCTION_STATE');
  if (!stateStr) return false;
  const state = JSON.parse(stateStr);
  if (state.status !== 'bidding' && state.status !== 'failed' && state.status !== 'failed_final') return false;
  state.endTime = Number(state.endTime) + Number(ms);
  const newStr = JSON.stringify(state);
  CacheService.getScriptCache().put('AUCTION_STATE', newStr, 21600);
  PropertiesService.getScriptProperties().setProperty('AUCTION_STATE', newStr);
  return true;
}


// ════════════════════════════════════════════════════════════════
// 6. 경매 낙찰 처리 (AuctionAdmin.html 에서 호출)
// ════════════════════════════════════════════════════════════════
function executeAuctionSold(studentInfo, itemDetails, price, roundNum) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const dateStr   = _todayStr();

  const curAsset = Number(mainSheet.getRange(studentInfo.rowIdx + 1, COL_ASSET).getValue()) || 0;
  // ── 자산 동결 체크
  const _emgA = _getActiveEmergency();
  if (_emgA && _emgA.type === '자산 동결') {
    const _usable = Math.floor(curAsset * (_emgA.freezeRate / 100));
    if (price > _usable) return { success: false, msg: `🔒 자산 동결 중! 사용 가능 금액: $${_usable.toLocaleString()} (보유액의 ${_emgA.freezeRate}%)` };
  }
  if (curAsset < price) return { success: false, msg: '잔액이 부족합니다!' };

  const newAsset = curAsset - price;
  mainSheet.getRange(studentInfo.rowIdx + 1, COL_ASSET).setValue(newAsset);

  const curValue = Number(mainSheet.getRange(studentInfo.rowIdx + 1, COL_VALUE).getValue()) || 0;

  // 자산사용 시트에 기록
  ss.getSheetByName(SHEET_SPEND).appendRow([
    dateStr, studentInfo.name, studentInfo.brand,
    `[경매낙찰] ${itemDetails.name}`, price, newAsset, '재판매 불가/무료 나눔만 가능'
  ]);
  // 히스토리 시트에 기록
  ss.getSheetByName(SHEET_HISTORY).appendRow([
    dateStr, studentInfo.name, studentInfo.brand,
    0, -price, curValue, newAsset, `[경매낙찰] ${itemDetails.name}`
  ]);

  // 경매관리 시트에 낙찰가 기록 (n차 경매 해당 열에)
  try {
    const mgmtSheet = ss.getSheetByName(SHEET_AUCTION);
    if (mgmtSheet && roundNum) {
      const parts      = itemDetails.name.split(' - ');
      const category   = parts[0].trim();
      const detailName = parts[1] ? parts[1].trim() : '';
      const mgmtData   = mgmtSheet.getDataRange().getValues();
      for (let i = 1; i < mgmtData.length; i++) {
        if (String(mgmtData[i][0]).trim() === category &&
            String(mgmtData[i][1]).trim() === detailName) {
          mgmtSheet.getRange(i + 1, roundNum + 2).setValue(price);
          break;
        }
      }
    }
  } catch (e) {
    console.log('경매관리 기록 오류: ' + e.message);
  }

  updateRankings();
  // 낙찰 애니메이션 상태 송출
  setAuctionState({
    status:     'sold',
    itemName:   itemDetails.name,
    winner:     studentInfo.name,
    finalPrice: price
  });
  return { success: true, newBalance: newAsset };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ════════════════════════════════════════════════════════════════
// 6-b. 유찰 기록 (AuctionAdmin.html 에서 호출)
// ════════════════════════════════════════════════════════════════
function recordAuctionFail(itemName, failCount, roundNum) {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const dateStr   = _todayStr();
    const timeStr   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');

    // 유찰로그 시트 자동 생성
    let logSheet = ss.getSheetByName('경매유찰로그');
    if (!logSheet) {
      logSheet = ss.insertSheet('경매유찰로그');
      logSheet.appendRow(['날짜', '시각', '회차', '상품명', '유찰구분', '누적유찰횟수']);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e74c3c').setFontColor('white');
    }

    // 유찰 구분 텍스트
    const failLabel = failCount === 1 ? '1차 유찰'
                    : failCount === 2 ? '2차 유찰(재경매)'
                    : '최종 유찰';

    logSheet.appendRow([dateStr, timeStr, roundNum + '차', itemName, failLabel, failCount]);
    return { success: true };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}

// 오늘의 경매 종료 결과 (전체 학생 포함 - 낙찰 없는 학생도 빈 배열로 포함)
function getTodayAuctionResults() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();

  // 전체 학생으로 초기화 (빈 배열)
  const results = {};
  for (let i = 1; i < mainData.length; i++) {
    const name = mainData[i][COL_NAME - 1];
    if (name) results[name] = [];
  }

  // 오늘 날짜의 경매낙찰 항목 채우기
  const spendSheet = ss.getSheetByName(SHEET_SPEND);
  if (spendSheet) {
    const data  = spendSheet.getDataRange().getValues();
    const today = _todayStr();
    for (let i = 1; i < data.length; i++) {
      let rowDate = data[i][0];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        rowDate = String(rowDate);
      }
      if (rowDate === today && String(data[i][3]).includes('[경매낙찰]')) {
        const studentName = String(data[i][1]).trim();
        const itemName    = String(data[i][3]).replace('[경매낙찰] ', '');
        if (results[studentName] !== undefined) {
          results[studentName].push(itemName);
        }
      }
    }
  }
  return results;
}

