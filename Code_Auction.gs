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
// 4-b. 슈퍼패스(우선입찰권) — 보유자 조회 & 인벤토리 차감
//      품목명 규칙: 경매 카테고리(자리/1인1역/급식순서) → "슈퍼패스(카테고리)"
//      예) "자리" → "슈퍼패스(자리)"
// ════════════════════════════════════════════════════════════════

// 공백 무시 비교용 (예: "슈퍼패스 (자리)" 같은 실수 방지)
function _spStripSpace_(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }

// 카테고리로부터 슈퍼패스 품목명 생성 (실제 인벤토리 품목명과 매핑)
//   경매관리 A열 카테고리 → 물품거래소/인벤토리 품목명
//   ⚠️ '급식순서' 카테고리는 패스 이름이 '급식'이므로 단순 치환 불가 → 명시 매핑
function _superPassItemName_(category) {
  var c = String(category == null ? '' : category).trim();
  var map = {
    '자리'     : '[경매형] 자리 슈퍼패스(1회권)',
    '1인1역'   : '[경매형] 1인1역 슈퍼패스(1회권)',
    '급식순서' : '[경매형] 급식 슈퍼패스(1회권)'
  };
  return map[c] || ('[경매형] ' + c + ' 슈퍼패스(1회권)');
}

// 특정 카테고리 슈퍼패스를 "사용 가능하게" 보유한 학생 목록
// (AuctionAdmin.html 에서 상품 선택/송출 시 호출)
function getSuperPassHolders(category) {
  try {
    const passItemName = _superPassItemName_(category);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName(SHEET_INVENTORY);
    if (!invSheet) return { success: true, holders: [], passItemName: passItemName };

    const data = invSheet.getDataRange().getValues();
    // 인벤토리 열(0-based): A0 구매시기, B1 구매자, C2 품목명,
    //   D3 구매가격, E4 구매수량, F5 사용수량, G6 사용여부, H7 사용시기, I8 교사확인
    const targetKey = _spStripSpace_(passItemName);
    const remainMap = {}; // 이름 → 잔여 수량 합산
    for (let i = 1; i < data.length; i++) {
      if (_spStripSpace_(data[i][2]) !== targetKey) continue;
      if (data[i][6] === true || data[i][6] === 'TRUE') continue; // 전량 사용 제외
      const buyQty  = Number(data[i][4]) || 0;
      const usedQty = Number(data[i][5]) || 0;
      const left    = buyQty - usedQty;
      if (left <= 0) continue;
      const name = String(data[i][1]).trim();
      remainMap[name] = (remainMap[name] || 0) + left;
    }

    const holders = Object.keys(remainMap).map(function(n) {
      return { name: n, count: remainMap[n] };
    });
    return { success: true, holders: holders, passItemName: passItemName };
  } catch (e) {
    return { success: false, msg: e.message, holders: [] };
  }
}

// 슈퍼패스 1개 차감 (executeAuctionSold 내부에서 호출 — 이미 락 보유 중이므로 락 미사용)
// 성공 시 true, 사용 가능한 패스가 없으면 false
function _consumeSuperPassInline_(ss, studentName, passItemName) {
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  if (!invSheet) return false;
  const data    = invSheet.getDataRange().getValues();
  const nameKey = String(studentName).trim();
  const passKey = _spStripSpace_(passItemName);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== nameKey) continue;
    if (_spStripSpace_(data[i][2]) !== passKey) continue;
    if (data[i][6] === true || data[i][6] === 'TRUE') continue;
    const buyQty  = Number(data[i][4]) || 0;
    const usedQty = Number(data[i][5]) || 0;
    if (buyQty - usedQty <= 0) continue;

    const newUsed = usedQty + 1;
    const fully   = newUsed >= buyQty;
    // F~H(사용수량/사용여부/사용시기) 3칸 한 번에 쓰기 (useItem과 동일 패턴)
    invSheet.getRange(i + 1, 6, 1, 3).setValues([[ newUsed, fully, fully ? new Date() : '' ]]);
    return true;
  }
  return false;
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
function executeAuctionSold(studentInfo, itemDetails, price, roundNum, superPassInfo) {
  // ── [진단용 타이머] 각 단계 소요시간(ms) 측정 ──
  const _t0 = new Date().getTime();
  const _T  = {};
  const _mark = function(label) { _T[label] = new Date().getTime() - _t0; };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  _mark('lock');
  try {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const dateStr   = _nowStr();
  _mark('open');

  const curAsset = Number(mainSheet.getRange(studentInfo.rowIdx + 1, COL_ASSET).getValue()) || 0;
  // ── 자산 동결 체크
  const _emgA = _getActiveEmergency();
  if (_emgA && _emgA.type === '자산 동결') {
    const _usable = Math.floor(curAsset * (_emgA.freezeRate / 100));
    if (price > _usable) return { success: false, msg: `🔒 자산 동결 중! 사용 가능 금액: $${_usable.toLocaleString()} (보유액의 ${_emgA.freezeRate}%)` };
  }
  _mark('emergency');
  if (curAsset < price) return { success: false, msg: '잔액이 부족합니다!' };

  const newAsset = curAsset - price;
  mainSheet.getRange(studentInfo.rowIdx + 1, COL_ASSET).setValue(newAsset);

  const curValue = Number(mainSheet.getRange(studentInfo.rowIdx + 1, COL_VALUE).getValue()) || 0;

  // 자산사용 시트에 기록
  ss.getSheetByName(SHEET_SPEND).appendRow([
    _todayStr(), studentInfo.name, studentInfo.brand,
    `[경매낙찰] ${itemDetails.name}`, price, newAsset, '재판매 불가/무료 나눔만 가능', dateStr
  ]);
  // 히스토리 시트에 기록
  ss.getSheetByName(SHEET_HISTORY).appendRow([
    dateStr, studentInfo.name, studentInfo.brand,
    0, -price, curValue, newAsset, `[경매낙찰] ${itemDetails.name}`
  ]);
  _mark('write_logs');

  // ── 슈퍼패스(우선입찰권) 사용 처리 ──────────────────────────────
  // 교사가 어드민에서 슈퍼패스 사용을 체크한 경우에만 동작.
  // 인벤토리에서 패스 1개를 차감하고 자산사용/히스토리에 사용 기록을 남긴다.
  // (금액 변동 0 — 구매 시 이미 결제됨. 낙찰 자체는 위에서 정상 처리됨)
  // 인벤토리에 사용 가능한 패스가 없어도 낙찰은 절대 막지 않고 경고만 반환.
  let _superPassConsumed = false;
  let _superPassWarning  = '';
  try {
    if (superPassInfo && superPassInfo.used && superPassInfo.passItemName) {
      _superPassConsumed = _consumeSuperPassInline_(ss, studentInfo.name, superPassInfo.passItemName);
      if (_superPassConsumed) {
        ss.getSheetByName(SHEET_SPEND).appendRow([
          _todayStr(), studentInfo.name, studentInfo.brand,
          `[슈퍼패스사용] ${superPassInfo.passItemName}`, 0, newAsset,
          `경매 우선입찰권 사용 → ${itemDetails.name}`, dateStr
        ]);
        ss.getSheetByName(SHEET_HISTORY).appendRow([
          dateStr, studentInfo.name, studentInfo.brand,
          0, 0, curValue, newAsset,
          `[슈퍼패스사용] ${superPassInfo.passItemName} → ${itemDetails.name}`
        ]);
      } else {
        _superPassWarning = `⚠️ ${studentInfo.name} 학생의 인벤토리에서 사용 가능한 `
          + `${superPassInfo.passItemName}을(를) 찾지 못해 슈퍼패스 사용 기록은 생략되었습니다. `
          + `(낙찰은 정상 처리되었습니다)`;
      }
    }
  } catch (e) {
    _superPassWarning = '슈퍼패스 처리 중 오류: ' + e.message + ' (낙찰은 정상 처리되었습니다)';
  }
  // ────────────────────────────────────────────────────────────────

  // 경매관리 시트에 낙찰가 기록 (n차 경매 해당 열에)
  try {
    const mgmtSheet = ss.getSheetByName(SHEET_AUCTION);
    if (mgmtSheet && roundNum) {
      const parts      = itemDetails.name.split(' - ');
      const category   = parts[0].trim();
      const detailName = parts[1] ? parts[1].trim() : '';
      // [성능] 전체 범위(C~L 회차열 포함) 대신 A:B(카테고리/상세명) 두 열만 읽어 매칭
      const mgmtLast = mgmtSheet.getLastRow();
      if (mgmtLast >= 2) {
        const mgmtKeys = mgmtSheet.getRange(2, 1, mgmtLast - 1, 2).getValues();
        for (let i = 0; i < mgmtKeys.length; i++) {
          if (String(mgmtKeys[i][0]).trim() === category &&
              String(mgmtKeys[i][1]).trim() === detailName) {
            mgmtSheet.getRange(i + 2, roundNum + 2).setValue(price);
            break;
          }
        }
      }
    }
  } catch (e) {
    console.log('경매관리 기록 오류: ' + e.message);
  }
  _mark('mgmt');

  // [성능] 랭킹 재계산(1.4초)·Firebase 동기화(2.5초)는 표시 최신화용일 뿐
  //   데이터 정확성(자산 차감/로그/경매관리 가격)은 위에서 이미 시트에 기록됨.
  //   매 낙찰마다 반복하지 않고, '경매 종료'(getTodayAuctionResults) 시 한 번에 일괄 처리한다.
  //   → 학생 대시보드의 잔액/순위는 경매 종료 후 일괄 갱신된다.
  //   (어드민 입찰표 잔액은 클라이언트에서 즉시 갱신되고, 서버는 매번 시트에서 재검증하므로 안전)
  _mark('rankings');  // (스킵)
  _mark('fb_sync');   // (스킵)
  // 낙찰 애니메이션 상태 송출
  setAuctionState({
    status:     'sold',
    itemName:   itemDetails.name,
    winner:     studentInfo.name,
    finalPrice: price,
    superPass:  !!(superPassInfo && superPassInfo.used)
  });
  _mark('state');

  // 단계별 소요시간을 문자열로 정리 (가장 오래 걸린 단계 식별용)
  const _timingStr = '⏱️ 진단 (총 ' + _T.state + 'ms)\n'
    + '· 락 대기: ' + _T.lock + 'ms\n'
    + '· 시트 열기: ' + (_T.open - _T.lock) + 'ms\n'
    + '· 잔액/동결 체크: ' + (_T.emergency - _T.open) + 'ms\n'
    + '· 로그 기록: ' + (_T.write_logs - _T.emergency) + 'ms\n'
    + '· 경매관리 기록: ' + (_T.mgmt - _T.write_logs) + 'ms\n'
    + '· (랭킹/Firebase는 종료 시 일괄 처리)\n'
    + '· 상태 송출: ' + (_T.state - _T.mgmt) + 'ms';
  Logger.log(_timingStr);

  return { success: true, newBalance: newAsset, __timing: _timingStr,
           superPassConsumed: _superPassConsumed, superPassWarning: _superPassWarning };
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
  // [성능] 경매 진행 중에는 낙찰마다 랭킹/Firebase 동기화를 생략했으므로,
  //   경매 종료(결과 집계) 시점에 전체 학생 랭킹 재계산 + Firebase 일괄 동기화를 1회 수행한다.
  //   (실패해도 결과 집계는 정상 진행되도록 try로 감쌈)
  try { updateRankings(); } catch (e) { Logger.log('[경매 종료 일괄 동기화 오류] ' + e.message); }

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