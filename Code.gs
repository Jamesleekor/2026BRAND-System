// ================================================================
// 학급 브랜드 포인트 시스템 - Code.gs
// ================================================================

// ── [설정] 시트 이름 ──────────────────────────────────────────────
const SHEET_MAIN    = '메인';
const SHEET_HISTORY = '히스토리';
const SHEET_SPEND   = '자산사용';
const SHEET_JOB     = '1인1역';
const SHEET_AUCTION = '경매관리';
const SHEET_TRACKER = '브랜드가치추적';
const SHEET_SNACK   = '간식관리';

// ── [설정] 열 번호 (1-indexed) ────────────────────────────────────
const COL_BRAND  = 1;  // A: 브랜드
const COL_NAME   = 2;  // B: 이름
const COL_VALUE  = 3;  // C: 브랜드가치
const COL_ASSET  = 4;  // D: 자산보유량
const COL_RANK_A = 5;  // E: 랭킹(자산)
const COL_RANK_V = 6;  // F: 랭킹(가치)
const COL_MVP    = 7;  // G: MVP포인트
const COL_TAX    = 8;  // H: 누적납세액


// ════════════════════════════════════════════════════════════════
// 1. 웹앱 진입점 (URL로 접속 시 어떤 화면을 보여줄지 결정)
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const page = e.parameter.page;
  const mode = e.parameter.mode;

  if (page === 'admin' || mode === 'admin') {
    return HtmlService.createTemplateFromFile('AuctionAdmin').evaluate()
      .setTitle('선생님용 경매 제어 패널')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'display') {
    return HtmlService.createTemplateFromFile('AuctionDisplay').evaluate()
      .setTitle('경매 실시간 중계')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('우리 반 경제 대시보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ════════════════════════════════════════════════════════════════
// 2. 구글 시트 상단 메뉴 등록
// ════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 B.R.A.N.D 관리')
    .addItem('📅 오늘 포인트 지급',              'openDailyInput')
    .addItem('💸 자산 사용 기록',                'openSpendDialog')
    .addItem('🍿 간식 판매 처리',                'openSnackDialog')
    .addSeparator()
    .addItem('🏆 MVP 포인트 지급',               'openMvpDialog')
    .addItem('📊 학생별 히스토리',               'openHistoryDialog')
    .addSeparator()
    .addItem('💾 [필수] 오늘 브랜드 가치 최종 기록', 'finalizeDailyTracker')
    .addItem('🔄 랭킹 새로고침',                  'updateRankings')
    .addItem('⚠️ 마지막 입력 취소(Undo)',          'undoLastHistory')
    .addToUi();
}

function finalizeDailyTracker() {
  _updateTracker(_todayStr(), null);
  SpreadsheetApp.getUi().alert('✅ 오늘의 브랜드 가치가 추적 시트에 최종 기록되었습니다.');
}


// ════════════════════════════════════════════════════════════════
// 3. 학생 대시보드 데이터 (Index.html 에서 호출)
// ════════════════════════════════════════════════════════════════
function getStudentData(studentName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();

  // 해당 학생 행 찾기
  let studentRow = null;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      studentRow = mainData[i];
      break;
    }
  }
  if (!studentRow) return { success: false, msg: '학생을 찾을 수 없습니다. 이름을 다시 확인해주세요.' };

  // 전체 반 누적 복지 기금 (H열 합산)
  let totalTax = 0;
  for (let i = 1; i < mainData.length; i++) {
    totalTax += Number(mainData[i][COL_TAX - 1]) || 0;
  }

  // 1인1역 데이터 (A: 이름, B: 직업명, C: 일급, D: 담당구역)
  const jobSheet = ss.getSheetByName(SHEET_JOB);
  const jobData  = jobSheet.getDataRange().getValues();
  let jobResult  = { title: '미배정', salary: 0, area: '-' };
  for (let j = 1; j < jobData.length; j++) {
    if (String(jobData[j][0]).trim() === String(studentName).trim()) {
      jobResult = {
        title:  jobData[j][1] || '미배정',
        salary: Number(jobData[j][2]) || 0,
        area:   jobData[j][3] || '-'
      };
      break;
    }
  }

  // 경매 시세 (A+B: 항목명 조합, L열: 다음달 입찰 시작가)
  const auctionSheet  = ss.getSheetByName(SHEET_AUCTION);
  const auctionPrices = [];
  if (auctionSheet) {
    const aData = auctionSheet.getDataRange().getValues();
    for (let m = 1; m < aData.length; m++) {
      if (!aData[m][0]) continue;
      auctionPrices.push({
        item:  `[${aData[m][0]}] ${aData[m][1] || ''}`.trim(),
        price: Number(aData[m][11]) || 0  // L열 = 인덱스 11
      });
    }
  }

  // 브랜드 등급 계산
  const honor = Number(studentRow[COL_VALUE - 1]) || 0;
  let tier = { name: '새싹', icon: '🌱', min: 0, max: 5000 };
  if      (honor >= 100000) tier = { name: '그랜드마스터', icon: '🏆', min: 100000, max: 100000 };
  else if (honor >= 75000)  tier = { name: '마스터',       icon: '👑', min: 75000,  max: 100000 };
  else if (honor >= 50000)  tier = { name: '다이아몬드',   icon: '💠', min: 50000,  max: 75000  };
  else if (honor >= 30000)  tier = { name: '플래티넘',     icon: '💎', min: 30000,  max: 50000  };
  else if (honor >= 20000)  tier = { name: '골드',         icon: '🥇', min: 20000,  max: 30000  };
  else if (honor >= 10000)  tier = { name: '실버',         icon: '🥈', min: 10000,  max: 20000  };
  else if (honor >= 5000)   tier = { name: '브론즈',       icon: '🥉', min: 5000,   max: 10000  };

  return {
    success:       true,
    personal: {
      name:        studentRow[COL_NAME - 1],
      brand:       studentRow[COL_BRAND - 1],
      honor:       honor,
      balance:     Number(studentRow[COL_ASSET - 1]) || 0,
      honorRank:   studentRow[COL_RANK_V - 1],
      balanceRank: studentRow[COL_RANK_A - 1]
    },
    personalTax:   Number(studentRow[COL_TAX - 1]) || 0,
    classTotalTax: totalTax,
    job:           jobResult,
    auctionPrices: auctionPrices,
    tierData:      tier,
    snacks:        getSnackData()
  };
}

// 간식 시세 계산 (재고 비율에 따라 최대 5배까지 비선형 상승)
function getSnackData() {
  const snackSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SNACK);
  if (!snackSheet) return [];
  const sData = snackSheet.getDataRange().getValues();
  const result = [];
  for (let n = 1; n < sData.length; n++) {
    if (!sData[n][0]) continue;
    const basePrice    = Number(sData[n][1]) || 0;
    const baseStock    = Number(sData[n][2]) || 1;
    const currentStock = Number(sData[n][3]);
    // 재고 0 → 5배, 재고 가득 → 1배
    const multiplier = (currentStock > 0)
      ? Math.max(1, Math.min(5, baseStock / currentStock))
      : 5;
    result.push({
      name:  sData[n][0],
      price: Math.round(basePrice * multiplier),
      stock: currentStock
    });
  }
  return result;
}


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
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const dateStr   = _todayStr();

  const curAsset = Number(mainSheet.getRange(studentInfo.rowIdx + 1, COL_ASSET).getValue()) || 0;
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
  return { students, snacks: getSnackData() };
}

// 간식 구매 실행 (잔액 차감 + 재고 감소 + 시트 기록)
function executeSnackPurchase(studentName, itemName, price) {
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
}


// ════════════════════════════════════════════════════════════════
// 8. 핵심 시스템 로직 (포인트, MVP, Undo, 랭킹)
// ════════════════════════════════════════════════════════════════

function undoLastHistory() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const main    = ss.getSheetByName(SHEET_MAIN);
  const hist    = ss.getSheetByName(SHEET_HISTORY);
  const lastRow = hist.getLastRow();
  if (lastRow < 2) return '❌ 취소할 기록이 없습니다.';

  const lastData  = hist.getRange(lastRow, 1, 1, 8).getValues()[0];
  const name      = lastData[1];
  const points    = Number(lastData[3]) || 0;
  const assetGain = Number(lastData[4]) || 0;

  const mainData = main.getDataRange().getValues();
  for (let i = 1; i < mainData.length; i++) {
    if (mainData[i][COL_NAME - 1] === name) {
      const curVal   = Number(main.getRange(i + 1, COL_VALUE).getValue());
      const curAsset = Number(main.getRange(i + 1, COL_ASSET).getValue());
      main.getRange(i + 1, COL_VALUE).setValue(curVal - points);
      main.getRange(i + 1, COL_ASSET).setValue(curAsset - assetGain);
      hist.deleteRow(lastRow);
      updateRankings();
      return `✅ [${name}] 마지막 기록이 취소되었습니다.`;
    }
  }
  return '❌ 학생을 찾지 못했습니다.';
}

function applyDailyPoints(date, entries, taxRate) {
  taxRate = taxRate || 0;
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const main      = ss.getSheetByName(SHEET_MAIN);
  const hist      = ss.getSheetByName(SHEET_HISTORY);
  const mainRange = main.getDataRange();
  const mainData  = mainRange.getValues();
  const histRows  = [];

  entries.forEach(e => {
    const rowIdx       = e.row;
    const curValue     = Number(mainData[rowIdx][COL_VALUE - 1]) || 0;
    const curAsset     = Number(mainData[rowIdx][COL_ASSET - 1]) || 0;
    const curTax       = Number(mainData[rowIdx][COL_TAX - 1])   || 0;
    const taxAmount    = Math.floor(e.points * (taxRate / 100));
    const netAssetGain = e.points - taxAmount;

    mainData[rowIdx][COL_VALUE - 1] = curValue + e.points;      // 브랜드가치: 세금 없이 전액
    mainData[rowIdx][COL_ASSET - 1] = curAsset + netAssetGain;  // 자산보유량: 세금 차감 후
    mainData[rowIdx][COL_TAX - 1]   = curTax + taxAmount;       // 누적납세액 증가

    histRows.push([
      date, e.name, e.brand,
      e.points, netAssetGain,
      curValue + e.points, curAsset + netAssetGain,
      e.note + (taxAmount > 0 ? ` (세금 ${taxAmount})` : '')
    ]);
  });

  mainRange.setValues(mainData);
  if (histRows.length > 0) {
    hist.getRange(hist.getLastRow() + 1, 1, histRows.length, 8).setValues(histRows);
  }
  updateRankings();
  return `✅ ${entries.length}명 포인트 지급 완료!`;
}

function recordSpend(date, rowIdx, name, brand, category, amount, note) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const main     = ss.getSheetByName(SHEET_MAIN);
  const curAsset = Number(main.getRange(rowIdx + 1, COL_ASSET).getValue()) || 0;
  if (curAsset < amount) return '❌ 자산이 부족합니다!';
  const newAsset = curAsset - amount;
  main.getRange(rowIdx + 1, COL_ASSET).setValue(newAsset);
  ss.getSheetByName(SHEET_SPEND).appendRow([date, name, brand, category, amount, newAsset, note]);
  updateRankings();
  return '✅ 자산 차감 완료!';
}

function grantMvp(date, rowIdx, name, brand, amount, note) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const curValue  = Number(mainSheet.getRange(rowIdx + 1, COL_VALUE).getValue()) || 0;
  const curAsset  = Number(mainSheet.getRange(rowIdx + 1, COL_ASSET).getValue()) || 0;
  mainSheet.getRange(rowIdx + 1, COL_VALUE).setValue(curValue + amount);
  mainSheet.getRange(rowIdx + 1, COL_ASSET).setValue(curAsset + amount);
  ss.getSheetByName(SHEET_HISTORY).appendRow([
    date, name, brand, amount, amount,
    curValue + amount, curAsset + amount, `[MVP] ${note}`
  ]);
  updateRankings();
  return '🏆 MVP 포인트 지급 완료!';
}

function updateRankings() {
  const main    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  const lastRow = main.getLastRow();
  if (lastRow < 2) return;
  const data = main.getRange(2, 1, lastRow - 1, COL_MVP).getValues();
  const vArr = data.map((r, i) => ({ idx: i, v: Number(r[COL_VALUE - 1]) || 0 }));
  const aArr = data.map((r, i) => ({ idx: i, v: Number(r[COL_ASSET - 1]) || 0 }));
  const rV   = _calcRank(vArr);
  const rA   = _calcRank(aArr);
  main.getRange(2, COL_RANK_A, rA.length, 1).setValues(rA.map(r => [r]));
  main.getRange(2, COL_RANK_V, rV.length, 1).setValues(rV.map(r => [r]));
}

function _calcRank(arr) {
  const sorted  = [...arr].sort((a, b) => b.v - a.v);
  const rankMap = {};
  let rank = 1;
  for (let i = 0; i < sorted.length; i++) {
    if (i > 0 && sorted[i].v < sorted[i - 1].v) rank = i + 1;
    rankMap[sorted[i].idx] = rank;
  }
  return arr.map(a => rankMap[a.idx]);
}

function getStudentHistory(name) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_HISTORY).getDataRange().getValues();
  return data.filter(r => String(r[1]) === String(name)).map(r => {
    let d = r[0];
    if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return [d, r[3], r[4], r[5], r[6], r[7]];
  }).reverse();
}


// ════════════════════════════════════════════════════════════════
// 9. 브랜드가치 일별 추적 (버튼으로 하루 마감 시 호출)
// ════════════════════════════════════════════════════════════════
function _updateTracker(date, ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getSheetByName(SHEET_TRACKER);
  const main    = ss.getSheetByName(SHEET_MAIN);
  if (!tracker || !main) return;

  const mainData    = main.getDataRange().getValues();
  let trackerData   = tracker.getDataRange().getValues();

  // 서식 무시하고 실제 데이터가 있는 마지막 행 계산
  let realLastRow = 0;
  for (let i = 0; i < trackerData.length; i++) {
    if (trackerData[i][0]) realLastRow = i + 1;
  }

  // 기존 등록된 학생 이름 목록
  const existingNames = {};
  for (let r = 1; r < trackerData.length; r++) {
    if (trackerData[r][0]) existingNames[trackerData[r][0]] = r + 1;
  }

  // 신규 학생 추가
  const newStudents = [];
  for (let i = 1; i < mainData.length; i++) {
    const name = mainData[i][COL_NAME - 1];
    if (name && !existingNames[name]) newStudents.push([name]);
  }
  if (newStudents.length > 0) {
    tracker.getRange(realLastRow + 1, 1, newStudents.length, 1).setValues(newStudents);
    trackerData = tracker.getDataRange().getValues();
    for (let r = 1; r < trackerData.length; r++) {
      if (trackerData[r][0]) existingNames[trackerData[r][0]] = r + 1;
    }
    realLastRow += newStudents.length;
  }

  // 날짜 열 확인 또는 추가
  const headerRow = tracker.getRange(1, 1, 1, tracker.getLastColumn() || 1).getValues()[0];
  let dateCol = headerRow.indexOf(date) + 1;
  if (dateCol === 0) {
    dateCol = (tracker.getLastColumn() || 1) + 1;
    tracker.getRange(1, dateCol).setValue(date).setBackground('#3d85c8').setFontColor('white');
  }

  // 브랜드가치 값 기록
  const valuesToWrite = new Array(realLastRow - 1).fill(['']);
  for (let i = 1; i < mainData.length; i++) {
    const name = mainData[i][COL_NAME - 1];
    if (name && existingNames[name]) {
      valuesToWrite[existingNames[name] - 2] = [mainData[i][COL_VALUE - 1] || 0];
    }
  }
  tracker.getRange(2, dateCol, valuesToWrite.length, 1).setValues(valuesToWrite);
}


// ════════════════════════════════════════════════════════════════
// 10. 유틸리티
// ════════════════════════════════════════════════════════════════
function getScriptUrl() { return ScriptApp.getService().getUrl(); }
function _todayStr()    { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }


// ════════════════════════════════════════════════════════════════
// 11. 다이얼로그 열기 함수들 (메뉴에서 호출)
// ════════════════════════════════════════════════════════════════
function openDailyInput()    { showModal(getDailyInputHtml(), '📅 오늘 포인트 지급',    700, 600); }
function openSpendDialog()   { showModal(getSpendHtml(),      '💸 자산 사용 기록',      500, 500); }
function openMvpDialog()     { showModal(getMvpHtml(),        '🏆 MVP 포인트 지급',     460, 380); }
function openHistoryDialog() { showModal(getHistoryHtml(),    '📊 학생별 히스토리',     750, 550); }

// 간식 판매 처리: Snackdialog.html 파일을 직접 불러옴
function openSnackDialog() {
  const html = HtmlService.createTemplateFromFile('Snackdialog').evaluate()
    .setWidth(500).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '🍿 간식 판매 처리');
}

function showModal(htmlText, title, w, h) {
  const html = HtmlService.createHtmlOutput(htmlText).setWidth(w).setHeight(h);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}


// ── 일일 포인트 지급 HTML ──────────────────────────────────────
function getDailyInputHtml() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(SHEET_MAIN);
  const data = main.getDataRange().getValues();
  let rowsHtml = '';
  for (let i = 1; i < data.length; i++) {
    const brand = data[i][COL_BRAND - 1], name = data[i][COL_NAME - 1];
    if (!name) continue;
    rowsHtml += `<tr>
      <td style="padding:4px;">${name}<br><small style="color:#666;">${brand}</small></td>
      <td style="padding:4px;"><input type="number" id="pt_${i}" class="ptInput" data-row="${i}" data-name="${name}" data-brand="${brand}" placeholder="점수" step="100" style="width:85px;padding:4px;text-align:right;"></td>
      <td style="padding:4px;"><input type="text" id="note_${i}" placeholder="비고" style="width:140px;padding:4px;"></td>
    </tr>`;
  }
  return `<!DOCTYPE html><html><head><style>
    body{font-family:'Noto Sans KR',sans-serif;margin:0;padding:12px;font-size:14px;}
    h3{margin:0 0 10px;color:#2c3e50;}
    .top-bar{background:#eaf2ff;border-radius:8px;padding:10px 14px;margin-bottom:12px;}
    .top-bar label{font-weight:bold;margin-right:8px;}
    table{border-collapse:collapse;width:100%;font-size:13px;}
    th{background:#3d85c8;color:white;padding:6px 8px;text-align:left;}
    tr:nth-child(even){background:#f5f9ff;}
    .btn-apply{background:#27ae60;color:white;border:none;border-radius:6px;padding:10px 28px;font-size:15px;cursor:pointer;margin-top:12px;width:100%;}
    .btn-fill{background:#3d85c8;color:white;border:none;border-radius:5px;padding:5px 12px;cursor:pointer;margin-left:6px;font-size:13px;}
  </style></head><body>
  <h3>📅 오늘의 포인트 지급</h3>
  <div class="top-bar">
    <label>날짜:</label><input type="date" id="today" value="${_todayStr()}" style="padding:4px;">
    <label style="margin-left:10px;">기본:</label>
    <select id="baseScore" style="padding:4px;">
      <option value="100">100</option><option value="200">200</option>
      <option value="300" selected>300</option><option value="400">400</option><option value="0">0</option>
    </select>
    <label style="margin-left:10px;color:#c0392b;">세금(%):</label>
    <input type="number" id="taxRate" value="10" style="width:40px;padding:4px;">
    <button class="btn-fill" onclick="fillAll()">전체적용</button>
  </div>
  <div style="max-height:300px;overflow-y:auto;border:1px solid #ddd;">
    <table><thead><tr><th>학생명</th><th>포인트</th><th>비고</th></tr></thead>
    <tbody>${rowsHtml}</tbody></table>
  </div>
  <button class="btn-apply" onclick="applyPoints()">🚀 포인트 및 세금 적용하기</button>
  <script>
    function fillAll() {
      var val = document.getElementById('baseScore').value;
      document.querySelectorAll('.ptInput').forEach(function(inp){ if(!inp.value) inp.value = val; });
    }
    function applyPoints() {
      var date = document.getElementById('today').value;
      var taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
      if (!date) return alert('날짜를 입력해주세요.');
      var entries = [];
      document.querySelectorAll('.ptInput').forEach(function(inp) {
        var val = inp.value.trim();
        if (val !== '' && !isNaN(val)) {
          var row = parseInt(inp.dataset.row);
          entries.push({ row: row, name: inp.dataset.name, brand: inp.dataset.brand,
            points: parseInt(val), note: document.getElementById('note_' + row).value.trim() });
        }
      });
      if (entries.length === 0) return alert('입력된 데이터가 없습니다.');
      google.script.run.withSuccessHandler(function(res){ alert(res); google.script.host.close(); })
        .applyDailyPoints(date, entries, taxRate);
    }
  </script></body></html>`;
}


// ── 자산 사용 기록 HTML ───────────────────────────────────────
function getSpendHtml() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
  let options = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_NAME - 1])
      options += `<option value="${i}" data-asset="${data[i][COL_ASSET-1]}">${data[i][COL_NAME-1]} (${data[i][COL_BRAND-1]})</option>`;
  }
  return `<!DOCTYPE html><html><head><style>
    body{font-family:'Noto Sans KR',sans-serif;padding:16px;font-size:14px;}
    label{display:block;margin:10px 0 4px;font-weight:bold;}
    select,input{width:100%;padding:7px;font-size:14px;box-sizing:border-box;}
    .btn{background:#e74c3c;color:white;border:none;border-radius:6px;padding:10px 24px;font-size:15px;cursor:pointer;margin-top:16px;width:100%;}
    .asset-info{background:#fff3cd;border-radius:6px;padding:8px;margin:8px 0;font-weight:bold;}
  </style></head><body>
  <h3>💸 자산 사용 기록</h3>
  <label>날짜</label><input type="date" id="spendDate" value="${_todayStr()}">
  <label>학생 선택</label><select id="studentSel" onchange="updateAsset()">${options}</select>
  <div class="asset-info" id="assetInfo">로딩 중...</div>
  <label>사용 항목</label>
  <select id="spendCategory">
    <option>자리 임대료</option><option>급식순서 변경</option>
    <option>1인1역 변경</option><option>간식 구매</option><option>기타</option>
  </select>
  <label>사용 금액</label><input type="number" id="spendAmt" placeholder="예: 200" step="50">
  <label>비고</label><input type="text" id="spendNote" placeholder="상세 내용">
  <button class="btn" onclick="recordSpend()">💸 자산 차감 적용</button>
  <script>
    function updateAsset() {
      var sel = document.getElementById('studentSel');
      var asset = sel.options[sel.selectedIndex].dataset.asset;
      document.getElementById('assetInfo').textContent = '현재 보유 자산: $' + Number(asset).toLocaleString();
    }
    function recordSpend() {
      var date = document.getElementById('spendDate').value;
      var sel = document.getElementById('studentSel');
      var row = parseInt(sel.value);
      var name = sel.options[sel.selectedIndex].text.split(' (')[0];
      var amt = parseInt(document.getElementById('spendAmt').value);
      if (!amt || amt <= 0) return alert('금액을 입력하세요.');
      google.script.run.withSuccessHandler(function(res){ alert(res); google.script.host.close(); })
        .recordSpend(date, row, name, '',
          document.getElementById('spendCategory').value, amt,
          document.getElementById('spendNote').value);
    }
    updateAsset();
  </script></body></html>`;
}


// ── MVP 지급 HTML ─────────────────────────────────────────────
function getMvpHtml() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN).getDataRange().getValues();
  let options = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_NAME - 1])
      options += `<option value="${i}">${data[i][COL_NAME-1]} (${data[i][COL_BRAND-1]})</option>`;
  }
  return `<!DOCTYPE html><html><head><style>
    body{font-family:'Noto Sans KR',sans-serif;padding:16px;font-size:14px;}
    label{display:block;margin:10px 0 4px;font-weight:bold;}
    select,input{width:100%;padding:7px;font-size:14px;box-sizing:border-box;}
    .btn{background:#f39c12;color:white;border:none;border-radius:6px;padding:10px;font-size:15px;cursor:pointer;margin-top:16px;width:100%;}
  </style></head><body>
  <h3>🏆 MVP 포인트 지급</h3>
  <label>날짜</label><input type="date" id="mvpDate" value="${_todayStr()}">
  <label>학생 선택</label><select id="mvpStudent">${options}</select>
  <label>포인트</label><input type="number" id="mvpAmt" value="1000">
  <button class="btn" onclick="grantMvp()">🏆 MVP 지급하기</button>
  <script>
    function grantMvp() {
      var date = document.getElementById('mvpDate').value;
      var sel  = document.getElementById('mvpStudent');
      google.script.run.withSuccessHandler(function(res){ alert(res); google.script.host.close(); })
        .grantMvp(date, parseInt(sel.value),
          sel.options[sel.selectedIndex].text.split(' (')[0], '',
          parseInt(document.getElementById('mvpAmt').value), 'MVP');
    }
  </script></body></html>`;
}


// ── 히스토리 조회 HTML ────────────────────────────────────────
function getHistoryHtml() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN).getDataRange().getValues();
  let options = '<option value="">-- 학생 선택 --</option>';
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_NAME - 1])
      options += `<option value="${data[i][COL_NAME-1]}">${data[i][COL_NAME-1]}</option>`;
  }
  return `<!DOCTYPE html><html><head><style>
    body{font-family:'Noto Sans KR',sans-serif;padding:16px;font-size:14px;}
    select{width:100%;padding:7px;font-size:14px;margin-bottom:10px;}
    table{border-collapse:collapse;width:100%;}
    th{background:#3d85c8;color:white;padding:6px;}
    td{padding:6px;border-bottom:1px solid #eee;}
    tr:nth-child(even){background:#f9f9f9;}
  </style></head><body>
  <h3>📊 학생별 히스토리</h3>
  <select id="hSel" onchange="loadHistory()">${options}</select>
  <div id="result"></div>
  <script>
    function loadHistory() {
      var name = document.getElementById('hSel').value;
      if (!name) return;
      google.script.run.withSuccessHandler(function(rows) {
        var h = '<table><thead><tr><th>날짜</th><th>획득PT</th><th>자산변동</th><th>비고</th></tr></thead><tbody>';
        rows.forEach(function(r) {
          h += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td>' + r[2] + '</td><td>' + r[5] + '</td></tr>';
        });
        h += '</tbody></table>';
        document.getElementById('result').innerHTML = h;
      }).getStudentHistory(name);
    }
  </script></body></html>`;
}
