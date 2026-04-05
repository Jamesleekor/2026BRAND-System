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
const SHEET_ACH_MASTER  = '업적마스터';
const SHEET_ACH_STUDENT = '학생업적달성';

// ── [설정] 열 번호 (1-indexed) ────────────────────────────────────
const COL_BRAND  = 1;  // A: 브랜드
const COL_NAME   = 2;  // B: 이름
const COL_VALUE  = 3;  // C: 브랜드가치
const COL_ASSET  = 4;  // D: 자산보유량
const COL_RANK_A = 5;  // E: 랭킹(자산)
const COL_RANK_V = 6;  // F: 랭킹(가치)
const COL_MVP    = 7;  // G: MVP포인트
const COL_TAX    = 8;  // H: 누적납세액
const COL_PASSWORD = 9;  // I: 비밀번호


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
    .addItem('🗑️ 업적 캐시 초기화',              'clearAchievementCache') 
    .addItem('🗑️ 전체 캐시 초기화',              'clearAllCache')          
    .addToUi();
}

function finalizeDailyTracker() {
  _updateTracker(_todayStr(), null);
  SpreadsheetApp.getUi().alert('✅ 오늘의 브랜드 가치가 추적 시트에 최종 기록되었습니다.');
}


// ════════════════════════════════════════════════════════════════
// 3. 학생 대시보드 데이터 (Index.html 에서 호출)
// ════════════════════════════════════════════════════════════════
function getStudentData(studentName, password) {
    const cache = CacheService.getScriptCache();
  const cacheKey = 'student_' + studentName;
  
  // 캐시에서 먼저 확인 (10분 유효)
  const cached = cache.get(cacheKey);
  if (cached) {
    const data = JSON.parse(cached);
    // 비밀번호만 재확인
    if (data.success && password === data._password) {
      delete data._password;
      // 복지기금 합계는 캐시를 무시하고 항상 실시간 계산 (기부 후 즉시 반영)
      try {
        const ss2      = SpreadsheetApp.getActiveSpreadsheet();
        const main2    = ss2.getSheetByName(SHEET_MAIN);
        if (main2) {
          const md2 = main2.getDataRange().getValues();
          let liveTax = 0;
          for (let i = 1; i < md2.length; i++) liveTax += Number(md2[i][COL_TAX - 1]) || 0;
          data.classTotalTax = liveTax;
        }
      } catch(e) {}
      return data;
    }
  }

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
  // 비밀번호 확인 (I열 = 인덱스 8)
  const correctPassword = String(studentRow[COL_PASSWORD - 1]).trim();
  const inputPassword = String(password).trim();
  
  if (correctPassword && inputPassword !== correctPassword) {
    return { success: false, msg: '비밀번호가 일치하지 않습니다.' };
  }

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
  else if (honor >= 85000)  tier = { name: '완성된 마스터',       icon: '👑', min: 85000,  max: 100000 };
  else if (honor >= 75000)  tier = { name: '마스터',       icon: '👑', min: 75000,  max: 85000 };
  else if (honor >= 65000)  tier = { name: '찬란한 다이아몬드',   icon: '💠', min: 50000,  max: 75000  };
  else if (honor >= 60000)  tier = { name: '진화한 다이아몬드',   icon: '💠', min: 50000,  max: 65000  };
  else if (honor >= 55000)  tier = { name: '성장한 다이아몬드',   icon: '💠', min: 50000,  max: 60000  };
  else if (honor >= 50000)  tier = { name: '거친 다이아몬드',   icon: '💠', min: 50000,  max: 55000  };
  else if (honor >= 45000)  tier = { name: '찬란한 루비',     icon: '💎', min: 30000,  max: 50000  };
  else if (honor >= 40000)  tier = { name: '진화한 루비',     icon: '💎', min: 30000,  max: 45000  };
  else if (honor >= 35000)  tier = { name: '성장한 루비',     icon: '💎', min: 30000,  max: 40000  };
  else if (honor >= 30000)  tier = { name: '루비 원석',     icon: '💎', min: 30000,  max: 35000  };
  else if (honor >= 27500)  tier = { name: '찬란한 골드',         icon: '🥇', min: 20000,  max: 30000  };
  else if (honor >= 25000)  tier = { name: '진화한 골드',         icon: '🥇', min: 20000,  max: 27500  };
  else if (honor >= 22500)  tier = { name: '성장한 골드',         icon: '🥇', min: 20000,  max: 25000  };
  else if (honor >= 20000)  tier = { name: '금 광석',         icon: '🥇', min: 20000,  max: 22500  };
  else if (honor >= 17500)  tier = { name: '찬란한 실버',         icon: '🥈', min: 10000,  max: 20000  };
  else if (honor >= 15000)  tier = { name: '진화한 실버',         icon: '🥈', min: 10000,  max: 17500  };
  else if (honor >= 12500)  tier = { name: '성장한 실버',         icon: '🥈', min: 10000,  max: 15000  };
  else if (honor >= 10000)  tier = { name: '거친 실버',         icon: '🥈', min: 10000,  max: 12500  };
  else if (honor >= 7500)   tier = { name: '빛나는 브론즈',       icon: '🥉', min: 7500,   max: 10000  };
  else if (honor >= 5000)   tier = { name: '브론즈',       icon: '🥉', min: 5000,   max: 7500  };

  // 업적 자동 체크 (로그인 시마다 조건 확인)
  checkAndGrantAchievements(studentName, Number(studentRow[COL_ASSET - 1]) || 0, Number(studentRow[COL_TAX - 1]) || 0, honor);

  const result = {
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
    snacks:        getSnackData(),
    achievements:  getStudentAchievements(studentName),
    job2:          getSecondaryJobForStudent(studentName),
    jobMarket:     getJobData()
  };
  
  // ── 캐시 저장 ──────────────────────────────────────────────
  result._password = correctPassword;
  cache.put(cacheKey, JSON.stringify(result), 600); // 10분
  
  delete result._password;
  return result;
  // ───────────────────────────────────────────────────────────
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
  return { students, snacks:        getSnackData(),
    achievements:  getStudentAchievements(studentName)
  };
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
    <button class="btn-fill" style="background:#8e44ad;" onclick="loadJobSalariesAndFill()">🤖 1인1역 자동계산</button>
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
    function autoFillFromJob() {
      if (!window.jobSalaries) { alert('직업 데이터가 아직 로딩되지 않았습니다. 잠시 후 다시 시도해주세요.'); return; }
      document.querySelectorAll('.ptInput').forEach(function(inp) {
        var row = inp.dataset.row;
        var salary = window.jobSalaries[row] || 0;
        inp.value = salary + 300;
      });
    }
    function loadJobSalariesAndFill() {
      var btn = document.querySelector('[onclick="loadJobSalariesAndFill()"]');
      btn.innerText = '로딩 중...';
      btn.disabled = true;
      google.script.run.withSuccessHandler(function(salaries) {
        window.jobSalaries = salaries; // { "1": 120, "2": 100, ... } (row → salary)
        autoFillFromJob();
        btn.innerText = '🤖 1인1역 자동계산';
        btn.disabled = false;
      }).getJobSalariesByRow();
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

// ════════════════════════════════════════════════════════════════
// 12. 1인1역 일급 데이터 반환 (행번호 → 일급 매핑)
// ════════════════════════════════════════════════════════════════
function getJobSalariesByRow() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
  const jobData  = ss.getSheetByName(SHEET_JOB).getDataRange().getValues();

  // 이름 → 일급 맵 만들기
  const salaryMap = {};
  for (let j = 1; j < jobData.length; j++) {
    const jName   = String(jobData[j][0]).trim();
    const jSalary = Number(jobData[j][2]) || 0;
    if (jName) salaryMap[jName] = jSalary;
  }

  // 행번호(rowIdx) → 일급 맵으로 변환
  const result = {};
  for (let i = 1; i < mainData.length; i++) {
    const name = String(mainData[i][COL_NAME - 1]).trim();
    if (name) result[String(i)] = salaryMap[name] || 0;
  }
  return result;
}


// ════════════════════════════════════════════════════════════════
// 13. 업적 시스템 서버 함수
// ════════════════════════════════════════════════════════════════

// 업적마스터 데이터 캐싱 (1시간 유효)
function getCachedAchievementMaster() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('achievement_master');
  
  if (cached) {
    return JSON.parse(cached);
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!masterSheet) return [];
  
  const data = masterSheet.getDataRange().getValues();
  const result = [];
  
  for (let m = 1; m < data.length; m++) {
    if (!data[m][0]) continue;
    result.push({
      achId:     String(data[m][0]).trim(),
      achName:   String(data[m][1]).trim(),
      condition: String(data[m][2]).trim(),
      isHidden:  String(data[m][3]).toUpperCase() === 'TRUE',
      hint:      String(data[m][4] || '').trim(),
      grade:     String(data[m][5] || '희귀').trim()
    });
  }
  
  cache.put('achievement_master', JSON.stringify(result), 3600); // 1시간
  return result;
}

// ════════════════════════════════════════════════════════════════
// 캐시 관리 함수
// ════════════════════════════════════════════════════════════════

// 업적마스터 캐시 초기화 (업적 수정 후 실행)
function clearAchievementCache() {
  CacheService.getScriptCache().remove('achievement_master');
  SpreadsheetApp.getUi().alert('✅ 업적마스터 캐시가 초기화되었습니다.');
}

// 전체 캐시 초기화 (디버깅용)
function clearAllCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['achievement_master']);
  
  // 학생별 캐시는 패턴으로 삭제 불가하므로 개별 삭제
  // (필요시 학생 목록 순회하며 삭제)
  
  SpreadsheetApp.getUi().alert('✅ 모든 캐시가 초기화되었습니다.');
}

// 특정 학생의 달성 업적 목록 반환
function getStudentAchievements(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACH_STUDENT);
  if (!sheet) return [];
  const data   = sheet.getDataRange().getValues();

  // 업적마스터에서 achId → grade 맵 생성
  const gradeMap = {};
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (!mData[m][0]) continue;
      gradeMap[String(mData[m][0]).trim()] = String(mData[m][5] || '희귀').trim();
    }
  }

  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentName).trim()) {
      let dateVal = data[i][4];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      const achId = String(data[i][1]);
      result.push({
        achId:     achId,
        achName:   String(data[i][2]),
        condition: String(data[i][3]),
        date:      String(dateVal),
        equipped:  data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE',
        sheetRow:  i + 1,
        grade:     gradeMap[achId] || '희귀'
      });
    }
  }
  return result;
}

// 칭호 장착 처리 (기존 장착 해제 → 새 칭호 장착)
function equipAchievement(studentName, targetSheetRow) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACH_STUDENT);
  if (!sheet) return { success: false, msg: '업적 시트를 찾을 수 없습니다.' };
  const data = sheet.getDataRange().getValues();

  // 해당 학생의 모든 행 탐색 → 기존 장착 FALSE로 초기화
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentName).trim()) {
      sheet.getRange(i + 1, 6).setValue(false); // F열: 장착여부
    }
  }
  // 새 칭호 TRUE로 설정
  sheet.getRange(targetSheetRow, 6).setValue(true);
  return { success: true };
}

// 칭호 해제
function unequipAchievement(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACH_STUDENT);
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentName).trim()) {
      sheet.getRange(i + 1, 6).setValue(false);
    }
  }
  return { success: true };
}

// 업적 달성 체크 및 자동 부여 (getStudentData 안에서 호출하거나 독립 호출 가능)
// 현재 자동 체크 조건: ① 자산 5000이상, ② 납세 500이상
function checkAndGrantAchievements(studentName, balance, totalTax, honor) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const achSheet   = ss.getSheetByName(SHEET_ACH_STUDENT);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!achSheet || !masterSheet) return;

  // 이미 달성한 업적 ID 목록
  const existing = new Set();
  const achData  = achSheet.getDataRange().getValues();
  for (let i = 1; i < achData.length; i++) {
    if (String(achData[i][0]).trim() === String(studentName).trim()) {
      existing.add(String(achData[i][1]).trim());
    }
  }

  const today      = _todayStr();
  const masterData = masterSheet.getDataRange().getValues();

  // ── 자동 조건 체크 맵 ──────────────────────────────────────
  // ── 자동 조건 체크 맵 ──────────────────────────────────────
  // ⚠️ 여기 없는 업적(ECO-001, ECO-002, RANK 시리즈, HID-004)은
  //    별도 함수에서 처리하므로 여기에 추가하지 않아도 됩니다.
  const conditionMap = {
    'ACH-001': balance >= 5000,
    'ACH-002': totalTax >= 500,
  };

  // ── ECO-001: 황금 절약가 (지난 30일 자산사용 1000 미만) ────
  const spendSheet = ss.getSheetByName(SHEET_SPEND);
  if (spendSheet && !existing.has('ECO-001')) {
    const spendData  = spendSheet.getDataRange().getValues();
    const cutoff     = new Date(); cutoff.setDate(cutoff.getDate() - 30);
    let recentSpend  = 0;
    for (let s = 1; s < spendData.length; s++) {
      if (String(spendData[s][1]).trim() !== studentName) continue;
      let rowDate = spendData[s][0];
      if (rowDate instanceof Date && rowDate >= cutoff) {
        recentSpend += Number(spendData[s][4]) || 0;
      }
    }
    if (recentSpend < 1000 && recentSpend >= 0) {
      // 마스터에서 ECO-001 정보 찾기
      for (let m = 1; m < masterData.length; m++) {
        if (String(masterData[m][0]).trim() === 'ECO-001') {
          achSheet.appendRow([studentName, 'ECO-001', String(masterData[m][1]), String(masterData[m][2]), today, false]);
          break;
        }
      }
    }
  }

  // ── ECO-002: 학급의 큰 손 (경매 낙찰가 학급 역대 최고가 경신) ──
  const auctionSheet2 = ss.getSheetByName(SHEET_AUCTION);
  if (auctionSheet2 && !existing.has('ECO-002')) {
    const aData2  = auctionSheet2.getDataRange().getValues();
    let classMax  = 0, myMax = 0;
    // C열(인덱스2)~K열(인덱스10): 1차~9차 낙찰가 열
    for (let a = 1; a < aData2.length; a++) {
      for (let c = 2; c <= 10; c++) {
        const v = Number(aData2[a][c]) || 0;
        if (v > classMax) classMax = v;
      }
    }
    // 학생 자신의 최고 낙찰가는 자산사용 시트에서 확인
    if (spendSheet) {
      const sd2 = spendSheet.getDataRange().getValues();
      for (let s = 1; s < sd2.length; s++) {
        if (String(sd2[s][1]).trim() !== studentName) continue;
        if (!String(sd2[s][3]).includes('[경매낙찰]')) continue;
        const v = Number(sd2[s][4]) || 0;
        if (v > myMax) myMax = v;
      }
    }
    if (myMax > 0 && myMax >= classMax) {
      for (let m = 1; m < masterData.length; m++) {
        if (String(masterData[m][0]).trim() === 'ECO-002') {
          achSheet.appendRow([studentName, 'ECO-002', String(masterData[m][1]), String(masterData[m][2]), today, false]);
          break;
        }
      }
    }
  }

  // ── HID-004: 업적 수집가 (달성 업적 10개 이상) ──────────────
  if (!existing.has('HID-004') && existing.size >= 10) {
    for (let m = 1; m < masterData.length; m++) {
      if (String(masterData[m][0]).trim() === 'HID-004') {
        achSheet.appendRow([studentName, 'HID-004', String(masterData[m][1]), String(masterData[m][2]), today, false]);
        break;
      }
    }
  }

  // ── RANK-001~006: 랭크 브레이커 ────────────────────────────
  const rankBreakers = {
    'RANK-001': ['거친 실버'],
    'RANK-002': ['금 광석'],
    'RANK-003': ['루비 원석'],
    'RANK-004': ['거친 다이아몬드'],
    'RANK-005': ['마스터'],
    'RANK-006': ['그랜드마스터']
  };
  // 현재 학생 티어명 계산 (honor 기반)
  const h = Number(honor) || 0;
  let currentTierName = '새싹';
  if      (h >= 100000) currentTierName = '그랜드마스터';
  else if (h >= 85000)  currentTierName = '완성된 마스터';
  else if (h >= 75000)  currentTierName = '마스터';
  else if (h >= 65000)  currentTierName = '찬란한 다이아몬드';
  else if (h >= 60000)  currentTierName = '진화한 다이아몬드';
  else if (h >= 55000)  currentTierName = '성장한 다이아몬드';
  else if (h >= 50000)  currentTierName = '거친 다이아몬드';
  else if (h >= 45000)  currentTierName = '찬란한 루비';
  else if (h >= 40000)  currentTierName = '진화한 루비';
  else if (h >= 35000)  currentTierName = '성장한 루비';
  else if (h >= 30000)  currentTierName = '루비 원석';
  else if (h >= 27500)  currentTierName = '찬란한 골드';
  else if (h >= 25000)  currentTierName = '진화한 골드';
  else if (h >= 22500)  currentTierName = '성장한 골드';
  else if (h >= 20000)  currentTierName = '금 광석';
  else if (h >= 17500)  currentTierName = '찬란한 실버';
  else if (h >= 15000)  currentTierName = '진화한 실버';
  else if (h >= 12500)  currentTierName = '성장한 실버';
  else if (h >= 10000)  currentTierName = '거친 실버';
  else if (h >= 7500)   currentTierName = '빛나는 브론즈';
  else if (h >= 5000)   currentTierName = '브론즈';

  Object.keys(rankBreakers).forEach(function(rankId) {
    if (existing.has(rankId)) return;
    if (rankBreakers[rankId].indexOf(currentTierName) === -1) return;
    // 학급 내 다른 학생이 이 rankId를 이미 달성했는지 확인 (최초 달성만)
    const allAchData = achSheet.getDataRange().getValues();
    let alreadyExists = false;
    for (let i = 1; i < allAchData.length; i++) {
      if (String(allAchData[i][1]).trim() === rankId) { alreadyExists = true; break; }
    }
    if (alreadyExists) return; // 이미 누군가 달성함 → 부여 안 함
    for (let m = 1; m < masterData.length; m++) {
      if (String(masterData[m][0]).trim() === rankId) {
        achSheet.appendRow([studentName, rankId, String(masterData[m][1]), String(masterData[m][2]), today, false]);
        break;
      }
    }
  });


  for (let m = 1; m < masterData.length; m++) {
    const achId   = String(masterData[m][0]).trim();
    const achName = String(masterData[m][1]).trim();
    const cond    = String(masterData[m][2]).trim();
    if (!achId) continue;
    if (existing.has(achId)) continue;
    if (conditionMap[achId] === true) {
      achSheet.appendRow([studentName, achId, achName, cond, today, false]);
    }
  }
}

// ════════════════════════════════════════════════════════════════
// 14. 업적 신청-승인 시스템 (v2)
// ════════════════════════════════════════════════════════════════

const SHEET_ACH_LOG    = '업적신청로그';
const SHEET_GLOBAL_NOTIFY = '전역알림';
const SHEET_JOB2_APP   = '2차직업신청';
const SHEET_JOB2_CURR  = '2차직업현황';

// ── 업적 도감 전체 데이터 반환 (학생 대시보드용) ─────────────────
// 반환값: { myAchievements, allAchievements, pendingIds, equippedTitle, globalNotices }
function getAchievementData(studentName) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  const achSheet    = ss.getSheetByName(SHEET_ACH_STUDENT);
  const logSheet    = ss.getSheetByName(SHEET_ACH_LOG);
  const notifySheet = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);

  // 1. 내가 달성한 업적 목록
  const myAchievements = [];
  let equippedTitle = null;
  if (achSheet) {
    const achData = achSheet.getDataRange().getValues();
    for (let i = 1; i < achData.length; i++) {
      if (String(achData[i][0]).trim() !== String(studentName).trim()) continue;
      let dateVal = achData[i][4];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      const equipped = achData[i][5] === true || String(achData[i][5]).toUpperCase() === 'TRUE';
      const ach = {
        achId:     String(achData[i][1]),
        achName:   String(achData[i][2]),
        condition: String(achData[i][3]),
        date:      String(dateVal),
        equipped:  equipped,
        sheetRow:  i + 1
      };
      myAchievements.push(ach);
      if (equipped) equippedTitle = ach.achName;
    }
  }
  const myAchIds = new Set(myAchievements.map(a => a.achId));

  // 2. 전체 업적 도감 (히든 처리 포함)
  const allAchievements = [];
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (!mData[m][0]) continue;
      const achId   = String(mData[m][0]).trim();
      const achName = String(mData[m][1]).trim();
      const cond    = String(mData[m][2]).trim();
      const isHidden = String(mData[m][3]).toUpperCase() === 'TRUE';
      const hint    = String(mData[m][4] || '');
      const earned  = myAchIds.has(achId);
      // 자동 부여 업적은 신청 드롭다운에서 제외
      const AUTO_GRANTED_IDS = new Set(['ACH-001','ACH-002','ECO-001','ECO-002','HID-004',
        'RANK-001','RANK-002','RANK-003','RANK-004','RANK-005','RANK-006']);
      allAchievements.push({
        achId,
        achName:     isHidden && !earned ? '???' : achName,
        condition:   isHidden && !earned ? '히든 업적입니다.' : cond,
        hint:        isHidden && !earned ? hint : '',
        isHidden,
        earned,
        autoGranted: AUTO_GRANTED_IDS.has(achId)
      });
    }
  }

  // 3. 현재 대기 중인 신청 업적ID 목록 (중복 신청 방지용)
  const pendingIds = new Set();
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues();
    for (let l = 1; l < logData.length; l++) {
      if (String(logData[l][1]).trim() === String(studentName).trim() &&
          String(logData[l][4]).trim() === '대기') {
        pendingIds.add(String(logData[l][2]).trim());
      }
    }
  }

  // 4. 전역 알림 (읽지 않은 공지 — 프론트에서 localStorage로 1회 처리)
  const globalNotices = [];
  if (notifySheet) {
    const nData = notifySheet.getDataRange().getValues();
    for (let n = 1; n < nData.length; n++) {
      if (nData[n][0]) {
        globalNotices.push({
          noticeId: String(nData[n][0]),
          message:  String(nData[n][1]),
          time:     String(nData[n][2])
        });
      }
    }
  }

  return {
    myAchievements,
    allAchievements,
    pendingIds:    [...pendingIds],
    equippedTitle,
    globalNotices
  };
}


// ── 업적 신청 / 특별 보고 제출 ────────────────────────────────────
// achievementId: 일반 신청 시 업적ID, 특별 보고 시 '특별보고'
function submitAchievement(studentName, achievementId, proofText) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  if (!logSheet) return { success: false, msg: '업적신청로그 시트를 찾을 수 없습니다.' };

  // 중복 대기 방지 (같은 업적ID가 이미 대기 중인지 확인)
  if (achievementId !== '특별보고') {
    const logData = logSheet.getDataRange().getValues();
    for (let i = 1; i < logData.length; i++) {
      if (String(logData[i][1]).trim() === String(studentName).trim() &&
          String(logData[i][2]).trim() === String(achievementId).trim() &&
          String(logData[i][4]).trim() === '대기') {
        return { success: false, msg: '이미 해당 업적이 승인 대기 중입니다.' };
      }
    }
    // 이미 달성한 업적인지 확인
    const achSheet = ss.getSheetByName(SHEET_ACH_STUDENT);
    if (achSheet) {
      const achData = achSheet.getDataRange().getValues();
      for (let i = 1; i < achData.length; i++) {
        if (String(achData[i][0]).trim() === String(studentName).trim() &&
            String(achData[i][1]).trim() === String(achievementId).trim()) {
          return { success: false, msg: '이미 달성한 업적입니다.' };
        }
      }
    }
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([timestamp, studentName, achievementId, proofText, '대기']);
  return { success: true, msg: '신청이 완료되었습니다. 선생님의 승인을 기다려주세요.' };
}


// ── 관리자: 업적 신청 승인/반려 ───────────────────────────────────
// rowNumber: 업적신청로그 시트의 실제 행 번호
// isApproved: true=승인, false=반려
// finalAchievementId: 특별보고를 승인할 때 선생님이 선택한 업적ID (일반 승인 시 null)
function approveAchievement(rowNumber, isApproved, finalAchievementId) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  const achSheet = ss.getSheetByName(SHEET_ACH_STUDENT);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!logSheet || !achSheet) return { success: false, msg: '시트를 찾을 수 없습니다.' };

  const row = logSheet.getRange(rowNumber, 1, 1, 5).getValues()[0];
  const studentName = String(row[1]).trim();
  const requestedId = String(row[2]).trim();

  if (!isApproved) {
    // 반려 처리
    logSheet.getRange(rowNumber, 5).setValue('반려');
    return { success: true, msg: '반려 처리되었습니다.' };
  }

  // 승인 처리
  logSheet.getRange(rowNumber, 5).setValue('승인');

  // 특별보고인 경우 선생님이 선택한 업적ID 사용, 일반 신청이면 원래 ID 사용
  const achId = (requestedId === '특별보고' && finalAchievementId)
    ? String(finalAchievementId).trim()
    : requestedId;

  // 마스터에서 업적명, 달성조건 찾기
  let achName = achId, achCond = '';
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (String(mData[m][0]).trim() === achId) {
        achName = String(mData[m][1]).trim();
        achCond = String(mData[m][2]).trim();
        break;
      }
    }
  }

  // 이미 달성한 업적인지 중복 체크
  const achData = achSheet.getDataRange().getValues();
  for (let i = 1; i < achData.length; i++) {
    if (String(achData[i][0]).trim() === studentName &&
        String(achData[i][1]).trim() === achId) {
      return { success: false, msg: '이미 달성 처리된 업적입니다.' };
    }
  }

  const today = _todayStr();

  // ★ 히든 업적 최초 달성 체크 → 전원 공지 + 히든 해제
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (String(mData[m][0]).trim() !== achId) continue;
      const isHidden = String(mData[m][3]).toUpperCase() === 'TRUE';
      if (!isHidden) break;

      // 이미 다른 학생이 달성했는지 확인
      let alreadyUnlocked = false;
      for (let i = 1; i < achData.length; i++) {
        if (String(achData[i][1]).trim() === achId) { alreadyUnlocked = true; break; }
      }

      if (!alreadyUnlocked) {
        // 최초 달성 → 히든여부 FALSE로 변경
        masterSheet.getRange(m + 1, 4).setValue('FALSE');

        // 전역 알림 시트에 공지 추가
        const notifySheet = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
        if (notifySheet) {
          const noticeId = 'HIDDEN_' + achId + '_' + new Date().getTime();
          const msg = `🎉 히든 업적 [${achName}]을(를) 달성한 사람이 최초로 등장했습니다! 지금부터 이 업적의 정체와 달성 조건이 모두에게 공개됩니다.`;
          const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          notifySheet.appendRow([noticeId, msg, ts]);
        }
      }
      break;
    }
  }

  // 학생업적달성 시트에 기록
  achSheet.appendRow([studentName, achId, achName, achCond, today, false]);

  // ★ 마일스톤 자산 보상 체크
  const finalAchData = achSheet.getDataRange().getValues();
  let totalCount = 0;
  for (let i = 1; i < finalAchData.length; i++) {
    if (String(finalAchData[i][0]).trim() === studentName) totalCount++;
  }
  grantMilestoneReward(studentName, totalCount);

  // ★ 전광판 알림 체크
  const achGradeForAlert = masterSheet ? (() => {
    const mData2 = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData2.length; m++) {
      if (String(mData2[m][0]).trim() === achId) return String(mData2[m][5] || '희귀').trim();
    }
    return '희귀';
  })() : '희귀';
  _checkAndPostGlobalAlert(studentName, achName, achGradeForAlert);

  return { success: true, msg: `[${studentName}] ${achName} 업적 승인 완료!` };
}


// ── 관리자: 업적 신청 대기 목록 반환 ─────────────────────────────
function getPendingAchievements() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!logSheet) return { pending: [], allMasterAchs: [] };

  // 업적마스터에서 업적ID → 업적명 맵 생성
  const achNameMap = {};
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (!mData[m][0]) continue;
      achNameMap[String(mData[m][0]).trim()] = String(mData[m][1]).trim();
    }
  }

  const logData = logSheet.getDataRange().getValues();
  const pending = [];
  for (let i = 1; i < logData.length; i++) {
    if (String(logData[i][4]).trim() !== '대기') continue;
    const achId = String(logData[i][2]).trim();
    pending.push({
      rowNumber:   i + 1,
      timestamp:   String(logData[i][0]),
      studentName: String(logData[i][1]),
      achId:       achId,
      achName:     achNameMap[achId] || '(알 수 없음)', // 업적명 추가
      proof:       String(logData[i][3])
    });
  }

  // 특별보고 승인 시 업적 선택용 전체 목록
  const allMasterAchs = [];
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (!mData[m][0]) continue;
      allMasterAchs.push({ achId: String(mData[m][0]), achName: String(mData[m][1]) });
    }
  }

  return { pending, allMasterAchs };
}

// ── 관리자: 업적 일괄 승인/반려 ─────────────────────────────────
function batchApproveAchievements(rowNumbers, isApproved) {
  if (!rowNumbers || rowNumbers.length === 0) {
    return { success: false, msg: '처리할 항목이 없습니다.' };
  }

  const results = [];
  let successCount = 0;
  let failCount = 0;

  for (let i = 0; i < rowNumbers.length; i++) {
    const res = approveAchievement(rowNumbers[i], isApproved, null);
    if (res.success) {
      successCount++;
    } else {
      failCount++;
    }
    results.push(res);
  }

  const action = isApproved ? '승인' : '반려';
  return {
    success: true,
    msg: `일괄 ${action} 완료: 성공 ${successCount}건, 실패 ${failCount}건`,
    details: results
  };
}


// ════════════════════════════════════════════════════════════════
// 15. 2차 직업 시스템
// ════════════════════════════════════════════════════════════════

// ── 전체 2차 직업 현황 반환 ───────────────────────────────────────
function getJobData() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_JOB2_CURR);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      studentName: String(data[i][0]),
      jobName:     String(data[i][1]),
      jobDesc:     String(data[i][2]),
      approvedDate: String(data[i][3])
    });
  }
  return result;
}

// ── 학생: 2차 직업 신청 ──────────────────────────────────────────
function submitJobApplication(studentName, jobName, jobDesc) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet = ss.getSheetByName(SHEET_JOB2_APP);
  if (!appSheet) return { success: false, msg: '2차직업신청 시트를 찾을 수 없습니다.' };

  // 이미 대기 중인 신청이 있는지 확인
  const appData = appSheet.getDataRange().getValues();
  for (let i = 1; i < appData.length; i++) {
    if (String(appData[i][1]).trim() === String(studentName).trim() &&
        String(appData[i][4]).trim() === '대기') {
      return { success: false, msg: '이미 승인 대기 중인 신청이 있습니다.' };
    }
  }

  // 이미 승인된 2차 직업이 있는지 확인
  const currSheet = ss.getSheetByName(SHEET_JOB2_CURR);
  if (currSheet) {
    const currData = currSheet.getDataRange().getValues();
    for (let i = 1; i < currData.length; i++) {
      if (String(currData[i][0]).trim() === String(studentName).trim()) {
        return { success: false, msg: '이미 2차 직업이 있습니다. 변경이 필요하면 선생님께 문의하세요.' };
      }
    }
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  appSheet.appendRow([ts, studentName, jobName, jobDesc, '대기']);
  return { success: true, msg: '신청이 완료되었습니다! 선생님의 승인을 기다려주세요.' };
}

// ── 관리자: 2차 직업 승인/반려 ───────────────────────────────────
function approveJob(rowNumber, isApproved) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet  = ss.getSheetByName(SHEET_JOB2_APP);
  const currSheet = ss.getSheetByName(SHEET_JOB2_CURR);
  if (!appSheet) return { success: false, msg: '시트를 찾을 수 없습니다.' };

  const row         = appSheet.getRange(rowNumber, 1, 1, 5).getValues()[0];
  const studentName = String(row[1]).trim();
  const jobName     = String(row[2]).trim();
  const jobDesc     = String(row[3]).trim();

  appSheet.getRange(rowNumber, 5).setValue(isApproved ? '승인' : '반려');

  if (isApproved && currSheet) {
    const today = _todayStr();
    currSheet.appendRow([studentName, jobName, jobDesc, today]);
  }

  // ── 우편함 발송 ──────────────────────────────────────────────
  if (isApproved) {
    _sendMail(
      studentName,
      `✅ 2차 직업 승인: ${jobName}`,
      `🎉 축하합니다! 2차 직업 [${jobName}] 신청이 승인되었습니다.\n\n직업 설명: ${jobDesc}\n\n대시보드의 '2차 직업 시스템'에서 확인해보세요!`,
      '승인'
    );
  } else {
    _sendMail(
      studentName,
      `❌ 2차 직업 반려: ${jobName}`,
      `[${jobName}] 2차 직업 신청이 반려되었습니다.\n\n직업 내용을 다듬어서 다시 신청해주세요.`,
      '반려'
    );
  }

  // ── 캐시 무효화 (학생 재로그인 시 최신 데이터 반영) ──────────
  CacheService.getScriptCache().remove('student_' + studentName);

  return { success: true, msg: isApproved ? `[${studentName}] 2차 직업 승인 완료!` : '반려 처리되었습니다.' };
}

// ── 관리자: 2차 직업 대기 목록 반환 ─────────────────────────────
function getPendingJobs() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const appSheet = ss.getSheetByName(SHEET_JOB2_APP);
  if (!appSheet) return [];
  const data = appSheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() !== '대기') continue;
    result.push({
      rowNumber:   i + 1,
      timestamp:   String(data[i][0]),
      studentName: String(data[i][1]),
      jobName:     String(data[i][2]),
      jobDesc:     String(data[i][3])
    });
  }
  return result;
}

// ── getStudentData에 2차 직업 정보 추가용 헬퍼 ──────────────────
// getStudentData()의 return 블록에 아래 필드를 추가해야 합니다:
//   job2: getSecondaryJobForStudent(studentName)
function getSecondaryJobForStudent(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_JOB2_CURR);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentName).trim()) {
      return { jobName: String(data[i][1]), jobDesc: String(data[i][2]) };
    }
  }
  return null;
}

// ════════════════════════════════════════════════════════════════
// 15. 로그인 화면용 - 전체 학생 업적 명예의 전당
// ════════════════════════════════════════════════════════════════

// 전체 학생의 칭호 및 업적 정보 반환 (로그인 화면용)
function getAllStudentsHonorBoard() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet   = ss.getSheetByName(SHEET_MAIN);
  const achSheet    = ss.getSheetByName(SHEET_ACH_STUDENT);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  
  if (!mainSheet || !achSheet || !masterSheet) return [];

  const mainData   = mainSheet.getDataRange().getValues();
  const achData    = achSheet.getDataRange().getValues();
  const masterData = masterSheet.getDataRange().getValues();

  // 업적마스터에서 업적ID별 등급 맵 생성 (F열 = 인덱스 5)
  const gradeMap = {};
  const emojiMap = {}; // 유니크 이상 업적에 이모지 추가
  for (let m = 1; m < masterData.length; m++) {
    const achId = String(masterData[m][0]).trim();
    const grade = String(masterData[m][5] || '희귀').trim(); // F열: 업적등급
    gradeMap[achId] = grade;
    
    // 유니크 이상 업적에 자동 이모지 할당
    if (grade === '유니크' || grade === '히든' || grade === '유일') {
      emojiMap[achId] = getEmojiForAchievement(achId);
    }
  }

  const result = [];

  // 학생별로 순회
  for (let i = 1; i < mainData.length; i++) {
    const studentName = String(mainData[i][COL_NAME - 1]).trim();
    if (!studentName) continue;

    // 해당 학생의 달성 업적 수집
    const achievements = [];
    let equippedTitle  = null;

    for (let j = 1; j < achData.length; j++) {
      if (String(achData[j][0]).trim() !== studentName) continue;
      
      const achId    = String(achData[j][1]).trim();
      const achName  = String(achData[j][2]).trim();
      const equipped = achData[j][5] === true || String(achData[j][5]).toUpperCase() === 'TRUE';
      const grade    = gradeMap[achId] || '희귀';
      const emoji    = emojiMap[achId] || '';

      achievements.push({
        achId:   achId,
        achName: achName,
        grade:   grade,
        emoji:   emoji
      });

      if (equipped) {
        equippedTitle = (emoji ? emoji + ' ' : '') + achName;
      }
    }

    result.push({
      name:            studentName,
      equippedTitle:   equippedTitle,
      achievementCount: achievements.length,
      achievements:    achievements
    });
  }

  // 업적 많은 순으로 정렬
  result.sort(function(a, b) {
    return b.achievementCount - a.achievementCount;
  });

  return result;
}

// 업적ID에 따라 적절한 이모지 반환 (유니크/히든/유일용)
function getEmojiForAchievement(achId) {
  const emojiMapping = {
    // 경제 관련
    'ECO-002': '💰', 'ECO-003': '💎', 'ECO-004': '🏆',
    // 생활 관련
    'LIFE-002': '🌟', 'LIFE-003': '⏰', 'LIFE-004': '📚',
    'LIFE-005': '🎯', 'LIFE-006': '🌈', 'LIFE-007': '💪',
    'LIFE-008': '🔥', 'LIFE-009': '✨', 'LIFE-010': '🎨',
    // MVP 관련
    'MVP-001': '👑', 'MVP-002': '🥇',
    // 학생 관련
    'STU-001': '🎓', 'STU-002': '📖', 'STU-003': '🌺',
    // 팀워크 관련
    'TEAM-001': '🤝', 'TEAM-002': '🎭',
    // 소비 관련
    'CONS-001': '🍪', 'CONS-002': '🎁',
    // 도전 과제
    'CHAL-001': '⚡', 'CHAL-002': '🚀', 'CHAL-003': '🌊',
    'CHAL-004': '🔮', 'CHAL-005': '🎪',
    // 히든
    'HID-001': '🕵️', 'HID-002': '🎩', 'HID-003': '💫', 'HID-005': '🏅',
    // 시작 업적
    'START-001': '🌱', 'START-002': '🌿', 'START-003': '🌳'
  };
  return emojiMapping[achId] || '⭐';
}


// ════════════════════════════════════════════════════════════════
// ★ 신규 시트 상수 (우편함 / 상점 / 전광판)
// ════════════════════════════════════════════════════════════════
const SHEET_MAILBOX   = '우편함_로그';      // 우편함 메시지 저장
const SHEET_SHOP_ITEMS = '상점_아이템';     // 상점 아이템 DB
const SHEET_SHOP_LOG   = '상점_구매로그';   // 구매 내역

// ════════════════════════════════════════════════════════════════
// ██ 기능 1: 우편함(Mailbox) ██
// 시트 컬럼: A=메시지ID, B=수신자, C=제목, D=내용, E=타입(승인/반려),
//            F=읽음여부(TRUE/FALSE), G=발송일시
// ════════════════════════════════════════════════════════════════

// 우편함 메시지 전송 (내부 헬퍼, approveAchievement에서 호출)
function _sendMail(recipientName, subject, body, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MAILBOX);
    sheet.appendRow(['메시지ID','수신자','제목','내용','타입','읽음','발송일시']);
  }
  const msgId = 'MSG_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,5);
  const ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([msgId, recipientName, subject, body, type, false, ts]);
}

// 학생의 읽지 않은 메시지 수 반환
function getUnreadMailCount(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(studentName).trim() &&
        String(data[i][5]).toUpperCase() !== 'TRUE') {
      count++;
    }
  }
  return count;
}

// 학생의 전체 메시지 목록 반환 + 읽음 처리
function getMailboxMessages(studentName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const msgs = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== String(studentName).trim()) continue;
    const isRead = String(data[i][5]).toUpperCase() === 'TRUE';
    msgs.push({
      msgId:   String(data[i][0]),
      subject: String(data[i][2]),
      body:    String(data[i][3]),
      type:    String(data[i][4]),   // '승인' | '반려'
      isRead:  isRead,
      sentAt:  String(data[i][6]),
      rowNum:  i + 1
    });
    // 읽음 처리
    if (!isRead) sheet.getRange(i + 1, 6).setValue(true);
  }
  return msgs.reverse(); // 최신순
}

// ════════════════════════════════════════════════════════════════
// approveAchievement 확장: 승인/반려 시 우편함 메시지 자동 발송
// 기존 함수를 덮어쓰지 않고, 래퍼 함수로 우편함 호출을 삽입합니다.
// → AuctionAdmin.html에서 approveAchievement 대신
//   approveAchievementWithMail 을 호출하도록 변경하세요.
// ════════════════════════════════════════════════════════════════
function approveAchievementWithMail(rowNumber, isApproved, finalAchievementId, rejectReason) {
  const result = approveAchievement(rowNumber, isApproved, finalAchievementId);
  if (!result.success) return result;

  // 수신자 이름 다시 조회
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  if (!logSheet) return result;
  const row         = logSheet.getRange(rowNumber, 1, 1, 5).getValues()[0];
  const studentName = String(row[1]).trim();
  const achId       = String(row[2]).trim();

  // 업적명 조회
  let achName = achId;
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (String(mData[m][0]).trim() === (finalAchievementId || achId)) {
        achName = String(mData[m][1]).trim();
        break;
      }
    }
  }

  if (isApproved) {
    _sendMail(
      studentName,
      `✅ 업적 승인: ${achName}`,
      `🎉 축하합니다! [${achName}] 업적 신청이 승인되었습니다. 나의 업적 창에서 확인해보세요!`,
      '승인'
    );
  } else {
    const reason = rejectReason ? rejectReason : '조건 미충족';
    _sendMail(
      studentName,
      `❌ 업적 반려: ${achName}`,
      `[${achName}] 업적 신청이 반려되었습니다.\n\n반려 사유: ${reason}\n\n조건을 다시 확인하고 재신청해주세요.`,
      '반려'
    );
  }
  return result;
}

// ════════════════════════════════════════════════════════════════
// ██ 기능 2: 업적 전광판 (Global Alert)
// 전광판 조건: ① 업적 10개 단위 ② 유일/초월 등급 획득
// 시트: 전역알림 (기존 SHEET_GLOBAL_NOTIFY) 재활용
// ════════════════════════════════════════════════════════════════

// 학생 업적 개수 & 등급 기반 전광판 메시지 생성 (approveAchievement 후 호출)
function _checkAndPostGlobalAlert(studentName, achName, achGrade) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const achSheet = ss.getSheetByName(SHEET_ACH_STUDENT);
  const notify   = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
  if (!achSheet || !notify) return;

  // 해당 학생 총 업적 수
  const achData = achSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < achData.length; i++) {
    if (String(achData[i][0]).trim() === String(studentName).trim()) count++;
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  let msg = null;
  let noticeId = null;

  // ① 10개 단위 달성
  if (count > 0 && count % 10 === 0) {
    noticeId = `MILESTONE_${studentName}_${count}_${new Date().getTime()}`;
    msg = `🏆 [${studentName}] 학생이 업적 ${count}개 달성! 대단한 업적 수집가가 탄생했습니다!`;
  }

  // ② 유일/초월 등급 획득 (더 우선순위 높음)
  if (achGrade === '유일' || achGrade === '초월') {
    const gradeLabel = achGrade === '유일' ? '🌌 유일' : '✨ 초월';
    noticeId = `GRADE_${achGrade}_${studentName}_${new Date().getTime()}`;
    msg = `${gradeLabel} 등급 업적 [${achName}] 최초 달성! [${studentName}] 학생이 역사에 이름을 남겼습니다!`;
  }

  if (msg && noticeId) {
    notify.appendRow([noticeId, msg, ts, 'ALERT']); // D열 = 'ALERT' 타입 표시
  }
}

// 전광판 최신 메시지 조회 (프론트에서 폴링)
function getLatestGlobalAlert(lastSeenId) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  // ALERT 타입만, lastSeenId 이후 것만 반환
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][3]) === 'ALERT' && String(data[i][0]) !== String(lastSeenId)) {
      return { noticeId: String(data[i][0]), msg: String(data[i][1]), ts: String(data[i][2]) };
    }
  }
  return null;
}

// ════════════════════════════════════════════════════════════════
// ██ 기능 3: 실시간 업적 현황판 (Wall of Fame)
// ════════════════════════════════════════════════════════════════
function getWallOfFame() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  const achSheet    = ss.getSheetByName(SHEET_ACH_STUDENT);
  if (!masterSheet || !achSheet) return [];

  // 업적마스터 목록
  const mData = masterSheet.getDataRange().getValues();
  const achList = [];
  const achMap  = {};  // achId → index
  for (let m = 1; m < mData.length; m++) {
    if (!mData[m][0]) continue;
    const id    = String(mData[m][0]).trim();
    const name  = String(mData[m][1]).trim();
    const isHid = String(mData[m][3]).toUpperCase() === 'TRUE';
    const grade = String(mData[m][5] || '희귀').trim();
    achList.push({ achId: id, achName: isHid ? '🔒 ???' : name, grade, isHidden: isHid, count: 0 });
    achMap[id]  = achList.length - 1;
  }

  // 달성 학생 집계
  const sData = achSheet.getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) {
    const id = String(sData[i][1]).trim();
    if (achMap[id] !== undefined) achList[achMap[id]].count++;
  }

  // count 내림차순 정렬
  achList.sort(function(a, b) { return b.count - a.count; });
  return achList;
}

// ════════════════════════════════════════════════════════════════
// ██ 기능 4: 상점 시스템
// 상점_아이템 시트 컬럼:
//   A=아이템ID, B=카테고리(스킨/폰트/캐릭터), C=아이템명,
//   D=가격(자산), E=구매조건설명, F=조건타입, G=조건값,
//   H=리소스URL(CSS값 또는 이미지URL), I=활성여부
// 조건타입: 'none' | 'ach_count' | 'ach_unique' | 'ach_grade:{등급명}'
// 상점_구매로그 컬럼:
//   A=구매ID, B=학생명, C=아이템ID, D=아이템명, E=가격, F=구매일시
// ════════════════════════════════════════════════════════════════

// 상점 초기화 (최초 1회 실행)
function initShopSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 상점_아이템 시트 생성
  let itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  if (!itemSheet) {
    itemSheet = ss.insertSheet(SHEET_SHOP_ITEMS);
    itemSheet.appendRow(['아이템ID','카테고리','아이템명','가격','구매조건설명','조건타입','조건값','리소스값','활성여부']);
    // ── 스킨 12종 ──
    itemSheet.appendRow(['SKIN-001','스킨','추상화',           500,  '업적 5개 이상 달성',         'ach_count',      '5',  'abstract',         true]);
    itemSheet.appendRow(['SKIN-002','스킨','컬러풀 글래스',    700,  '업적 7개 이상 달성',         'ach_count',      '7',  'colorful_glass',   true]);
    itemSheet.appendRow(['SKIN-003','스킨','붉은 벽돌벽',      500,  '업적 10개 이상 달성',        'ach_count',      '10', 'red_brick',        true]);
    itemSheet.appendRow(['SKIN-004','스킨','숲 속의 비밀기지', 700,  '업적 10개 이상 달성',        'ach_count',      '10', 'hideout',          true]);
    itemSheet.appendRow(['SKIN-005','스킨','바다요정의 쉼터',  1000, '업적 10개 이상 달성',        'ach_count',      '10', 'ocean_fairy',      true]);
    itemSheet.appendRow(['SKIN-006','스킨','눈꽃',             1500, '업적 15개 이상 달성',        'ach_count',      '15', 'snowflower',       true]);
    itemSheet.appendRow(['SKIN-007','스킨','환상의 핑크레이크',1500, '유니크 업적 3개 이상 달성',  'ach_grade:유니크','3',  'pink_lake',        true]);
    itemSheet.appendRow(['SKIN-008','스킨','풀문',             1500, '업적 15개 이상 달성',        'ach_count',      '15', 'full_moon',        true]);
    itemSheet.appendRow(['SKIN-009','스킨','화이트 드래곤',    2000, '유니크 업적 5개 이상 달성',  'ach_grade:유니크','5',  'white_dragon',     true]);
    itemSheet.appendRow(['SKIN-010','스킨','전설의 소원나무',  2000, '업적 20개 이상 달성',        'ach_count',      '20', 'wish_tree',        true]);
    itemSheet.appendRow(['SKIN-011','스킨','밀키 웨이',        3000, '유니크 업적 7개 이상 달성',  'ach_grade:유니크','7',  'milky_way',        true]);
    itemSheet.appendRow(['SKIN-012','스킨','마법사왕의 궁전',  3000, '업적 30개 이상 달성',        'ach_count',      '30', 'palace_of_wizard', true]);
    // ── 샘플 폰트 3종 ──
    itemSheet.appendRow(['FONT-001','폰트','✏️ 귀여운 손글씨', 500, '조건 없음',              'none',          '0',  'Gaegu',         true]);
    itemSheet.appendRow(['FONT-002','폰트','📐 모던 고딕',     800, '업적 3개 이상 달성',      'ach_count',     '3',  'Black Han Sans', true]);
    itemSheet.appendRow(['FONT-003','폰트','👑 프리미엄 세리프', 2000,'초월 업적 1개 이상 달성','ach_grade:초월', '1',  'Nanum Myeongjo', true]);
    // ── 샘플 캐릭터 3종 ──
    itemSheet.appendRow(['CHAR-001','캐릭터','🐱 고양이 마스코트', 600, '조건 없음',           'none',          '0',  '🐱', true]);
    itemSheet.appendRow(['CHAR-002','캐릭터','🦊 여우 탐정',      1200,'업적 7개 이상 달성',   'ach_count',     '7',  '🦊', true]);
    itemSheet.appendRow(['CHAR-003','캐릭터','🐲 황금 드래곤',    3000,'유일 업적 1개 이상 달성','ach_grade:유일','1',  '🐲', true]);
  }

  // 상점_구매로그 시트 생성
  let logSheet = ss.getSheetByName(SHEET_SHOP_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_SHOP_LOG);
    logSheet.appendRow(['구매ID','학생명','아이템ID','아이템명','가격','구매일시']);
  }

  // 우편함_로그 시트 생성
  let mailSheet = ss.getSheetByName(SHEET_MAILBOX);
  if (!mailSheet) {
    mailSheet = ss.insertSheet(SHEET_MAILBOX);
    mailSheet.appendRow(['메시지ID','수신자','제목','내용','타입','읽음','발송일시']);
  }

  SpreadsheetApp.getUi().alert('✅ 상점 시트 초기화 완료! 상점_아이템 시트에서 아이템을 수정하세요.');
}

// 상점 아이템 목록 + 학생별 구매가능 여부 반환
function getShopItems(studentName) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  const logSheet  = ss.getSheetByName(SHEET_SHOP_LOG);
  const achSheet  = ss.getSheetByName(SHEET_ACH_STUDENT);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!itemSheet) return { items: [], owned: [] };

  // 학생 보유 자산 조회
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();
  let balance = 0;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME-1]).trim() === studentName) {
      balance = Number(mainData[i][COL_ASSET-1]) || 0;
      break;
    }
  }

  // 학생의 업적 목록 (등급 포함)
  const gradeMap = {};
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (mData[m][0]) gradeMap[String(mData[m][0]).trim()] = String(mData[m][5] || '희귀').trim();
    }
  }
  let totalAch = 0;
  const gradeCount = {};  // { '유니크': 2, '초월': 1, ... }
  if (achSheet) {
    const aData = achSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      if (String(aData[i][0]).trim() !== studentName) continue;
      totalAch++;
      const g = gradeMap[String(aData[i][1]).trim()] || '희귀';
      gradeCount[g] = (gradeCount[g] || 0) + 1;
    }
  }

  // 이미 구매한 아이템 목록
  const owned = [];
  if (logSheet) {
    const lData = logSheet.getDataRange().getValues();
    for (let i = 1; i < lData.length; i++) {
      if (String(lData[i][1]).trim() === studentName) owned.push(String(lData[i][2]).trim());
    }
  }

  // 아이템 목록 + 구매가능여부 판별
  const iData = itemSheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < iData.length; i++) {
    if (!iData[i][0] || String(iData[i][8]).toUpperCase() !== 'TRUE') continue;
    const itemId      = String(iData[i][0]).trim();
    const category    = String(iData[i][1]).trim();
    const itemName    = String(iData[i][2]).trim();
    const price       = Number(iData[i][3]) || 0;
    const condDesc    = String(iData[i][4]).trim();
    const condType    = String(iData[i][5]).trim();
    const condVal     = String(iData[i][6]).trim();
    const resourceVal = String(iData[i][7]).trim();
    const isOwned     = owned.includes(itemId);

    // 구매 조건 충족 여부
    let condMet = true;
    if (condType === 'ach_count') {
      condMet = totalAch >= Number(condVal);
    } else if (condType.startsWith('ach_grade:')) {
      const targetGrade = condType.split(':')[1];
      condMet = (gradeCount[targetGrade] || 0) >= Number(condVal);
    }
    // 'none' 이면 condMet = true

    const canBuy = !isOwned && condMet && balance >= price;

    items.push({ itemId, category, itemName, price, condDesc, condType, condVal, resourceVal, isOwned, condMet, canBuy });
  }

  return { items, owned, balance };
}

// 상점 아이템 목록 + 장착 여부 포함 반환 (Index.html openShopModal에서 호출)
function getShopItemsWithEquip(studentName) {
  const base = getShopItems(studentName);
  if (!base || !base.items) return base;

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_SHOP_LOG);
  const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  if (!logSheet || !itemSheet) return base;

  const iData = itemSheet.getDataRange().getValues();
  const lData = logSheet.getDataRange().getValues();

  // 장착 중인 아이템 ID 집합
  const equippedSet = new Set();
  for (let i = 1; i < lData.length; i++) {
    if (String(lData[i][1]).trim() !== studentName) continue;
    const isEq = lData[i][6] === true || String(lData[i][6]).toUpperCase() === 'TRUE';
    if (isEq) equippedSet.add(String(lData[i][2]).trim());
  }

  // 각 아이템에 isEquipped 필드 추가
  base.items = base.items.map(function(item) {
    item.isEquipped = equippedSet.has(item.itemId);
    return item;
  });
  base.equippedSet = Array.from(equippedSet);
  return base;
}

// 상점 구매 처리
function purchaseShopItem(studentName, itemId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  const logSheet  = ss.getSheetByName(SHEET_SHOP_LOG);
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!itemSheet || !logSheet || !mainSheet) return { success: false, msg: '시트 오류' };

  // 아이템 정보 조회
  const iData = itemSheet.getDataRange().getValues();
  let itemRow = null;
  let itemRowNum = -1;
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][0]).trim() === itemId) { itemRow = iData[i]; itemRowNum = i + 1; break; }
  }
  if (!itemRow) return { success: false, msg: '아이템을 찾을 수 없습니다.' };

  const price    = Number(itemRow[3]) || 0;
  const itemName = String(itemRow[2]).trim();

  // 이미 구매했는지 체크
  const lData = logSheet.getDataRange().getValues();
  for (let i = 1; i < lData.length; i++) {
    if (String(lData[i][1]).trim() === studentName && String(lData[i][2]).trim() === itemId) {
      return { success: false, msg: '이미 구매한 아이템입니다.' };
    }
  }

  // 잔액 차감
  const mData = mainSheet.getDataRange().getValues();
  for (let i = 1; i < mData.length; i++) {
    if (String(mData[i][COL_NAME-1]).trim() === studentName) {
      const current = Number(mData[i][COL_ASSET-1]) || 0;
      if (current < price) return { success: false, msg: `잔액 부족 (현재: $${current}, 필요: $${price})` };
      mainSheet.getRange(i + 1, COL_ASSET).setValue(current - price);

      // 자산사용 시트에 기록 (A=날짜, B=학생명, C=브랜드, D=구분, E=금액, F=사용후잔액, G=비고)
      const spendSheet = ss.getSheetByName(SHEET_SPEND);
      if (spendSheet) {
        const today = _todayStr();
        const newAsset = current - price;
        spendSheet.appendRow([today, studentName, mData[i][COL_BRAND-1], '상점구매', price, newAsset, `[${itemName}] 구매`]);
      }
      break;
    }
  }

  // 구매 로그 기록 (A=구매ID, B=학생명, C=아이템ID, D=아이템명, E=가격, F=구매일시, G=장착여부)
  const purchaseId = 'PUR_' + new Date().getTime();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([purchaseId, studentName, itemId, itemName, price, ts, true]);

  // 캐시 무효화
  CacheService.getScriptCache().remove('student_' + studentName);
  updateRankings();

  return { success: true, msg: `[${itemName}] 구매 완료! $${price} 차감되었습니다.`, itemId, resourceVal: String(itemRow[7]).trim(), category: String(itemRow[1]).trim() };
}

// 학생의 구매 아이템 목록 반환 (로그인 시 호출 → 스킨 복원용)
function getOwnedItems(studentName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_SHOP_LOG);
  const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  if (!logSheet || !itemSheet) return [];

  // 아이템 리소스값 맵
  const resourceMap = {};
  const iData = itemSheet.getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    if (iData[i][0]) {
      resourceMap[String(iData[i][0]).trim()] = {
        category: String(iData[i][1]).trim(),
        itemName: String(iData[i][2]).trim(),
        resourceVal: String(iData[i][7]).trim()
      };
    }
  }

  const lData = logSheet.getDataRange().getValues();
  const owned = [];
  for (let i = 1; i < lData.length; i++) {
    if (String(lData[i][1]).trim() !== studentName) continue;
    const itemId = String(lData[i][2]).trim();
    const info   = resourceMap[itemId] || {};
    owned.push({ itemId, itemName: String(lData[i][3]), category: info.category || '', resourceVal: info.resourceVal || '' });
  }
  return owned;
}

// 장착된 아이템만 반환 (로그인 시 복원용) - G열 장착여부 TRUE인 것만
function getEquippedItems(studentName) {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet  = ss.getSheetByName(SHEET_SHOP_LOG);
    const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
    if (!logSheet || !itemSheet) return [];

    const resourceMap = {};
    const iData = itemSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      if (iData[i][0]) {
        resourceMap[String(iData[i][0]).trim()] = {
          category:    String(iData[i][1]).trim(),
          itemName:    String(iData[i][2]).trim(),
          resourceVal: String(iData[i][7]).trim()
        };
      }
    }

    const lData    = logSheet.getDataRange().getValues();
    const equipped = [];
    for (let i = 1; i < lData.length; i++) {
      if (String(lData[i][1]).trim() !== studentName) continue;
      const isEquipped = lData[i][6] === true || String(lData[i][6]).toUpperCase() === 'TRUE';
      if (!isEquipped) continue;
      const itemId = String(lData[i][2]).trim();
      const info   = resourceMap[itemId] || {};
      equipped.push({
        itemId,
        itemName:    String(lData[i][3]),
        category:    info.category    || '',
        resourceVal: info.resourceVal || ''
      });
    }
    return equipped;
  } catch(e) {
    Logger.log('getEquippedItems 오류: ' + e.toString());
    return [];
  }
}

// 아이템 장착 처리 (보유 중인 아이템을 장착으로 변경)
function equipShopItem(studentName, itemId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet  = ss.getSheetByName(SHEET_SHOP_LOG);
  const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  if (!logSheet || !itemSheet) return { success: false, msg: '시트 오류' };

  // 아이템 정보 조회
  const iData = itemSheet.getDataRange().getValues();
  let category    = '';
  let resourceVal = '';
  for (let i = 1; i < iData.length; i++) {
    if (String(iData[i][0]).trim() === itemId) {
      category    = String(iData[i][1]).trim();
      resourceVal = String(iData[i][7]).trim();
      break;
    }
  }
  if (!category) return { success: false, msg: '아이템을 찾을 수 없습니다.' };

  // 같은 카테고리의 기존 장착 아이템 전체 해제 후 해당 아이템 장착
  const lData = logSheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < lData.length; i++) {
    if (String(lData[i][1]).trim() !== studentName) continue;
    const rowItemId = String(lData[i][2]).trim();
    // 같은 카테고리 여부 확인
    for (let j = 1; j < iData.length; j++) {
      if (String(iData[j][0]).trim() === rowItemId &&
          String(iData[j][1]).trim() === category) {
        logSheet.getRange(i + 1, 7).setValue(false);
        break;
      }
    }
    if (rowItemId === itemId) targetRow = i + 1;
  }

  if (targetRow === -1) return { success: false, msg: '보유하지 않은 아이템입니다.' };
  logSheet.getRange(targetRow, 7).setValue(true);

  return { success: true, msg: '장착 완료!', category, resourceVal };
}

// 아이템 장착 해제 (카테고리 기준 전체 해제)
function unequipShopItem(studentName, category) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet  = ss.getSheetByName(SHEET_SHOP_LOG);
  const itemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
  if (!logSheet || !itemSheet) return { success: false, msg: '시트 오류' };

  const iData = itemSheet.getDataRange().getValues();
  const lData = logSheet.getDataRange().getValues();

  for (let i = 1; i < lData.length; i++) {
    if (String(lData[i][1]).trim() !== studentName) continue;
    const rowItemId = String(lData[i][2]).trim();
    for (let j = 1; j < iData.length; j++) {
      if (String(iData[j][0]).trim() === rowItemId &&
          String(iData[j][1]).trim() === category) {
        logSheet.getRange(i + 1, 7).setValue(false);
        break;
      }
    }
  }
  return { success: true, msg: '장착 해제 완료!', category, resourceVal: 'default' };
}

// ════════════════════════════════════════════════════════════════
// ██ 기능 4-b: 업적 마일스톤 자산 보상 트리거
// checkAndGrantAchievements 내부 또는 approveAchievementWithMail 이후 호출
// ════════════════════════════════════════════════════════════════
function grantMilestoneReward(studentName, achCount) {
  // 마일스톤별 보상 정의 (업적 개수 → 자산 보상)
  const MILESTONES = { 5: 500, 10: 1000, 15: 1500, 20: 3000, 30: 5000 };
  const reward = MILESTONES[achCount];
  if (!reward) return;

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  const data = mainSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NAME-1]).trim() !== studentName) continue;
    const current = Number(data[i][COL_ASSET-1]) || 0;
    mainSheet.getRange(i + 1, COL_ASSET).setValue(current + reward);

    // 히스토리 기록
    const histSheet = ss.getSheetByName(SHEET_HISTORY);
    if (histSheet) {
      const ts = _todayStr();
      histSheet.appendRow([ts, studentName, data[i][COL_BRAND-1], '업적보상', reward, data[i][COL_VALUE-1], current + reward, `업적 ${achCount}개 달성 자동 보상`]);
    }


    // 우편함 알림 발송
    _sendMail(
      studentName,
      `🎁 업적 ${achCount}개 달성 보상!`,
      `축하합니다! 업적 ${achCount}개를 달성하여 자동 보상 $${reward}이 지급되었습니다! 계속 도전하세요! 🚀`,
      '보상'
    );
    break;
  }
  CacheService.getScriptCache().remove('student_' + studentName);
}

// onOpen 메뉴에 상점 초기화 추가 (기존 onOpen 대신 별도 등록)
function addShopMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🏪 상점 관리')
    .addItem('상점 시트 초기화 (최초 1회)', 'initShopSheet')
    .addToUi();
}

// ════════════════════════════════════════════════════════════════
// 일괄 승인/반려 + 우편함 발송 버전
// ════════════════════════════════════════════════════════════════
function batchApproveAchievementsWithMail(rowNumbers, isApproved, rejectReason) {
  let successCount = 0;
  let failCount    = 0;
  const msgs = [];

  rowNumbers.forEach(function(rowNum) {
    try {
      const res = approveAchievementWithMail(rowNum, isApproved, null, rejectReason || '조건 미충족');
      if (res.success) {
        successCount++;
        // 승인인 경우 마일스톤 체크
        if (isApproved) {
          const ss       = SpreadsheetApp.getActiveSpreadsheet();
          const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
          if (logSheet) {
            const row         = logSheet.getRange(rowNum, 1, 1, 5).getValues()[0];
            const studentName = String(row[1]).trim();
            const achSheet    = ss.getSheetByName(SHEET_ACH_STUDENT);
            const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);

            // 총 업적 수 집계
            let count = 0;
            const gradeMap = {};
            if (masterSheet) {
              const mData = masterSheet.getDataRange().getValues();
              for (let m = 1; m < mData.length; m++) {
                if (mData[m][0]) gradeMap[String(mData[m][0]).trim()] = String(mData[m][5] || '희귀').trim();
              }
            }
            let achGrade = '희귀';
            if (achSheet) {
              const aData = achSheet.getDataRange().getValues();
              for (let i = 1; i < aData.length; i++) {
                if (String(aData[i][0]).trim() === studentName) {
                  count++;
                  const id = String(aData[i][1]).trim();
                  if (gradeMap[id]) achGrade = gradeMap[id];
                }
              }
            }

            // 마일스톤 자산 보상
            grantMilestoneReward(studentName, count);

            // 전광판 체크
            const achNameRow = logSheet.getRange(rowNum, 1, 1, 5).getValues()[0];
            let achName = String(achNameRow[2]).trim();
            _checkAndPostGlobalAlert(studentName, achName, achGrade);
          }
        }
      } else {
        failCount++;
      }
    } catch(e) {
      failCount++;
    }
  });

  return {
    success: true,
    msg: `일괄 처리 완료: 성공 ${successCount}건, 실패/중복 ${failCount}건`
  };
}

// ════════════════════════════════════════════════════════════════
// ██ 기부 시스템
// 학생이 자신의 자산을 복지 기금으로 자발 기부
// ════════════════════════════════════════════════════════════════
function donateToWelfare(studentName, amount, message) {
  if (!amount || amount <= 0) return { success: false, msg: '금액이 올바르지 않습니다.' };

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };

  const data = mainSheet.getDataRange().getValues();
  let studentRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      studentRowIdx = i;
      break;
    }
  }
  if (studentRowIdx === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

  const curAsset = Number(data[studentRowIdx][COL_ASSET - 1]) || 0;
  if (curAsset < amount) {
    return { success: false, msg: `잔액이 부족합니다. (현재: $${curAsset.toLocaleString()})` };
  }

  const curTax    = Number(data[studentRowIdx][COL_TAX - 1]) || 0;
  const curValue  = Number(data[studentRowIdx][COL_VALUE - 1]) || 0;
  const newAsset  = curAsset - amount;
  const newTax    = curTax + amount;  // 복지 기금 누적에 합산

  // 자산 차감 + 복지기금(세금) 누적
  mainSheet.getRange(studentRowIdx + 1, COL_ASSET).setValue(newAsset);
  mainSheet.getRange(studentRowIdx + 1, COL_TAX).setValue(newTax);

  // 히스토리 기록
  const today    = _todayStr();
  const memo     = message ? `[기부] ${message}` : '[복지 기금 기부]';
  const histSheet = ss.getSheetByName(SHEET_HISTORY);
  if (histSheet) {
    histSheet.appendRow([
      today,
      studentName,
      data[studentRowIdx][COL_BRAND - 1],
      0,          // 브랜드가치 변동 없음
      -amount,    // 자산 변동
      curValue,   // 브랜드가치 (변동 없음)
      newAsset,   // 새 자산
      memo
    ]);
  }

  // 자산사용 시트 기록
  const spendSheet = ss.getSheetByName(SHEET_SPEND);
  if (spendSheet) {
    spendSheet.appendRow([today, studentName, data[studentRowIdx][COL_BRAND - 1], '기부', amount, newAsset, memo]);
  }

  // 캐시 무효화: 기부자 본인 + 전체 학생 캐시 삭제
  // (복지기금 합계는 전 학생에게 동일하게 보여야 하므로 전체 무효화)
  const cache = CacheService.getScriptCache();
  cache.remove('student_' + studentName);
  // 다른 학생들의 캐시도 무효화 (메인 시트에서 이름 목록 조회)
  const allNames = mainSheet.getDataRange().getValues()
    .slice(1)
    .map(r => String(r[COL_NAME - 1]).trim())
    .filter(n => n && n !== studentName);
  allNames.forEach(n => cache.remove('student_' + n));

  updateRankings();

  return {
    success: true,
    msg: `$${amount.toLocaleString()} 기부 완료! 따뜻한 마음 감사합니다 💚`
  };
}

function testWallOfFame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());
  Logger.log('전체 시트 목록: ' + JSON.stringify(sheets));
  Logger.log('업적마스터 시트: ' + (ss.getSheetByName('업적마스터') ? '있음' : '없음'));
  Logger.log('학생업적달성 시트: ' + (ss.getSheetByName('학생업적달성') ? '있음' : '없음'));
}

function testWallOfFame2() {
  try {
    const result = getWallOfFame();
    Logger.log('결과 개수: ' + result.length);
    if (result.length > 0) {
      Logger.log('첫 번째 항목: ' + JSON.stringify(result[0]));
    }
  } catch(e) {
    Logger.log('오류 발생: ' + e.toString());
    Logger.log('오류 위치: ' + e.stack);
  }
}

// ════════════════════════════════════════════════════════════════
// ██ P2P 거래 시스템 추가
// 시트: P2P거래로그
//   A=거래ID, B=날짜, C=보내는학생, D=받는학생, E=금액,
//   F=태그, G=거래설명, H=상태(정상/이상거래)
// ════════════════════════════════════════════════════════════════

const SHEET_P2P = 'P2P거래로그';

// ── 거래 가능한 학생 목록 반환 (본인 제외) ───────────────────────
function getP2PReceiverList(studentName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
  const result   = [];
  for (let i = 1; i < mainData.length; i++) {
    const name = String(mainData[i][COL_NAME - 1]).trim();
    if (!name) continue;
    if (name === String(studentName).trim()) continue;  // 본인 제외
    result.push({
      name:    name,
      brand:   String(mainData[i][COL_BRAND - 1]).trim(),
      balance: Number(mainData[i][COL_ASSET - 1]) || 0
    });
  }
  return result;
}

// ── P2P 거래 실행 ────────────────────────────────────────────────
// senderName   : 보내는 학생 이름
// receiverName : 받는 학생 이름
// amount       : 거래 금액
// tag          : 태그 (#학습도움 / #정서적지지 / #재능판매 / #권리 및 기회 / #기타)
// description  : 거래 설명
function p2pTransfer(senderName, receiverName, amount, tag, description) {
  // ── 기본 유효성 검사 ─────────────────────────────────────────
  if (!receiverName || !receiverName.trim()) {
    return { success: false, msg: '받는 학생을 선택해주세요.' };
  }
  if (String(senderName).trim() === String(receiverName).trim()) {
    return { success: false, msg: '자기 자신에게는 거래할 수 없습니다.' };
  }
  amount = Number(amount);
  if (!amount || amount <= 0 || !Number.isInteger(amount)) {
    return { success: false, msg: '금액은 1 이상의 정수로 입력해주세요.' };
  }
  if (amount > 10000) {
    return { success: false, msg: '1회 거래 한도는 $10,000 입니다.' };
  }
  if (!tag) {
    return { success: false, msg: '거래 태그를 선택해주세요.' };
  }
  if (!description || !description.trim()) {
    return { success: false, msg: '거래 설명을 입력해주세요.' };
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };

  const mainData = mainSheet.getDataRange().getValues();

  // 보내는 학생 행 찾기
  let senderIdx = -1, receiverIdx = -1;
  for (let i = 1; i < mainData.length; i++) {
    const name = String(mainData[i][COL_NAME - 1]).trim();
    if (name === String(senderName).trim())   senderIdx   = i;
    if (name === String(receiverName).trim()) receiverIdx = i;
  }

  if (senderIdx   === -1) return { success: false, msg: '보내는 학생을 찾을 수 없습니다.' };
  if (receiverIdx === -1) return { success: false, msg: '받는 학생을 찾을 수 없습니다.' };

  const senderBalance = Number(mainData[senderIdx][COL_ASSET - 1]) || 0;
  if (senderBalance < amount) {
    return { success: false, msg: `잔액이 부족합니다. (현재: $${senderBalance.toLocaleString()})` };
  }

  // ── 이상 거래 감지 ───────────────────────────────────────────
  // 기준: 오늘 동일인에게 3회 이상 / 단일 거래 $2000 초과
  const p2pSheet = ss.getSheetByName(SHEET_P2P);
  let isAnomaly  = false;
  let anomalyReason = '';

  if (p2pSheet) {
    const p2pData  = p2pSheet.getDataRange().getValues();
    const today    = _todayStr();
    let todaySameCount = 0;
    let todaySameTotal = 0;

    for (let i = 1; i < p2pData.length; i++) {
      const rowDate   = String(p2pData[i][1]).substring(0, 10); // B열: 날짜
      const rowSender = String(p2pData[i][2]).trim();           // C열: 보내는 학생
      const rowRecv   = String(p2pData[i][3]).trim();           // D열: 받는 학생
      if (rowDate === today && rowSender === String(senderName).trim() && rowRecv === String(receiverName).trim()) {
        todaySameCount++;
        todaySameTotal += Number(p2pData[i][4]) || 0;
      }
    }

    if (todaySameCount >= 3) {
      isAnomaly     = true;
      anomalyReason = `오늘 동일인 ${todaySameCount + 1}회 거래`;
    }
    if (amount >= 2000) {
      isAnomaly     = true;
      anomalyReason = (anomalyReason ? anomalyReason + ' / ' : '') + `단일 거래 $${amount.toLocaleString()}`;
    }
  }

  // ── 자산 이동 ────────────────────────────────────────────────
  const newSenderBalance   = senderBalance - amount;
  const receiverBalance    = Number(mainData[receiverIdx][COL_ASSET - 1]) || 0;
  const newReceiverBalance = receiverBalance + amount;

  // 소득세 계산 (받는 학생 기준 10%)
  const taxAmount      = Math.floor(amount * 0.1);
  const netReceived    = amount - taxAmount;
  const afterTaxBalance = receiverBalance + netReceived;
  const receiverTax    = Number(mainData[receiverIdx][COL_TAX - 1]) || 0;

  mainSheet.getRange(senderIdx   + 1, COL_ASSET).setValue(newSenderBalance);
  mainSheet.getRange(receiverIdx + 1, COL_ASSET).setValue(afterTaxBalance);
  mainSheet.getRange(receiverIdx + 1, COL_TAX  ).setValue(receiverTax + taxAmount);

  // ── 히스토리 기록 (보내는 쪽) ───────────────────────────────
  const today     = _todayStr();
  const histSheet = ss.getSheetByName(SHEET_HISTORY);
  if (histSheet) {
    histSheet.appendRow([
      today,
      senderName,
      mainData[senderIdx][COL_BRAND - 1],
      0,           // 브랜드가치 변동 없음
      -amount,
      mainData[senderIdx][COL_VALUE - 1],
      newSenderBalance,
      `[P2P송금→${receiverName}] ${tag} ${description}`
    ]);
    histSheet.appendRow([
      today,
      receiverName,
      mainData[receiverIdx][COL_BRAND - 1],
      0,
      netReceived,  // 세후 수령액
      mainData[receiverIdx][COL_VALUE - 1],
      afterTaxBalance,
      `[P2P수령←${senderName}] ${tag} ${description} (세금 $${taxAmount} 자동 차감)`
    ]);
  }

  // ── 자산사용 시트 기록 (보내는 쪽) ──────────────────────────
  const spendSheet = ss.getSheetByName(SHEET_SPEND);
  if (spendSheet) {
    spendSheet.appendRow([
      today, senderName, mainData[senderIdx][COL_BRAND - 1],
      `[P2P송금] ${tag}`, amount, newSenderBalance,
      `→${receiverName}: ${description}`
    ]);
  }

  // ── P2P 거래 로그 기록 ───────────────────────────────────────
  if (p2pSheet) {
    const txnId = 'TXN_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 4);
    p2pSheet.appendRow([
      txnId,
      today,
      senderName,
      receiverName,
      amount,
      tag,
      description.trim(),
      isAnomaly ? '이상거래' : '정상'
    ]);
  }

  // ── 우편함 알림 (받는 학생에게) ──────────────────────────────
  _sendMail(
    receiverName,
    `💸 P2P 거래 수령 알림`,
    `[${senderName}] 학생에게 $${amount.toLocaleString()}을 받았습니다.\n태그: ${tag}\n내용: ${description}\n\n소득세 $${taxAmount} 자동 차감 후 실수령액: $${netReceived.toLocaleString()}`,
    '거래'
  );

  // ── 랭킹 갱신 + 캐시 무효화 ─────────────────────────────────
  updateRankings();
  const cache = CacheService.getScriptCache();
  cache.remove('student_' + senderName);
  cache.remove('student_' + receiverName);

  return {
    success:        true,
    msg:            `거래 완료! $${amount.toLocaleString()} 송금 (상대방 세후 수령: $${netReceived.toLocaleString()})`,
    newBalance:     newSenderBalance,
    isAnomaly:      isAnomaly,
    anomalyReason:  anomalyReason
  };
}

// ── 나의 P2P 거래 내역 반환 ──────────────────────────────────────
function getMyP2PHistory(studentName) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return [];

  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const sender   = String(data[i][2]).trim();
    const receiver = String(data[i][3]).trim();
    const name     = String(studentName).trim();
    if (sender !== name && receiver !== name) continue;

    result.push({
      txnId:       String(data[i][0]),
      date:        String(data[i][1]),
      sender:      sender,
      receiver:    receiver,
      amount:      Number(data[i][4]) || 0,
      tag:         String(data[i][5]),
      description: String(data[i][6]),
      status:      String(data[i][7]),
      isSent:      sender === name  // true=내가 보낸 것, false=내가 받은 것
    });
  }
  return result.reverse(); // 최신순
}

// ── 교사용: 이상 거래 목록 반환 ──────────────────────────────────
function getP2PAlerts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return [];

  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7]).trim() !== '이상거래') continue;
    result.push({
      rowNum:      i + 1,
      txnId:       String(data[i][0]),
      date:        String(data[i][1]),
      sender:      String(data[i][2]),
      receiver:    String(data[i][3]),
      amount:      Number(data[i][4]) || 0,
      tag:         String(data[i][5]),
      description: String(data[i][6])
    });
  }
  return result.reverse(); // 최신순
}

// ── 교사용: 이상 거래 상태 수동 변경 ('이상거래' → '정상 확인됨') ─
function resolveP2PAlert(rowNum) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return { success: false, msg: 'P2P거래로그 시트를 찾을 수 없습니다.' };
  sheet.getRange(rowNum, 8).setValue('정상 확인됨');
  return { success: true, msg: '정상 확인 처리되었습니다.' };
}
