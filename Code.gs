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
const SHEET_ACH_LOG       = '업적신청로그';
const SHEET_GLOBAL_NOTIFY = '전역알림';
const SHEET_JOB2_APP      = '2차직업신청';
const SHEET_JOB2_CURR     = '2차직업현황';
const SHEET_MAILBOX       = '우편함_로그';
const SHEET_SHOP_ITEMS    = '상점_아이템';
const SHEET_SHOP_LOG      = '상점_구매로그';
const SHEET_P2P           = 'P2P거래로그';
const SHEET_DEPOSIT_PROD  = '예금상품';
const SHEET_DEPOSIT_LOG   = '학생별가입예금';
const SHEET_GUARD_PENALTY = '수호대적발로그';
const SHEET_LOAN_STATUS   = '대출현황';
const SHEET_LOAN_LOG      = '대출신청로그';
const SHEET_CREDIT_HISTORY = '신용점수이력';
const SHEET_INVENTORY     = '인벤토리';
const SHEET_EMERGENCY     = '비상사태현황';
const SHEET_ASSIGN_LIST   = '과제목록';
const SHEET_ASSIGN_SUBMIT = '과제제출';
const ASSIGN_DRIVE_FOLDER = '2026BRAND_과제제출';   // Google Drive 자동 생성 폴더명
const BACKUP_FOLDER_NAME  = 'B.R.A.N.D 자동백업';
const WEEKLY_BUY_LIMIT    = 5;

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
const TIER_ORDER = [
  '새싹',        // 1
  '브론즈',       // 2
  '빛나는 브론즈', // 3
  '거친 실버',    // 4
  '성장한 실버',   // 5
  '진화한 실버',   // 6
  '은빛 극점',    // 7
  '금 광석',      // 8
  '제련된 골드',   // 9
  '정련된 골드',   // 10
  '태양의 황금',   // 11
  '루비 원석',    // 12
  '연마된 루비',   // 13
  '각성한 루비',   // 14
  '홍염의 정점',   // 15
  '다이아 원석',   // 16
  '세공된 다이아', // 17
  '무결 다이아',   // 18
  '영원의 결정',   // 19
  '마스터',       // 20
  '천상의 마스터', // 21
  '그랜드마스터'   // 22
];

// ════════════════════════════════════════════════════════════════
// 1. 웹앱 진입점 (URL로 접속 시 어떤 화면을 보여줄지 결정)
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const param = (e && e.parameter) ? e.parameter : {};
  const page = param.page;
  const mode = param.mode;

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
  if (page === 'guard') {
    return HtmlService.createTemplateFromFile('GuardDashboard').evaluate()
      .setTitle('경제 수호대 대시보드')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'verify') {
    return HtmlService.createTemplateFromFile('AchievementVerify').evaluate()
      .setTitle('업적 검증 도우미')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
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
    .addSeparator()
    .addItem('🚀 [배포] 웹앱 URL 안내',           'showDeployInfo')
    .addItem('💾 [백업] 지금 즉시 백업',           'runManualBackup')
    .addItem('⏰ [백업] 자동 백업 스케줄 설정',    'setupDailyBackupTrigger')
    .addItem('⏰ [추적] 브랜드가치 자동 기록 설정', 'setupDailyTrackerTrigger')
    .addItem('📋 [백업] 백업 목록 확인',           'showBackupList')
    .addItem('⏰ [예금] 만기 자동 처리 트리거 설정', 'setupDepositTrigger')
    .addSeparator()
    .addItem('🔥 [Firebase] 전체 학생 스냅샷 동기화', 'syncAllStudentsToFirebase')
    .addSeparator()
    .addItem('⚡ [속도] 워밍업 트리거 설정 (수업시간 자동 유지)', 'setupWarmupTrigger')
    .addItem('🛑 [속도] 워밍업 트리거 삭제',                    'removeWarmupTrigger')
    .addToUi();
}

function finalizeDailyTracker() {
  _updateTracker(_todayStr(), null);
  SpreadsheetApp.getUi().alert('✅ 오늘의 브랜드 가치가 추적 시트에 최종 기록되었습니다.');
}


// ════════════════════════════════════════════════════════════════
// 3-0. 비밀번호 검증만 수행 (Firebase 캐시 히트 시 호출)
// ════════════════════════════════════════════════════════════════
function verifyStudentPassword(studentName, password) {
  if (!_validateStudentName(studentName)) {
    return { success: false, msg: '유효하지 않은 이름입니다.' };
  }
  if (!_validatePassword(password)) {
    return { success: false, msg: '유효하지 않은 비밀번호입니다.' };
  }
  studentName = String(studentName).trim();

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();

  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === studentName) {
      const correctPw  = String(mainData[i][COL_PASSWORD - 1]).trim();
      const inputPw    = (password === null || password === undefined) ? null : String(password).trim();
      const masterPw   = _getMasterPassword();
      const isMaster   = (masterPw !== null && inputPw === masterPw);
      if (!isMaster && inputPw !== null && correctPw && inputPw !== correctPw) {
        return { success: false, msg: '비밀번호가 일치하지 않습니다.' };
      }
      try {
        const loginLog = ss.getSheetByName('로그인_로그');
        if (loginLog) {
          const now     = new Date();
          const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
          loginLog.appendRow([dateStr, studentName, timeStr, isMaster ? '마스터접속' : '']);
        }
      } catch(e) {}
      return { success: true };
    }
  }
  return { success: false, msg: '학생을 찾을 수 없습니다.' };
}

// ════════════════════════════════════════════════════════════════
// 3. 학생 대시보드 데이터 (Index.html 에서 호출)
// ════════════════════════════════════════════════════════════════
function getStudentData(studentName, password) {
  if (!_validateStudentName(studentName)) {
    return { success: false, msg: '유효하지 않은 이름입니다.' };
  }
  if (!_validatePassword(password)) {
    return { success: false, msg: '유효하지 않은 비밀번호입니다.' };
  }
  studentName = String(studentName).trim();

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();

  let studentRow = null;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      studentRow = mainData[i];
      break;
    }
  }
  if (!studentRow) return { success: false, msg: '학생을 찾을 수 없습니다. 이름을 다시 확인해주세요.' };

  const correctPassword = String(studentRow[COL_PASSWORD - 1]).trim();
  const inputPassword   = (password === null || password === undefined) ? null : String(password).trim();
  const masterPassword  = _getMasterPassword();
  const isMasterLogin   = (masterPassword !== null && inputPassword === masterPassword);
  if (!isMasterLogin && inputPassword !== null && correctPassword && inputPassword !== correctPassword) {
    return { success: false, msg: '비밀번호가 일치하지 않습니다.' };
  }

  try {
    const loginLog = ss.getSheetByName('로그인_로그');
    if (loginLog) {
      const now     = new Date();
      const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
      loginLog.appendRow([dateStr, studentName, timeStr, isMasterLogin ? '마스터접속' : '']);
    }
  } catch(e) {}

  let totalTax = 0;
  for (let i = 1; i < mainData.length; i++) {
    totalTax += Number(mainData[i][COL_TAX - 1]) || 0;
  }

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

  const auctionSheet  = ss.getSheetByName(SHEET_AUCTION);
  const auctionPrices = [];
  if (auctionSheet) {
    const aData = auctionSheet.getDataRange().getValues();
    for (let m = 1; m < aData.length; m++) {
      if (!aData[m][0]) continue;
      auctionPrices.push({
        item:  `[${aData[m][0]}] ${aData[m][1] || ''}`.trim(),
        price: Number(aData[m][11]) || 0
      });
    }
  }

  const honor = Number(studentRow[COL_VALUE - 1]) || 0;
  const tier  = _calcTier(honor);
  const emergency = getEmergencyStatus();

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
    myDonation:    0,
    classTotalTax: totalTax,
    job:           jobResult,
    auctionPrices: auctionPrices,
    tierData:      tier,
    emergency:     emergency
  };
}

// ════════════════════════════════════════════════════════════════
// 3-1. 학생 대시보드 부가 데이터
// ════════════════════════════════════════════════════════════════
function getStudentDataSub(studentName) {
  if (!_validateStudentName(studentName)) {
    return { success: false, msg: '유효하지 않은 이름입니다.' };
  }
  studentName = String(studentName).trim();

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();

  let asset = 0, tax = 0, honor = 0;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === studentName) {
      asset = Number(mainData[i][COL_ASSET - 1]) || 0;
      tax   = Number(mainData[i][COL_TAX - 1])   || 0;
      honor = Number(mainData[i][COL_VALUE - 1])  || 0;
      break;
    }
  }

  checkAndGrantAchievements(studentName, asset, tax, honor);

  let myDonation = 0;
  try {
    const spendSh = ss.getSheetByName(SHEET_SPEND);
    if (spendSh && spendSh.getLastRow() >= 2) {
      const spendData = spendSh.getRange(2, 1, spendSh.getLastRow() - 1, 5).getValues();
      myDonation = spendData.reduce(function(sum, row) {
        return (row[1] === studentName && row[3] === '기부') ? sum + (Number(row[4]) || 0) : sum;
      }, 0);
    }
  } catch(e) {}

  return {
    success:    true,
    snacks:     getSnackData(),
    achievements: getStudentAchievements(studentName),
    job2:       getSecondaryJobForStudent(studentName),
    jobMarket:  getJobData(),
    myDonation: myDonation
  };
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
  const nowStr    = _nowStr();

  entries.forEach(e => {
    const rowIdx       = e.row;
    const curValue     = Number(mainData[rowIdx][COL_VALUE - 1]) || 0;
    const curAsset     = Number(mainData[rowIdx][COL_ASSET - 1]) || 0;
    const curTax       = Number(mainData[rowIdx][COL_TAX - 1])   || 0;
    const taxAmount    = e.points > 0 ? Math.floor(e.points * (taxRate / 100)) : 0;
    const netAssetGain = e.points - taxAmount;

    mainData[rowIdx][COL_VALUE - 1] = curValue + e.points;
    mainData[rowIdx][COL_ASSET - 1] = curAsset + netAssetGain;
    mainData[rowIdx][COL_TAX - 1]   = curTax + taxAmount;

    histRows.push([
      date, e.name, e.brand,
      e.points, netAssetGain,
      curValue + e.points, curAsset + netAssetGain,
      e.note + (taxAmount > 0 ? ` (세금 ${taxAmount})` : ''),
      nowStr
    ]);
  });

  mainRange.setValues(mainData);
  if (histRows.length > 0) {
    hist.getRange(hist.getLastRow() + 1, 1, histRows.length, 9).setValues(histRows);
  }
  updateRankings();
  return `✅ ${entries.length}명 포인트 지급 완료!`;
}

// ════════════════════════════════════════════════════════════════
// ★ 신규: 자산만 지급 (브랜드가치 변동 없음, 세금 적용)
// ════════════════════════════════════════════════════════════════
function applyAssetOnly(date, entries, taxRate) {
  taxRate = taxRate || 0;
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const main      = ss.getSheetByName(SHEET_MAIN);
  const hist      = ss.getSheetByName(SHEET_HISTORY);
  const mainRange = main.getDataRange();
  const mainData  = mainRange.getValues();
  const histRows  = [];
  const nowStr    = _nowStr();

  entries.forEach(e => {
    const rowIdx       = e.row;
    const curValue     = Number(mainData[rowIdx][COL_VALUE - 1]) || 0;  // 변동 없음
    const curAsset     = Number(mainData[rowIdx][COL_ASSET - 1]) || 0;
    const curTax       = Number(mainData[rowIdx][COL_TAX - 1])   || 0;
    const taxAmount    = e.points > 0 ? Math.floor(e.points * (taxRate / 100)) : 0;
    const netAssetGain = e.points - taxAmount;

    // ★ COL_VALUE 는 건드리지 않음
    mainData[rowIdx][COL_ASSET - 1] = curAsset + netAssetGain;
    mainData[rowIdx][COL_TAX - 1]   = curTax + taxAmount;

    histRows.push([
      date, e.name, e.brand,
      0,                          // BV변동 0
      netAssetGain,
      curValue,                   // BV 기존값 유지
      curAsset + netAssetGain,
      '[자산지급] ' + e.note + (taxAmount > 0 ? ` (세금 ${taxAmount})` : ''),
      nowStr
    ]);
  });

  mainRange.setValues(mainData);
  if (histRows.length > 0) {
    hist.getRange(hist.getLastRow() + 1, 1, histRows.length, 9).setValues(histRows);
  }
  updateRankings();
  return `✅ ${entries.length}명 자산 지급 완료! (브랜드가치 변동 없음)`;
}

function recordSpend(date, rowIdx, name, brand, category, amount, note) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const main     = ss.getSheetByName(SHEET_MAIN);
  const curAsset = Number(main.getRange(rowIdx + 1, COL_ASSET).getValue()) || 0;
  if (curAsset < amount) return '❌ 자산이 부족합니다!';
  const newAsset = curAsset - amount;
  main.getRange(rowIdx + 1, COL_ASSET).setValue(newAsset);
  ss.getSheetByName(SHEET_SPEND).appendRow([date, name, brand, category, amount, newAsset, note, _nowStr()]);
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
    curValue + amount, curAsset + amount, `[MVP] ${note}`, _nowStr()
  ]);
  updateRankings();
  return '🏆 MVP 포인트 지급 완료!';
}

// ════════════════════════════════════════════════════════════════
// ★ 신규: MVP 다중 지급 (여러 학생 동시)
// ════════════════════════════════════════════════════════════════
function grantMvpMulti(date, entries, note) {
  if (!entries || entries.length === 0) return '❌ 선택된 학생이 없습니다.';
  note = note || 'MVP';

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const main      = ss.getSheetByName(SHEET_MAIN);
  const hist      = ss.getSheetByName(SHEET_HISTORY);
  const mainRange = main.getDataRange();
  const mainData  = mainRange.getValues();
  const histRows  = [];
  const nowStr    = _nowStr();

  entries.forEach(e => {
    const rowIdx   = e.row;
    const amount   = Number(e.amount) || 0;
    if (amount === 0) return;
    const curValue = Number(mainData[rowIdx][COL_VALUE - 1]) || 0;
    const curAsset = Number(mainData[rowIdx][COL_ASSET - 1]) || 0;

    mainData[rowIdx][COL_VALUE - 1] = curValue + amount;
    mainData[rowIdx][COL_ASSET - 1] = curAsset + amount;

    histRows.push([
      date, e.name, e.brand,
      amount, amount,
      curValue + amount, curAsset + amount,
      `[MVP] ${note}`,
      nowStr
    ]);
  });

  mainRange.setValues(mainData);
  if (histRows.length > 0) {
    hist.getRange(hist.getLastRow() + 1, 1, histRows.length, 9).setValues(histRows);
  }
  updateRankings();
  return `🏆 ${histRows.length}명에게 MVP 포인트 지급 완료!`;
}

function _updateRankingsOnly() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const main    = ss.getSheetByName(SHEET_MAIN);
  const lastRow = main.getLastRow();
  if (lastRow < 2) return;
  const data = main.getRange(2, 1, lastRow - 1, COL_MVP).getValues();

  const depositSheet = ss.getSheetByName('학생별가입예금');
  const depositMap   = {};
  if (depositSheet) {
    const depData = depositSheet.getDataRange().getValues();
    for (let i = 1; i < depData.length; i++) {
      const dName     = String(depData[i][1]).trim();
      const principal = Number(depData[i][2]) || 0;
      const dStatus   = String(depData[i][7]).trim();
      if (!dName) continue;
      if (dStatus !== '진행중') continue;
      depositMap[dName] = (depositMap[dName] || 0) + principal;
    }
  }

  const loanSheet = ss.getSheetByName('대출현황');
  const loanMap   = {};
  if (loanSheet) {
    const loanData = loanSheet.getDataRange().getValues();
    for (let i = 1; i < loanData.length; i++) {
      const lName   = String(loanData[i][1]).trim();
      const balance = Number(loanData[i][10]) || 0;
      if (!lName) continue;
      loanMap[lName] = (loanMap[lName] || 0) + balance;
    }
  }

  const vArr = data.map((r, i) => ({ idx: i, v: Number(r[COL_VALUE - 1]) || 0 }));
  const aArr = data.map((r, i) => {
    const name      = String(r[COL_NAME - 1]).trim();
    const realAsset = (Number(r[COL_ASSET - 1]) || 0)
                    + (depositMap[name] || 0)
                    - (loanMap[name]    || 0);
    return { idx: i, v: realAsset };
  });
  const rV = _calcRank(vArr);
  const rA = _calcRank(aArr);
  main.getRange(2, COL_RANK_A, rA.length, 1).setValues(rA.map(r => [r]));
  main.getRange(2, COL_RANK_V, rV.length, 1).setValues(rV.map(r => [r]));
}

function updateRankings() {
  _updateRankingsOnly();
  try { syncAllStudentsToFirebase(); } catch(e) {
    Logger.log('[Firebase 동기화 실패] ' + e.message);
  }
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
// 9. 브랜드가치 일별 추적
// ════════════════════════════════════════════════════════════════
function _updateTracker(date, ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const tracker = ss.getSheetByName(SHEET_TRACKER);
  const main    = ss.getSheetByName(SHEET_MAIN);
  if (!tracker || !main) return;

  const mainData    = main.getDataRange().getValues();
  let trackerData   = tracker.getDataRange().getValues();

  let realLastRow = 0;
  for (let i = 0; i < trackerData.length; i++) {
    if (trackerData[i][0]) realLastRow = i + 1;
  }

  const existingNames = {};
  for (let r = 1; r < trackerData.length; r++) {
    if (trackerData[r][0]) existingNames[trackerData[r][0]] = r + 1;
  }

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

  const headerRow = tracker.getRange(1, 1, 1, tracker.getLastColumn() || 1).getValues()[0];
  let dateCol = headerRow.indexOf(date) + 1;
  if (dateCol === 0) {
    dateCol = (tracker.getLastColumn() || 1) + 1;
    tracker.getRange(1, dateCol).setValue(date).setBackground('#3d85c8').setFontColor('white');
  }

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
function _nowStr()      { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }

// ════════════════════════════════════════════════════════════════
// 11. 다이얼로그 열기 함수들 (메뉴에서 호출)
// ════════════════════════════════════════════════════════════════
function openDailyInput()    { showModal(getDailyInputHtml(), '📅 오늘 포인트 지급',    700, 600); }
function openSpendDialog()   { showModal(getSpendHtml(),      '💸 자산 사용 기록',      500, 500); }
function openMvpDialog()     { showModal(getMvpHtml(),        '🏆 MVP 포인트 지급',     520, 560); }
function openHistoryDialog() { showModal(getHistoryHtml(),    '📊 학생별 히스토리',     750, 550); }

function openSnackDialog() {
  const html = HtmlService.createTemplateFromFile('Snackdialog').evaluate()
    .setWidth(500).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '🍿 간식 판매 처리');
}

function showModal(htmlText, title, w, h) {
  const html = HtmlService.createHtmlOutput(htmlText).setWidth(w).setHeight(h);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}


// ════════════════════════════════════════════════════════════════
// ★ [교체] 일일 포인트 지급 HTML — 지급 방식 선택 추가
//   "브랜드가치+자산(기본)" vs "💰 자산만 지급"
// ════════════════════════════════════════════════════════════════
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
    .mode-bar{background:#fff8e1;border:1px solid #ffe0a0;border-radius:8px;padding:8px 14px;margin-bottom:10px;font-size:13px;}
    .mode-bar label{font-weight:bold;margin-right:14px;cursor:pointer;}
    table{border-collapse:collapse;width:100%;font-size:13px;}
    th{background:#3d85c8;color:white;padding:6px 8px;text-align:left;}
    tr:nth-child(even){background:#f5f9ff;}
    .btn-apply{background:#27ae60;color:white;border:none;border-radius:6px;padding:10px 28px;font-size:15px;cursor:pointer;margin-top:12px;width:100%;}
    .btn-fill{background:#3d85c8;color:white;border:none;border-radius:5px;padding:5px 12px;cursor:pointer;margin-left:6px;font-size:13px;}
  </style></head><body>
  <h3>📅 오늘의 포인트 지급</h3>

  <div class="mode-bar">
    <label><input type="radio" name="payMode" value="both" checked> 📈 브랜드가치 + 자산 (기본)</label>
    <label><input type="radio" name="payMode" value="asset"> 💰 자산만 지급 <span style="color:#e67e22;">(브랜드가치 변동 없음, 세금 적용)</span></label>
  </div>

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
  <button class="btn-apply" id="applyBtn" onclick="applyPoints()">🚀 포인트 및 세금 적용하기</button>
  <script>
    document.querySelectorAll('input[name="payMode"]').forEach(function(r){
      r.addEventListener('change', function(){
        var btn = document.getElementById('applyBtn');
        btn.innerText = (this.value === 'asset')
          ? '💰 자산만 지급하기 (세금 적용)'
          : '🚀 포인트 및 세금 적용하기';
        btn.style.background = (this.value === 'asset') ? '#e67e22' : '#27ae60';
      });
    });
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
        var noteEl = document.getElementById('note_' + row);
        if (noteEl && !noteEl.value.trim()) noteEl.value = '일일퀘스트';
      });
    }
    function loadJobSalariesAndFill() {
      var btn = document.querySelector('[onclick="loadJobSalariesAndFill()"]');
      btn.innerText = '로딩 중...';
      btn.disabled = true;
      google.script.run.withSuccessHandler(function(salaries) {
        window.jobSalaries = salaries;
        autoFillFromJob();
        btn.innerText = '🤖 1인1역 자동계산';
        btn.disabled = false;
      }).getJobSalariesByRow();
    }
    function applyPoints() {
      var date = document.getElementById('today').value;
      var taxRate = parseFloat(document.getElementById('taxRate').value) || 0;
      if (!date) return alert('날짜를 입력해주세요.');
      var mode = document.querySelector('input[name="payMode"]:checked').value;
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
      var btn = document.getElementById('applyBtn');
      btn.disabled = true; btn.innerText = '처리 중...';
      var handler = function(res){ alert(res); google.script.host.close(); };
      if (mode === 'asset') {
        google.script.run.withSuccessHandler(handler).applyAssetOnly(date, entries, taxRate);
      } else {
        google.script.run.withSuccessHandler(handler).applyDailyPoints(date, entries, taxRate);
      }
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


// ════════════════════════════════════════════════════════════════
// ★ [교체] MVP 지급 HTML — 체크박스 다중선택 + 개별/일괄 금액
// ════════════════════════════════════════════════════════════════
function getMvpHtml() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN).getDataRange().getValues();
  let rowsHtml = '';
  for (let i = 1; i < data.length; i++) {
    const name = data[i][COL_NAME - 1], brand = data[i][COL_BRAND - 1];
    if (!name) continue;
    rowsHtml += `<tr>
      <td style="padding:5px 8px;"><label style="cursor:pointer;display:flex;align-items:center;gap:6px;">
        <input type="checkbox" class="mvpChk" data-row="${i}" data-name="${name}" data-brand="${brand}" style="width:16px;height:16px;">
        ${name} <small style="color:#888;">${brand}</small>
      </label></td>
      <td style="padding:5px 8px;"><input type="number" class="mvpAmt" data-row="${i}" placeholder="(일괄값)" step="100" style="width:90px;padding:4px;text-align:right;"></td>
    </tr>`;
  }
  return `<!DOCTYPE html><html><head><style>
    body{font-family:'Noto Sans KR',sans-serif;padding:16px;font-size:14px;}
    h3{margin:0 0 12px;color:#2c3e50;}
    .ctrl{background:#fff3e0;border:1px solid #ffe0a0;border-radius:8px;padding:10px 14px;margin-bottom:10px;}
    .ctrl label{font-weight:bold;}
    input[type=number]{padding:6px;font-size:14px;box-sizing:border-box;}
    table{border-collapse:collapse;width:100%;font-size:13px;}
    th{background:#f39c12;color:white;padding:6px 8px;text-align:left;}
    tr:nth-child(even){background:#fffaf2;}
    .btn{background:#f39c12;color:white;border:none;border-radius:6px;padding:11px;font-size:15px;cursor:pointer;margin-top:14px;width:100%;}
    .btn-mini{background:#e67e22;color:white;border:none;border-radius:5px;padding:5px 12px;cursor:pointer;font-size:13px;margin-left:6px;}
    .cnt{color:#e67e22;font-weight:bold;}
  </style></head><body>
  <h3>🏆 MVP 포인트 지급 <span style="font-size:13px;color:#888;">(여러 명 동시 지급 가능)</span></h3>
  <div class="ctrl">
    <label>날짜:</label> <input type="date" id="mvpDate" value="${_todayStr()}" style="margin-right:12px;">
    <label>일괄 금액:</label> <input type="number" id="bulkAmt" value="1000" step="100" style="width:90px;">
    <button class="btn-mini" onclick="applyBulk()">✔ 선택 학생에 적용</button>
    <button class="btn-mini" style="background:#95a5a6;" onclick="toggleAll()">전체 선택/해제</button>
    <div style="margin-top:8px;font-size:13px;">
      선택: <span class="cnt" id="selCnt">0</span>명 &nbsp;|&nbsp;
      비고: <input type="text" id="mvpNote" value="MVP" style="width:120px;padding:4px;">
      <small style="color:#999;margin-left:6px;">개별 금액 비워두면 일괄값 사용</small>
    </div>
  </div>
  <div style="max-height:340px;overflow-y:auto;border:1px solid #eee;border-radius:6px;">
    <table><thead><tr><th>학생 (체크로 선택)</th><th>개별 지급액</th></tr></thead>
    <tbody>${rowsHtml}</tbody></table>
  </div>
  <button class="btn" onclick="grantSelected()">🏆 선택한 학생에게 MVP 지급</button>
  <script>
    function updateCnt(){
      document.getElementById('selCnt').textContent =
        document.querySelectorAll('.mvpChk:checked').length;
    }
    document.querySelectorAll('.mvpChk').forEach(function(c){ c.addEventListener('change', updateCnt); });
    function applyBulk(){
      var amt = document.getElementById('bulkAmt').value;
      document.querySelectorAll('.mvpChk:checked').forEach(function(chk){
        var amtInput = document.querySelector('.mvpAmt[data-row="'+chk.dataset.row+'"]');
        if (amtInput) amtInput.value = amt;
      });
    }
    function toggleAll(){
      var boxes = document.querySelectorAll('.mvpChk');
      var anyUnchecked = Array.prototype.some.call(boxes, function(b){ return !b.checked; });
      boxes.forEach(function(b){ b.checked = anyUnchecked; });
      updateCnt();
    }
    function grantSelected(){
      var date = document.getElementById('mvpDate').value;
      var note = document.getElementById('mvpNote').value.trim() || 'MVP';
      var bulk = parseInt(document.getElementById('bulkAmt').value) || 0;
      var entries = [];
      document.querySelectorAll('.mvpChk:checked').forEach(function(chk){
        var row = parseInt(chk.dataset.row);
        var amtInput = document.querySelector('.mvpAmt[data-row="'+row+'"]');
        var amt = (amtInput && amtInput.value.trim() !== '') ? parseInt(amtInput.value) : bulk;
        if (amt && amt !== 0) {
          entries.push({ row: row, name: chk.dataset.name, brand: chk.dataset.brand, amount: amt });
        }
      });
      if (entries.length === 0) return alert('지급할 학생을 선택하고 금액을 입력해주세요.');
      if (!confirm(entries.length + '명에게 MVP 포인트를 지급합니다. 계속할까요?')) return;
      var btn = document.querySelector('.btn');
      btn.disabled = true; btn.innerText = '지급 중...';
      google.script.run
        .withSuccessHandler(function(res){ alert(res); google.script.host.close(); })
        .grantMvpMulti(date, entries, note);
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
// ██ 기부 시스템
// ════════════════════════════════════════════════════════════════
function donateToWelfare(studentName, amount, message) {
  if (!amount || amount <= 0) return { success: false, msg: '금액이 올바르지 않습니다.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    if (!mainSheet) return { success: false, msg: '메인 시트를 찾을 수 없습니다.' };

    const data = mainSheet.getDataRange().getValues();
    let studentRowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
        studentRowIdx = i; break;
      }
    }
    if (studentRowIdx === -1) return { success: false, msg: '학생을 찾을 수 없습니다.' };

    const curAsset = Number(data[studentRowIdx][COL_ASSET - 1]) || 0;
    const _emgD = _getActiveEmergency();
    if (_emgD && _emgD.type === '자산 동결') {
      const _usable = Math.floor(curAsset * (_emgD.freezeRate / 100));
      if (amount > _usable) return { success: false, msg: `🔒 자산 동결 중! 사용 가능 금액: $${_usable.toLocaleString()} (보유액의 ${_emgD.freezeRate}%)` };
    }
    if (curAsset < amount) {
      return { success: false, msg: `잔액이 부족합니다. (현재: $${curAsset.toLocaleString()})` };
    }

    const curTax   = Number(data[studentRowIdx][COL_TAX - 1]) || 0;
    const curValue = Number(data[studentRowIdx][COL_VALUE - 1]) || 0;
    const newAsset = curAsset - amount;
    const newTax   = curTax + amount;

    mainSheet.getRange(studentRowIdx + 1, COL_ASSET).setValue(newAsset);
    mainSheet.getRange(studentRowIdx + 1, COL_TAX).setValue(newTax);

    const today    = _nowStr();
    const memo     = message ? `[기부] ${message}` : '[복지 기금 기부]';
    const histSheet = ss.getSheetByName(SHEET_HISTORY);
    if (histSheet) {
      histSheet.appendRow([
        _todayStr(), studentName, data[studentRowIdx][COL_BRAND - 1],
        0, -amount, curValue, newAsset, memo, today
      ]);
    }

    const spendSheet = ss.getSheetByName(SHEET_SPEND);
    if (spendSheet) {
      spendSheet.appendRow([_todayStr(), studentName, data[studentRowIdx][COL_BRAND - 1], '기부', amount, newAsset, memo, today]);
    }

    try { syncOneStudentToFirebase(studentName); } catch(e) {
      Logger.log('[Firebase 동기화 실패] ' + e.message);
    }

    return {
      success:    true,
      newBalance: newAsset,
      msg: `$${amount.toLocaleString()} 기부 완료! 따뜻한 마음 감사합니다 💚`
    };
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ════════════════════════════════════════════════════════════════
// ██ 보안 강화 헬퍼 함수
// ════════════════════════════════════════════════════════════════
function _sanitizeString(input) {
  if (input === null || input === undefined) return '';
  return String(input)
    .replace(/&/g,  '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#x27;').trim();
}

function _sanitizeNumber(input, min, max) {
  const n = Number(input);
  if (isNaN(n)) return null;
  if (min !== undefined && n < min) return null;
  if (max !== undefined && n > max) return null;
  return Math.floor(n);
}

function _validateStudentName(name) {
  if (!name || typeof name !== 'string') return false;
  const clean = name.trim();
  if (clean.length === 0 || clean.length > 20) return false;
  if (/[<>"'&;=()]/.test(clean)) return false;
  return true;
}

function _getMasterPassword() {
  try {
    const mp = PropertiesService.getScriptProperties().getProperty('MASTER_PASSWORD');
    return (mp && String(mp).trim()) ? String(mp).trim() : null;
  } catch (e) { return null; }
}

function setupMasterPassword() {
  const MASTER_PW = 'masterpassword';
  PropertiesService.getScriptProperties()
    .setProperty('MASTER_PASSWORD', String(MASTER_PW).trim());
  try { SpreadsheetApp.getUi().alert('✅ 마스터 비밀번호가 설정되었습니다.'); } catch (e) {}
}

function _validatePassword(password) {
  if (password === null || password === undefined) return true;
  const clean = String(password).trim();
  if (clean.length > 20) return false;
  return true;
}

// ════════════════════════════════════════════════════════════════
// 로렌츠 곡선 복원
// ════════════════════════════════════════════════════════════════
function buildLorenzData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('히스토리');
  if (!historySheet) { SpreadsheetApp.getUi().alert('히스토리 시트를 찾을 수 없습니다.'); return; }

  const SNAPSHOTS = [
    { label: '1회차_직전 (03-12)', date: '2026-03-12' },
    { label: '1회차_직후 (03-13)', date: '2026-03-13' },
    { label: '2회차_직전 (03-26)', date: '2026-03-26' },
    { label: '2회차_직후 (03-27)', date: '2026-03-27' },
    { label: '3회차_직전 (04-23)', date: '2026-04-23' },
    { label: '3회차_직후 (04-24)', date: '2026-04-24' },
    { label: '4회차_직전 (05-27)', date: '2026-05-27' },
    { label: '4회차_직후 (05-28)', date: '2026-05-28' },
    { label: '5회차_직전 (06-24)', date: '2026-06-24' },
    { label: '4회차_직후 (06-27)', date: '2026-06-27' },
  ];

  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('히스토리 데이터가 없습니다.'); return; }
  const rawData = historySheet.getRange(2, 1, lastRow - 1, 8).getValues();

  const OUTPUT_SHEET_NAME = '로렌츠_분석';
  let outSheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (outSheet) { outSheet.clearContents(); } else { outSheet = ss.insertSheet(OUTPUT_SHEET_NAME); }

  let currentOutputRow = 1;
  SNAPSHOTS.forEach(function(snap) {
    const snapshotData = _getSnapshot(rawData, snap.date);
    const students = Object.keys(snapshotData);
    if (students.length === 0) return;

    const assets = students.map(function(n){ return snapshotData[n].asset; }).sort(function(a,b){ return a-b; });
    const brands = students.map(function(n){ return snapshotData[n].brand; }).sort(function(a,b){ return a-b; });
    const lorenzAsset = _calcLorenz(assets);
    const lorenzBrand = _calcLorenz(brands);
    const giniAsset  = _calcGini(lorenzAsset);
    const giniBrand  = _calcGini(lorenzBrand);

    outSheet.getRange(currentOutputRow, 1).setValue('【' + snap.label + '】').setFontWeight('bold');
    outSheet.getRange(currentOutputRow, 2).setValue('자산 지니계수: ' + giniAsset.toFixed(4));
    outSheet.getRange(currentOutputRow, 3).setValue('브랜드 지니계수: ' + giniBrand.toFixed(4));
    outSheet.getRange(currentOutputRow, 4).setValue('학생 수: ' + students.length);
    currentOutputRow++;
    outSheet.getRange(currentOutputRow, 1).setValue('순위(하위→상위)');
    outSheet.getRange(currentOutputRow, 2).setValue('인구 누적비율');
    outSheet.getRange(currentOutputRow, 3).setValue('자산 누적비율');
    outSheet.getRange(currentOutputRow, 4).setValue('브랜드 누적비율');
    outSheet.getRange(currentOutputRow, 5).setValue('완전평등선');
    currentOutputRow++;
    outSheet.getRange(currentOutputRow, 1).setValue(0);
    outSheet.getRange(currentOutputRow, 2).setValue(0);
    outSheet.getRange(currentOutputRow, 3).setValue(0);
    outSheet.getRange(currentOutputRow, 4).setValue(0);
    outSheet.getRange(currentOutputRow, 5).setValue(0);
    currentOutputRow++;
    for (let i = 0; i < lorenzAsset.length; i++) {
      outSheet.getRange(currentOutputRow, 1).setValue(i + 1);
      outSheet.getRange(currentOutputRow, 2).setValue(lorenzAsset[i].popShare);
      outSheet.getRange(currentOutputRow, 3).setValue(lorenzAsset[i].cumShare);
      outSheet.getRange(currentOutputRow, 4).setValue(lorenzBrand[i].cumShare);
      outSheet.getRange(currentOutputRow, 5).setValue(lorenzAsset[i].popShare);
      currentOutputRow++;
    }
    currentOutputRow += 2;
  });

  SpreadsheetApp.getUi().alert('로렌츠 분석 완료!\n"로렌츠_분석" 시트를 확인해주세요.');
}

function _getSnapshot(rawData, targetDateStr) {
  const targetDate = new Date(targetDateStr);
  targetDate.setHours(23, 59, 59, 999);
  const studentMap = {};
  rawData.forEach(function(row, idx) {
    const name = row[1], brand = row[5], asset = row[6], rowDate = row[0];
    if (!name || String(name).toLowerCase().startsWith('test')) return;
    if (!rowDate || !(rowDate instanceof Date) || isNaN(rowDate.getTime())) return;
    if (rowDate > targetDate) return;
    studentMap[name] = {
      asset: (typeof asset === 'number') ? asset : 0,
      brand: (typeof brand === 'number') ? brand : 0,
      rowIdx: idx
    };
  });
  return studentMap;
}

function _calcLorenz(sortedValues) {
  const n = sortedValues.length;
  const total = sortedValues.reduce(function(sum, v){ return sum + v; }, 0);
  if (total === 0) return sortedValues.map(function(_, i){ return { popShare: (i+1)/n, cumShare: 0 }; });
  let cumSum = 0;
  return sortedValues.map(function(v, i) {
    cumSum += v;
    return { popShare: Math.round((i+1)/n*10000)/10000, cumShare: Math.round(cumSum/total*10000)/10000 };
  });
}

function _calcGini(lorenzPoints) {
  const points = [{ popShare: 0, cumShare: 0 }].concat(lorenzPoints);
  let areaUnderLorenz = 0;
  for (let i = 1; i < points.length; i++) {
    const dx = points[i].popShare - points[i-1].popShare;
    const avgY = (points[i].cumShare + points[i-1].cumShare) / 2;
    areaUnderLorenz += dx * avgY;
  }
  return Math.max(0, Math.min(1, (0.5 - areaUnderLorenz) / 0.5));
}

// ════════════════════════════════════════════════════════════════
// GAS 워밍업
// ════════════════════════════════════════════════════════════════
function keepAlive() {}

function setupWarmupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'keepAlive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('keepAlive').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert('✅ 워밍업 트리거가 설정되었습니다.\n매 1분마다 keepAlive()가 실행되어 GAS 인스턴스를 활성 상태로 유지합니다.\n수업이 없는 주말·방학에는 [워밍업 트리거 삭제]로 중단하세요.');
}

function removeWarmupTrigger() {
  var count = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'keepAlive') { ScriptApp.deleteTrigger(t); count++; }
  });
  SpreadsheetApp.getUi().alert(count > 0 ? '✅ 워밍업 트리거 ' + count + '개를 삭제했습니다.' : 'ℹ️ 삭제할 워밍업 트리거가 없습니다.');
}

// ════════════════════════════════════════════════════════════════
// Firebase 동기화
// ════════════════════════════════════════════════════════════════
const FIREBASE_URL = 'https://brand-503cd-default-rtdb.asia-southeast1.firebasedatabase.app';
const FIREBASE_API_KEY = 'AIzaSyAspOwJq6u54YBDWIx_WVOjhQHCupmriNc';

function _calcTier(honor) {
  if      (honor >= 100000) return { name: '그랜드마스터', icon: '🏆', min: 100000, max: 100000 };
  else if (honor >= 85000)  return { name: '천상의 마스터', icon: '👑', min: 85000,  max: 100000 };
  else if (honor >= 75000)  return { name: '마스터',        icon: '👑', min: 75000,  max: 85000  };
  else if (honor >= 65000)  return { name: '영원의 결정',   icon: '💠', min: 50000,  max: 75000  };
  else if (honor >= 60000)  return { name: '무결 다이아',   icon: '💠', min: 50000,  max: 65000  };
  else if (honor >= 55000)  return { name: '세공된 다이아', icon: '💠', min: 50000,  max: 60000  };
  else if (honor >= 50000)  return { name: '다이아 원석',   icon: '💠', min: 50000,  max: 55000  };
  else if (honor >= 45000)  return { name: '홍염의 정점',   icon: '💎', min: 30000,  max: 50000  };
  else if (honor >= 40000)  return { name: '각성한 루비',   icon: '💎', min: 30000,  max: 45000  };
  else if (honor >= 35000)  return { name: '연마된 루비',   icon: '💎', min: 30000,  max: 40000  };
  else if (honor >= 30000)  return { name: '루비 원석',     icon: '💎', min: 30000,  max: 35000  };
  else if (honor >= 27500)  return { name: '태양의 황금',   icon: '🥇', min: 20000,  max: 30000  };
  else if (honor >= 25000)  return { name: '정련된 골드',   icon: '🥇', min: 20000,  max: 27500  };
  else if (honor >= 22500)  return { name: '제련된 골드',   icon: '🥇', min: 20000,  max: 25000  };
  else if (honor >= 20000)  return { name: '금 광석',       icon: '🥇', min: 20000,  max: 22500  };
  else if (honor >= 17500)  return { name: '은빛 극점',     icon: '🥈', min: 17500,  max: 20000  };
  else if (honor >= 15000)  return { name: '진화한 실버',   icon: '🥈', min: 10000,  max: 17500  };
  else if (honor >= 12500)  return { name: '성장한 실버',   icon: '🥈', min: 10000,  max: 15000  };
  else if (honor >= 10000)  return { name: '거친 실버',     icon: '🥈', min: 10000,  max: 12500  };
  else if (honor >= 7500)   return { name: '빛나는 브론즈', icon: '🥉', min: 7500,   max: 10000  };
  else if (honor >= 5000)   return { name: '브론즈',        icon: '🥉', min: 5000,   max: 7500   };
  else                      return { name: '새싹',          icon: '🌱', min: 0,      max: 5000   };
}

function _firebasePut(path, data) {
  const url = FIREBASE_URL + '/' + encodeURI(path) + '.json?key=' + FIREBASE_API_KEY;
  const options = {
    method: 'put', contentType: 'application/json',
    payload: JSON.stringify(data), muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() !== 200) {
    throw new Error('Firebase PUT 실패 [' + path + '] ' + res.getContentText());
  }
}

function syncAllStudentsToFirebase() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();

  const jobSheet = ss.getSheetByName(SHEET_JOB);
  const jobMap   = {};
  if (jobSheet) {
    const jobData = jobSheet.getDataRange().getValues();
    for (let j = 1; j < jobData.length; j++) {
      const jName = String(jobData[j][0]).trim();
      if (jName) jobMap[jName] = { title: jobData[j][1] || '미배정', salary: Number(jobData[j][2]) || 0, area: jobData[j][3] || '-' };
    }
  }

  const auctionSheet  = ss.getSheetByName(SHEET_AUCTION);
  const auctionPrices = [];
  if (auctionSheet) {
    const aData = auctionSheet.getDataRange().getValues();
    for (let m = 1; m < aData.length; m++) {
      if (!aData[m][0]) continue;
      auctionPrices.push({ item: '[' + aData[m][0] + '] ' + (aData[m][1] || ''), price: Number(aData[m][11]) || 0 });
    }
  }

  let classTotalTax = 0;
  for (let i = 1; i < mainData.length; i++) classTotalTax += Number(mainData[i][COL_TAX - 1]) || 0;

  const emergency  = getEmergencyStatus();
  const updatedAt  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const errors = [];

  for (let i = 1; i < mainData.length; i++) {
    const row  = mainData[i];
    const name = String(row[COL_NAME - 1]).trim();
    if (!name) continue;
    const honor = Number(row[COL_VALUE - 1]) || 0;
    const snapshot = {
      name, brand: row[COL_BRAND - 1] || '', honor,
      balance: Number(row[COL_ASSET - 1]) || 0,
      honorRank: Number(row[COL_RANK_V - 1]) || 0,
      balanceRank: Number(row[COL_RANK_A - 1]) || 0,
      personalTax: Number(row[COL_TAX - 1]) || 0,
      classTotalTax,
      job: jobMap[name] || { title: '미배정', salary: 0, area: '-' },
      tierData: _calcTier(honor),
      auctionPrices, emergency, updatedAt
    };
    try { _firebasePut('students/' + name + '/snapshot', snapshot); }
    catch(e) { errors.push(name + ': ' + e.message); Logger.log('[Firebase 동기화 오류] ' + name + ' - ' + e.message); }
  }

  if (errors.length === 0) Logger.log('[Firebase] 전체 학생 스냅샷 동기화 완료 (' + (mainData.length - 1) + '명)');
  else Logger.log('[Firebase] 동기화 완료 (오류 ' + errors.length + '건): ' + errors.join(', '));
}

function syncOneStudentToFirebase(studentName) {
  if (!studentName) return;
  studentName = String(studentName).trim();
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet.getDataRange().getValues();

  let studentRow = null;
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === studentName) { studentRow = mainData[i]; break; }
  }
  if (!studentRow) return;

  const jobSheet = ss.getSheetByName(SHEET_JOB);
  let jobResult  = { title: '미배정', salary: 0, area: '-' };
  if (jobSheet) {
    const jobData = jobSheet.getDataRange().getValues();
    for (let j = 1; j < jobData.length; j++) {
      if (String(jobData[j][0]).trim() === studentName) {
        jobResult = { title: jobData[j][1] || '미배정', salary: Number(jobData[j][2]) || 0, area: jobData[j][3] || '-' }; break;
      }
    }
  }

  const auctionSheet  = ss.getSheetByName(SHEET_AUCTION);
  const auctionPrices = [];
  if (auctionSheet) {
    const aData = auctionSheet.getDataRange().getValues();
    for (let m = 1; m < aData.length; m++) {
      if (!aData[m][0]) continue;
      auctionPrices.push({ item: '[' + aData[m][0] + '] ' + (aData[m][1] || ''), price: Number(aData[m][11]) || 0 });
    }
  }

  let classTotalTax = 0;
  for (let i = 1; i < mainData.length; i++) classTotalTax += Number(mainData[i][COL_TAX - 1]) || 0;

  const honor    = Number(studentRow[COL_VALUE - 1]) || 0;
  const snapshot = {
    name: studentName, brand: studentRow[COL_BRAND - 1] || '', honor,
    balance: Number(studentRow[COL_ASSET - 1]) || 0,
    honorRank: Number(studentRow[COL_RANK_V - 1]) || 0,
    balanceRank: Number(studentRow[COL_RANK_A - 1]) || 0,
    personalTax: Number(studentRow[COL_TAX - 1]) || 0,
    classTotalTax, job: jobResult, tierData: _calcTier(honor),
    auctionPrices, emergency: getEmergencyStatus(),
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };

  try {
    _firebasePut('students/' + studentName + '/snapshot', snapshot);
    Logger.log('[Firebase] ' + studentName + ' 스냅샷 동기화 완료');
  } catch(e) {
    Logger.log('[Firebase 동기화 오류] ' + studentName + ' - ' + e.message);
  }
}