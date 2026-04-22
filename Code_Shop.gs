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
  let balance        = 0;
  let studentTierName  = '새싹';  // 티어 조건 판별용
  let studentTaxPaid   = 0;       // 누적 납세+기부 조건 판별용
  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME-1]).trim() === studentName) {
      balance        = Number(mainData[i][COL_ASSET-1]) || 0;
      studentTaxPaid = Number(mainData[i][COL_TAX-1])   || 0;  // H열: 누적 납세
      // 브랜드가치로 티어명 계산
      const honor = Number(mainData[i][COL_VALUE-1]) || 0;
      if      (honor >= 100000) studentTierName = '그랜드마스터';
      else if (honor >= 85000)  studentTierName = '천상의 마스터';
      else if (honor >= 75000)  studentTierName = '마스터';
      else if (honor >= 65000)  studentTierName = '영원의 결정';
      else if (honor >= 60000)  studentTierName = '무결 다이아';
      else if (honor >= 55000)  studentTierName = '세공된 다이아';
      else if (honor >= 50000)  studentTierName = '다이아 원석';
      else if (honor >= 45000)  studentTierName = '홍염의 정점';
      else if (honor >= 40000)  studentTierName = '각성한 루비';
      else if (honor >= 35000)  studentTierName = '연마된 루비';
      else if (honor >= 30000)  studentTierName = '루비 원석';
      else if (honor >= 25000)  studentTierName = '태양의 황금';
      else if (honor >= 21000)  studentTierName = '정련된 골드';
      else if (honor >= 17000)  studentTierName = '제련된 골드';
      else if (honor >= 13000)  studentTierName = '금 광석';
      else if (honor >= 10000)  studentTierName = '은빛 극점';
      else if (honor >= 7500)   studentTierName = '진화한 실버';
      else if (honor >= 5500)   studentTierName = '성장한 실버';
      else if (honor >= 3500)   studentTierName = '거친 실버';
      else if (honor >= 2000)   studentTierName = '빛나는 브론즈';
      else if (honor >= 800)    studentTierName = '브론즈';
      else                      studentTierName = '새싹';
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

    // 구매 조건 관련
    // 구매 조건 충족 여부
    // ── J열: 조건타입2, K열: 조건값2, L열: 한정판_종료일 ──────────
    const condType2   = iData[i][9]  ? String(iData[i][9]).trim()  : '';
    const condVal2    = iData[i][10] ? String(iData[i][10]).trim() : '';
    const limitedDate = iData[i][11]
      ? ((iData[i][11] instanceof Date)
          ? Utilities.formatDate(iData[i][11], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(iData[i][11]).trim().substring(0, 10))
      : '';

    // 한정판 만료 체크 (시트 날짜는 Date 객체로 읽힘 → yyyy-MM-dd 문자열로 변환 후 비교)
    let isExpired = false;
    if (iData[i][11]) {
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const limitedDateStr = (iData[i][11] instanceof Date)
        ? Utilities.formatDate(iData[i][11], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(iData[i][11]).trim().substring(0, 10);
      isExpired = today > limitedDateStr;
    }

    // 조건 판별 함수 (condType + condVal 한 쌍을 받아서 true/false 반환)
    function checkOneCond(cType, cVal) {
      if (!cType || cType === 'none') return true;
      if (cType === 'ach_count') {
        return totalAch >= Number(cVal);
      }
      if (cType.startsWith('ach_grade:')) {
        const targetGrade = cType.split(':')[1];
        return (gradeCount[targetGrade] || 0) >= Number(cVal);
      }
      if (cType === 'tier') {
        // condVal은 티어 번호(1~22). 학생 현재 티어 번호와 비교
        const studentTierNum = TIER_ORDER.indexOf(studentTierName) + 1; // 못 찾으면 0
        return studentTierNum >= Number(cVal);
      }
      if (cType === 'asset') {
        return balance >= Number(cVal);
      }
      if (cType === 'tax_paid') {
        return studentTaxPaid >= Number(cVal);
      }
      return true; // 알 수 없는 타입은 통과
    }

    // 조건1 OR 조건2 (조건2가 없으면 조건1만 체크)
    const cond1Met = checkOneCond(condType, condVal);
    const cond2Met = condType2 ? checkOneCond(condType2, condVal2) : false;
    const condMet  = condType2 ? (cond1Met && cond2Met) : cond1Met;

    const canBuy = !isOwned && condMet && balance >= price && !isExpired;

    items.push({
      itemId, category, itemName, price,
      condDesc, condType, condVal, condType2, condVal2,
      limitedDate, isExpired,
      resourceVal, isOwned, condMet, canBuy
    });
  }  // ← for 루프 닫기

  return { items, owned, balance };
}  // ← getShopItems 함수 닫기

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

  // 한정판 만료 체크 (시트 날짜는 Date 객체로 읽힘 → yyyy-MM-dd 문자열로 변환 후 비교)
  if (itemRow[11]) {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const limitedDateStr = (itemRow[11] instanceof Date)
      ? Utilities.formatDate(itemRow[11], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(itemRow[11]).trim().substring(0, 10);
    if (today > limitedDateStr) {
      return { success: false, msg: `[${itemName}]은 한정판 기간이 종료된 아이템입니다.` };
    }
  }

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
