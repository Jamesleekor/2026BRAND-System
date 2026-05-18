// ════════════════════════════════════════════════════════════════
// 15. 로그인 화면용 - 전체 학생 업적 명예의 전당
// ════════════════════════════════════════════════════════════════

function getAllStudentsHonorBoard() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet   = ss.getSheetByName(SHEET_MAIN);
  const achSheet    = ss.getSheetByName(SHEET_ACH_STUDENT);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);

  if (!mainSheet || !achSheet || !masterSheet) return [];

  const mainData   = mainSheet.getDataRange().getValues();
  const achData    = achSheet.getDataRange().getValues();
  const masterData = masterSheet.getDataRange().getValues();

  // ── 캐릭터 정보 추출: 상점_아이템 + 상점_구매로그 조인 ──────────
  // [왜 두 시트를 조인하는가]
  //   purchaseShopItem()이 상점_구매로그에 기록하는 실제 컬럼 구조:
  //     A(0):구매ID, B(1):학생명, C(2):아이템ID, D(3):아이템명,
  //     E(4):가격,   F(5):구매일시, G(6):장착여부(TRUE/FALSE)
  //   ※ 카테고리(캐릭터/스킨/폰트)와 리소스값은 상점_구매로그에 없습니다.
  //     상점_아이템 시트에서 아이템ID로 조회해야 합니다.
  //
  // 상점_아이템 컬럼 구조 (SHEET_SHOP_ITEMS):
  //     A(0):아이템ID, B(1):카테고리, C(2):아이템명, D(3):가격,
  //     E(4):조건설명, F(5):조건타입,  G(6):조건값,  H(7):리소스값
  const charMap = {}; // { 학생명: resourceVal } — 장착 중인 캐릭터만
  try {
    const shopLog  = ss.getSheetByName(SHEET_SHOP_LOG);
    const shopItem = ss.getSheetByName(SHEET_SHOP_ITEMS);
    if (shopLog && shopItem && shopLog.getLastRow() >= 2) {
      // 1단계: 상점_아이템에서 아이템ID → {카테고리, 리소스값} 맵 생성
      const itemData = shopItem.getDataRange().getValues();
      const itemMap  = {};
      for (let i = 1; i < itemData.length; i++) {
        const iId  = String(itemData[i][0]).trim(); // A열: 아이템ID
        const iCat = String(itemData[i][1]).trim(); // B열: 카테고리
        const iRes = String(itemData[i][7]).trim(); // H열: 리소스값
        if (iId) itemMap[iId] = { category: iCat, resVal: iRes };
      }
      // 2단계: 상점_구매로그에서 G열=TRUE(장착)인 캐릭터 아이템만 추출
      const logData = shopLog.getRange(2, 1, shopLog.getLastRow() - 1, 7).getValues();
      for (let i = 0; i < logData.length; i++) {
        const sName   = String(logData[i][1]).trim(); // B열: 학생명
        const itemId  = String(logData[i][2]).trim(); // C열: 아이템ID
        const equipped = logData[i][6] === true || String(logData[i][6]).toUpperCase() === 'TRUE'; // G열
        const info    = itemMap[itemId];
        if (equipped && info && info.category === '캐릭터' && info.resVal && info.resVal !== 'default') {
          charMap[sName] = info.resVal; // 가장 마지막으로 장착된 캐릭터가 덮어씀
        }
      }
    }
  } catch(e) {
    // 캐릭터 로드 실패 시 기본 이니셜 아바타로 대체 (기능 중단 없음)
    Logger.log('캐릭터 로드 실패(무시): ' + e.message);
  }

  // ── 업적마스터에서 업적ID별 등급/이모지 맵 생성 ─────────────────
  const gradeMap = {};
  const emojiMap = {};
  for (let m = 1; m < masterData.length; m++) {
    const achId = String(masterData[m][0]).trim();
    const grade = String(masterData[m][5] || '희귀').trim(); // F열: 업적등급
    gradeMap[achId] = grade;
    // 유니크 이상 등급에만 이모지 적용 (희귀/언커먼은 제외)
    if (grade === '유니크' || grade === '에픽' || grade === '히든' || grade === '유일' || grade === '초월') {
      emojiMap[achId] = getEmojiForAchievement(achId);
    }
  }

  const result = [];

  // ── 학생별 순회: 업적 + 칭호 + 캐릭터 정보 수집 ────────────────
  for (let i = 1; i < mainData.length; i++) {
    const studentName = String(mainData[i][COL_NAME - 1]).trim();
    if (!studentName) continue;

    const achievements = [];
    let equippedTitle  = null;

    for (let j = 1; j < achData.length; j++) {
      if (String(achData[j][0]).trim() !== studentName) continue;
      const achId    = String(achData[j][1]).trim();
      const achName  = String(achData[j][2]).trim();
      const equipped = achData[j][5] === true || String(achData[j][5]).toUpperCase() === 'TRUE';
      const grade    = gradeMap[achId] || '희귀';
      const emoji    = emojiMap[achId] || '';
      achievements.push({ achId, achName, grade, emoji });
      if (equipped) equippedTitle = (emoji ? emoji + ' ' : '') + achName;
    }

    // ── 캐릭터 resourceVal 분류: 이미지 URL vs 이모지 문자 ──────
    const charVal  = charMap[studentName] || null;
    let charImgUrl = null;
    let charEmoji  = null;
    if (charVal) {
      if (charVal.startsWith('http://') || charVal.startsWith('https://')) {
        charImgUrl = charVal; // Google Drive 이미지 URL 등
      } else {
        charEmoji = charVal;  // 이모지 문자 (예: 🐶, 🦊)
      }
    }

    result.push({
      name:             studentName,
      equippedTitle:    equippedTitle,
      achievementCount: achievements.length,
      achievements:     achievements,
      charImgUrl:       charImgUrl, // Index.html getAvatarHtml()에서 사용
      charEmoji:        charEmoji   // Index.html getAvatarHtml()에서 사용
    });
  }

  // 업적 많은 순으로 정렬 (동점 시 가나다 순)
  result.sort(function(a, b) {
    if (b.achievementCount !== a.achievementCount) return b.achievementCount - a.achievementCount;
    return a.name.localeCompare(b.name, 'ko');
  });

  return result;
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
    const hint  = String(mData[m][4] || '').trim(); // E열: 달성 힌트 (성좌 맵 툴팁용)
    const emoji = getEmojiForAchievement(id);
    achList.push({ achId: id, achName: isHid ? '🔒 ???' : name, grade, isHidden: isHid, count: 0, emoji: isHid ? '' : emoji, hint, firstAchiever: null });
    achMap[id]  = achList.length - 1;
  }

  // 달성 학생 집계 + 최초 달성자 기록
  // [왜 이 방식으로 최초 달성자를 찾는가]
  //   학생업적달성 시트는 업적 달성 시 appendRow로 순서대로 행이 추가된다.
  //   따라서 해당 achId가 시트에서 가장 먼저 등장하는 행의 A열(학생명)이 곧 최초 달성자다.
  //   firstAchiever가 이미 설정됐으면 덮어쓰지 않고 건너뛴다.
  const sData = achSheet.getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) {
    const id = String(sData[i][1]).trim(); // B열: 업적ID
    if (achMap[id] !== undefined) {
      achList[achMap[id]].count++;
      if (!achList[achMap[id]].firstAchiever) {
        achList[achMap[id]].firstAchiever = String(sData[i][0]).trim(); // A열: 학생명
      }
    }
  }

  // count 내림차순 정렬
  achList.sort(function(a, b) { return b.count - a.count; });
  return achList;
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

// ══════════════════════════════════════════════════════════
// 불평등 지수: 지니계수 + 로렌츠 곡선 데이터 반환
// ══════════════════════════════════════════════════════════
function getInequalityData() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return { success: false, msg: '메인 시트 없음' };

  const mainData = mainSheet.getDataRange().getValues();

  // 학생별 자산보유량 + 브랜드가치 수집 (1행은 헤더라 스킵)
  const students = [];
  for (let i = 1; i < mainData.length; i++) {
    const name  = String(mainData[i][COL_NAME  - 1]).trim();
    const asset = Number(mainData[i][COL_ASSET - 1]) || 0;
    const value = Number(mainData[i][COL_VALUE - 1]) || 0;
    if (!name) continue;
    students.push({ name, asset, value });
  }

  if (students.length === 0) return { success: false, msg: '학생 데이터 없음' };
  // ── 예금 원금 합산 ──────────────────────────────────────────
  const depositSheet = ss.getSheetByName('학생별가입예금');
  const depositMap   = {};
  if (depositSheet) {
    const depData = depositSheet.getDataRange().getValues();
    for (let i = 1; i < depData.length; i++) {
      const dName     = String(depData[i][1]).trim(); // B열
      const principal = Number(depData[i][2]) || 0;  // C열
      if (!dName) continue;
      if (String(depData[i][7]).trim() !== '진행중') continue;
      depositMap[dName] = (depositMap[dName] || 0) + principal;
    }
  }

  // ── 대출 잔액 합산 ──────────────────────────────────────────
  const loanSheet = ss.getSheetByName('대출현황');
  const loanMap   = {};
  if (loanSheet) {
    const loanData = loanSheet.getDataRange().getValues();
    for (let i = 1; i < loanData.length; i++) {
      const lName   = String(loanData[i][1]).trim();  // B열
      const balance = Number(loanData[i][10]) || 0;  // K열
      if (!lName) continue;
      loanMap[lName] = (loanMap[lName] || 0) + balance;
    }
  }

  // ── 지니계수 계산 함수 ──────────────────────────────────────
  // G = (Σi Σj |xi - xj|) / (2 * n^2 * x̄)
  function calcGini(values) {
    const n    = values.length;
    const mean = values.reduce(function(s, v) { return s + v; }, 0) / n;
    if (mean === 0) return 0;
    let sumDiff = 0;
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < n; j++) {
        sumDiff += Math.abs(values[i] - values[j]);
      }
    }
    return sumDiff / (2 * n * n * mean);
  }

  // ── 로렌츠 곡선 데이터 계산 함수 ──────────────────────────
  // 오름차순 정렬 후 누적 비율 계산
  // 반환값: [{x: 인구누적비율(0~1), y: 자산누적비율(0~1)}] 배열
  function calcLorenz(values) {
    const sorted = values.slice().sort(function(a, b) { return a - b; });
    const total  = sorted.reduce(function(s, v) { return s + v; }, 0);
    const n      = sorted.length;
    const points = [{ x: 0, y: 0 }]; // 시작점
    let cumSum = 0;
    for (let i = 0; i < n; i++) {
      cumSum += sorted[i];
      points.push({
        x: (i + 1) / n,
        y: total > 0 ? cumSum / total : 0,
        studentCount: i + 1,
        cumAsset: cumSum
      });
    }
    return points;
  }

  // ── 자산보유량 기준 계산 ───────────────────────────────────
  const assetValues  = students.map(function(s) { return s.asset; });
  const giniAsset    = calcGini(assetValues);
  const lorenzAsset  = calcLorenz(assetValues);

  // ── 브랜드가치 기준 계산 (보조 지표) ─────────────────────
  const valueValues  = students.map(function(s) { return s.value; });
  const giniValue    = calcGini(valueValues);
  const lorenzValue  = calcLorenz(valueValues);

  // ── 분위별 자산 점유율 (하위 20%, 중위 60%, 상위 20%) ──────
  const sortedAsset  = assetValues.slice().sort(function(a, b) { return a - b; });
  const n            = sortedAsset.length;
  const totalAsset   = sortedAsset.reduce(function(s, v) { return s + v; }, 0);
  const bot20idx     = Math.floor(n * 0.2);
  const top20idx     = Math.floor(n * 0.8);
  const bot20sum     = sortedAsset.slice(0, bot20idx).reduce(function(s, v) { return s + v; }, 0);
  const mid60sum     = sortedAsset.slice(bot20idx, top20idx).reduce(function(s, v) { return s + v; }, 0);
  const top20sum     = sortedAsset.slice(top20idx).reduce(function(s, v) { return s + v; }, 0);

  const shareBot20   = totalAsset > 0 ? bot20sum / totalAsset : 0;
  const shareMid60   = totalAsset > 0 ? mid60sum / totalAsset : 0;
  const shareTop20   = totalAsset > 0 ? top20sum / totalAsset : 0;

  // ── 오늘 날짜 기록 (지니계수_로그 시트에 저장) ─────────────
  _recordGiniLog(ss, giniAsset, giniValue, totalAsset);

  // ── 지니계수_로그 시트에서 히스토리 읽기 ──────────────────
  const history = _readGiniHistory(ss);

  // ── 학생 순위 (자산보유량 내림차순, 프론트 표시용) ─────────
  const ranked = students.slice()
    .map(function(s) {
      const realAsset = s.asset + (depositMap[s.name] || 0) - (loanMap[s.name] || 0);
      return Object.assign({}, s, { realAsset: realAsset });
    })
    .sort(function(a, b) { return b.realAsset - a.realAsset; });

  return {
    success:    true,
    studentCount: students.length,
    giniAsset:  Math.round(giniAsset  * 1000) / 1000,
    giniValue:  Math.round(giniValue  * 1000) / 1000,
    lorenzAsset,
    lorenzValue,
    shareBot20: Math.round(shareBot20 * 1000) / 1000,
    shareMid60: Math.round(shareMid60 * 1000) / 1000,
    shareTop20: Math.round(shareTop20 * 1000) / 1000,
    totalAsset,
    history,
    ranked
  };
}

// ── 지니계수 로그 기록 헬퍼 ───────────────────────────────────
function _recordGiniLog(ss, giniAsset, giniValue, totalAsset) {
  try {
    let logSheet = ss.getSheetByName('지니계수_로그');
    if (!logSheet) {
      logSheet = ss.insertSheet('지니계수_로그');
      logSheet.appendRow(['날짜', '자산_지니계수', '브랜드_지니계수', '총자산합계', '기록시각']);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const timeStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
    const existing = logSheet.getDataRange().getValues();

    // 오늘 날짜 행이 이미 있으면 덮어쓰기 (중복 방지)
    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][0]) === today) {
        logSheet.getRange(i + 1, 1, 1, 5).setValues([[today, giniAsset, giniValue, totalAsset, timeStr]]);
        return;
      }
    }
    logSheet.appendRow([today, giniAsset, giniValue, totalAsset, timeStr]);
  } catch(e) {
    Logger.log('지니계수 로그 기록 실패: ' + e.message);
  }
}

// ── 지니계수 히스토리 읽기 헬퍼 ──────────────────────────────
function _readGiniHistory(ss) {
  try {
    const logSheet = ss.getSheetByName('지니계수_로그');
    if (!logSheet) return [];
    const data = logSheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        date:       String(data[i][0]),
        giniAsset:  Number(data[i][1]) || 0,
        giniValue:  Number(data[i][2]) || 0,
        totalAsset: Number(data[i][3]) || 0
      });
    }
    // 최신순 정렬 후 최근 12개만
    return result.reverse().slice(0, 12).reverse();
  } catch(e) {
    return [];
  }
}