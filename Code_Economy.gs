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

  // ── 지니계수 계산 함수 (업로드된 공식 그대로) ──────────────
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
        // 이 점에 해당하는 학생 수 (툴팁용)
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
    // 지니계수
    giniAsset:  Math.round(giniAsset  * 1000) / 1000,
    giniValue:  Math.round(giniValue  * 1000) / 1000,
    // 로렌츠 곡선 포인트
    lorenzAsset,
    lorenzValue,
    // 분위별 점유율
    shareBot20: Math.round(shareBot20 * 1000) / 1000,
    shareMid60: Math.round(shareMid60 * 1000) / 1000,
    shareTop20: Math.round(shareTop20 * 1000) / 1000,
    totalAsset,
    // 히스토리 (차트용)
    history,
    // 학생 자산 순위
    ranked
  };
}

// ── 지니계수 로그 기록 헬퍼 ───────────────────────────────────
function _recordGiniLog(ss, giniAsset, giniValue, totalAsset) {
  try {
    let logSheet = ss.getSheetByName('지니계수_로그');
    // 시트가 없으면 자동 생성
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
    // 없으면 새 행 추가
    logSheet.appendRow([today, giniAsset, giniValue, totalAsset, timeStr]);
  } catch(e) {
    // 로그 실패해도 메인 기능은 계속 동작
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
    // 최근 12주치만 반환 (헤더 제외, 역순)
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

