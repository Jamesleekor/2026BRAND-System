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

  // ── 상점_구매로그에서 학생별 장착 캐릭터 추출 ────────────────
  // 상점_구매로그 열 구조 (Code_Shop.gs 기준):
  //   A(0):날짜, B(1):학생이름, C(2):아이템ID, D(3):아이템명,
  //   E(4):카테고리, F(5):resourceVal, G(6):장착여부(TRUE/FALSE)
  const charMap = {}; // studentName → resourceVal
  try {
    const shopLog = ss.getSheetByName(SHEET_SHOP_LOG);
    if (shopLog && shopLog.getLastRow() >= 2) {
      const logData = shopLog.getRange(2, 1, shopLog.getLastRow() - 1, 7).getValues();
      for (let i = 0; i < logData.length; i++) {
        const row      = logData[i];
        const sName    = String(row[1]).trim();
        const category = String(row[4]).trim();
        const resVal   = String(row[5]).trim();
        const equipped = row[6] === true || String(row[6]).toUpperCase() === 'TRUE';
        if (category === '캐릭터' && equipped && resVal && resVal !== 'default') {
          charMap[sName] = resVal; // 장착 중인 캐릭터만
        }
      }
    }
  } catch(e) {
    Logger.log('캐릭터 로드 실패(무시): ' + e.message);
  }

  // ── 업적마스터에서 등급/이모지 맵 생성 ───────────────────────
  const gradeMap = {};
  const emojiMap = {};
  for (let m = 1; m < masterData.length; m++) {
    const achId = String(masterData[m][0]).trim();
    const grade = String(masterData[m][5] || '희귀').trim();
    gradeMap[achId] = grade;
    if (grade === '유니크' || grade === '에픽' || grade === '히든' || grade === '유일' || grade === '초월') {
      emojiMap[achId] = getEmojiForAchievement(achId);
    }
  }

  const result = [];

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
      if (equipped) {
        equippedTitle = (emoji ? emoji + ' ' : '') + achName;
      }
    }

    // 캐릭터 resourceVal 분류
    const charVal = charMap[studentName] || null;
    let charImgUrl  = null;
    let charEmoji   = null;
    if (charVal) {
      if (charVal.startsWith('http://') || charVal.startsWith('https://')) {
        charImgUrl = charVal;
      } else {
        charEmoji = charVal;
      }
    }

    result.push({
      name:             studentName,
      equippedTitle:    equippedTitle,
      achievementCount: achievements.length,
      achievements:     achievements,
      charImgUrl:       charImgUrl,   // ← 캐릭터 이미지 URL (없으면 null)
      charEmoji:        charEmoji     // ← 캐릭터 이모지 (없으면 null)
    });
  }

  result.sort(function(a, b) { return b.achievementCount - a.achievementCount; });
  return result;
}