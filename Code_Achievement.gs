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

  // 1) 학생별 캐시 삭제 — 메인 시트에서 이름 목록을 읽어 개별 삭제
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(SHEET_MAIN);
    if (mainSheet) {
      const mainData = mainSheet.getDataRange().getValues();
      const keys = [];
      for (let i = 1; i < mainData.length; i++) {
        const name = String(mainData[i][COL_NAME - 1]).trim();
        if (name) keys.push('student_' + name);
      }
      if (keys.length > 0) cache.removeAll(keys);
    }
  } catch(e) {}

  // 2) 업적 캐시 삭제
  cache.remove('achievement_master');

  SpreadsheetApp.getUi().alert('✅ 모든 캐시가 초기화되었습니다. (학생 캐시 포함)');
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
        // 히든 업적 최초 달성 전역 알림
        const notifySheet = ss.getSheetByName(SHEET_GLOBAL_NOTIFY);
        if (notifySheet) {
          let alreadyUnlocked = false;
          const achDataNow = achSheet.getDataRange().getValues();
          for (let i = 1; i < achDataNow.length; i++) {
            if (String(achDataNow[i][1]).trim() === 'HID-004' && String(achDataNow[i][0]).trim() !== studentName) {
              alreadyUnlocked = true; break;
            }
          }
          if (!alreadyUnlocked) {
            const noticeId = 'HIDDEN_HID-004_' + new Date().getTime();
            const msg = `🎉 히든 업적 [${String(masterData[m][1])}]을(를) 달성한 사람이 최초로 등장했습니다! 지금부터 이 업적의 정체와 달성 조건이 모두에게 공개됩니다.`;
            const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
            notifySheet.appendRow([noticeId, msg, ts, 'ALERT']);
            masterSheet.getRange(m + 1, 4).setValue('FALSE'); // 히든 해제
          }
        }
        break;
      }
    }
  }

  // ── RANK-001~007: 랭크 브레이커 ────────────────────────────
  const rankBreakers = {
    'RANK-001': ['거친 실버'],
    'RANK-002': ['금 광석'],
    'RANK-003': ['루비 원석'],
    'RANK-004': ['다이아 원석'],
    'RANK-005': ['마스터'],
    'RANK-006': ['천상의 마스터'],
    'RANK-007': ['그랜드마스터']
  };
  // 전역 알림 대상 티어 (최초 진입 시 전체 공지)
  const TIER_ALERT_TARGETS = ['금 광석', '루비 원석', '다이아 원석', '마스터', '천상의 마스터', '그랜드마스터'];
  // 현재 학생 티어명 계산 (honor 기반)
  const h = Number(honor) || 0;
  let currentTierName = '새싹';
  if      (h >= 100000) currentTierName = '그랜드마스터';
  else if (h >= 85000)  currentTierName = '천상의 마스터';
  else if (h >= 75000)  currentTierName = '마스터';
  else if (h >= 65000)  currentTierName = '영원의 결정';
  else if (h >= 60000)  currentTierName = '무결 다이아';
  else if (h >= 55000)  currentTierName = '세공된 다이아';
  else if (h >= 50000)  currentTierName = '다이아 원석';
  else if (h >= 45000)  currentTierName = '홍염의 정점';
  else if (h >= 40000)  currentTierName = '각성한 루비';
  else if (h >= 35000)  currentTierName = '연마된 루비';
  else if (h >= 30000)  currentTierName = '루비 원석';
  else if (h >= 27500)  currentTierName = '태양의 황금';
  else if (h >= 25000)  currentTierName = '정련된 골드';
  else if (h >= 22500)  currentTierName = '제련된 골드';
  else if (h >= 20000)  currentTierName = '금 광석';
  else if (h >= 17500)  currentTierName = '은빛 극점';
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
    let grantedAchName = '';
    for (let m = 1; m < masterData.length; m++) {
      if (String(masterData[m][0]).trim() === rankId) {
        grantedAchName = String(masterData[m][1]).trim();
        achSheet.appendRow([studentName, rankId, grantedAchName, String(masterData[m][2]), today, false]);
        break;
      }
    }

    // ★ 전역 알림 대상 티어 최초 진입 시 알림 발송
    const tierName = rankBreakers[rankId][0]; // 해당 RANK의 티어명
    if (TIER_ALERT_TARGETS.indexOf(tierName) !== -1) {
      _postTierFirstAlert(studentName, tierName);
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
      // 자동 부여 업적도 유일/초월 등급이면 전역 알림
      const achGrade = String(masterData[m][5] || '희귀').trim();
      _checkAndPostGlobalAlert(studentName, achName, achGrade);
    }
  }
}

// ════════════════════════════════════════════════════════════════
// 14. 업적 신청-승인 시스템 (v2)
// ════════════════════════════════════════════════════════════════


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
        'RANK-001','RANK-002','RANK-003','RANK-004','RANK-005','RANK-006', 'RANK-007']);
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
          notifySheet.appendRow([noticeId, msg, ts, 'ALERT']);
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
// ██ 반려 정정함 시스템
// ════════════════════════════════════════════════════════════════

// ── 반려된 업적 신청 목록 반환 ───────────────────────────────────
function getRejectedAchievements() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet    = ss.getSheetByName(SHEET_ACH_LOG);
  const masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!logSheet) return [];

  const achNameMap = {};
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let m = 1; m < mData.length; m++) {
      if (!mData[m][0]) continue;
      achNameMap[String(mData[m][0]).trim()] = String(mData[m][1]).trim();
    }
  }

  const logData = logSheet.getDataRange().getValues();
  const result  = [];
  for (let i = 1; i < logData.length; i++) {
    if (String(logData[i][4]).trim() !== '반려') continue;
    const achId = String(logData[i][2]).trim();
    result.push({
      rowNumber:   i + 1,
      timestamp:   String(logData[i][0]),
      studentName: String(logData[i][1]).trim(),
      achId:       achId,
      achName:     achNameMap[achId] || '(알 수 없음)',
      proof:       String(logData[i][3]).trim()
    });
  }
  return result.reverse();
}

// ── 반려 → 대기로 되돌리기 ──────────────────────────────────────
function correctRejection(rowNumber) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  if (!logSheet) return { success: false, msg: '업적신청로그 시트를 찾을 수 없습니다.' };
  if (rowNumber < 2) return { success: false, msg: '유효하지 않은 행 번호입니다.' };

  const row    = logSheet.getRange(rowNumber, 1, 1, 5).getValues()[0];
  const status = String(row[4]).trim();
  if (status !== '반려') return { success: false, msg: '반려 상태인 신청만 정정할 수 있습니다.' };

  logSheet.getRange(rowNumber, 5).setValue('대기');
  const studentName = String(row[1]).trim();
  const achId       = String(row[2]).trim();
  return { success: true, msg: `[${studentName}] ${achId} 신청이 대기 상태로 복원되었습니다.` };
}
