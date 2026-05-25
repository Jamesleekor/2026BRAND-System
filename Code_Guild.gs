// ============================================================
// Code_Guild.gs — B.R.A.N.D 길드 시스템 v2
// 작성일: 2026-04-28
// 담당: Master Lee
//
// ※ 주의사항
//   - 기존 Code.gs 함수를 절대 수정·삭제하지 않음
//   - 이 파일은 Apps Script 프로젝트에 별도 파일로 추가
//   - 시트명은 아래 SHEET_NAMES 상수를 통해서만 참조
// ============================================================


// ============================================================
// 섹션 1 — 상수 및 설정
// ============================================================

// --- GS 공식 가중치 (나중에 조정 시 여기만 수정) ---
// 월간 GS = α(0.50) + β(0.25) + 미션(0.25) → 월 최대 1.00
// 시즌 GS = 5월 + 6월 + 7월 (최대 3.00) + 학기프로젝트 (최대 1.00) → 시즌 최대 4.00
var GS_ALPHA        = 0.50;  // 인당 자산 증가량 가중치
var GS_MISSION_PARTICIPATION = 0.15;  // 길드 미션 참여 횟수 가중치
var GS_SESSION_ATTENDANCE    = 0.10;  // 세션 참석 횟수 가중치
var GS_MISSION_RATE = 0.25;  // 미션항 가중치
var GS_PROJECT_RATE = 1.00;  // 학기 프로젝트 가중치 (시즌 종료 시 1회 별도 입력)

// --- 캡 시스템: 개인 브랜드가치 기여 상한 = 길드 평균의 N배 ---
var GS_CAP_MULTIPLIER = 2.0; // 평균의 2배 초과분은 잘라냄

// --- 미션별 가중치 (합계가 반드시 1.0이 되도록 설정) ---
// 마스터가 직접 수정. 미사용 미션은 0으로 두면 됨.
var MISSION_WEIGHTS = {
  "M01": 0.05,
  "M02": 0.20,
  "M03": 0.10,
  "M04": 0.10,
  "M07": 0.20,
  "M12": 0.25,
  "M14": 0.05,
  "M15": 0.05
};

// --- 길드 ID 목록 ---
var GUILD_IDS = ["GUILD_01", "GUILD_02", "GUILD_03", "GUILD_04", "GUILD_05"];

// --- 시트명 상수 ---
var SHEET_NAMES = {
  GUILD_MEMBERS  : "길드_구성",
  MISSION_LOG    : "길드_미션로그",
  GS_MONTHLY     : "길드_GS_월간",
  MANUAL_EVAL    : "길드_정성평가",
  MAIN           : "메인",
  ASSET_USE      : "자산사용",
  P2P_LOG        : "P2P거래로그",
  ACHIEVEMENT    : "학생업적달성",
  ACHIEVEMENT_MASTER : "업적마스터",
  PEER_EVAL    : "길드_동료평가",
  ACTIVITY_LOG : "길드활동로그"
};

// --- 미션 상태값 상수 ---
var MISSION_STATUS = {
  ACTIVE    : "진행중",
  CLEARED   : "클리어",
  FAILED    : "실패",
  PENDING   : "대기"
};

// --- 자동 검증 대상 미션 목록 ---
var AUTO_VERIFY_MISSIONS = ["M02", "M04", "M07", "M12"];

// --- M07 간식 가격 상한 배율 ---
var M07_PRICE_LIMIT = 1.5;


// ============================================================
// 섹션 2 — 미션 발표 (길드_미션로그 행 생성)
// ============================================================

/**
 * 미션을 발표하고 길드_미션로그에 5개 길드 행을 자동 생성합니다.
 *
 * 사용법 (Apps Script 편집기에서 직접 실행):
 *   announceGuildMission("M02", "2026-05-01", "2026-05-31")
 *
 * @param {string} missionId  - 미션 코드 (예: "M02")
 * @param {string} startDate  - 발표일 (YYYY-MM-DD)
 * @param {string} endDate    - 마감일 (YYYY-MM-DD)
 */
function announceGuildMission(missionId, startDate, endDate) {
  if (!missionId || !startDate || !endDate) {
    throw new Error("announceGuildMission: missionId, startDate, endDate 모두 필수입니다.");
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  if (!sheet) throw new Error("시트를 찾을 수 없습니다: " + SHEET_NAMES.MISSION_LOG);

  // 중복 발표 방지 — 같은 missionId가 이미 존재하면 에러
  var existing = _getMissionLogRows(sheet, missionId);
  if (existing.length > 0) {
    throw new Error("이미 발표된 미션입니다: " + missionId + " (" + existing.length + "개 행 존재)");
  }

  var startDateObj = new Date(startDate);
  var endDateObj   = new Date(endDate);

  // 길드별 1행씩 생성
  // 컬럼: A=미션ID, B=길드ID, C=발표일, D=마감일, E=상태, F=자동검증결과,
  //        G=정성평가점수, H=시너지점수, I=최종점수, J=비고
  GUILD_IDS.forEach(function(guildId) {
    sheet.appendRow([
      missionId,              // A: 미션ID
      guildId,                // B: 길드ID
      startDateObj,           // C: 발표일
      endDateObj,             // D: 마감일
      MISSION_STATUS.ACTIVE,  // E: 상태
      "",                     // F: 자동검증결과 (검증 후 채워짐)
      "",                     // G: 정성평가점수
      "",                     // H: 시너지점수
      "",                     // I: 최종점수
      ""                      // J: 비고
    ]);
  });

  Logger.log("[announceGuildMission] " + missionId + " 발표 완료. 길드 " + GUILD_IDS.length + "개 행 생성.");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    missionId + " 발표 완료. " + GUILD_IDS.length + "개 길드 행 생성됨.",
    "길드 미션 발표", 4
  );
}


// ============================================================
// 섹션 3 — 자동 검증
// ============================================================

// ----------------------------------------------------------
// 3-0. 전체 자동 검증 일괄 실행 (매일 트리거에서 호출)
// ----------------------------------------------------------

/**
 * 매일 트리거에서 호출되는 자동 검증 메인 함수.
 * 진행중 상태인 자동 검증 미션을 찾아 각 검증 함수를 실행합니다.
 */
function runDailyGuildVerification() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  if (!sheet) {
    Logger.log("[runDailyGuildVerification] 미션 로그 시트 없음. 종료.");
    return;
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  AUTO_VERIFY_MISSIONS.forEach(function(missionId) {
    var rows = _getMissionLogRows(sheet, missionId);
    rows.forEach(function(rowInfo) {
      // 진행중 상태인 경우만 검증
      if (rowInfo.status !== MISSION_STATUS.ACTIVE) return;

      var result = _runVerifyForMission(missionId, rowInfo.guildId, rowInfo.startDate, rowInfo.endDate);

      // F열(자동검증결과) 업데이트
      sheet.getRange(rowInfo.rowIndex, 6).setValue(result.resultText);

      // 마감일이 지났으면 상태 확정
      var endDate = new Date(rowInfo.endDate);
      endDate.setHours(0, 0, 0, 0);
      if (today > endDate) {
        var finalStatus = result.cleared ? MISSION_STATUS.CLEARED : MISSION_STATUS.FAILED;
        sheet.getRange(rowInfo.rowIndex, 5).setValue(finalStatus);
        Logger.log("[runDailyGuildVerification] " + missionId + " / " + rowInfo.guildId + " → " + finalStatus);
      }
    });
  });

  Logger.log("[runDailyGuildVerification] 완료: " + today.toISOString());
}

/**
 * 미션 코드에 따라 올바른 검증 함수를 호출합니다.
 * @returns {object} { cleared: boolean, resultText: string }
 */
function _runVerifyForMission(missionId, guildId, startDate, endDate) {
  switch (missionId) {
    case "M02": return verifyMissionM02(guildId, startDate, endDate);
    case "M04": return verifyMissionM04(guildId, startDate, endDate);
    case "M07": return verifyMissionM07(guildId, startDate, endDate);
    case "M12": return verifyMissionM12(guildId, startDate, endDate);
    default:
      return { cleared: false, resultText: "검증 함수 없음: " + missionId };
  }
}

// ----------------------------------------------------------
// 3-1. M02 — 다섯 개의 다른 길
// 조건: 기간 내 길드원 전원이 신규 업적 1개 이상 달성
//        단, 길드원 간 업적 종류(카테고리)가 모두 달라야 함
// ----------------------------------------------------------

/**
 * @param {string} guildId
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {object} { cleared: boolean, resultText: string, detail: object }
 */
function verifyMissionM02(guildId, startDate, endDate) {
  var members = _getGuildMembers(guildId);
  if (members.length === 0) {
    return { cleared: false, resultText: "길드 멤버 없음", detail: {} };
  }

  var start = _toMidnight(startDate);
  var end   = _toEndOfDay(endDate);

  var achieveSheet = SpreadsheetApp.getActiveSpreadsheet()
                       .getSheetByName(SHEET_NAMES.ACHIEVEMENT);
  var masterSheet  = SpreadsheetApp.getActiveSpreadsheet()
                       .getSheetByName(SHEET_NAMES.ACHIEVEMENT_MASTER);
  if (!achieveSheet || !masterSheet) {
    return { cleared: false, resultText: "업적 시트 없음", detail: {} };
  }

  // 미션 시작일 고정: 5월 8일 이후 달성분만 인정
  var missionStart = new Date("2026-05-08T00:00:00");
  var effectiveStart = (start > missionStart) ? start : missionStart;

  // 기간 내 각 멤버의 신규 업적 달성 목록 조회 (업적ID 기준)
  // 학생업적달성 컬럼 구조: A=학생명, B=업적ID, C=달성일시, ...
  var achieveData      = achieveSheet.getDataRange().getValues();
  var memberAchieveMap = {}; // { 학생명: [업적ID, ...] }
  members.forEach(function(name) { memberAchieveMap[name] = []; });

  for (var j = 1; j < achieveData.length; j++) {
    var studentName = String(achieveData[j][0]).trim();
    var achId       = String(achieveData[j][1]).trim();
    var achDate     = new Date(achieveData[j][4]);

    if (!memberAchieveMap.hasOwnProperty(studentName)) continue;
    if (achDate < effectiveStart || achDate > end) continue; // 5/8 이후만 인정

    memberAchieveMap[studentName].push(achId); // 카테고리 아닌 업적ID 저장
  }

  // 검증 1: 전원 최소 1개 달성 여부
  var notAchieved = members.filter(function(name) {
    return memberAchieveMap[name].length === 0;
  });

  if (notAchieved.length > 0) {
    return {
      cleared: false,
      resultText: "미달성: " + notAchieved.join(", "),
      detail: memberAchieveMap
    };
  }

  // 검증 2: 업적ID 중복 여부 (선착순 — 첫 번째 달성 업적ID 기준)
  var usedAchIds    = [];
  var duplicateFound = false;
  var duplicateInfo  = "";

  members.forEach(function(name) {
    var firstId = memberAchieveMap[name][0];
    if (usedAchIds.indexOf(firstId) !== -1) {
      duplicateFound = true;
      duplicateInfo  = "업적 중복: " + firstId + " (" + name + ")";
    } else {
      usedAchIds.push(firstId);
    }
  });

  if (duplicateFound) {
    return {
      cleared: false,
      resultText: duplicateInfo,
      detail: memberAchieveMap
    };
  }

  return {
    cleared: true,
    resultText: "클리어 — 전원 달성, 카테고리 모두 상이",
    detail: memberAchieveMap
  };
}

// ----------------------------------------------------------
// 3-2. M04 — 우리 길드의 장보기
// 조건: 기간 내 길드원 전원이 P2P 거래 5건 이상
//        (보낸 건 + 받은 건 합산)
// ----------------------------------------------------------

/**
 * @param {string} guildId
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {object} { cleared: boolean, resultText: string, detail: object }
 */
function verifyMissionM04(guildId, startDate, endDate) {
  var members = _getGuildMembers(guildId);
  if (members.length === 0) {
    return { cleared: false, resultText: "길드 멤버 없음", detail: {} };
  }

  var start = _toMidnight(startDate);
  var end   = _toEndOfDay(endDate);

  // P2P거래로그 컬럼 구조: A=타임스탬프, B=보낸사람, C=받는사람, D=금액, ...
  var p2pSheet = SpreadsheetApp.getActiveSpreadsheet()
                   .getSheetByName(SHEET_NAMES.P2P_LOG);
  if (!p2pSheet) {
    return { cleared: false, resultText: "P2P거래로그 시트 없음", detail: {} };
  }

  var p2pData = p2pSheet.getDataRange().getValues();
  var countMap = {}; // { 학생명: 건수 }
  members.forEach(function(name) { countMap[name] = 0; });

  for (var i = 1; i < p2pData.length; i++) {
    var ts     = new Date(p2pData[i][0]);
    var sender = String(p2pData[i][1]).trim();
    var recvr  = String(p2pData[i][2]).trim();

    if (ts < start || ts > end) continue;

    if (countMap.hasOwnProperty(sender)) countMap[sender]++;
    if (countMap.hasOwnProperty(recvr))  countMap[recvr]++;
  }

  var underFive = members.filter(function(name) { return countMap[name] < 5; });

  if (underFive.length > 0) {
    var detail = underFive.map(function(name) {
      return name + "(" + countMap[name] + "건)";
    }).join(", ");
    return {
      cleared: false,
      resultText: "5건 미달: " + detail,
      detail: countMap
    };
  }

  return {
    cleared: true,
    resultText: "클리어 — 전원 5건 이상 달성",
    detail: countMap
  };
}

// ----------------------------------------------------------
// 3-3. M07 — 미식 탐험가 (6월 월간)
// 조건: 기간 내 길드원 전원이 간식 시장에서 각자 다른 상품 구매
//        각 구매 시점의 가격 배율이 1.5배 이하일 것
// ----------------------------------------------------------

/**
 * @param {string} guildId
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {object} { cleared: boolean, resultText: string, detail: object }
 */
function verifyMissionM07(guildId, startDate, endDate) {
  var members = _getGuildMembers(guildId);
  if (members.length === 0) {
    return { cleared: false, resultText: "길드 멤버 없음", detail: {} };
  }

  var start = _toMidnight(startDate);
  var end   = _toEndOfDay(endDate);

  // 자산사용 시트에서 간식 구매 내역 조회
  // 자산사용 컬럼 구조: A=타임스탬프, B=학생명, C=항목유형, D=상품명, E=결제금액, F=기준가, G=배율, ...
  // ※ 항목유형이 "간식" 또는 "snack"인 행만 대상
  var assetSheet = SpreadsheetApp.getActiveSpreadsheet()
                     .getSheetByName(SHEET_NAMES.ASSET_USE);
  if (!assetSheet) {
    return { cleared: false, resultText: "자산사용 시트 없음", detail: {} };
  }

  var assetData = assetSheet.getDataRange().getValues();

  // { 학생명: { itemName: string, ratio: number } } — 조건 충족한 첫 구매만 기록
  var validPurchaseMap = {};
  members.forEach(function(name) { validPurchaseMap[name] = null; });

  for (var i = 1; i < assetData.length; i++) {
    var ts       = new Date(assetData[i][0]);
    var student  = String(assetData[i][1]).trim();
    var itemType = String(assetData[i][2]).trim();
    var itemName = String(assetData[i][3]).trim();
    var ratio    = parseFloat(assetData[i][6]); // G열: 배율

    if (ts < start || ts > end) continue;
    if (!validPurchaseMap.hasOwnProperty(student)) continue;
    if (itemType !== "간식" && itemType !== "snack") continue;
    if (isNaN(ratio) || ratio > M07_PRICE_LIMIT) continue;
    if (validPurchaseMap[student] !== null) continue; // 이미 기록됨

    validPurchaseMap[student] = { itemName: itemName, ratio: ratio };
  }

  // 검증 1: 전원 유효 구매 여부
  var notBought = members.filter(function(name) {
    return validPurchaseMap[name] === null;
  });
  if (notBought.length > 0) {
    return {
      cleared: false,
      resultText: "미구매(또는 1.5배 초과): " + notBought.join(", "),
      detail: validPurchaseMap
    };
  }

  // 검증 2: 상품 중복 여부
  var usedItems = [];
  var dupItem   = "";
  members.forEach(function(name) {
    var item = validPurchaseMap[name].itemName;
    if (usedItems.indexOf(item) !== -1) {
      dupItem = "상품 중복: " + item + " (" + name + ")";
    } else {
      usedItems.push(item);
    }
  });

  if (dupItem) {
    return { cleared: false, resultText: dupItem, detail: validPurchaseMap };
  }

  return {
    cleared: true,
    resultText: "클리어 — 전원 다른 상품 1.5배 이하 구매",
    detail: validPurchaseMap
  };
}

// ----------------------------------------------------------
// 3-4. M12 — 길드 명예의 일격
// 조건: M02와 동일하나 기간이 한 주 (검증 로직 재사용)
// ----------------------------------------------------------

/**
 * @param {string} guildId
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {object} { cleared: boolean, resultText: string, detail: object }
 */
function verifyMissionM12(guildId, startDate, endDate) {
  // M02 검증 로직과 동일 — 기간만 다름
  var result = verifyMissionM02(guildId, startDate, endDate);
  // resultText 앞에 [M12] 태그 추가해서 구분
  result.resultText = "[M12] " + result.resultText;
  return result;
}


// ============================================================
// 섹션 4 — 정성평가 입력
// ============================================================

/**
 * 마스터가 미션 정성평가 점수를 입력합니다.
 * 점수는 길드_정성평가 시트에 기록되고,
 * 길드_미션로그의 G열(정성평가점수)도 함께 업데이트됩니다.
 *
 * 사용법:
 *   setGuildMissionManualScore("M01", "GUILD_01", 90, "첫 모임 활기차게 진행")
 *
 * @param {string} missionId
 * @param {string} guildId
 * @param {number} score      - 0~100
 * @param {string} comment    - 선택 사항
 */
function setGuildMissionManualScore(missionId, guildId, score, comment) {
  if (!missionId || !guildId || score === undefined || score === null) {
    throw new Error("setGuildMissionManualScore: missionId, guildId, score 필수입니다.");
  }
  if (score < 0 || score > 100) {
    throw new Error("score는 0~100 사이여야 합니다. 입력값: " + score);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) 길드_정성평가 시트에 기록
  // 컬럼: A=미션ID, B=길드ID, C=점수, D=입력일시, E=코멘트
  var evalSheet = ss.getSheetByName(SHEET_NAMES.MANUAL_EVAL);
  if (!evalSheet) throw new Error("시트 없음: " + SHEET_NAMES.MANUAL_EVAL);

  evalSheet.appendRow([
    missionId,
    guildId,
    score,
    new Date(),
    comment || ""
  ]);

  // 2) 길드_미션로그 G열 업데이트
  var logSheet = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  if (logSheet) {
    var rows = _getMissionLogRows(logSheet, missionId);
    rows.forEach(function(rowInfo) {
      if (rowInfo.guildId === guildId) {
        logSheet.getRange(rowInfo.rowIndex, 7).setValue(score); // G열
      }
    });
  }

  Logger.log("[setGuildMissionManualScore] " + missionId + " / " + guildId + " → " + score + "점");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    guildId + " " + missionId + " 정성평가: " + score + "점 입력 완료",
    "정성평가 입력", 3
  );
}


// ============================================================
// 섹션 5 — 최종점수 계산
// ============================================================

/**
 * 길드_미션로그의 특정 행 최종점수(I열)를 계산하고 기록합니다.
 *
 * 최종점수 = 자동검증결과(클리어=100/실패=0) × (1 - 정성비율)
 *            + 정성평가점수 × 정성비율
 *            (정성평가가 없는 자동 검증 미션은 자동검증결과만 사용)
 *
 * 정성평가 전용 미션(M01, M03, M14, M15)은 정성평가점수를 그대로 사용.
 *
 * 사용법:
 *   calcGuildMissionFinalScore("M02", "GUILD_01")
 *
 * @param {string} missionId
 * @param {string} guildId
 */
function calcGuildMissionFinalScore(missionId, guildId) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  if (!logSheet) throw new Error("시트 없음: " + SHEET_NAMES.MISSION_LOG);

  var rows = _getMissionLogRows(logSheet, missionId);
  var targetRow = null;
  rows.forEach(function(r) {
    if (r.guildId === guildId) targetRow = r;
  });

  if (!targetRow) {
    Logger.log("[calcGuildMissionFinalScore] 행 없음: " + missionId + " / " + guildId);
    return;
  }

  var autoResult  = String(logSheet.getRange(targetRow.rowIndex, 6).getValue()); // F열
  var manualScore = logSheet.getRange(targetRow.rowIndex, 7).getValue();          // G열
  var synergyPts  = logSheet.getRange(targetRow.rowIndex, 8).getValue() || 0;    // H열

  var autoScore;
  // 자동 검증 미션
  if (AUTO_VERIFY_MISSIONS.indexOf(missionId) !== -1) {
    autoScore = (autoResult.indexOf("클리어") !== -1) ? 100 : 0;
  } else {
    // 정성평가 전용 미션 — 자동점수 없음
    autoScore = null;
  }

  var finalScore;
  if (autoScore === null) {
    // 정성평가 전용
    finalScore = (manualScore !== "" && manualScore !== null) ? Number(manualScore) : 0;
  } else if (manualScore !== "" && manualScore !== null) {
    // 자동 + 정성 혼합 (M04: 자동 70% + 정성 30%)
    finalScore = autoScore * 0.7 + Number(manualScore) * 0.3;
  } else {
    // 자동만
    finalScore = autoScore;
  }

  // 시너지 점수 가산
  finalScore = Math.min(100, finalScore + Number(synergyPts));

  logSheet.getRange(targetRow.rowIndex, 9).setValue(Math.round(finalScore)); // I열

  Logger.log("[calcGuildMissionFinalScore] " + missionId + " / " + guildId +
             " → 최종점수: " + Math.round(finalScore));
}

/**
 * 모든 미션, 모든 길드의 최종점수를 일괄 계산합니다.
 * 월간 GS 산출 직전에 호출됩니다.
 */
function calcAllGuildMissionFinalScores() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  if (!logSheet) return;

  var data = logSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var missionId = String(data[i][0]).trim();
    var guildId   = String(data[i][1]).trim();
    if (!missionId || !guildId) continue;
    calcGuildMissionFinalScore(missionId, guildId);
  }
  Logger.log("[calcAllGuildMissionFinalScores] 전체 최종점수 계산 완료.");
}


// ============================================================
// 섹션 6 — 월간 GS 산출
// ============================================================

/**
 * 매월 30일 13:00에 트리거로 자동 실행됩니다.
 * calcAllGuildMissionFinalScores 선행 호출 후 GS 계산.
 */
function calcMonthlyGS() {
  // 최종점수 먼저 갱신
  calcAllGuildMissionFinalScores();

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet   = ss.getSheetByName(SHEET_NAMES.MAIN);
  var logSheet    = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  var gsSheet     = ss.getSheetByName(SHEET_NAMES.GS_MONTHLY);
  var actSheet    = ss.getSheetByName("길드활동로그"); // 길드활동로그 시트

  if (!mainSheet || !logSheet || !gsSheet) {
    Logger.log("[calcMonthlyGS] 필수 시트 없음. 종료.");
    return;
  }

  var now       = new Date();
  var yearMonth = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM");

  // 이미 이번 달 계산됐으면 중복 방지
  var gsData = gsSheet.getDataRange().getValues();
  for (var i = 1; i < gsData.length; i++) {
    if (String(gsData[i][0]).trim() === yearMonth) {
      Logger.log("[calcMonthlyGS] 이번 달 이미 계산됨: " + yearMonth);
      return;
    }
  }

  // ── 메인 시트에서 학생별 자산 조회 ───────────────────────────
  // ── 메인 시트에서 학생별 현재 브랜드가치 조회 (종료값용) ───────
var mainData = mainSheet.getDataRange().getValues();
var assetMap = {}; // { 학생명: 현재브랜드가치 }
for (var j = 1; j < mainData.length; j++) {
  var sName  = String(mainData[j][1]).trim();
  var sAsset = parseFloat(mainData[j][2]) || 0;
  if (sName) assetMap[sName] = sAsset;
}

// ── 브랜드가치추적 시트에서 월 시작값 조회 ───────────────────
// 컬럼: A=학생명, B~=날짜별 브랜드가치 (헤더가 날짜)
var trackSheet   = ss.getSheetByName("브랜드가치추적");
var startAssetMapByStudent = {}; // { 학생명: 월시작 브랜드가치 }

if (trackSheet) {
  var trackData  = trackSheet.getDataRange().getValues();
  var trackHeaders = trackData[0]; // 1행: 날짜 헤더

  // 이번 달 1일 찾기 (없으면 이번 달 내 가장 첫 번째 날짜)
  var monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  var monthEnd   = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  var startColIdx = -1;
  var closestDate = null;

  for (var h = 1; h < trackHeaders.length; h++) {
    var hDate = new Date(trackHeaders[h]);
    if (isNaN(hDate)) continue;
    hDate.setHours(0, 0, 0, 0);
    if (hDate >= monthStart && hDate <= monthEnd) {
      if (closestDate === null || hDate < closestDate) {
        closestDate  = hDate;
        startColIdx  = h;
      }
    }
  }

  if (startColIdx !== -1) {
    for (var t = 1; t < trackData.length; t++) {
      var tName  = String(trackData[t][0]).trim();
      var tAsset = parseFloat(trackData[t][startColIdx]) || 0;
      if (tName) startAssetMapByStudent[tName] = tAsset;
    }
  }
}

  // ── 미션 점수 집계 (길드별 가중 합산) ─────────────────────────
  var missionScoreMap = {}; // { guildId: 합산점수 }
  GUILD_IDS.forEach(function(g) { missionScoreMap[g] = 0; });

  var logData = logSheet.getDataRange().getValues();
  for (var k = 1; k < logData.length; k++) {
    var mId    = String(logData[k][0]).trim();
    var gId    = String(logData[k][1]).trim();
    var mScore = parseFloat(logData[k][8]) || 0; // I열
    var weight = MISSION_WEIGHTS[mId] || 0;
    if (missionScoreMap.hasOwnProperty(gId)) {
      missionScoreMap[gId] += mScore * weight;
    }
  }

  // ── 길드활동로그에서 참여/출석 횟수 집계 ──────────────────────
  // 길드활동로그 컬럼: A=날짜, B=유형(미션/세션), C=미션ID, D=학생명, E=참여여부(O/X)
  // 길드별로 집계: { guildId: { missionCount: N, sessionCount: N } }
  var actMap = {}; // { guildId: { missionCount, sessionCount } }
  GUILD_IDS.forEach(function(g) {
    actMap[g] = { missionCount: 0, sessionCount: 0 };
  });

  // ── 길드활동로그 → 미션 참여 횟수만 집계 ─────────────────────
  if (actSheet) {
    var actData         = actSheet.getDataRange().getValues();
    var mbSheet         = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
    var studentGuildMap = _buildStudentGuildMap(mbSheet);

    for (var a = 1; a < actData.length; a++) {
      var aType   = String(actData[a][1]).trim(); // B열: 유형
      var aName   = String(actData[a][3]).trim(); // D열: 학생명
      var aResult = String(actData[a][4]).trim(); // E열: O/X

      if (aResult !== "O" || aType !== "미션") continue;
      var aGuildId = studentGuildMap[aName];
      if (!aGuildId || !actMap[aGuildId]) continue;
      actMap[aGuildId].missionCount++;
    }
  }

  // ── 길드_세션출석 → 세션 참석 횟수 집계 ─────────────────────
  // 컬럼: A=날짜, B=길드ID, C=학생명, D=출석여부(O/X)
  var sesSheet = ss.getSheetByName("길드_세션출석");
  if (sesSheet) {
    var sesData = sesSheet.getDataRange().getValues();
    for (var s = 1; s < sesData.length; s++) {
      var sGuildId = String(sesData[s][1]).trim(); // B열: 길드ID
      var sResult  = String(sesData[s][3]).trim(); // D열: O/X
      if (sResult !== "O") continue;
      if (!actMap[sGuildId]) continue;
      actMap[sGuildId].sessionCount++;
    }
  }
  // ── 길드별 GS 계산 ────────────────────────────────────────────
  var gsResults = [];

  GUILD_IDS.forEach(function(guildId) {
    var members = _getGuildMembers(guildId);
    if (members.length === 0) return;

    // 자산 배열 추출
    var assets = members.map(function(name) {
      return assetMap[name] || 0;
    });

    // 시작 자산: 전월 GS 시트에서, 없으면 현재값
    var prevRow     = _getPrevMonthGSRow(gsSheet, guildId, yearMonth);
    // 브랜드가치추적에서 월 시작값 합산, 없으면 전월 종료값, 그것도 없으면 현재값
    var startAssets;
    if (Object.keys(startAssetMapByStudent).length > 0) {
      startAssets = members.reduce(function(sum, name) {
        return sum + (startAssetMapByStudent[name] || assetMap[name] || 0);
      }, 0);
    } else if (prevRow) {
      startAssets = prevRow.endAssetTotal;
    } else {
      startAssets = _sum(assets);
    }

    var endAssetTotal   = _sum(assets);
    var memberCount     = members.length;
    var perCapitaGrowth = (endAssetTotal - startAssets) / memberCount;

    // 캡 시스템: 고자산 멤버 독점 방지
    var avgAsset = endAssetTotal / memberCount;
    var capLimit = avgAsset * GS_CAP_MULTIPLIER;
    if (perCapitaGrowth > capLimit) {
      Logger.log("[calcMonthlyGS] " + guildId + " 캡 적용: perCapitaGrowth " +
                 Math.round(perCapitaGrowth) + " → " + Math.round(capLimit));
      perCapitaGrowth = capLimit;
    }

    // ── 참여/출석 점수 계산 ─────────────────────────────────────
    // 길드 전체 참여 횟수 합산 → 인당 평균으로 정규화
    var actData_guild    = actMap[guildId] || { missionCount: 0, sessionCount: 0 };
    // 인당 평균 참여 횟수 (소수점 허용)
    var avgMissionPart   = memberCount > 0 ? actData_guild.missionCount / memberCount : 0;
    var avgSessionAttend = memberCount > 0 ? actData_guild.sessionCount / memberCount : 0;

    var missionPartVal  = GS_MISSION_PARTICIPATION * avgMissionPart;
    var sessionAttendVal = GS_SESSION_ATTENDANCE   * avgSessionAttend;

    // ── GS 합산 ──────────────────────────────────────────────────
    var alphaVal   = GS_ALPHA        * perCapitaGrowth;
    var missionVal = GS_MISSION_RATE * (missionScoreMap[guildId] || 0);
    var totalGS    = alphaVal + missionPartVal + sessionAttendVal + missionVal;

    gsResults.push({
      guildId          : guildId,
      startAssetTotal  : startAssets,
      endAssetTotal    : endAssetTotal,
      perCapitaGrowth  : Math.round(perCapitaGrowth),
      avgMissionPart   : Math.round(avgMissionPart * 100) / 100,
      avgSessionAttend : Math.round(avgSessionAttend * 100) / 100,
      missionScore     : Math.round(missionScoreMap[guildId] || 0),
      alphaVal         : Math.round(alphaVal * 100) / 100,
      missionPartVal   : Math.round(missionPartVal * 100) / 100,
      sessionAttendVal : Math.round(sessionAttendVal * 100) / 100,
      missionVal       : Math.round(missionVal * 100) / 100,
      totalGS          : Math.round(totalGS * 100) / 100
    });
  });

  // 순위 산정
  // ── 교체 후 ───────────────────────────────────────────────────

  // 알파 정규화: 전체 길드 중 최대 인당 증가량 기준
  var maxPerCapita = Math.max.apply(null, gsResults.map(function(r) {
    return r.perCapitaGrowth > 0 ? r.perCapitaGrowth : 0;
  }));

  gsResults.forEach(function(r) {
    var normalizedAlpha = maxPerCapita > 0 ? (r.perCapitaGrowth / maxPerCapita) : 0;
    r.alphaVal        = Math.round(normalizedAlpha * GS_ALPHA * 1000) / 1000;
    r.totalGS         = Math.round((r.alphaVal + r.missionPartVal + r.sessionAttendVal + r.missionVal) * 1000); // ← 1000 곱해서 정수
  });

  // 순위 산정
  gsResults.sort(function(a, b) { return b.totalGS - a.totalGS; });
  gsResults.forEach(function(r, idx) { r.rank = idx + 1; });

  // ── 길드_GS_월간 시트에 기록 ──────────────────────────────────
  // 컬럼 구조 (기존과 동일한 위치 유지, 베타 항목 내용만 교체):
  //   A=월, B=길드ID,
  //   C=시작_자산합계, D=종료_자산합계,
  //   E=인당_증가량(캡적용),
  //   F=인당_미션참여횟수, G=인당_세션출석횟수,   ← 기존 F=정규화증가, G=시작표준편차 자리 재활용
  //   H=미션점수합계,
  //   I=알파_적용값, J=미션참여_적용값, K=세션출석_적용값, L=미션항_적용값,
  //   M=월간_GS_합계, N=월간_순위
  //   (P, Q, R, S 열은 setGuildProjectScore 가 나중에 채움 — 위치 고정)
  gsResults.forEach(function(r) {
    gsSheet.appendRow([
      yearMonth,            // A
      r.guildId,            // B
      r.startAssetTotal,    // C
      r.endAssetTotal,      // D
      r.perCapitaGrowth,    // E
      r.avgMissionPart,     // F  ← 인당 미션 참여 횟수
      r.avgSessionAttend,   // G  ← 인당 세션 출석 횟수
      r.missionScore,       // H
      r.alphaVal,           // I
      r.missionPartVal,     // J  ← 미션참여 적용값
      r.sessionAttendVal,   // K  ← 세션출석 적용값
      r.missionVal,         // L
      r.totalGS,            // M
      r.rank                // N
    ]);
  });

  Logger.log("[calcMonthlyGS] " + yearMonth + " GS 산출 완료. " + gsResults.length + "개 길드.");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    yearMonth + " 월간 GS 산출 완료.",
    "GS 산출", 5
  );

  // 길드 GS 산출 직후 개인 기여점수도 함께 계산
  calcMonthlyIndividualGS();
}


/**
 * 수동으로 월간 GS를 즉시 실행할 때 사용합니다.
 * (Apps Script 편집기에서 직접 실행)
 */
function runMonthlyGSManually() {
  calcMonthlyGS();
}

/**
 * 시즌 종료 후 학기 프로젝트 점수를 입력합니다.
 * P열(프로젝트점수)과 Q열(적용값 = score × 1.00)을 업데이트하고
 * R열(월간_GS)도 재계산합니다.
 *
 * 시즌 GS 구조:
 *   월간 GS (5월) + 월간 GS (6월) + 월간 GS (7월) = 최대 3.00
 *   + 학기 프로젝트 (score × 1.00)                = 최대 1.00
 *   ─────────────────────────────────────────────────────────
 *   시즌 GS 합계                                   = 최대 4.00
 *
 * 사용법 (Apps Script 편집기에서 직접 실행):
 *   setGuildProjectScore("2026-07", "GUILD_01", 88)
 *
 * @param {string} yearMonth  - 대상 월 (yyyy-MM)
 * @param {string} guildId
 * @param {number} score      - 0~100
 */
function setGuildProjectScore(yearMonth, guildId, score) {
  if (!yearMonth || !guildId || score === undefined) {
    throw new Error("setGuildProjectScore: yearMonth, guildId, score 모두 필수입니다.");
  }
  if (score < 0 || score > 100) {
    throw new Error("score는 0~100 사이여야 합니다. 입력값: " + score);
  }

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var gsSheet = ss.getSheetByName(SHEET_NAMES.GS_MONTHLY);
  if (!gsSheet) throw new Error("길드_GS_월간 시트를 찾을 수 없습니다.");

  var data       = gsSheet.getDataRange().getValues();
  var projectVal = Math.round(GS_PROJECT_RATE * score * 1000); // 동일 스케일
  var updated    = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== yearMonth) continue;
    if (String(data[i][1]).trim() !== guildId)   continue;

    // 현재 totalGS(M열=인덱스12) + 프로젝트 적용값
    var currentGS = parseFloat(data[i][12]) || 0;
    var newTotalGS = Math.round((currentGS + projectVal) * 100) / 100;

    gsSheet.getRange(i + 1, 15).setValue(score);                              // O열: 학기프로젝트점수
    gsSheet.getRange(i + 1, 16).setValue(Math.round(projectVal * 100) / 100); // P열: 프로젝트적용값
    gsSheet.getRange(i + 1, 13).setValue(newTotalGS);                         // M열: 월간_GS 업데이트
    updated = true;

    Logger.log("[setGuildProjectScore] " + yearMonth + " / " + guildId +
               " → 프로젝트: " + score + "점, 새 GS: " + newTotalGS);
    break;
  }

  if (!updated) {
    throw new Error("해당 월/길드 행을 찾을 수 없습니다: " + yearMonth + " / " + guildId);
  }

  // 순위 재산정
  _recalcGSRank(gsSheet, yearMonth);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    guildId + " 학기 프로젝트 " + score + "점 입력 완료.",
    "프로젝트 점수 입력", 4
  );
}


/**
 * 특정 월의 GS 순위를 재산정합니다. (R열 기준)
 */
function _recalcGSRank(gsSheet, yearMonth) {
  var data = gsSheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== yearMonth) continue;
    rows.push({ rowIndex: i + 1, totalGS: parseFloat(data[i][12]) || 0 }); // M열 = 인덱스 12
  }

  rows.sort(function(a, b) { return b.totalGS - a.totalGS; });
  rows.forEach(function(r, idx) {
    gsSheet.getRange(r.rowIndex, 14).setValue(idx + 1); // N열 = 14번째
  });
}


// ============================================================
// 섹션 7 — 트리거 등록
// ============================================================

/**
 * 이 함수를 한 번만 실행하면 자동 트리거 2개가 등록됩니다.
 *   1) runDailyGuildVerification — 매일 09:00
 *   2) calcMonthlyGS            — 매월 30일 13:00
 *
 * ※ 중복 등록 방지: 기존 트리거를 먼저 삭제하고 새로 등록합니다.
 * ※ Apps Script 편집기에서 한 번만 실행하면 됩니다.
 */
function setupGuildTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 기존 길드 관련 트리거 삭제
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === "runDailyGuildVerification" || fn === "calcMonthlyGS") {
      ScriptApp.deleteTrigger(t);
      Logger.log("[setupGuildTriggers] 기존 트리거 삭제: " + fn);
    }
  });

  // 1) 매일 09:00 자동 검증
  ScriptApp.newTrigger("runDailyGuildVerification")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // 2) 매월 30일 13:00 GS 산출
  ScriptApp.newTrigger("calcMonthlyGS")
    .timeBased()
    .onMonthDay(30)
    .atHour(13)
    .create();

  Logger.log("[setupGuildTriggers] 트리거 2개 등록 완료.");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "트리거 등록 완료: 일일 검증(09:00) + 월간 GS(매월 30일 13:00)",
    "트리거 설정", 5
  );
}


// ============================================================
// 섹션 8 — 길드 세션 출석 기록
// ============================================================

/**
 * 길드 세션 출석을 O/X로 기록합니다.
 * 별도 시트('길드_세션출석')를 사용합니다.
 * 시트가 없으면 자동 생성합니다.
 *
 * 사용법:
 *   recordGuildSession("GUILD_01", "2026-05-09", { "홍길동": "O", "김철수": "X", ... })
 *
 * @param {string} guildId
 * @param {string} sessionDate  - YYYY-MM-DD
 * @param {object} attendanceMap - { 학생명: "O" | "X" }
 */
function recordGuildSession(guildId, sessionDate, attendanceMap) {
  if (!guildId || !sessionDate || !attendanceMap) {
    throw new Error("recordGuildSession: guildId, sessionDate, attendanceMap 모두 필수입니다.");
  }

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName   = "길드_세션출석";
  var sheet       = ss.getSheetByName(sheetName);

  // 시트 없으면 자동 생성
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["날짜", "길드ID", "학생명", "출석여부", "입력일시"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    Logger.log("[recordGuildSession] 길드_세션출석 시트 자동 생성.");
  }

  var inputTime = new Date();
  var dateObj   = new Date(sessionDate);

  Object.keys(attendanceMap).forEach(function(studentName) {
    var status = attendanceMap[studentName]; // "O" or "X"
    sheet.appendRow([dateObj, guildId, studentName, status, inputTime]);
  });

  Logger.log("[recordGuildSession] " + guildId + " / " + sessionDate +
             " 출석 기록 완료. " + Object.keys(attendanceMap).length + "명.");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    guildId + " " + sessionDate + " 세션 출석 기록 완료.",
    "세션 출석", 3
  );
}

/**
 * 특정 길드의 세션 출석률을 조회합니다.
 *
 * @param {string} guildId
 * @returns {object} { totalSessions, avgAttendanceRate, memberStats }
 */
function getGuildSessionStats(guildId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("길드_세션출석");
  if (!sheet) {
    Logger.log("[getGuildSessionStats] 길드_세션출석 시트 없음.");
    return null;
  }

  var data = sheet.getDataRange().getValues();
  var sessionDates = []; // 이 길드의 세션 날짜 목록
  var memberMap    = {}; // { 학생명: { O: 수, total: 수 } }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== guildId) continue;
    var dateStr = Utilities.formatDate(new Date(data[i][0]), "Asia/Seoul", "yyyy-MM-dd");
    var name    = String(data[i][2]).trim();
    var status  = String(data[i][3]).trim();

    if (sessionDates.indexOf(dateStr) === -1) sessionDates.push(dateStr);
    if (!memberMap[name]) memberMap[name] = { O: 0, total: 0 };
    memberMap[name].total++;
    if (status === "O") memberMap[name].O++;
  }

  var memberStats = {};
  Object.keys(memberMap).forEach(function(name) {
    var m = memberMap[name];
    memberStats[name] = {
      attended : m.O,
      total    : m.total,
      rate     : m.total > 0 ? Math.round(m.O / m.total * 100) : 0
    };
  });

  var totalSessions = sessionDates.length;
  var rateSum       = 0;
  var memberCount   = Object.keys(memberStats).length;
  Object.keys(memberStats).forEach(function(n) { rateSum += memberStats[n].rate; });

  return {
    totalSessions       : totalSessions,
    avgAttendanceRate   : memberCount > 0 ? Math.round(rateSum / memberCount) : 0,
    memberStats         : memberStats
  };
}


// ============================================================
// 섹션 9 — 유틸리티 함수
// ============================================================

/**
 * 길드 멤버 이름 목록을 반환합니다.
 * 길드_구성 시트: A=길드ID, B=길드명, C=멤버명, D=포트, E=가입일, F=탈퇴일
 * 탈퇴일이 비어있는 멤버만 현재 활성 멤버로 간주합니다.
 *
 * @param {string} guildId
 * @returns {string[]} 멤버 이름 배열
 */
function _getGuildMembers(guildId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
  if (!sheet) return [];

  var data    = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    var rowGuildId  = String(data[i][0]).trim();
    var memberName  = String(data[i][2]).trim();
    var leaveDate   = data[i][5]; // F열: 탈퇴일

    if (rowGuildId !== guildId) continue;
    if (!memberName) continue;
    if (leaveDate && leaveDate !== "") continue; // 탈퇴 멤버 제외

    members.push(memberName);
  }

  return members;
}

/**
 * 길드_미션로그에서 특정 missionId에 해당하는 행 정보를 모두 반환합니다.
 *
 * @param {Sheet} sheet
 * @param {string} missionId
 * @returns {object[]} [{ rowIndex, guildId, startDate, endDate, status }, ...]
 */
function _getMissionLogRows(sheet, missionId) {
  var data   = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var rowMissionId = String(data[i][0]).trim();
    if (rowMissionId !== missionId) continue;

    result.push({
      rowIndex  : i + 1,           // Sheets는 1-based, data는 0-based
      guildId   : String(data[i][1]).trim(),
      startDate : data[i][2],       // C열
      endDate   : data[i][3],       // D열
      status    : String(data[i][4]).trim()  // E열
    });
  }

  return result;
}

/**
 * GS 시트에서 전월 해당 길드 데이터 행을 반환합니다.
 * 없으면 null.
 */
function _getPrevMonthGSRow(gsSheet, guildId, currentYearMonth) {
  var data = gsSheet.getDataRange().getValues();
  var prev = null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() !== guildId) continue;
    var ym = String(data[i][0]).trim();
    if (ym < currentYearMonth) {
      // 가장 최근 이전 달
      if (!prev || ym > prev.yearMonth) {
        prev = {
          yearMonth     : ym,
          endAssetTotal : parseFloat(data[i][3]) || 0, // D열
          endStdDev     : parseFloat(data[i][6]) || 0  // G열
        };
      }
    }
  }

  return prev;
}

/** Date를 당일 00:00:00으로 변환 */
function _toMidnight(d) {
  var date = new Date(d);
  date.setHours(0, 0, 0, 0);
  return date;
}

/** Date를 당일 23:59:59로 변환 */
function _toEndOfDay(d) {
  var date = new Date(d);
  date.setHours(23, 59, 59, 999);
  return date;
}

/** 배열 합계 */
function _sum(arr) {
  return arr.reduce(function(acc, v) { return acc + (parseFloat(v) || 0); }, 0);
}

/** 표본 표준편차 (n-1) */
function _stdDev(arr) {
  if (arr.length < 2) return 0;
  var mean = _sum(arr) / arr.length;
  var variance = arr.reduce(function(acc, v) {
    return acc + Math.pow((parseFloat(v) || 0) - mean, 2);
  }, 0) / (arr.length - 1);
  return Math.sqrt(variance);
}


// ============================================================
// 섹션 10 — AuctionAdmin 길드 현황 탭용 GAS 함수
// ============================================================

/**
 * 이번 달 길드 GS 순위를 반환합니다.
 * AuctionAdmin.html의 renderGuildGsRank()에서 호출.
 *
 * @returns {object[]} [ { guildId, guildName, rank, totalGS, yearMonth }, ... ]
 */
function getGuildGsRank() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var gsSheet = ss.getSheetByName(SHEET_NAMES.GS_MONTHLY);
  var mbSheet = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
  if (!gsSheet) return [];

  var now       = new Date();
  var yearMonth = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM");

  // 이번 달 행 추출
  // 컬럼: A=월, B=길드ID, R=월간_GS_합계(17), S=월간_순위(18)
  var data    = gsSheet.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== yearMonth) continue;
    var guildId = String(data[i][1]).trim();
    results.push({
      guildId   : guildId,
      guildName : _getGuildName(mbSheet, guildId),
      rank      : parseInt(data[i][18]) || 0,  // S열
      totalGS   : parseFloat(data[i][17]) || 0, // R열
      yearMonth : yearMonth
    });
  }

  // 순위순 정렬
  results.sort(function(a, b) { return a.rank - b.rank; });

  // 이번 달 데이터가 없으면 전월 최신 데이터로 대체
  if (results.length === 0) {
    var prevMonth = _getPrevYearMonth(yearMonth);
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][0]).trim() !== prevMonth) continue;
      var gId = String(data[j][1]).trim();
      results.push({
        guildId   : gId,
        guildName : _getGuildName(mbSheet, gId),
        rank      : parseInt(data[j][18]) || 0,
        totalGS   : parseFloat(data[j][17]) || 0,
        yearMonth : prevMonth + " (전월)"
      });
    }
    results.sort(function(a, b) { return a.rank - b.rank; });
  }

  return results;
}


/**
 * 길드별 미션 클리어 현황을 반환합니다.
 * AuctionAdmin.html의 renderGuildMissionStatus()에서 호출.
 *
 * @returns {object} {
 *   missions: [ { missionId, missionName, star }, ... ],
 *   guilds:   [ { guildId, guildName, results: { M01: '클리어'|'실패'|'진행중'|'대기' } }, ... ]
 * }
 */
function getGuildMissionStatus() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  var mbSheet  = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
  if (!logSheet) return { missions: [], guilds: [] };

  // 미션 메타 (표시 이름 + 별 등급)
  var missionMeta = [
    { missionId: "M01", missionName: "첫 깃발",           star: 1 },
    { missionId: "M02", missionName: "다섯 개의 다른 길", star: 3 },
    { missionId: "M03", missionName: "길드 정체성 만들기", star: 2 },
    { missionId: "M04", missionName: "우리 길드의 장보기", star: 2 },
    { missionId: "M07", missionName: "미식 탐험가",        star: 3 },
    { missionId: "M12", missionName: "길드 명예의 일격",   star: 4 },
    { missionId: "M14", missionName: "미래에 남기는 편지", star: 2 },
    { missionId: "M15", missionName: "시즌 결산",          star: 1 }
  ];

  // 미션로그에서 길드 × 미션별 상태 추출
  // 컬럼: A=미션ID, B=길드ID, E=상태
  var logData    = logSheet.getDataRange().getValues();
  var statusMap  = {}; // { "GUILD_01|M02": "클리어" }

  for (var i = 1; i < logData.length; i++) {
    var mId    = String(logData[i][0]).trim();
    var gId    = String(logData[i][1]).trim();
    var status = String(logData[i][4]).trim();
    if (mId && gId) statusMap[gId + "|" + mId] = status;
  }

  // 길드별 results 객체 생성
  var guilds = GUILD_IDS.map(function(guildId) {
    var results = {};
    missionMeta.forEach(function(m) {
      results[m.missionId] = statusMap[guildId + "|" + m.missionId] || "대기";
    });
    return {
      guildId   : guildId,
      guildName : _getGuildName(mbSheet, guildId),
      results   : results
    };
  });

  return { missions: missionMeta, guilds: guilds };
}


/**
 * 길드별 세션 출석률을 반환합니다.
 * AuctionAdmin.html의 renderGuildAttendance()에서 호출.
 *
 * @returns {object[]} [
 *   { guildId, guildName, totalSessions, avgAttendanceRate,
 *     members: [ { name, attended, total, rate }, ... ] }, ...
 * ]
 */
function getGuildAttendanceStats() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var mbSheet = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);

  return GUILD_IDS.map(function(guildId) {
    var stats     = getGuildSessionStats(guildId); // 섹션 8 함수 재사용
    var guildName = _getGuildName(mbSheet, guildId);

    if (!stats) {
      return {
        guildId           : guildId,
        guildName         : guildName,
        totalSessions     : 0,
        avgAttendanceRate : 0,
        members           : []
      };
    }

    // members 배열로 변환 (이름순 정렬)
    var memberArr = Object.keys(stats.memberStats).map(function(name) {
      var m = stats.memberStats[name];
      return { name: name, attended: m.attended, total: m.total, rate: m.rate };
    });
    memberArr.sort(function(a, b) { return a.name < b.name ? -1 : 1; });

    return {
      guildId           : guildId,
      guildName         : guildName,
      totalSessions     : stats.totalSessions,
      avgAttendanceRate : stats.avgAttendanceRate,
      members           : memberArr
    };
  });
}


// ── 섹션 10 내부 유틸 ──────────────────────────────────────

/**
 * 길드_구성 시트에서 길드명(B열)을 반환합니다.
 * 없으면 guildId를 그대로 반환.
 */
function _getGuildName(mbSheet, guildId) {
  if (!mbSheet) return guildId;
  var data = mbSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === guildId) {
      var name = String(data[i][1]).trim();
      return name || guildId;
    }
  }
  return guildId;
}

/**
 * "2026-05" → "2026-04" (전월 yyyy-MM 반환)
 */
function _getPrevYearMonth(yearMonth) {
  var parts = yearMonth.split("-");
  var y = parseInt(parts[0]);
  var m = parseInt(parts[1]) - 1;
  if (m === 0) { m = 12; y--; }
  return y + "-" + (m < 10 ? "0" + m : String(m));
}


// ============================================================
// 섹션 11 — Index.html 학생 길드 카드용 GAS 함수
// ============================================================

/**
 * 학생 한 명을 위한 길드 카드 데이터를 반환합니다.
 * Index.html의 loadGuildCardForStudent()에서 호출.
 *
 * 학생 인터페이스 비공개 원칙:
 *   - 다른 길드의 GS 점수는 절대 노출 안 됨 (등수만)
 *   - 길드 멤버의 자산·브랜드가치 노출 안 됨 (이름 + 포트만)
 *
 * @param {string} studentName
 * @returns {object} {
 *   hasGuild: boolean,
 *   guildId, guildName, slogan,
 *   members: [ { name, port } ],
 *   myRank,
 *   missionResults: { M01: '클리어'|...|'대기' },
 *   missionClearedCount,
 *   allRanks: [ { guildName, rank } ]   // 점수 없음
 * }
 */
function getGuildCardDataForStudent(studentName) {
  if (!studentName) return { hasGuild: false };

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var mbSheet = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
  if (!mbSheet) return { hasGuild: false };

  // 학생의 소속 길드 찾기
  // 길드_구성: A=길드ID, B=길드명, C=멤버명, D=포트, E=가입일, F=탈퇴일, G=비고
  var mbData      = mbSheet.getDataRange().getValues();
  var myGuildId   = null;
  var myGuildName = null;
  var mySlogan    = "";

  for (var i = 1; i < mbData.length; i++) {
    var rowName    = String(mbData[i][2]).trim();
    var leaveDate  = mbData[i][5];
    if (rowName === studentName && (!leaveDate || leaveDate === "")) {
      myGuildId   = String(mbData[i][0]).trim();
      myGuildName = String(mbData[i][1]).trim();
      // G열(비고)에 슬로건이 있으면 활용 (선택)
      mySlogan = String(mbData[i][6] || "").trim();
      break;
    }
  }

  if (!myGuildId) return { hasGuild: false };

  // 같은 길드의 모든 활성 멤버 (이름 + 포트만)
  var members = [];
  for (var j = 1; j < mbData.length; j++) {
    var gId       = String(mbData[j][0]).trim();
    var leaveD    = mbData[j][5];
    if (gId !== myGuildId) continue;
    if (leaveD && leaveD !== "") continue;
    var name = String(mbData[j][2]).trim();
    var port = String(mbData[j][3]).trim();
    if (name) members.push({ name: name, port: port });
  }

  // ── 길드원 장착 캐릭터 조회 ─────────────────────────────────────
  // 상점_아이템 + 상점_구매로그를 한 번씩만 읽어서 멤버별 캐릭터(이모지/URL) 매핑
  try {
    var shopItemSheet = ss.getSheetByName(SHEET_SHOP_ITEMS);
    var shopLogSheet  = ss.getSheetByName(SHEET_SHOP_LOG);
    if (shopItemSheet && shopLogSheet) {
      // 캐릭터 아이템 ID → resourceVal(이모지 또는 이미지 URL) 매핑
      var charResourceMap = {};
      var iData = shopItemSheet.getDataRange().getValues();
      for (var ci = 1; ci < iData.length; ci++) {
        if (String(iData[ci][1]).trim() === '캐릭터') {
          charResourceMap[String(iData[ci][0]).trim()] = String(iData[ci][7]).trim();
        }
      }
      // 구매로그에서 현재 장착(TRUE) 상태인 캐릭터 찾기 (나중 행이 최신)
      var memberCharMap = {};
      var slData = shopLogSheet.getDataRange().getValues();
      for (var sl = 1; sl < slData.length; sl++) {
        var slName   = String(slData[sl][1]).trim();
        var isEq     = (slData[sl][6] === true) || (String(slData[sl][6]).toUpperCase() === 'TRUE');
        var slItemId = String(slData[sl][2]).trim();
        if (isEq && charResourceMap[slItemId]) {
          memberCharMap[slName] = charResourceMap[slItemId];
        }
      }
      // members 배열에 charVal 필드 추가
      members = members.map(function(m) {
        return { name: m.name, port: m.port, charVal: memberCharMap[m.name] || '' };
      });
    }
  } catch (charErr) {
    Logger.log('[getGuildCardDataForStudent] 캐릭터 조회 오류: ' + charErr.message);
    // 오류 시 charVal 없이 기존 동작 유지
  }

  // 내 길드 미션 클리어 현황
  var logSheet  = ss.getSheetByName(SHEET_NAMES.MISSION_LOG);
  var results   = {};
  var clearedCt = 0;

  if (logSheet) {
    var logData = logSheet.getDataRange().getValues();
    for (var k = 1; k < logData.length; k++) {
      var lmId    = String(logData[k][0]).trim();
      var lgId    = String(logData[k][1]).trim();
      var lstatus = String(logData[k][4]).trim();
      if (lgId === myGuildId && lmId) {
        results[lmId] = lstatus || "대기";
        if (lstatus === "클리어") clearedCt++;
      }
    }
  }

  // 전체 길드 순위 (이번 달 GS, 점수 제외 — 길드명 + 등수만)
  var gsSheet   = ss.getSheetByName(SHEET_NAMES.GS_MONTHLY);
  var allRanks  = [];
  var myRank    = null;
  var myGS      = null;

  if (gsSheet) {
    var now       = new Date();
    var yearMonth = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM");
    var gsData    = gsSheet.getDataRange().getValues();
    var rawRanks  = []; // 이번 달 행 추출

    for (var x = 1; x < gsData.length; x++) {
      var cellVal = gsData[x][0];
      var cellYM  = cellVal instanceof Date
        ? Utilities.formatDate(cellVal, "Asia/Seoul", "yyyy-MM")
        : String(cellVal).trim().substring(0, 7);
      if (cellYM !== yearMonth) continue;
      rawRanks.push({
        guildId : String(gsData[x][1]).trim(),
        rank    : parseInt(gsData[x][13]) || 0,   // N열
        totalGS : parseFloat(gsData[x][12]) || 0  // M열
      });
    }

    // 이번 달 데이터 없으면 전월로 fallback
    if (rawRanks.length === 0) {
      var prev = _getPrevYearMonth(yearMonth);
      for (var y = 1; y < gsData.length; y++) {
        if (String(gsData[y][0]).trim() !== prev) continue;
        rawRanks.push({
          guildId : String(gsData[y][1]).trim(),
          rank    : parseInt(gsData[y][18]) || 0,   // S열
          totalGS : parseFloat(gsData[y][17]) || 0  // R열
        });
      }
    }

    // 정렬 + 길드명 부여 (내 길드 GS만 반환, 타 길드 점수 비공개)
    rawRanks.sort(function(a, b) { return a.rank - b.rank; });
    rawRanks.forEach(function(r) {
      var gName = _getGuildName(mbSheet, r.guildId);
      allRanks.push({ guildName: gName, rank: r.rank });
      if (r.guildId === myGuildId) {
        myRank = r.rank;
        myGS   = r.totalGS;
      }
    });
  }

  return {
    hasGuild            : true,
    guildId             : myGuildId,
    guildName           : myGuildName || myGuildId,
    slogan              : mySlogan,
    members             : members,
    myRank              : myRank,
    myGS                : myGS,
    missionResults      : results,
    missionClearedCount : clearedCt,
    allRanks            : allRanks
  };
}


// ============================================================
// 섹션 12 — AuctionAdmin 길드 모달용 UI 래퍼 함수
// ============================================================
//
// 기존 섹션 2·4·8의 핵심 함수들은 throw 또는 toast로 결과를 알리기 때문에,
// 모달 UI에서 깔끔하게 success/msg 형식으로 받기 위해 래퍼를 둡니다.
// 기존 함수는 그대로 유지 (편집기에서 직접 호출하는 용도).
//

/**
 * 미션 발표 UI 래퍼.
 * @returns {object} { success: boolean, msg: string }
 */
function announceGuildMissionUI(missionId, startDate, endDate) {
  try {
    announceGuildMission(missionId, startDate, endDate);
    return { success: true, msg: missionId + " 발표 완료. " + GUILD_IDS.length + "개 길드 행 생성됨." };
  } catch (e) {
    return { success: false, msg: e.message || String(e) };
  }
}

/**
 * 정성평가 저장 UI 래퍼.
 * @returns {object} { success: boolean, msg: string }
 */
function saveGuildManualScoreUI(missionId, guildId, score, comment) {
  try {
    setGuildMissionManualScore(missionId, guildId, score, comment || "");
    return { success: true, msg: guildId + " " + missionId + " 정성평가 " + score + "점 저장 완료." };
  } catch (e) {
    return { success: false, msg: e.message || String(e) };
  }
}

/**
 * 세션 출석 입력용 — 길드 멤버 이름 배열 반환.
 * 모달에서 길드 선택 시 호출.
 * @returns {string[]}
 */
function getGuildMembersForSession(guildId) {
  return _getGuildMembers(guildId); // 섹션 9 유틸 재사용
}

/**
 * 세션 출석 저장 UI 래퍼.
 * @param {string} guildId
 * @param {string} sessionDate  YYYY-MM-DD
 * @param {object} attendanceMap  { 학생명: "O" | "X" }
 * @returns {object} { success: boolean, msg: string }
 */
function saveGuildSessionAttendanceUI(guildId, sessionDate, attendanceMap) {
  try {
    recordGuildSession(guildId, sessionDate, attendanceMap);
    var n = Object.keys(attendanceMap).length;
    return { success: true, msg: guildId + " " + sessionDate + " 출석 " + n + "명 저장 완료." };
  } catch (e) {
    return { success: false, msg: e.message || String(e) };
  }
}

/**
 * 미션 참여 기록 저장 — AuctionAdmin 탭 5에서 호출
 * 길드활동로그 시트에 개인별 O/X 기록
 *
 * @param {string} guildId
 * @param {string} missionId   - M01~M15
 * @param {string} date        - YYYY-MM-DD
 * @param {object} attendanceMap - { 학생명: "O" | "X" }
 * @returns {object} { success, msg }
 */
function saveMissionParticipationUI(guildId, missionId, date, attendanceMap) {
  try {
    if (!guildId || !missionId || !date || !attendanceMap) {
      throw new Error("필수 파라미터가 누락되었습니다.");
    }

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var sheet    = ss.getSheetByName("길드활동로그");
    if (!sheet) throw new Error("길드활동로그 시트를 찾을 수 없습니다. 시트를 먼저 생성해주세요.");

    var dateObj   = new Date(date);
    var inputTime = new Date();

    Object.keys(attendanceMap).forEach(function(studentName) {
      sheet.appendRow([
        dateObj,                // A: 날짜
        "미션",                 // B: 유형
        missionId,              // C: 미션ID
        studentName,            // D: 학생명
        attendanceMap[studentName] // E: O/X
      ]);
    });

    var n = Object.keys(attendanceMap).length;
    Logger.log("[saveMissionParticipationUI] " + guildId + " / " + missionId + " / " + date + " → " + n + "명 기록.");
    return { success: true, msg: guildId + " " + missionId + " " + date + " 참여 " + n + "명 저장 완료." };
  } catch (e) {
    return { success: false, msg: e.message || String(e) };
  }
}


// ============================================================
// 섹션 13 — 동료 평가 시스템
// ============================================================

/**
 * 동료 평가를 제출합니다.
 * 같은 미션에서 같은 평가자→피평가자 조합은 중복 제출 불가.
 *
 * @param {string} missionId    - 미션 코드 또는 "PROJECT"
 * @param {string} evaluatorName - 평가자 학생명
 * @param {string} targetName    - 피평가자 학생명
 * @param {number} score         - 1~10 정수
 * @param {string} comment       - 후기 (자유 텍스트)
 */
function submitPeerEval(missionId, evaluatorName, targetName, score, comment) {
  if (!missionId || !evaluatorName || !targetName || score === undefined) {
    throw new Error("submitPeerEval: 필수 파라미터가 누락되었습니다.");
  }
  if (evaluatorName === targetName) {
    throw new Error("자기 자신은 평가할 수 없습니다.");
  }
  score = parseInt(score);
  if (isNaN(score) || score < 1 || score > 10) {
    throw new Error("점수는 1~10 사이의 정수여야 합니다.");
  }

  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var evalSheet  = ss.getSheetByName("길드_동료평가");
  var mbSheet    = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);

  if (!evalSheet) throw new Error("길드_동료평가 시트를 찾을 수 없습니다.");

  // 평가자의 길드 확인
  var studentGuildMap = _buildStudentGuildMap(mbSheet);
  var guildId         = studentGuildMap[evaluatorName];
  if (!guildId) throw new Error("평가자의 소속 길드를 찾을 수 없습니다.");

  // 피평가자도 같은 길드인지 확인
  if (studentGuildMap[targetName] !== guildId) {
    throw new Error("같은 길드원만 평가할 수 있습니다.");
  }

  // 중복 제출 방지 (A+C+D 조합)
  var existing = evalSheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim() === missionId &&
        String(existing[i][2]).trim() === evaluatorName &&
        String(existing[i][3]).trim() === targetName) {
      throw new Error("이미 이 미션에서 해당 길드원을 평가했습니다.");
    }
  }

  // 기록
  evalSheet.appendRow([
    missionId,                                                         // A: 미션ID
    guildId,                                                           // B: 길드ID
    evaluatorName,                                                     // C: 평가자
    targetName,                                                        // D: 피평가자
    score,                                                             // E: 점수
    comment || "",                                                     // F: 후기
    Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"), // G: 작성일시
    false                                                              // H: 공개여부 (FALSE)
  ]);

  // 전원 작성 완료 여부 체크 → 완료 시 자동 공개
  _checkAndRevealPeerEval(missionId, guildId, evalSheet, mbSheet);

  Logger.log("[submitPeerEval] " + missionId + " / " + evaluatorName + " → " + targetName + " : " + score + "점");
}


/**
 * 해당 미션+길드의 전원 작성 완료 여부를 확인하고,
 * 완료 시 H열(공개여부)을 TRUE로 일괄 전환합니다.
 *
 * 완료 조건: 길드 내 N명이 각자 (N-1)명을 평가했을 때
 */
function _checkAndRevealPeerEval(missionId, guildId, evalSheet, mbSheet) {
  var members    = _getGuildMembers(guildId);      // 활성 멤버 목록
  var memberCount = members.length;
  if (memberCount < 2) return;

  var required   = memberCount * (memberCount - 1); // 총 필요 평가 수
  var data       = evalSheet.getDataRange().getValues();

  // 이 미션+길드의 현재 제출 건수 카운트 (공개 여부 무관)
  var submitted  = 0;
  var targetRows = []; // 행 번호 (1-based)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === missionId &&
        String(data[i][1]).trim() === guildId) {
      submitted++;
      targetRows.push(i + 1);
    }
  }

  if (submitted >= required) {
    // 전원 완료 → H열 TRUE로 일괄 전환
    targetRows.forEach(function(rowNum) {
      evalSheet.getRange(rowNum, 8).setValue(true);
    });
    Logger.log("[_checkAndRevealPeerEval] " + missionId + "/" + guildId + " 전원 완료 → 공개 전환");
  }
}


/**
 * 특정 학생이 특정 미션에서 받은 동료 평가 결과를 반환합니다.
 * H열=TRUE(공개)인 행만 반환 → 평가자 이름은 포함하지 않음(익명).
 *
 * @param {string} studentName
 * @param {string} missionId
 * @returns {object} { revealed: boolean, myAvg: number|null, guildAvg: number|null, reviews: string[] }
 */
function getPeerEvalResult(studentName, missionId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var evalSheet = ss.getSheetByName("길드_동료평가");
  if (!evalSheet) return { revealed: false, myAvg: null, guildAvg: null, reviews: [] };

  var mbSheet         = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
  var studentGuildMap = _buildStudentGuildMap(mbSheet);
  var guildId         = studentGuildMap[studentName];
  if (!guildId) return { revealed: false, myAvg: null, guildAvg: null, reviews: [] };

  var data         = evalSheet.getDataRange().getValues();
  var myScores     = [];
  var myReviews    = [];
  var guildScores  = []; // 이 미션에서 길드 전체 점수 (공개된 것만)
  var isRevealed   = false;

  for (var i = 1; i < data.length; i++) {
    var rowMission = String(data[i][0]).trim();
    var rowGuild   = String(data[i][1]).trim();
    var rowTarget  = String(data[i][3]).trim();
    var rowScore   = parseFloat(data[i][4]) || 0;
    var rowComment = String(data[i][5]).trim();
    var rowPublic  = data[i][7]; // H열: 공개여부

    if (rowMission !== missionId || rowGuild !== guildId) continue;
    if (!rowPublic) continue; // 아직 비공개

    isRevealed = true;
    guildScores.push(rowScore);

    if (rowTarget === studentName) {
      myScores.push(rowScore);
      if (rowComment) myReviews.push(rowComment);
    }
  }

  var myAvg    = myScores.length    > 0 ? Math.round((_sum(myScores)    / myScores.length)    * 10) / 10 : null;
  var guildAvg = guildScores.length > 0 ? Math.round((_sum(guildScores) / guildScores.length) * 10) / 10 : null;

  return {
    revealed  : isRevealed,
    myAvg     : myAvg,
    guildAvg  : guildAvg,
    reviews   : myReviews
  };
}


/**
 * 특정 학생이 제출해야 할 동료 평가 목록과 이미 제출한 목록을 반환합니다.
 * (Index 대시보드에서 "아직 평가 안 한 항목" 표시용)
 *
 * @param {string} studentName
 * @param {string} missionId
 * @returns {object} { pending: string[], done: string[] }  // 피평가자 이름 배열
 */
function getPeerEvalStatus(studentName, missionId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mbSheet   = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);
  var evalSheet = ss.getSheetByName("길드_동료평가");

  var studentGuildMap = _buildStudentGuildMap(mbSheet);
  var guildId         = studentGuildMap[studentName];
  if (!guildId) return { pending: [], done: [] };

  var members = _getGuildMembers(guildId).filter(function(n) { return n !== studentName; });

  var done = [];
  if (evalSheet) {
    var data = evalSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === missionId &&
          String(data[i][2]).trim() === studentName) {
        done.push(String(data[i][3]).trim());
      }
    }
  }

  var pending = members.filter(function(n) { return done.indexOf(n) === -1; });
  return { pending: pending, done: done };
}


/**
 * 학생 대시보드용: 모든 미션(+PROJECT)의 동료 평가 결과를 한 번에 반환합니다.
 *
 * @param {string} studentName
 * @returns {object} { missionId: { revealed, myAvg, guildAvg, reviews }, ... }
 */
function getAllPeerEvalResults(studentName) {
  var missionIds = Object.keys(MISSION_WEIGHTS).concat(["PROJECT"]);
  var result     = {};
  missionIds.forEach(function(mid) {
    result[mid] = getPeerEvalResult(studentName, mid);
  });
  return result;
}


/**
 * UI 래퍼: Index.html에서 동료 평가 제출 시 호출
 * @returns {object} { success: boolean, msg: string }
 */
function submitPeerEvalUI(missionId, targetName, score, comment) {
  try {
    var studentName = Session.getActiveUser().getEmail(); // 필요 시 학생명으로 교체
    // ※ B.R.A.N.D는 학생명 기반이므로 실제로는 getStudentData()로 이름을 가져와야 합니다.
    // Index.html에서 studentName을 파라미터로 넘기는 방식 권장.
    submitPeerEval(missionId, studentName, targetName, score, comment);
    return { success: true, msg: "평가가 제출되었습니다." };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}

/**
 * UI 래퍼 (학생명 명시 버전) — Index.html에서 호출
 */
function submitPeerEvalByName(missionId, evaluatorName, targetName, score, comment) {
  try {
    submitPeerEval(missionId, evaluatorName, targetName, score, comment);
    return { success: true, msg: "평가가 제출되었습니다." };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}

/**
 * UI 래퍼: 학생 대시보드에서 동료 평가 결과 조회
 */
function getAllPeerEvalResultsUI(studentName) {
  try {
    return { success: true, data: getAllPeerEvalResults(studentName) };
  } catch (e) {
    return { success: false, msg: e.message, data: {} };
  }
}

/**
 * UI 래퍼: 특정 미션의 평가 현황 조회 (누구를 아직 안 했는지)
 */
function getPeerEvalStatusUI(studentName, missionId) {
  try {
    return { success: true, data: getPeerEvalStatus(studentName, missionId) };
  } catch (e) {
    return { success: false, msg: e.message, data: { pending: [], done: [] } };
  }
}


// ============================================================
// 헬퍼: 학생명 → 길드ID 매핑 캐시 생성
// ============================================================

/**
 * 길드_구성 시트를 읽어 { 학생명: 길드ID } 객체를 반환합니다.
 * 탈퇴일이 있는 행은 제외.
 */
function _buildStudentGuildMap(mbSheet) {
  var map = {};
  if (!mbSheet) return map;
  var data = mbSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var gId      = String(data[i][0]).trim();
    var name     = String(data[i][2]).trim();
    var leaveDate = data[i][5];
    if (name && gId && (!leaveDate || leaveDate === "")) {
      map[name] = gId;
    }
  }
  return map;
}

// ============================================================
// 섹션 14 — 개인 GS 기여점수 산출
// ============================================================
//
// 목적: 월별로 각 학생이 길드 GS에 얼마나 기여했는지 수치화하여
//       길드 내 1~N위를 가릴 수 있도록 '길드_개인기여도' 시트에 기록.
//
// 공식 (정규화 기준: 길드 소속 전체 학생 중 1등):
//   α_개인  = (내 브랜드가치 증가량  / 전체 최대 증가량)  × 0.50
//   참여점수 = (내 미션참여 횟수     / 전체 최다 참여횟수) × 0.15
//   출석점수 = (내 세션출석 횟수     / 전체 최다 출석횟수) × 0.10
//   합계     = α_개인 + 참여점수 + 출석점수  (최대 0.75)
//
// 길드내순위: 같은 길드원 사이에서 합계 기준 내림차순 (1위가 최고)
//
// 자동 실행: calcMonthlyGS() 내부에서 월간 GS 산출 직후 호출됨
// 수동 실행: runIndividualGSManually() 함수 직접 실행
// ============================================================

/**
 * 개인 GS 기여점수를 계산하여 '길드_개인기여도' 시트에 기록합니다.
 * calcMonthlyGS() 끝에서 자동 호출됩니다.
 */
function calcMonthlyIndividualGS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var mainSheet  = ss.getSheetByName(SHEET_NAMES.MAIN);
  var trackSheet = ss.getSheetByName('브랜드가치추적');
  var actSheet   = ss.getSheetByName('길드활동로그');
  var sesSheet   = ss.getSheetByName('길드_세션출석');
  var mbSheet    = ss.getSheetByName(SHEET_NAMES.GUILD_MEMBERS);

  if (!mainSheet || !mbSheet) {
    Logger.log('[calcMonthlyIndividualGS] 필수 시트(메인/길드_구성) 없음. 종료.');
    return;
  }

  var now       = new Date();
  var yearMonth = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM');

  // ── 시트 준비: 없으면 자동 생성 ──────────────────────────────────
  var indSheet = ss.getSheetByName('길드_개인기여도');
  if (!indSheet) {
    indSheet = ss.insertSheet('길드_개인기여도');
    indSheet.appendRow([
      '월', '길드ID', '학생명',
      '브랜드가치증가량', '미션참여횟수', '세션출석횟수',
      'α점수', '참여점수', '출석점수',
      '기여점수합계', '길드내순위'
    ]);
    indSheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    Logger.log('[calcMonthlyIndividualGS] 길드_개인기여도 시트 자동 생성.');
  }

  // ── 이번 달 중복 방지 ─────────────────────────────────────────────
  var existingData = indSheet.getDataRange().getValues();
  for (var ex = 1; ex < existingData.length; ex++) {
    var cellVal = existingData[ex][0];
    var cellYM  = cellVal instanceof Date
      ? Utilities.formatDate(cellVal, 'Asia/Seoul', 'yyyy-MM')
      : String(cellVal).trim().substring(0, 7);
    if (cellYM === yearMonth) {
      Logger.log('[calcMonthlyIndividualGS] 이번 달 이미 계산됨: ' + yearMonth);
      return;
    }
  }

  // ── 1. 전체 활성 길드원 목록 수집 ────────────────────────────────
  // { 학생명: 길드ID }
  var studentGuildMap = _buildStudentGuildMap(mbSheet);
  var allStudents     = Object.keys(studentGuildMap);

  if (allStudents.length === 0) {
    Logger.log('[calcMonthlyIndividualGS] 활성 길드원 없음. 종료.');
    return;
  }

  // ── 2. 학생별 브랜드가치 증가량 계산 ─────────────────────────────
  // 현재값: 메인 시트 C열(브랜드가치)
  var mainData  = mainSheet.getDataRange().getValues();
  var curValMap = {}; // { 학생명: 현재 브랜드가치 }
  for (var m = 1; m < mainData.length; m++) {
    var mName = String(mainData[m][COL_NAME  - 1]).trim();
    var mVal  = parseFloat(mainData[m][COL_VALUE - 1]) || 0;
    if (mName) curValMap[mName] = mVal;
  }

  // 월초값: 브랜드가치추적 시트에서 해당 월 첫 번째 날짜 컬럼
  var startValMap = {}; // { 학생명: 월초 브랜드가치 }
  if (trackSheet) {
    var trackData    = trackSheet.getDataRange().getValues();
    var trackHeaders = trackData[0];
    var monthStart   = new Date(now.getFullYear(), now.getMonth(), 1);
    var monthEnd     = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    var startColIdx  = -1;
    var closestDate  = null;

    for (var h = 1; h < trackHeaders.length; h++) {
      var hDate = new Date(trackHeaders[h]);
      if (isNaN(hDate.getTime())) continue;
      hDate.setHours(0, 0, 0, 0);
      if (hDate >= monthStart && hDate <= monthEnd) {
        if (closestDate === null || hDate < closestDate) {
          closestDate = hDate;
          startColIdx = h;
        }
      }
    }

    if (startColIdx !== -1) {
      for (var t = 1; t < trackData.length; t++) {
        var tName = String(trackData[t][0]).trim();
        var tVal  = parseFloat(trackData[t][startColIdx]) || 0;
        if (tName) startValMap[tName] = tVal;
      }
    }
  }

  // 증가량 맵: 음수(자산 감소)는 0으로 처리
  var growthMap = {};
  allStudents.forEach(function(name) {
    var cur   = curValMap[name]   || 0;
    var start = startValMap[name] !== undefined ? startValMap[name] : cur;
    growthMap[name] = Math.max(0, cur - start);
  });

  // ── 3. 학생별 미션 참여 횟수 집계 (이번 달만) ─────────────────────
  // 길드활동로그: A=날짜, B=유형('미션'), C=미션ID, D=학생명, E=참여여부('O'/'X')
  var missionCntMap = {};
  allStudents.forEach(function(name) { missionCntMap[name] = 0; });

  if (actSheet) {
    var actData     = actSheet.getDataRange().getValues();
    var actMStart   = new Date(now.getFullYear(), now.getMonth(), 1);
    var actMEnd     = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    actMEnd.setHours(23, 59, 59, 999);

    for (var a = 1; a < actData.length; a++) {
      var aDate   = new Date(actData[a][0]);
      var aType   = String(actData[a][1]).trim();
      var aName   = String(actData[a][3]).trim();
      var aResult = String(actData[a][4]).trim();

      if (isNaN(aDate.getTime()) || aDate < actMStart || aDate > actMEnd) continue;
      if (aType !== '미션') continue;
      if (aResult !== 'O') continue;
      if (!missionCntMap.hasOwnProperty(aName)) continue;
      missionCntMap[aName]++;
    }
  }

  // ── 4. 학생별 세션 출석 횟수 집계 (이번 달만) ─────────────────────
  // 길드_세션출석: A=날짜, B=길드ID, C=학생명, D=출석여부('O'/'X')
  var sessionCntMap = {};
  allStudents.forEach(function(name) { sessionCntMap[name] = 0; });

  if (sesSheet) {
    var sesData   = sesSheet.getDataRange().getValues();
    var sesMStart = new Date(now.getFullYear(), now.getMonth(), 1);
    var sesMEnd   = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    sesMEnd.setHours(23, 59, 59, 999);

    for (var s = 1; s < sesData.length; s++) {
      var sDate   = new Date(sesData[s][0]);
      var sName   = String(sesData[s][2]).trim();
      var sResult = String(sesData[s][3]).trim();

      if (isNaN(sDate.getTime()) || sDate < sesMStart || sDate > sesMEnd) continue;
      if (sResult !== 'O') continue;
      if (!sessionCntMap.hasOwnProperty(sName)) continue;
      sessionCntMap[sName]++;
    }
  }

  // ── 5. 전체 학생 기준 최대값 (정규화 분모) ───────────────────────
  var maxGrowth  = Math.max.apply(null, allStudents.map(function(n) { return growthMap[n]     || 0; }));
  var maxMission = Math.max.apply(null, allStudents.map(function(n) { return missionCntMap[n] || 0; }));
  var maxSession = Math.max.apply(null, allStudents.map(function(n) { return sessionCntMap[n] || 0; }));

  Logger.log('[calcMonthlyIndividualGS] 정규화 기준 — 최대증가량: ' + maxGrowth +
             ', 최다미션참여: ' + maxMission + ', 최다세션출석: ' + maxSession);

  // ── 6. 학생별 점수 계산 ───────────────────────────────────────────
  var studentScores = [];

  allStudents.forEach(function(name) {
    var growth  = growthMap[name]     || 0;
    var mCnt    = missionCntMap[name] || 0;
    var sCnt    = sessionCntMap[name] || 0;
    var guildId = studentGuildMap[name];

    // 각 항목 정규화 후 가중치 적용. 최대값=0이면 해당 항목 0점 처리.
    var alphaScore  = maxGrowth  > 0 ? (growth / maxGrowth)  * GS_ALPHA                : 0;
    var partScore   = maxMission > 0 ? (mCnt   / maxMission) * GS_MISSION_PARTICIPATION : 0;
    var attendScore = maxSession > 0 ? (sCnt   / maxSession) * GS_SESSION_ATTENDANCE   : 0;
    var totalScore  = alphaScore + partScore + attendScore;

    studentScores.push({
      guildId     : guildId,
      name        : name,
      growth      : Math.round(growth),
      mCnt        : mCnt,
      sCnt        : sCnt,
      alphaScore  : Math.round(alphaScore  * 1000) / 1000,
      partScore   : Math.round(partScore   * 1000) / 1000,
      attendScore : Math.round(attendScore * 1000) / 1000,
      totalScore  : Math.round(totalScore  * 1000) / 1000,
      guildRank   : 0  // 다음 단계에서 계산
    });
  });

  // ── 7. 길드 내 순위 계산 ─────────────────────────────────────────
  // 동점자는 같은 순위 부여 (공동 n위)
  GUILD_IDS.forEach(function(guildId) {
    var guildMembers = studentScores.filter(function(s) { return s.guildId === guildId; });
    guildMembers.sort(function(a, b) { return b.totalScore - a.totalScore; });

    var prevScore = null;
    var prevRank  = 0;
    guildMembers.forEach(function(s, idx) {
      if (s.totalScore !== prevScore) {
        prevRank  = idx + 1;
        prevScore = s.totalScore;
      }
      s.guildRank = prevRank;
    });
  });

  // ── 8. 시트에 기록 (길드ID → 길드내순위 순 정렬) ──────────────────
  studentScores.sort(function(a, b) {
    if (a.guildId < b.guildId) return -1;
    if (a.guildId > b.guildId) return  1;
    return a.guildRank - b.guildRank;
  });

  studentScores.forEach(function(s) {
    indSheet.appendRow([
      yearMonth,    // A: 월
      s.guildId,    // B: 길드ID
      s.name,       // C: 학생명
      s.growth,     // D: 브랜드가치증가량
      s.mCnt,       // E: 미션참여횟수
      s.sCnt,       // F: 세션출석횟수
      s.alphaScore, // G: α점수 (최대 0.50)
      s.partScore,  // H: 참여점수 (최대 0.15)
      s.attendScore,// I: 출석점수 (최대 0.10)
      s.totalScore, // J: 기여점수합계 (최대 0.75)
      s.guildRank   // K: 길드내순위
    ]);
  });

  Logger.log('[calcMonthlyIndividualGS] ' + yearMonth +
             ' 개인 기여점수 산출 완료. ' + studentScores.length + '명 처리.');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    yearMonth + ' 개인 기여점수 산출 완료 (' + studentScores.length + '명)',
    '개인 GS 기여점수', 4
  );
}


/**
 * 수동으로 즉시 실행할 때 사용합니다.
 * (Apps Script 편집기에서 직접 실행)
 *
 * 사용법:
 *   1. Apps Script 편집기 열기
 *   2. 함수 선택 드롭다운에서 runIndividualGSManually 선택
 *   3. ▶ 실행
 */
function runIndividualGSManually() {
  calcMonthlyIndividualGS();
}


function debugGSSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("길드_GS_월간");
  var data  = sheet.getDataRange().getValues();
  Logger.log("A2값: " + JSON.stringify(data[1][0]));
  Logger.log("M2값: " + JSON.stringify(data[1][12]));
  Logger.log("N2값: " + JSON.stringify(data[1][13]));
  Logger.log("yearMonth비교: " + Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM"));
}
