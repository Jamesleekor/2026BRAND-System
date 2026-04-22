// ── 신용점수이력 시트가 없으면 자동 생성 ─────────────────────────
function _ensureCreditHistorySheet(ss) {
  let sheet = ss.getSheetByName(SHEET_CREDIT_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CREDIT_HISTORY);
    sheet.appendRow(['기준일', '학생명', '트랙A', '투자성실도', '사회기여', '규범준수', '총점', '등급']);
  }
  return sheet;
}

// ── 신용점수 계산 (개별 학생) ────────────────────────────────────────
function calcCreditScore(studentName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 메인 시트: 내 브랜드가치 + 학급 최고 브랜드가치
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const mainData  = mainSheet ? mainSheet.getDataRange().getValues() : [];
  let myHonor = 0, maxHonor = 0;
  for (let i = 1; i < mainData.length; i++) {
    const h = Number(mainData[i][COL_VALUE - 1]) || 0;
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) myHonor = h;
    if (h > maxHonor) maxHonor = h;
  }
  const trackA = maxHonor > 0 ? Math.floor((myHonor / maxHonor) * 400) : 0;

  // 예금 시트: 투자 성실도 (최대 200점)
  const depositSheet = ss.getSheetByName(SHEET_DEPOSIT_LOG);
  const depositData  = depositSheet ? depositSheet.getDataRange().getValues() : [];
  const todayMs      = new Date().setHours(0, 0, 0, 0);
  let totalDepositDays = 0;
  for (let i = 1; i < depositData.length; i++) {
    if (String(depositData[i][1]).trim() !== String(studentName).trim()) continue;
    const status = String(depositData[i][7]).trim();
    const weeks  = Number(depositData[i][4]) || 0;
    if (status === '진행중') {
      const startMs = new Date(String(depositData[i][5]).trim()).setHours(0, 0, 0, 0);
      totalDepositDays += Math.max(0, Math.floor((todayMs - startMs) / 86400000));
    } else if (status === '만기' || status === '중도해지') {
      totalDepositDays += weeks * 7;
    }
  }
  const invest = Math.min(200, Math.floor((totalDepositDays / 28) * 200));

  // 자산사용 시트: 기부 점수 (최대 100점)
  const spendSheet = ss.getSheetByName(SHEET_SPEND);
  const spendData  = spendSheet ? spendSheet.getDataRange().getValues() : [];
  const donationMap = {};
  for (let i = 1; i < spendData.length; i++) {
    if (String(spendData[i][3]).trim() !== '기부') continue;
    const nm = String(spendData[i][1]).trim();
    donationMap[nm] = (donationMap[nm] || 0) + (Number(spendData[i][4]) || 0);
  }
  const myDonation  = donationMap[studentName] || 0;
  const donVals     = Object.values(donationMap);
  const maxDonation = donVals.length > 0 ? Math.max.apply(null, donVals) : 0;
  const donationScore = maxDonation > 0 ? Math.floor((myDonation / maxDonation) * 100) : 50;

  // P2P 거래로그: 평판 점수 (최대 100점) + 이상거래 건수
  const p2pSheet = ss.getSheetByName(SHEET_P2P);
  const p2pData  = p2pSheet ? p2pSheet.getDataRange().getValues() : [];
  let ratingSum = 0, ratingCount = 0, myAnomalyCount = 0;
  for (let i = 1; i < p2pData.length; i++) {
    const receiver = String(p2pData[i][3]).trim();
    const sender   = String(p2pData[i][2]).trim();
    const rating   = Number(p2pData[i][9]) || 0;
    const status   = String(p2pData[i][7]).trim();
    if (receiver === String(studentName).trim() && rating > 0) {
      ratingSum += rating; ratingCount++;
    }
    if ((sender === String(studentName).trim() || receiver === String(studentName).trim()) &&
        (status === '이상거래' || status === '최종적발')) {
      myAnomalyCount++;
    }
  }
  const ratingScore = ratingCount > 0 ? Math.min(100, Math.floor((ratingSum / ratingCount) * 10)) : 50;
  const social = donationScore + ratingScore;

  // 수호대적발로그: 적발 횟수
  const penaltySheet = ss.getSheetByName(SHEET_GUARD_PENALTY);
  const penaltyData  = penaltySheet ? penaltySheet.getDataRange().getValues() : [];
  let myPenaltyCount = 0;
  for (let i = 1; i < penaltyData.length; i++) {
    if (String(penaltyData[i][2]).trim() === String(studentName).trim()) myPenaltyCount++;
  }

  const compliance = Math.max(0, 200 - (myPenaltyCount * 50) - (myAnomalyCount * 20));
  const total      = trackA + invest + social + compliance;
  const gradeInfo  = getCreditGrade(total);

  return {
    trackA, invest, social, compliance, total,
    grade: gradeInfo.grade,
    detail: { donationScore, ratingScore, ratingCount,
              penaltyCount: myPenaltyCount, anomalyCount: myAnomalyCount, totalDepositDays }
  };
}

// ── 점수 → 신용등급 변환 ─────────────────────────────────────────
function getCreditGrade(score) {
  if      (score >= 900) return { grade: 'S',      maxLoanRate: 1.2, weeklyRate: 3  };
  else if (score >= 750) return { grade: 'A+',     maxLoanRate: 1.0, weeklyRate: 4  };
  else if (score >= 600) return { grade: 'A',      maxLoanRate: 0.8, weeklyRate: 5  };
  else if (score >= 450) return { grade: 'B+',     maxLoanRate: 0.6, weeklyRate: 7  };
  else if (score >= 300) return { grade: 'B',      maxLoanRate: 0.4, weeklyRate: 9  };
  else if (score >= 200) return { grade: 'C',      maxLoanRate: 0.2, weeklyRate: 12 };
  else                   return { grade: '대출불가', maxLoanRate: 0,   weeklyRate: 0  };
}

// ── 전체 학생 신용점수 주간 스냅샷 저장 (매주 월요일 오전 트리거) ──
function recordCreditScoreSnapshot() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  const mainData  = mainSheet.getDataRange().getValues();
  const today     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const histSheet = _ensureCreditHistorySheet(ss);

  for (let i = 1; i < mainData.length; i++) {
    const name = String(mainData[i][COL_NAME - 1]).trim();
    if (!name) continue;
    try {
      const result = calcCreditScore(name);
      histSheet.appendRow([
        today,              // A: 기준일
        name,               // B: 학생명
        result.trackA,      // C: 트랙A
        result.invest,      // D: 투자성실도
        result.social,      // E: 사회기여
        result.compliance,  // F: 규범준수
        result.total,       // G: 총점
        result.grade        // H: 등급
      ]);
    } catch(e) {
      Logger.log('신용점수 스냅샷 오류 - ' + name + ': ' + e.message);
    }
  }
}

// ── 대시보드용: 학생 신용점수 + 최근 7일 변동 반환 ─────────────────
function getCreditScoreForStudent(studentName) {
  const current = calcCreditScore(studentName);

  let delta   = null;
  let history = [];
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const histSheet = ss.getSheetByName(SHEET_CREDIT_HISTORY);
  if (histSheet) {
    const histData       = histSheet.getDataRange().getValues();
    const sevenDaysAgo   = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const sevenDaysAgoStr = Utilities.formatDate(sevenDaysAgo, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    let oldScore = null;
    for (let i = histData.length - 1; i >= 1; i--) {
      if (String(histData[i][1]).trim() !== String(studentName).trim()) continue;
      const rowDate = String(histData[i][0]).substring(0, 10);
      history.push({ date: rowDate, total: Number(histData[i][6]) || 0 });
      if (rowDate <= sevenDaysAgoStr && oldScore === null) {
        oldScore = Number(histData[i][6]) || 0;
      }
    }
    history.reverse();
    if (oldScore !== null) delta = current.total - oldScore;
  }

  return {
    score:      current.total,
    grade:      current.grade,
    trackA:     current.trackA,
    invest:     current.invest,
    social:     current.social,
    compliance: current.compliance,
    detail:     current.detail,
    delta,
    history
  };
}
