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
    if (String(data[i][0]).trim() !== String(studentName).trim()) continue;

    // ── 평균 평점 계산 (구매자 5명 이상일 때만 공개) ──────────
    let ratingAvg    = null;
    let ratingCount  = 0;
    const p2pSheet   = ss.getSheetByName(SHEET_P2P);
    if (p2pSheet) {
      const p2pData = p2pSheet.getDataRange().getValues();
      let ratingSum = 0;
      for (let j = 1; j < p2pData.length; j++) {
        // receiver가 본인 = 판매자로서 받은 거래
        if (String(p2pData[j][3]).trim() !== String(studentName).trim()) continue;
        const r = Number(p2pData[j][9]) || 0;  // J열: 평점
        if (r > 0) {
          ratingSum += r;
          ratingCount++;
        }
      }
      // 5명 이상 평가한 경우만 평균 공개
      if (ratingCount >= 5) {
        ratingAvg = Math.round((ratingSum / ratingCount) * 10) / 10; // 소수점 1자리
      }
    }

    return {
      jobName:     String(data[i][1]),
      jobDesc:     String(data[i][2]),
      ratingAvg:   ratingAvg,
      ratingCount: ratingCount
    };
  }
  return null;
}
