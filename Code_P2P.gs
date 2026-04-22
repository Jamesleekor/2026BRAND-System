// ════════════════════════════════════════════════════════════════
// ██ P2P 거래 시스템 추가
// 시트: P2P거래로그
//   A=거래ID, B=날짜, C=보내는학생, D=받는학생, E=금액,
//   F=태그, G=거래설명, H=상태(정상/이상거래)
// ════════════════════════════════════════════════════════════════

// ── 거래 가능한 학생 목록 반환 (본인 제외) ───────────────────────
function getP2PReceiverList(studentName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();

  // 2차직업현황 시트에서 승인된 학생 이름 집합 생성
  const approvedSet = new Set();
  const currSheet   = ss.getSheetByName(SHEET_JOB2_CURR);
  if (currSheet) {
    const currData = currSheet.getDataRange().getValues();
    for (let i = 1; i < currData.length; i++) {
      const n = String(currData[i][0]).trim();
      if (n) approvedSet.add(n);
    }
  }

  const result = [];
  for (let i = 1; i < mainData.length; i++) {
    const name  = String(mainData[i][COL_NAME  - 1]).trim();
    const honor = Number(mainData[i][COL_VALUE  - 1]) || 0;
    const brand = String(mainData[i][COL_BRAND  - 1]).trim();

    if (!name) continue;
    if (name === String(studentName).trim()) continue;   // 본인 제외
    if (honor < 20000) continue;                         // 금 광석 미만 제외
    if (!approvedSet.has(name)) continue;                // 2차 직업 미승인 제외

    result.push({ name, brand });
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

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { success: false, msg: '다른 처리 중입니다. 잠시 후 다시 시도해주세요.' }; }
  try {
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
  // ── 자산 동결 체크
  const _emgP = _getActiveEmergency();
  if (_emgP && _emgP.type === '자산 동결') {
    const _usable = Math.floor(senderBalance * (_emgP.freezeRate / 100));
    if (amount > _usable) return { success: false, msg: `🔒 자산 동결 중! 사용 가능 금액: $${_usable.toLocaleString()} (보유액의 ${_emgP.freezeRate}%)` };
  }
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
    const todayStr = _todayStr();
    let todaySameCount = 0;
    // 오늘 동일 sender→receiver 기존 행 번호 수집 (소급 업데이트용)
    const todaySameRows = [];

    for (let i = 1; i < p2pData.length; i++) {
      const rowDate   = String(p2pData[i][1]).substring(0, 10);
      const rowSender = String(p2pData[i][2]).trim();
      const rowRecv   = String(p2pData[i][3]).trim();
      if (rowDate === todayStr &&
          rowSender === String(senderName).trim() &&
          rowRecv   === String(receiverName).trim()) {
        todaySameCount++;
        todaySameRows.push(i + 1); // 시트 행 번호 (1-indexed)
      }
    }

    // 현재 거래 포함 3회 이상이면 이상거래
    // todaySameCount는 이미 시트에 저장된 건수 → +1이 현재 거래
    if (todaySameCount + 1 >= 3) {
      isAnomaly     = true;
      anomalyReason = `오늘 동일인 ${todaySameCount + 1}회 거래`;

      // 기존 행이 '정상'으로 저장된 경우 소급해서 '이상거래'로 업데이트
      // (3번째 거래 발생 시점에 1·2번째 행도 이상거래로 변경)
      for (let r = 0; r < todaySameRows.length; r++) {
        const existingStatus = String(p2pData[todaySameRows[r] - 1][7]).trim();
        if (existingStatus === '정상') {
          p2pSheet.getRange(todaySameRows[r], 8).setValue('이상거래');
        }
      }
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
  } catch(e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ── 나의 P2P 거래 내역 반환 ──────────────────────────────────────
// ██ getMyP2PHistory 함수 교체본
// 기존 함수 전체를 아래로 교체하세요.
// 변경 내용: rating(J열) 필드 추가, canRate 필드 추가
// ════════════════════════════════════════════════════════════════

// ── 나의 P2P 거래 내역 반환 ──────────────────────────────────────
function getMyP2PHistory(studentName) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return [];

  const data   = sheet.getDataRange().getValues();
  const result = [];
  const name   = String(studentName).trim();

  for (let i = 1; i < data.length; i++) {
    const sender   = String(data[i][2]).trim();
    const receiver = String(data[i][3]).trim();
    if (sender !== name && receiver !== name) continue;

    const isSent = sender === name;
    const rating = Number(data[i][9]) || 0;  // J열: 평점 (0=미평가)

    result.push({
      txnId:       String(data[i][0]).trim(),
      date:        String(data[i][1]).substring(0, 10),
      sender:      sender,
      receiver:    receiver,
      amount:      Number(data[i][4]) || 0,
      tag:         String(data[i][5]).trim(),
      description: String(data[i][6]).trim(),
      status:      String(data[i][7]).trim(),
      isSent:      isSent,
      rating:      rating,
      // 평점 가능 여부: sender(서비스 구매자)이고 아직 미평가인 경우만 true
      canRate:     (isSent && rating === 0)
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



// ════════════════════════════════════════════════════════════════
// ██ P2P 거래 평점 시스템
// J열: 0 = 미평가, 1~10 = 평점
// ════════════════════════════════════════════════════════════════

// ── P2P 거래 평점 저장 ───────────────────────────────────────────
// txnId   : 거래ID (A열 값)
// rater   : 평점을 남기는 학생 (반드시 receiver여야 함)
// rating  : 1~10 정수
function rateP2PTransaction(txnId, rater, rating) {
  rating = Number(rating);
  if (!txnId || !rater) return { success: false, msg: '거래 정보가 올바르지 않습니다.' };
  if (!Number.isInteger(rating) || rating < 1 || rating > 10)
    return { success: false, msg: '평점은 1~10 사이의 정수로 입력해주세요.' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return { success: false, msg: 'P2P거래로그 시트를 찾을 수 없습니다.' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowTxnId   = String(data[i][0]).trim();
    const rowSender = String(data[i][2]).trim();  // C열: 보낸 학생 = 서비스 구매자

    if (rowTxnId !== String(txnId).trim()) continue;

    // 평점 권한 확인: 돈을 보낸 쪽(sender = 서비스 구매자)만 평가 가능
    if (rowSender !== String(rater).trim())
      return { success: false, msg: '서비스 구매자(송금한 학생)만 평점을 남길 수 있습니다.' };

    // 이미 평가한 경우 중복 방지
    const existing = Number(data[i][9]) || 0;  // J열: 인덱스 9
    if (existing > 0)
      return { success: false, msg: '이미 평점을 남긴 거래입니다.' };

    // J열(10번째 열)에 평점 저장
    sheet.getRange(i + 1, 10).setValue(rating);
    return { success: true, msg: `⭐ 평점 ${rating}점이 저장되었습니다.` };
  }
  return { success: false, msg: '해당 거래를 찾을 수 없습니다.' };
}
