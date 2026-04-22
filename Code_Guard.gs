// ════════════════════════════════════════════════════════════════
// ██ 경제 수호대 시스템
// P2P거래로그 시트 I열(수호대메모) 추가 필요
// ════════════════════════════════════════════════════════════════

// ── 수호대 비밀번호 설정 (AuctionAdmin에서 호출) ─────────────────
function setGuardPassword(pw) {
  if (!pw || !String(pw).trim()) return { success: false, msg: '비밀번호를 입력해주세요.' };
  PropertiesService.getScriptProperties().setProperty('GUARD_PASSWORD', String(pw).trim());
  return { success: true, msg: '✅ 수호대 비밀번호가 설정되었습니다.' };
}

// ── 수호대 비밀번호 검증 (GuardDashboard 로그인 시 호출) ─────────
function verifyGuardPassword(pw) {
  const stored = PropertiesService.getScriptProperties().getProperty('GUARD_PASSWORD');
  if (!stored) return { success: false, msg: '비밀번호가 설정되지 않았습니다. 선생님께 문의하세요.' };
  if (String(pw).trim() === stored) return { success: true };
  return { success: false, msg: '비밀번호가 올바르지 않습니다.' };
}

// ── 수호대 대시보드 통합 데이터 반환 ────────────────────────────
// period: 'week'(이번 주) | 'month'(이번 달) | 'all'(전체)
function getGuardDashboardData(period) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const p2pSheet = ss.getSheetByName(SHEET_P2P);
  if (!p2pSheet) return { transactions: [], stats: {}, network: [] };

  const allData = p2pSheet.getDataRange().getValues();

  // ── 기간 필터 기준일 계산 ────────────────────────────────────
  const now   = new Date();
  let cutoff  = null;
  if (period === 'week') {
    const day  = now.getDay(); // 0=일, 1=월
    const diff = (day === 0 ? -6 : 1 - day);
    cutoff = new Date(now);
    cutoff.setDate(now.getDate() + diff);
  } else if (period === 'month') {
    cutoff = new Date(now.getFullYear(), now.getMonth(), 1);
  }
  // cutoffStr은 루프 안에서 매번 계산하므로 여기선 cutoff 객체만 유지

  // ── 거래 데이터 파싱 ─────────────────────────────────────────
  const transactions = [];
  const tagCount     = {};  // 태그별 건수
  const tagAmount    = {};  // 태그별 금액
  const sellerMap    = {};  // 학생별 판매 건수 및 금액 (sender)
  const buyerMap     = {};  // 학생별 구매 건수 (receiver)
  // 네트워크: { "A→B": { from, to, count, total } }
  const edgeMap      = {};

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0]) continue; // 빈 행 스킵

    const dateStr = String(row[1]).substring(0, 10);
    // 기간 필터 적용
    // 기간 필터 적용 (문자열 비교 — timezone 문제 없음)
    if (cutoff) {
      const cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (dateStr < cutoffStr) continue;
    }

    const sender  = String(row[2]).trim();
    const recv    = String(row[3]).trim();
    const amount  = Number(row[4]) || 0;
    const tag     = String(row[5]).trim();
    const desc    = String(row[6]).trim();
    const status  = String(row[7]).trim();
    const memo    = row[8] ? String(row[8]).trim() : '';

    // 이상거래 사유 재계산 (프론트에서 강조 표시용)
    const anomalyReasons = [];
    if (amount >= 2000)        anomalyReasons.push('고액 거래');
    if (desc.length < 10)      anomalyReasons.push('사유 불충분');
    if (tag === '#기타' && desc.length < 20) anomalyReasons.push('태그 불일치 의심');

    transactions.push({
      rowNum:    i + 1,
      txnId:     String(row[0]),
      date:      dateStr,
      sender,
      receiver:  recv,
      amount,
      tag,
      description: desc,
      status,
      memo,
      anomalyReasons  // 빈 배열이면 강조 없음
    });

    // 태그 통계
    tagCount[tag]  = (tagCount[tag]  || 0) + 1;
    tagAmount[tag] = (tagAmount[tag] || 0) + amount;

    // 판매자(sender) 통계
    if (!sellerMap[sender]) sellerMap[sender] = { count: 0, total: 0 };
    sellerMap[sender].count++;
    sellerMap[sender].total += amount;

    // 구매자(receiver) 통계
    if (!buyerMap[recv]) buyerMap[recv] = { count: 0, total: 0 };
    buyerMap[recv].count++;
    buyerMap[recv].total += amount;

    // 네트워크 엣지
    const edgeKey = sender + '→' + recv;
    if (!edgeMap[edgeKey]) edgeMap[edgeKey] = { from: sender, to: recv, count: 0, total: 0 };
    edgeMap[edgeKey].count++;
    edgeMap[edgeKey].total += amount;
  }

  // ── 이번 주 동일인 간 반복 거래 감지 (별도 패스) ─────────────
  // 현재 필터 기간 내 sender+receiver 조합별 건수 집계
  const pairCount = {};
  transactions.forEach(function(tx) {
    const key = tx.sender + '|' + tx.receiver;
    pairCount[key] = (pairCount[key] || 0) + 1;
  });
  // 3회 이상인 거래에 '반복 거래' 사유 추가
  transactions.forEach(function(tx) {
    const key = tx.sender + '|' + tx.receiver;
    if (pairCount[key] >= 3 && tx.anomalyReasons.indexOf('반복 거래') === -1) {
      tx.anomalyReasons.push('반복 거래');
    }
    // status가 이상거래인데 anomalyReasons가 비어있으면 원본 상태 반영
    if (tx.status === '이상거래' && tx.anomalyReasons.length === 0) {
      tx.anomalyReasons.push('시스템 감지');
    }
  });

  // ── 통계 요약 ────────────────────────────────────────────────
  const totalCount  = transactions.length;
  const totalAmount = transactions.reduce(function(s, t) { return s + t.amount; }, 0);
  const anomalyCount = transactions.filter(function(t) {
    return t.status === '이상거래' || t.anomalyReasons.length > 0;
  }).length;

  // Top 판매자 3명
  const topSellers = Object.keys(sellerMap)
    .map(function(name) { return { name, count: sellerMap[name].count, total: sellerMap[name].total }; })
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 3);

  // Top 구매자 3명
  const topBuyers = Object.keys(buyerMap)
    .map(function(name) { return { name, count: buyerMap[name].count, total: buyerMap[name].total }; })
    .sort(function(a, b) { return b.count - a.count; })
    .slice(0, 3);

  // 태그별 통계 배열
  const tagStats = Object.keys(tagCount).map(function(tag) {
    return { tag, count: tagCount[tag], amount: tagAmount[tag] };
  }).sort(function(a, b) { return b.count - a.count; });

  // 주간 요약 텍스트 자동 생성
  const topTag   = tagStats.length > 0 ? tagStats[0].tag : '-';
  const weekSummary = `이번 기간 총 ${totalCount}건, 총 $${totalAmount.toLocaleString()} 거래 발생. ` +
    `최다 태그: ${topTag}. 이상 거래 ${anomalyCount}건 감지. ` +
    (topSellers.length > 0 ? `최다 판매자: ${topSellers[0].name}(${topSellers[0].count}건).` : '');

  // ── 네트워크 노드/엣지 (시각화용) ───────────────────────────
  // 노드: 거래에 등장한 모든 학생
  const nodeSet = new Set();
  transactions.forEach(function(tx) {
    nodeSet.add(tx.sender);
    nodeSet.add(tx.receiver);
  });
  const nodes = Array.from(nodeSet).map(function(name) {
    const sell  = sellerMap[name]  || { count: 0, total: 0 };
    const buy   = buyerMap[name]   || { count: 0, total: 0 };
    return {
      name,
      sellCount: sell.count,
      buyCount:  buy.count,
      totalActivity: sell.count + buy.count
    };
  });
  const edges = Object.values(edgeMap);

  return {
    transactions: transactions.reverse(), // 최신순
    stats: {
      totalCount,
      totalAmount,
      anomalyCount,
      topSellers,
      topBuyers,
      tagStats,
      weekSummary
    },
    network: { nodes, edges }
  };
}

// ── 수호대: 이상 거래 목록 반환 (메모 포함) ─────────────────────
function getP2PAlertsForGuard() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return [];

  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][7]).trim();
    if (status !== '이상거래') continue;
    result.push({
      rowNum:      i + 1,
      txnId:       String(data[i][0]),
      date:        String(data[i][1]).substring(0, 10),
      sender:      String(data[i][2]).trim(),
      receiver:    String(data[i][3]).trim(),
      amount:      Number(data[i][4]) || 0,
      tag:         String(data[i][5]).trim(),
      description: String(data[i][6]).trim(),
      memo:        data[i][8] ? String(data[i][8]).trim() : ''
    });
  }
  return result.reverse();
}

// ── 수호대: 이상 거래에 메모 저장 (I열) ─────────────────────────
function saveGuardMemo(rowNum, memo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return { success: false, msg: 'P2P거래로그 시트를 찾을 수 없습니다.' };
  if (rowNum < 2) return { success: false, msg: '유효하지 않은 행 번호입니다.' };
  try {
    sheet.getRange(rowNum, 9).setValue(String(memo || '').trim()); // I열
    return { success: true, msg: '메모가 저장되었습니다.' };
  } catch(e) {
    return { success: false, msg: '저장 오류: ' + e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// ██ 수호대 최종 적발 시스템
// 시트: 수호대적발로그
//   A=적발ID, B=적발일, C=피적발학생, D=거래ID,
//   E=적발사유, F=수호대메모, G=처리일
// ════════════════════════════════════════════════════════════════

// ── 수호대적발로그 시트가 없으면 자동 생성 ───────────────────────
function _ensureGuardPenaltySheet(ss) {
  let sheet = ss.getSheetByName(SHEET_GUARD_PENALTY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_GUARD_PENALTY);
    sheet.appendRow(['적발ID', '적발일', '피적발학생', '거래ID', '적발사유', '수호대메모', '처리일']);
  }
  return sheet;
}

// ── 수호대 최종 적발 기록 ────────────────────────────────────────
// txnRowNum : P2P거래로그의 행 번호 (2 이상)
// reason    : 적발 사유 문자열
// memo      : 수호대 추가 메모
function recordGuardPenalty(txnRowNum, reason, memo) {
  if (!txnRowNum || txnRowNum < 2) {
    return { success: false, msg: '유효하지 않은 거래 행 번호입니다.' };
  }
  if (!reason || !String(reason).trim()) {
    return { success: false, msg: '적발 사유를 선택해주세요.' };
  }

  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const p2pSheet = ss.getSheetByName(SHEET_P2P);
    if (!p2pSheet) return { success: false, msg: 'P2P거래로그 시트를 찾을 수 없습니다.' };

    // ── P2P거래로그에서 해당 행 정보 읽기 ──────────────────────
    const p2pRow    = p2pSheet.getRange(txnRowNum, 1, 1, 10).getValues()[0];
    const txnId     = String(p2pRow[0]).trim();   // A: 거래ID
    const victim    = String(p2pRow[2]).trim();   // C: 보내는학생(서비스 판매자 = 피적발 대상)
    const curStatus = String(p2pRow[7]).trim();   // H: 현재 상태

    if (!txnId) return { success: false, msg: '해당 행에 거래 데이터가 없습니다.' };

    // 이미 최종 적발된 거래는 중복 처리 방지
    if (curStatus === '최종적발') {
      return { success: false, msg: '이미 최종 적발 처리된 거래입니다.' };
    }

    const now     = new Date();
    const nowStr  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // ── P2P거래로그 H열 상태를 '최종적발'로 변경 ───────────────
    p2pSheet.getRange(txnRowNum, 8).setValue('최종적발');

    // ── 수호대적발로그에 기록 ───────────────────────────────────
    const penaltySheet = _ensureGuardPenaltySheet(ss);
    const penaltyId    = 'PNL_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
    penaltySheet.appendRow([
      penaltyId,                    // A: 적발ID
      dateStr,                      // B: 적발일
      victim,                       // C: 피적발학생
      txnId,                        // D: 거래ID
      String(reason).trim(),        // E: 적발사유
      String(memo || '').trim(),    // F: 수호대메모
      nowStr                        // G: 처리일
    ]);

    // ── 피적발 학생에게 우편 발송 ───────────────────────────────
    _sendMail(
      victim,
      '🚨 경제 수호대 적발 통보',
      '경제 수호대에 의해 거래(ID: ' + txnId + ')가 최종 적발되었습니다.\n사유: ' + String(reason).trim() + '\n이의가 있을 경우 선생님께 문의하세요.',
      'penalty'
    );

    return { success: true, msg: '✅ 최종 적발이 기록되었습니다.' };

  } catch (e) {
    return { success: false, msg: '오류가 발생했습니다: ' + e.message };
  }
}

// ── 수호대: 적발 로그 전체 반환 (신용점수 계산용 + 관리용) ────────
function getGuardPenaltyLog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_GUARD_PENALTY);
  if (!sheet) return [];

  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      penaltyId:   String(data[i][0]),
      date:        String(data[i][1]),
      victim:      String(data[i][2]).trim(),
      txnId:       String(data[i][3]).trim(),
      reason:      String(data[i][4]).trim(),
      memo:        String(data[i][5]).trim(),
      processedAt: String(data[i][6])
    });
  }
  return result.reverse(); // 최신순
}
