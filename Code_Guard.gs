// ════════════════════════════════════════════════════════════════
// ██ 경제 수호대 시스템
// P2P거래로그 시트 I열(수호대메모) 추가 필요
// ════════════════════════════════════════════════════════════════

// ── 수호대 이상거래 고액 기준 ────────────────────────────────────
// ※ 이 상수 하나를 바꾸면 getGuardDashboardData()와 getP2PAlertsForGuard()
//   두 함수 모두 동시에 반영됩니다. 기준 변경 시 여기만 수정하세요.
const GUARD_HIGH_AMOUNT = 1000; // $1,000 이상 거래 = 고액 거래로 분류

// ── AI 심층 브리핑에 사용할 Claude 모델 ──────────────────────────
// 비용/품질 균형이 좋은 Sonnet을 기본값으로 사용합니다.
// 더 깊은 분석이 필요하면 'claude-opus-4-8' 로 바꿀 수 있습니다.
const GUARD_AI_MODEL = 'claude-sonnet-4-6';

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
// period: 'week'(이번 주) | 'month'(이번 달) | 'all'(전체) | 'custom'(사용자 지정)
// startDate, endDate: period가 'custom'일 때만 사용 ('yyyy-MM-dd' 문자열)
function getGuardDashboardData(period, startDate, endDate) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const p2pSheet = ss.getSheetByName(SHEET_P2P);
  if (!p2pSheet) return { transactions: [], stats: {}, network: [] };

  const allData = p2pSheet.getDataRange().getValues();

  // ── 기간 필터 기준일 계산 ────────────────────────────────────
  const now   = new Date();
  let cutoff  = null;
  // 사용자 지정 기간(custom): startDate ~ endDate (둘 다 'yyyy-MM-dd' 문자열)
  let customStart = null, customEnd = null;
  if (period === 'custom') {
    customStart = startDate ? String(startDate).substring(0, 10) : null;
    customEnd   = endDate   ? String(endDate).substring(0, 10)   : null;
  } else if (period === 'week') {
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
  const tagQuantity  = {};  // 태그별 총 수량 (K열 기반, K열 없으면 1로 간주)
  const sellerMap    = {};  // 학생별 판매 건수 및 금액 (판매자 = 돈을 '받는' receiver)
  const buyerMap     = {};  // 학생별 구매 건수 및 금액 (구매자 = 돈을 '보내는' sender)
  const dateCountMap = {};  // 날짜별 거래 건수 (라인 차트용)
  // 네트워크: { "A→B": { from, to, count, total } }
  const edgeMap      = {};

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0]) continue; // 빈 행 스킵

    // row[1]이 Date 객체면 formatDate로 안전하게 변환, 문자열이면 앞 10자 사용
    const dateStr = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[1]).substring(0, 10);
    // 기간 필터 적용 (문자열 비교 — timezone 문제 없음)
    if (period === 'custom') {
      if (customStart && dateStr < customStart) continue;
      if (customEnd   && dateStr > customEnd)   continue;
    } else if (cutoff) {
      const cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (dateStr < cutoffStr) continue;
    }

    const sender   = String(row[2]).trim();
    const recv     = String(row[3]).trim();
    const amount   = Number(row[4]) || 0;
    const tag      = String(row[5]).trim();
    const desc     = String(row[6]).trim();
    const status   = String(row[7]).trim();
    const memo     = row[8] ? String(row[8]).trim() : '';
    // K열(인덱스10): 수량. 빈칸이면 1로 간주 (기존 데이터 호환)
    const quantity = (row[10] && Number(row[10]) > 0) ? Number(row[10]) : 1;

    // 이상거래 사유 재계산 (프론트에서 강조 표시용)
    const anomalyReasons = [];
    if (amount >= GUARD_HIGH_AMOUNT) anomalyReasons.push('고액 거래');
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
      quantity,
      anomalyReasons  // 빈 배열이면 강조 없음
    });

    // 태그 통계
    tagCount[tag]    = (tagCount[tag]    || 0) + 1;
    tagAmount[tag]   = (tagAmount[tag]   || 0) + amount;
    tagQuantity[tag] = (tagQuantity[tag] || 0) + quantity;

    // 날짜별 거래 건수
    dateCountMap[dateStr] = (dateCountMap[dateStr] || 0) + 1;

    // 판매자 통계: 서비스를 제공하고 자산을 '받는' 쪽(receiver)이 판매자
    if (!sellerMap[recv]) sellerMap[recv] = { count: 0, total: 0 };
    sellerMap[recv].count++;
    sellerMap[recv].total += amount;

    // 구매자 통계: 서비스를 구매하고 자산을 '보내는' 쪽(sender)이 구매자
    if (!buyerMap[sender]) buyerMap[sender] = { count: 0, total: 0 };
    buyerMap[sender].count++;
    buyerMap[sender].total += amount;

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
  const anomalySuspectCount = transactions.filter(function(t) {
  return t.status !== '최종적발' &&
         (t.status === '이상거래' || t.anomalyReasons.length > 0);
  }).length;
  const anomalyFinalCount = transactions.filter(function(t) {
    return t.status === '최종적발';
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

  // 태그별 통계 배열 — 수량 기반 단가 계산
  const FIXED_TAGS = ['#학습도움', '#정서적지지', '#재능판매', '#권리 및 기회'];
  const tagStats = Object.keys(tagCount).map(function(t) {
    const isFixed   = FIXED_TAGS.indexOf(t) !== -1;
    const totalQty  = tagQuantity[t] || tagCount[t]; // 수량 없으면 건수로 fallback
    const unitPrice = isFixed ? Math.round(tagAmount[t] / totalQty) : null;
    return {
      tag:        t,
      count:      tagCount[t],
      amount:     tagAmount[t],
      quantity:   totalQty,
      unitPrice:  unitPrice,   // 수량 기반 단가 (건당 평균 금액)
      avgAmount:  unitPrice    // 기존 필드명 호환용
    };
  }).sort(function(a, b) { return b.count - a.count; });

  // 날짜별 거래 건수 배열 (라인 차트용) — 날짜순 정렬
  const dateStats = Object.keys(dateCountMap)
    .sort()
    .map(function(d) { return { date: d, count: dateCountMap[d] }; });

  // 주간 요약 텍스트 자동 생성
  const topTag   = tagStats.length > 0 ? tagStats[0].tag : '-';
  const weekSummary = `이번 기간 총 ${totalCount}건, 총 $${totalAmount.toLocaleString()} 거래 발생. ` +
    `최다 태그: ${topTag}. 이상거래 의심 ${anomalySuspectCount}건 / 최종 적발 ${anomalyFinalCount}건. ` +
    (topSellers.length > 0 ? `최다 판매자: ${topSellers[0].name}(${topSellers[0].count}건).` : '');

  // ── 브리핑 리포트 생성 ────────────────────────────────────────
  // 불평등 지수 데이터 가져오기 (getInequalityData 재사용)
  let ineqSummary = '';
  try {
    const ineq = getInequalityData();
    if (ineq && ineq.success) {
      const g = ineq.giniAsset;
      const top20pct = Math.round(ineq.shareTop20 * 100);
      let giniDesc = '';
      if      (g < 0.20) giniDesc = '매우 평등한 상태로, 북유럽 복지국가 수준입니다.';
      else if (g < 0.30) giniDesc = '비교적 평등한 편으로, 독일·일본과 비슷한 수준입니다.';
      else if (g < 0.35) giniDesc = '우리나라 평균과 비슷한 수준입니다.';
      else if (g < 0.45) giniDesc = '불평등이 다소 심한 편으로, 미국과 비슷한 수준입니다.';
      else if (g < 0.55) giniDesc = '불평등이 심각한 수준으로, 남미 국가들과 비슷합니다.';
      else               giniDesc = '극심한 불평등 상태입니다.';

      // 지니계수 이력에서 직전 값 비교
      let prevGiniStr = '';
      if (ineq.history && ineq.history.length >= 2) {
        const prev = ineq.history[ineq.history.length - 2].giniAsset;
        const diff = g - prev;
        prevGiniStr = diff > 0
          ? `지난 브리핑(${prev.toFixed(3)})보다 불평등이 소폭 심화되었습니다.`
          : diff < 0
          ? `지난 브리핑(${prev.toFixed(3)})보다 불평등이 소폭 완화되었습니다.`
          : `지난 브리핑과 동일한 수준을 유지하고 있습니다.`;
      }

      ineqSummary =
        `현재 우리 반 자산 지니계수는 ${g.toFixed(3)}입니다.\n` +
        `이를 쉽게 설명하면, 상위 20% 학생이 우리 반 전체 자산의 약 ${top20pct}%를 보유하고 있는 상태입니다.\n` +
        (prevGiniStr ? prevGiniStr + '\n' : '') +
        `${giniDesc}`;
    }
  } catch(e) {
    ineqSummary = '(불평등 지수 데이터를 불러오지 못했습니다.)';
  }

  // 태그별 시세 요약 (단가 있는 태그만)
  const priceLines = tagStats
    .filter(function(t) { return t.unitPrice !== null; })
    .map(function(t) {
      return `${t.tag.padEnd(8)} — 거래 건수 ${t.count}건 / 총 수량 ${t.quantity}건 / 건당 단가 $${t.unitPrice.toLocaleString()}`;
    }).join('\n');

  // 기간 문자열
  let periodLabel = '이번 주';
  if (period === 'month')  periodLabel = '이번 달';
  if (period === 'all')    periodLabel = '전체 기간';
  if (period === 'custom') periodLabel = (customStart || '처음') + ' ~ ' + (customEnd || '오늘');

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd');

  const briefingReport =
    `📊 B.R.A.N.D 경제 브리핑 — ${periodLabel} (${today} 기준)\n` +
    `\n━━ 1. 거래 현황 ━━\n` +
    `이번 기간 우리 반에서는 총 ${totalCount}건, 총 $${totalAmount.toLocaleString()}의 거래가 발생했습니다.\n` +
    (tagStats.length > 0
      ? `가장 활발한 거래 유형은 ${tagStats[0].tag}(${tagStats[0].count}건)이며,\n` +
        tagStats.slice(1, 3).map(function(t) { return `${t.tag}(${t.count}건)`; }).join(', ') +
        (tagStats.length > 1 ? '이 뒤를 이었습니다.' : '')
      : '') +
    `\n\n━━ 2. 주목할 학생 ━━\n` +
    (topSellers.length > 0 ? `가장 활발히 판매한 학생: ${topSellers[0].name} (${topSellers[0].count}건, $${topSellers[0].total.toLocaleString()})\n` : '') +
    (topBuyers.length  > 0 ? `가장 활발히 구매한 학생: ${topBuyers[0].name}  (${topBuyers[0].count}건, $${topBuyers[0].total.toLocaleString()})` : '') +
    `\n\n━━ 3. 이상거래 모니터링 ━━\n` +
    `이상거래 의심 ${anomalySuspectCount}건이 감지되었으며, 이 중 최종 적발은 ${anomalyFinalCount}건입니다.\n` +
    `경제 수호대가 지속적으로 모니터링하고 있습니다.` +
    (ineqSummary
      ? `\n\n━━ 4. 우리 반 경제 불평등 리포트 ━━\n${ineqSummary}`
      : '') +
    (priceLines
      ? `\n\n━━ 5. 태그별 거래 시세 현황 ━━\n${priceLines}`
      : '');

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
      anomalySuspectCount,
      anomalyFinalCount,
      topSellers,
      topBuyers,
      tagStats,
      weekSummary,
      briefingReport,
      dateStats
    },
    network: { nodes, edges }
  };
}

// ════════════════════════════════════════════════════════════════
// ██ AI 심층 경제 브리핑 (#6)
// P2P거래로그 G열(거래 설명)을 AI가 직접 읽고 정성적으로 분석합니다.
//
// [사전 준비 — 선생님만 1회 설정]
//   Apps Script 편집기 → 좌측 톱니바퀴(프로젝트 설정)
//   → '스크립트 속성' → 속성 추가
//   → 속성 이름: ANTHROPIC_API_KEY / 값: 발급받은 API 키
//   (https://console.anthropic.com 에서 발급)
// ════════════════════════════════════════════════════════════════
function getGuardBriefingAI(period, startDate, endDate) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    return {
      success: false,
      msg: 'AI 브리핑을 사용하려면 먼저 선생님이 API 키를 등록해야 합니다.\n' +
           '(Apps Script → 프로젝트 설정 → 스크립트 속성 → ANTHROPIC_API_KEY)'
    };
  }

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const p2pSheet = ss.getSheetByName(SHEET_P2P);
  if (!p2pSheet) return { success: false, msg: 'P2P거래로그 시트를 찾을 수 없습니다.' };
  const allData = p2pSheet.getDataRange().getValues();

  // ── 기간 필터 기준 계산 (getGuardDashboardData와 동일 규칙) ──
  const now = new Date();
  let cutoffStr = null, customStart = null, customEnd = null;
  if (period === 'custom') {
    customStart = startDate ? String(startDate).substring(0, 10) : null;
    customEnd   = endDate   ? String(endDate).substring(0, 10)   : null;
  } else if (period === 'week') {
    const day  = now.getDay();
    const diff = (day === 0 ? -6 : 1 - day);
    const c    = new Date(now); c.setDate(now.getDate() + diff);
    cutoffStr  = Utilities.formatDate(c, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (period === 'month') {
    const c   = new Date(now.getFullYear(), now.getMonth(), 1);
    cutoffStr = Utilities.formatDate(c, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  // ── 거래 수집 (G열 설명·J열 평점·K열 수량 포함) ──────────────
  const txns = [];
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0]) continue;
    const dateStr = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[1]).substring(0, 10);

    if (period === 'custom') {
      if (customStart && dateStr < customStart) continue;
      if (customEnd   && dateStr > customEnd)   continue;
    } else if (cutoffStr) {
      if (dateStr < cutoffStr) continue;
    }

    txns.push({
      date:   dateStr,
      buyer:  String(row[2]).trim(),   // C열: 보내는 학생 = 구매자
      seller: String(row[3]).trim(),   // D열: 받는 학생  = 판매자
      amount: Number(row[4]) || 0,     // E열: 금액
      tag:    String(row[5]).trim(),   // F열: 태그
      desc:   String(row[6]).trim(),   // G열: 거래 설명 ★
      status: String(row[7]).trim(),   // H열: 상태
      rating: (row[9] !== '' && row[9] != null) ? Number(row[9]) : null, // J열: 평점
      qty:    (row[10] && Number(row[10]) > 0) ? Number(row[10]) : 1      // K열: 수량
    });
  }

  if (txns.length === 0) {
    return { success: false, msg: '선택한 기간에 분석할 거래 데이터가 없습니다.' };
  }

  // ── 기간 라벨 ──────────────────────────────────────────────
  let periodLabel = '이번 주';
  if (period === 'month')  periodLabel = '이번 달';
  if (period === 'all')    periodLabel = '전체 기간';
  if (period === 'custom') periodLabel = (customStart || '처음') + ' ~ ' + (customEnd || '오늘');

  // ── AI에게 전달할 거래 텍스트 구성 (토큰 절약: 최근 150건 제한) ──
  const MAX_TX = 150;
  const used   = txns.slice(-MAX_TX);
  const txLines = used.map(function(t) {
    return `[${t.date}] 구매자:${t.buyer} → 판매자:${t.seller} | ${t.tag} | $${t.amount} | 수량:${t.qty}` +
           (t.rating != null ? ` | 평점:${t.rating}/10` : '') +
           (t.status && t.status !== '정상' && t.status !== '정상 확인됨' ? ` | 상태:${t.status}` : '') +
           ` | 설명: ${t.desc || '(설명 없음)'}`;
  }).join('\n');

  const promptText =
    `당신은 초등학교 학급 모의경제 시스템 'B.R.A.N.D'의 경제 분석 보조 AI입니다.\n` +
    `아래는 학생들 사이에서 일어난 P2P(학생 간) 서비스 거래 기록입니다. ` +
    `각 거래는 한 학생(구매자)이 다른 학생(판매자)에게 자산(달러)을 보내고 서비스를 받은 내역이며, ` +
    `'설명'란에 실제로 어떤 서비스를 주고받았는지가 적혀 있습니다.\n\n` +
    `[분석 기간] ${periodLabel}\n` +
    `[총 거래 ${used.length}건]\n` +
    txLines +
    `\n\n위 자료를 바탕으로, 경제 수호대 학생들이 학급 친구들 앞에서 발표할 '경제 동향 브리핑'을 작성해 주세요.\n` +
    `다음 조건을 반드시 지켜주세요:\n` +
    `1. 단순히 숫자만 나열하지 말고, 각 거래의 '설명'을 실제로 읽고 어떤 종류의 서비스가 인기 있었는지, ` +
    `어떤 거래 흐름이나 패턴(예: 특정 학생이 어떤 분야를 전문적으로 판매, 서비스 가격 변화 등)이 보이는지 분석할 것.\n` +
    `2. 특정 학생을 비난하거나 망신 주지 말 것. 긍정적이고 교육적인 어조를 유지할 것.\n` +
    `3. 초등학생이 이해할 수 있는 쉬운 말로, 따뜻하면서도 분석적으로 쓸 것.\n` +
    `4. 아래 네 가지 소제목으로 구성할 것:\n` +
    `   ① 이번 기간 한눈에 보기\n   ② 인기 있었던 서비스와 거래 흐름\n   ③ 흥미로운 발견\n   ④ 우리 반에 주는 제안\n` +
    `5. 전체 길이는 600~900자 내외. 굵게/제목 같은 마크다운 기호(#, *, ** 등)는 쓰지 말고 일반 텍스트로 작성할 것.`;

  // ── Claude API 호출 ────────────────────────────────────────
  try {
    const payload = {
      model: GUARD_AI_MODEL,
      max_tokens: 1500,
      messages: [{ role: 'user', content: promptText }]
    };
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code !== 200) {
      return { success: false, msg: 'AI 호출 실패 (HTTP ' + code + ')\n' + body.substring(0, 300) };
    }

    const json = JSON.parse(body);
    const aiText = (json.content || [])
      .filter(function(b) { return b.type === 'text'; })
      .map(function(b) { return b.text; })
      .join('\n').trim();

    if (!aiText) {
      return { success: false, msg: 'AI 응답이 비어 있습니다. 잠시 후 다시 시도해주세요.' };
    }

    const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy.MM.dd');
    const report =
      `🤖 B.R.A.N.D AI 심층 경제 브리핑 — ${periodLabel} (${today} 작성)\n` +
      `※ 본 분석은 AI가 학생 간 거래 설명을 직접 읽고 작성한 참고용 리포트입니다.\n` +
      `────────────────────────────────────\n\n` +
      aiText;

    return { success: true, report: report };

  } catch (e) {
    return { success: false, msg: 'AI 호출 중 오류가 발생했습니다: ' + e.message };
  }
}

// ── 수호대: 이상 거래 목록 반환 (메모 포함) ─────────────────────
function getP2PAlertsForGuard(period, startDate, endDate) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();

  // ── [추가] 기간 필터 기준 계산 (스트리밍 탭 getGuardDashboardData와 동일 규칙) ──
  // period가 'all' 또는 미지정이면 cutoff=null → 전체 조회(기존 동작과 동일)
  const now = new Date();
  let cutoff = null;
  let customStart = null, customEnd = null;
  if (period === 'custom') {
    customStart = startDate ? String(startDate).substring(0, 10) : null;
    customEnd   = endDate   ? String(endDate).substring(0, 10)   : null;
  } else if (period === 'week') {
    const day  = now.getDay();             // 0=일, 1=월
    const diff = (day === 0 ? -6 : 1 - day);
    cutoff = new Date(now);
    cutoff.setDate(now.getDate() + diff);
  } else if (period === 'month') {
    cutoff = new Date(now.getFullYear(), now.getMonth(), 1);
  }

  // ─1단계: 전체 행 파싱 (+ [추가] 기간 필터 적용)
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    // [수정] 날짜 문자열을 먼저 계산해서 기간 필터에 사용
    const dateStr = (data[i][1] instanceof Date)
      ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(data[i][1]).substring(0, 10);

    // [추가] 기간 필터 (스트리밍 탭과 동일하게 적용 — 문자열 비교라 timezone 문제 없음)
    if (period === 'custom') {
      if (customStart && dateStr < customStart) continue;
      if (customEnd   && dateStr > customEnd)   continue;
    } else if (cutoff) {
      const cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (dateStr < cutoffStr) continue;
    }

    rows.push({
      rowNum:         i + 1,
      txnId:          String(data[i][0]),
      date:           dateStr,
      sender:         String(data[i][2]).trim(),
      receiver:       String(data[i][3]).trim(),
      amount:         Number(data[i][4]) || 0,
      tag:            String(data[i][5]).trim(),
      description:    String(data[i][6]).trim(),
      status:         String(data[i][7]).trim(),
      memo:           data[i][8] ? String(data[i][8]).trim() : '',
      anomalyReasons: []
    });
  }

  // ─2단계: 개별 행 기준 이상거래 재계산
  rows.forEach(function(tx) {
    if (tx.amount >= GUARD_HIGH_AMOUNT)
      tx.anomalyReasons.push('고액 거래');
    if (tx.description.length < 10)
      tx.anomalyReasons.push('사유 불충분');
    if (tx.tag === '#기타' && tx.description.length < 20)
      tx.anomalyReasons.push('태그 불일치 의심');
  });

  // ─3단계: 반복 거래 / 금액 집중 감지
  // [수정] '반복 거래'는 스트리밍 탭과 동일하게 '기간 전체' 기준(보낸사람|받는사람)으로 집계
  //        → 같은 두 사람이 기간 내 3회 이상 거래하면 반복 거래로 표시
  //        ('금액 집중'은 기존대로 '하루(date)' 기준 그대로 유지)
  const pairCount     = {};   // [추가] 기간 전체 동일 페어 건수 (날짜 무시)
  const dayPairAmount = {};   // [原본 유지] 하루 기준 페어 금액 합계
  rows.forEach(function(tx) {
    const pairKey = tx.sender + '|' + tx.receiver;                  // [수정] 날짜 제거
    const dayKey  = tx.date + '|' + tx.sender + '|' + tx.receiver;  // [原본 유지]
    pairCount[pairKey]    = (pairCount[pairKey]    || 0) + 1;
    dayPairAmount[dayKey] = (dayPairAmount[dayKey] || 0) + tx.amount;
  });
  rows.forEach(function(tx) {
    const pairKey = tx.sender + '|' + tx.receiver;
    const dayKey  = tx.date + '|' + tx.sender + '|' + tx.receiver;
    if (pairCount[pairKey] >= 3 &&
        tx.anomalyReasons.indexOf('반복 거래') === -1) {
      tx.anomalyReasons.push('반복 거래');
    }
    if (dayPairAmount[dayKey] >= 500 &&
        tx.anomalyReasons.indexOf('금액 집중') === -1) {
      tx.anomalyReasons.push('금액 집중');
    }
    // 시트에 이상거래로 기록됐지만 사유가 없으면 "시스템 감지"
    if (tx.status === '이상거래' && tx.anomalyReasons.length === 0) {
      tx.anomalyReasons.push('시스템 감지');
    }
  });

  // ─4단계: 이상거래 해당 행만 필터 후 반환
  const result = rows.filter(function(tx) {
    if (tx.status === '정상 확인됨') return false;
    return tx.status === '이상거래' ||
           tx.status === '최종적발' ||
           tx.anomalyReasons.length > 0;
});
  return result.reverse();
}

// ── 수호대: 이상 거래에 메모 저장 (I열) ─────────────────────────
function saveGuardMemo(rowNum, memo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_P2P);
  if (!sheet) return { success: false, msg: 'P2P거래로그 시트를 찾을 수 없습니다.' };
  if (rowNum < 2) return { success: false, msg: '유효하지 않은 행 번호입니다.' };
  try {
    sheet.getRange(rowNum, 9).setValue(String(memo || '').trim()); // I열: 메모
    sheet.getRange(rowNum, 8).setValue('정상 확인됨');              // H열: 상태 변경
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