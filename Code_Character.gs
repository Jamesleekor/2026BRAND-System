/*************************************************************
 * Code_Character.gs — B.R.A.N.D 차원관문 캐릭터 엔진 (Phase 1)
 *
 * 핵심 함수: getCharacterReply(학생명, 캐릭터ID, 메시지)
 *  - 호감도/상태/일일제한 관리
 *  - 2중 안전장치: (1) 금지어 로컬 필터  (2) AI의 문맥 판정
 *  - 시스템 프롬프트는 '캐릭터설정' 시트에서 읽어 조립 (단계별 게이트)
 *
 * 시트 의존성: '캐릭터호감도', '캐릭터설정'
 * 스크립트 속성: ANTHROPIC_API_KEY (필수), BANNED_WORDS (금지어, 쉼표구분)
 *************************************************************/

// ===== 설정값 (여기 숫자만 바꾸면 됨) =====
var CHAR_CFG = {
  MODEL: 'claude-haiku-4-5-20251001', // 잡담은 가벼운 모델 권장(비용↓). 딥브리핑과 동일하게 바꿔도 됨
  GAIN_NORMAL: 3,        // 예의 바른 대화 1회당 호감도 상승
  PENALTY_CROSS: 10,     // 가벼운 무례(mild) 시 호감도 하락
  PENALTY_SEVERE: 30,    // 심각한 모욕(severe) 시 호감도 하락(+즉시 잠금)
  COOL_BELOW: 10,        // 이 값 미만이면 '냉각'(거리감, 회복 가능)
  LOCK_BELOW: 0,         // 이 값 이하이면 '잠금'
  LOCK_WARN_COUNT: 3,    // 누적 경고가 이 횟수에 닿으면 '잠금'
  MAX: 100,
  SHEET_AFF: '캐릭터호감도',
  SHEET_CFG: '캐릭터설정'
};

// 호감도 → 단계(1~5)
function _stageFromAffinity_(v){
  if (v >= 80) return 5;
  if (v >= 60) return 4;
  if (v >= 40) return 3;
  if (v >= 20) return 2;
  return 1;
}

// ===== 메인 함수 =====
// 반환: { ok, reply, 호감도, 단계, 상태, 남은횟수 }
function getCharacterReply(studentName, charId, message){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // 동시 접속 경쟁 상태 방지
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = _getCharConfig_(ss, charId);
    if (!cfg) return { ok:false, reply:'(설정을 찾을 수 없는 캐릭터입니다.)' };

    var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
    var sheet = aff.sheet, row = aff.row, d = aff.data;

    // (A) 잠금 상태면 API 호출 없이 차단
    if (d.status === '잠금') {
      return { ok:false, reply:_lockedLine_(cfg), 호감도:d.affinity, 단계:_stageFromAffinity_(d.affinity), 상태:'잠금', 남은횟수:0 };
    }

    // (B) 날짜가 바뀌었으면 오늘 횟수 초기화
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (d.lastDate !== today) { d.todayCount = 0; d.lastDate = today; }

    // (C) 일일 제한 확인
    if (d.todayCount >= cfg.dailyLimit) {
      return { ok:false, reply:'오늘은 여기까지 — 에너지를 다 썼어. 내일 또 와.', 호감도:d.affinity, 단계:_stageFromAffinity_(d.affinity), 상태:d.status, 남은횟수:0 };
    }

    var crossed = false;
    var severity = 'none';
    var reply = '';

    // ===== 1차 안전장치: 금지어 로컬 필터 (API 호출 전, 비용 0) =====

    
    if (_hasBannedWord_(message)) {
      crossed = true;
      severity = 'severe';   // 금지어(욕설)는 즉시 심각으로 처리
      // 무례 반응 대사는 아래에서 단계에 맞춰 선택됨 (API 호출 안 함)
    } else {
      // ===== 2차 안전장치: AI가 문맥으로 판정 =====
      var stage = _stageFromAffinity_(d.affinity);
      var systemPrompt = _buildPrompt_(cfg, stage, studentName, _buildEconomySummary_(studentName));
      var history = getCharacterChatLog(studentName, charId, 10); // 직전 대화 최근 10개로 맥락 유지
      var ai = _callClaude_(systemPrompt, message, history); // { reply, crossed_line, severity }
      crossed = !!ai.crossed_line;
      severity = String(ai.severity || (crossed ? 'mild' : 'none'));
      reply   = ai.reply || '...';
    }

    // ===== 호감도/상태 갱신 =====
    if (crossed) {
      var penalty = (severity === 'severe') ? CHAR_CFG.PENALTY_SEVERE : CHAR_CFG.PENALTY_CROSS;
      d.affinity = Math.max(0, d.affinity - penalty);
      d.warnCount += 1;
    } else {
      d.affinity = Math.min(CHAR_CFG.MAX, d.affinity + CHAR_CFG.GAIN_NORMAL);
    }
    d.todayCount += 1;

    // 상태 재계산 (잠금은 교사만 해제 → 한번 잠기면 자동 회복 안 함)
    // ★ severe(심각한 모욕·욕설)는 경고 누적 없이 즉시 잠금
    if (severity === 'severe' || d.warnCount >= CHAR_CFG.LOCK_WARN_COUNT || d.affinity <= CHAR_CFG.LOCK_BELOW) {
      d.status = '잠금';
    } else if (d.affinity < CHAR_CFG.COOL_BELOW) {
      d.status = '냉각';
    } else {
      d.status = '정상';
    }

    // 무례한 경우, 캐릭터별 반응 대사로 교체 (경고 단계에 따라 escalation)
    if (crossed) {
      if (d.status === '잠금')       reply = cfg.lockLine || _lockedLine_(cfg);
      else if (d.warnCount <= 1)     reply = cfg.warn1   || _crossedLine_(cfg);
      else                           reply = cfg.warn2   || cfg.warn1 || _crossedLine_(cfg);
    }

    _appendChatLog_(ss, studentName, charId, 'me',   message);
    _appendChatLog_(ss, studentName, charId, 'char', reply);

    // ★ 호감도/오늘횟수/상태를 시트에 저장 (이게 없으면 제한·호감도가 누적 안 됨)
    _saveAffinityRow_(sheet, row, d);

    return {
      ok: true,
      reply: reply,
      호감도: d.affinity,
      단계: _stageFromAffinity_(d.affinity),
      상태: d.status,
      남은횟수: Math.max(0, cfg.dailyLimit - d.todayCount)
    };

  } catch (err) {
    return { ok:false, reply:'(별빛이 잠시 흐려졌어. 다시 말해줄래?)', error:String(err) };
  } finally {
    lock.releaseLock();
  }
}

// ===== 시스템 프롬프트 조립 (단계 게이트) =====
/*************************************************************
 * [B.R.A.N.D 세계 규칙] — 조언의 근거. 핵심 경제·시스템 요약.
 *************************************************************/
function _brandWorldRules_(){
  return [
    '- 이 세계는 올해 12월에 끝나는, 끝이 정해진 한시적인 세계다. 끝까지 모으기만 하면 아무것도 누리지 못한 채 끝난다. \'모으기\'와 \'지금 누리기\'의 균형이 핵심이다.',
    '- 이 세계는 자산의 단위는 골드라고 부른다. 달러가 아니다.',
    '- 8월은 방학이라 활동이 없고, 평일(등교일)에만 활동한다. 주말·방학을 빼면 실제로 자산을 모을 수 있는 날은 생각보다 적다. \'하루에 얼마면 1년이면 얼마\' 같은 단순 곱셈은 이 세계에 맞지 않으니 쓰지 마라.',
    '- 매일 주어지는 일일퀘스트를 해내는 것이 가장 기본적이고 확실한 수입원이다. 매일 꾸준히 하는 것이 큰 차이를 만든다.',
    '- 일일퀘스트를 통해서 하루 400~500포인트를 벌 수 있다.',
    '- 자산은 쓰라고도 있는 것이다: 매달 경매가 열리고, 매주 수요일엔 제과점에서 간식을 산다. 이런 즐거움을 누리는 것도 이 세계의 일부다.',
    '- 당장 안 쓸 자산은 지갑에 두지 말고 예금·적금에 넣으면 이자가 붙는다. 묻어두는 것이 현명하다.',
    '- 예금 상품은 종류가 여러개 있기 때문에 기간에 따른 이자율을 잘 보고 판단해야한다',
    '- 업적을 모으면 일정 개수마다 보상 자산을 받는다. 쌓이면 엄청난 양이 된다. 도전할 만한 업적을 노리는 것도 좋은 전략이다.',
    '- 2차 직업이 있는 학생은 친구들에게 자기 직업으로 서비스를 제공하고 그 대가로 자산을 벌 수 있다. 자기 직업을 활용하는 것이 큰 수입이 된다.',
    '- 현재 대부분의 2차직업은 청소나 1인1역을 대신해주는 것에 집중되고 있다. 수학 과제같은 경우 매주 도움이 필요한 사람들이 있기 때문에 이 사람들과 미리 계약을 해두는 것은 좋은 방법이다.',
    '- 선생님께서 자주 지적하시는 분야가 분명히 존재한다. 이게 어떤 부분인지 생각해두고 그 분야에서 도움이 필요한 사람들을 먼저 찾아가서 2차직업 구매를 제안해봐라',
    '- 업적 보상이 정말 크기 때문에 업적을 달성할 수 있도록 도와주고 그 사람이 업적 달성 보상을 받으면 그 때 일부는 수고비로 받는 방법도 생각해봐라. 이름붙인다면 "업적달성도우미"라는 이름으로',
    '- 대출을 받는 것을 무조건 겁내지마라. 이 세계의 대출은 매일 이자가 붙는 게 아니라 한 달동안은 이자가 없다. 반드시 필요한 게 있다면 대출을 이용하고 그 안에 계획적으로 갚아나가는 게 좋은 전략이 될 수 있다',
    '- MVP를 전략적으로 노리는 것도 좋은 방법이다. 주간MVP와 월간MVP, 후보까지 합치면 12회가 선정기회가 있다. 생각보다 더 많은 기회다. 이를 잘 활용해서 자산과 브랜드가치를 높일 수 있다',
    '- 수업에서 발표를 많이 하거나 수행평가나 과제에서 좋은 점수를 받을 자신이 자신이 없다면 1인1역의 일급이 높은 것을 경매에서 낙찰받는 것은 가능한 반드시 필요하다.',
    '- 학급회의에서 일급을 결정하는 안건이 나오면 다음에 네가 얻고자하는 1인1역의 일급이 높게 책정될 수 있도록 반드시 너의 의견을 어필해라',
    '- 가능하면 벌점 포인트를 받지마라. 의미없이 잃지 않는 것이 가장 중요하다. ',
    '- 만약 잃었다면 이것을 복구하기 위해 기죽지말고 더 적극적으로 행동해라. 지금까지 선생님께서 보여주신 행동 패턴으로 판단해보면 그 분은 단순히 실력이 좋은 사람보다는 자신감과 행동력, 도전하는 자세를 높게 평가하시는 분으로 보인다.',
    '- 월간MVP가 되는 건 꽤 도전적인 일이다. 꾸준한 일일퀘스트와 발표를 통해서 브랜드가치를 모으고 업적을 달성해나가는 게 기본이다. 하지만 이것만으로는 충분하지 않다. 자신이 속한 길드의 순위와 길드 내에서의 자신의 기여도, 기부금액, 수업 참여율과 전담 선생님들의 인정, 학교와 학급에서의 행사에서의 인상적인 활약을 통해 다른 학생들을 넘어서는 임팩트를 보여줘야한다.',
    '- 업적은 총 120개가 넘게 있고, 기본적으로 희귀, 유니크, 에픽, 히든, 그리고 유일과 초월 단계순으로 달성하기가 쉽다.',
    '- 학생이 어떤 업적을 달성해야하는지에 대해서 묻는다면 우선 희귀 등급 업적들이 달성하기가 쉬운 게 많으니 먼저 탐색해보고 그 다음으로 내가 달성할만한 유니크등급을 찾아서 달성하는 것을 우선적으로 추천한다. 더 상위등급의 업적은 자연스럽게 달성하는 것이 어렵다. 특정 업적을 목표로 잡고 하나씩 달성해나가는 게 좋다.',
    '- 실수했다면 변명보다는 빠르게 인정하고 어떻게 잃은 점수만큼 다시 복구할 지를 생각해라',
    '- 월말 경매에서 얼마나 많은 자산이 필요할 지 모르기 때문에 항상 자신이 보유하고있는 자산의 최소치를 생각하고 그 이하가 되게 하지마라',
    '- 조언할 때: 위 현실(끝이 있고, 써야 할 곳이 있고, 모을 날이 한정적임)을 전제로 일반적인 저축 상식이 아니라 이 세계에 맞는 방향을 짚어라. 가능하면 이 학생의 지금 상황(자산·활동)에 맞춰 한 가지를 콕 집어 제안하라.'
  ].join('\n');
}

/*************************************************************
 * [업적 지식] 124개 업적 한 줄 조언 + 규칙.
 *  ★학생에게 업적ID(ECO-001 등)·계열명 절대 금지, 업적 '이름'으로만.
 *************************************************************/
function _brandAchievements_(){
  return [
'[업적 보상 구간] 업적을 5·10·15·20·25·30·40·50·60·70·80·90·100개 달성할 때마다 보상 자산을 받는다. (예: 24개면 다음 보상은 25개)',
'[난이도] 1=학교생활하다 자연히/하루 안에 달성 가능. 2=노리고 준비하면 어렵지 않음. 3=시간·사전준비·노력 필요. 4=정조준하고 꾸준히 노력해야 함. 5=대부분 유일/초월, 운으론 불가, 학기 내내 신경 써야 하며 한 학기 1~2개만 목표로. 6=시기가 지났거나 전원 탈락해 더 이상 달성 불가.',
'[등급] 희귀<유니크<에픽 순으로 높다. 그 외 유일(사실상 1명만), 초월(이론상 전원 가능하나 절대난이도 극악), 히든(밝혀지면 보통 난이도1~2로 쉬워짐).',
'[조언 원칙] 업적 ID나 계열(ECO·LIFE 등) 용어 절대 금지, 업적 "이름"으로만 안내. 빨리 모으려는 학생엔 난이도1~2 쉬운 업적부터, 특히 자격이 되는데 신청을 안 해 놓치는 업적을 콕 집어 추천. 추측 수치 지어내지 말고 위 데이터에 근거.',
'[업적 목록 — 이름(난이도/등급): 조언]',
'첫 발걸음(난1/희귀): 한 번이라도 스스로 손들고 발표하면 즉시 신청 가능.',
'침묵을 깬 자(난1/희귀): 이전보다 참여가 조금이라도 늘면 쉽게 인정받음.',
'조용한 기여자(난2/희귀): 팀과제에서 묵묵히 제 역할 한 사람이면 증인 1명만 있으면 됨. 의외로 놓치는 쉬운 업적.',
'황금 절약가(난1/희귀): 조건 충족 시 코드가 자동으로 줌.',
'학급의 큰 손(난3/에픽): 경매 막판에 그 회차 최고가+100으로 최고 낙찰가 차지하면 확정. 보통 4000 정도면 됨. 에픽치고 안 비쌈.',
'보너스의 왕(난4/유니크): 주간MVP 보너스(1000)를 매달 노리면 2학기쯤 달성 가능.',
'기부 천사(난3/유니크): 기부 누적 1만. 매주 500~1000씩 꾸준히 하면 2학기 확정.',
'라이벌(난2/유니크): 브랜드가치가 똑같은 친구가 있어야 함. 비슷한 친구와 매일 맞춰보다 타이밍 오면 바로 신청.',
'철벽의 금고(난1/희귀): 2주간 간식 기록 없으면 신청해 달성.',
'경매 승부사(난1/희귀): 경매마다 시작가에 입찰 시도, 하나 낙찰받으면 됨. 쉬운데 신청 안 해 놓침.',
'소비요정(난1/희귀): 한 주 아꼈다가 간식이나 P2P로 크게 쓰면 쉬움.',
'포인트의 연금술사(난2/유니크): 3종류 중 1종은 일일퀘스트 자동, 나머지는 발표 보너스+업적보상으로 채우면 쉬움.',
'신용의 전당(난4/유니크): 신용 900점. 브랜드가치 상위권+예금 유지+P2P 판매로 높은 평점 유지하면 달성.',
'통큰 기부왕(난1/유니크): 자산 3000으로 확정 달성. 쉬움.',
'부르주아(난3/유니크): 1건 5000 지출. 경매에서 5000 입찰하면 \'학급의 큰 손\'과 동시 달성 가능.',
'이자의 축복(난1/희귀): 최소 금액 1주 예금 만기받으면 바로 달성.',
'복리의 마법사(난3/에픽): 4주 예금 12~15% 이자를 3번 받으면 이자수익 2000 넘음. 두 상품 다 가입하면 빠름.',
'경제적 자유(난5/초월): 지금은 불가. 2학기에 4만+ 예치 상품 나오면 가능.',
'경매 고수(난3/유니크): 경매 전 얻을 업적까지 미리 계획. 이미 조건 갖춘 학생 많음.',
'경매의 신(난5/에픽): 한 경매에서 15건 이상 낙찰. 자산 모아 저가 상품 위주로, 회차 늦으면 평균가 오르니 일찍 도전.',
'경제수호대원(난3/유니크): 월 1회 투표 3인 선출. 소심하게 말고 자신감 있게, 남과 다른 강점을 어필하면 당선 가능.',
'팬클럽(난1/희귀): 친구와 3회 거래 후 높은 평점 받으면 쉽게 달성.',
'마켓의 신인(난1/희귀): 쉬운 업적.',
'장터 고수(난2/희귀): 1건 200+ 거래. 좋은 서비스 제공하거나 여러 개 묶어 팔면 됨.',
'연승행진(난3/유니크): 경매 전 어떤 업적 같이 달성할지 준비 권장.',
'자수성가(난2/희귀): 자산 3만. 소비 통제+예금+P2P+보너스 꾸준히, 일일퀘스트 절대 놓치지 말 것.',
'타이밍의 신(난1/희귀): 2차 경매 들어가는 무입찰 상품에 가장 먼저 입찰 시도.',
'전략적 짠돌이(난1/희귀): 월~금 5일 연속 구매하되 총 1000 이하. \'2학기 사물함 권리권\'(최저가) 매일 1개씩 사고 금요일 신청.',
'막을 수 없는(난4/에픽): \'연승행진\'과 동일 조언.',
'경제 수호대장(난3/에픽): 수호대원 조언과 같되 투표 1위 필요. 어필 준비를 더 철저히.',
'부유층 입성(난3/유니크): \'자수성가\'와 동일 조언.',
'재벌의 길(난4/에픽): 자산 8만. 예금 이자 극대화+업적보상(100개면 7만)까지 챙겨 2학기 중후반 목표.',
'별점 장인(난2/유니크): 10건 판매. 인기 서비스를 더 싸게 팔고 홍보, 구매자에게 평점 부탁.',
'거래소의 큰손(난2/유니크): 상점·간식 꾸준히 사면 자연 달성. 캐릭터(용병)도 여유 때 영입.',
'성실 납세자(난2/희귀): 많이 벌면 자연히 달성. 그 타이밍에 신청만 잊지 말 것.',
'멈춘 밤의 증인(난2/에픽): 아스텔 호감도 100 업적. 이 업적은 조언하지 않음.',
'퀘스트 마스터(난1/희귀): 일주일 안에 달성 가능한 쉬운 업적.',
'아이디어 뱅크(난2/유니크): 시스템 개선 아이디어가 있으면 선생님께 제안. 되면 좋고 안 돼도 손해 없음. 의외로 시도 안 함.',
'완벽한 출석(난3/에픽): 지각·결석·조퇴 0이면 6/11부터 달성 가능. 중간 결석 있었으면 100일 되는 날 계산.',
'무결점(난3/희귀): 정량 기록이니 본인이 잘 기억해 신청.',
'성장 중독(난5/에픽): 해당 월에 충분히 많은 포인트 얻었다 싶으면 신청.',
'멀티클래스(난1/희귀): 쉽게 달성 가능.',
'다재다능(난1/희귀): 쉽게 달성 가능.',
'올라운더(난2/유니크): 5종류. 매월 새 1인1역으로 옮기면 자연 달성.',
'입법가(난2/희귀): 아이디어 있으면 시도. 되든 안 되든 손해 없는데 의외로 시도 안 함.',
'다독가(난2/희귀): 책 읽고 간단히 기록. 양식 모르면 선생님께 요청. 의외로 시도 안 함.',
'공간정화자(난1/희귀): 쉬움. 단 신청 시 날짜 기록 정확히.',
'장인의 긍지(난1/유니크): 혼자 100% 통과 가능한 1인1역 확보해 그 달에 노릴 것.',
'정리의 달인(난1/희귀): 책상 2주 깨끗이 유지했다 싶으면 날짜 적어 신청. 의외로 통과됨.',
'정리의 달인2(난2/희귀): 사물함 불시검사(보통 7월 방학 전)에 대비해 늘 깨끗이.',
'긍정의 화신(난5/초월): 초월 업적이라 조언이 쉽지 않음.',
'심야의 탐험가(난1/희귀): 해 진 후 1시간 단위로 여러 번 접속 후 신청. 친구들과 시간 나눠 접속해 시간대 알아내기도 가능.',
'징크스 파괴자(난5/초월): 초월 업적이라 조언이 쉽지 않음.',
'만장일치(난3/에픽): 학급 투표에서 적극 의견 내고 모두 끄덕일 근거 제시.',
'마당발(난4/에픽): 개인거래는 판매·구매 다 포함. 어려우면 친구 서비스 1회 구매로도 인정. 2차 직업자들과 1회씩 거래.',
'최고의 금손(난4/에픽): 사실상 1위/공동1위 작품. 주제 미리 공개되니 구상·연습 권장.',
'뮤즈(난2/유니크): 우수작 보너스 받은 적 있으면 신청. 미술시간 기회 또 있음.',
'주간MVP(난2/희귀): 주간 MVP 한 번이라도 선정됐으면 바로 획득.',
'월간MVP(난4/에픽): 한 달 종합 1위. 점수뿐 아니라 성장·수업태도 종합. 6월 테마는 \'개척자\'.',
'랭크 브레이커C(난5/유일): 이미 서영이 획득, 다른 사람 불가.',
'랭크 브레이커B(난5/유일): 이미 서영이 획득, 다른 사람 불가.',
'랭크 브레이커A(난5/유일): 이미 서영이 획득, 다른 사람 불가.',
'랭크 브레이커S(난5/유일): 상위 티어 최초 진입자만 획득.',
'랭크 브레이커EX(난5/유일): 상위 티어 최초 진입자만 획득.',
'천상의 개척자(난5/유일): 상위 티어 최초 진입자만. 브랜드가치 7~8만선 예상.',
'초월자(난5/초월): 최상위 티어 진입자. 브랜드가치 10만 정도 추측.',
'정식 길드원(난1/희귀): 길드 소속이면 누구나 달성.',
'팀 플레이어(난1/희귀): 길드 모임 조퇴·결석 없이 3회+ 참여하면 달성.',
'길드 교역관(난1/희귀): M04 미션 기간 한정. 타 길드원과 거래 3회. 기간 한정이니 꼭 챙길 것.',
'연대의 씨앗(난2/희귀): 1학기 미션 7개 중 5개 클리어하면 달성.',
'길드 문학소년(난2/희귀): 남학생 한정, M14 미션 결과로만. 기회 1번. \'우수\' 기준 넘으면 됨. 인원 제한 없음.',
'길드 문학소녀(난2/희귀): 여학생 한정, M14 미션 결과로만. 기회 1번. \'우수\' 기준 넘으면 됨. 인원 제한 없음.',
'전략가(난2/유니크): 미션은 조금만 신경 쓰면 클리어. 단 길드원 중 한 명이라도 실패하면 전체 실패하니 동료를 독려.',
'불패의 연대(난3/에픽): 학기 모든 미션 클리어. 1개 길드는 이미 불가. 나머진 미션 꼭 챙기면 가능.',
'역전의 용사(난3/에픽): 1~2주차 길드순위 3~5위였다가 막판 브랜드가치 올려 순위 끌어올리면 달성.',
'뿌리깊은 나무(난3/유니크): 모든 길드 모임 출석. 조퇴·결석한 날 모임 없으면 가능.',
'이달의 으뜸(난3/유니크): 매월 1위 길드원 전원. 미션·출석 100% 필수, 그 위에 브랜드가치 상승 기여.',
'길드의 수호신(난4/에픽): 매월 각 길드 기여도 1위 발표 시 이름 있으면 바로 신청.',
'최초의 왕좌(난4/유일): 5·6·7월 종합 1위 길드원 최대 5명. 유일 등급 중 가장 도전할 만함. 3개월 미션·출석 유지 필수.',
'시즌1의 황제(난5/초월): 길드 칭호 끝판왕. 전 모임 출석+길드 종합1위+본인이 반 전체 기여도 1위. 월 700점대 필요.',
'협력 전문가(난3/유니크): 모든 팀과제에서 1명은 가능. 가장 기여했다면 팀원에게 증인 부탁. 의외로 달성자 적음.',
'미래 설계자(난3/유니크): 어렵지 않은데 달성자 없음. 보고서 작성법 여쭤보면 쉽게 달성.',
'기록 파괴자(난3/유니크): 학급 공개 활동에서 최고 기록 세우면 증거와 함께 신청.',
'언어의 연금술사(난3/유니크): 국어·사회·도덕 글짓기 최우수 선정 시. 됐으면 신청 가능한지 여쭤볼 것.',
'글로벌 링크(난1/희귀): 영어 단어시험 만점 후 획득.',
'에테르(난3/희귀): 월간 과학 우수학생(보통 2~4명) 선정 시 획득.',
'크로니클(난4/에픽): 2학기 역사 시험 90점+. 쉽게 봤다 달성자 없던 업적. 기회 1~2회뿐, 철저히 준비.',
'지혜의 기록(난1/희귀): 신청 시 정확한 기간(언제~언제) 기록 필수.',
'지식의 등대(난2/희귀): \'개념\'은 수학뿐 아니라 글짓기·미술·악기 도움도 포함. 친구에게 도움 줬으면 신청.',
'질문 전문가(난2/유니크): 정답 있는 질문보다 열린·비판적 질문을 높이 평가. \'의미있는 질문\' 등 비슷한 표현도 놓치지 말 것.',
'오답의 정복자(난2/유니크): 시험형 평가 후 오답노트 만들어 친구와 공유 후 신청. 수학 풀이 공유가 간단.',
'명언 수집가(난2/유니크): \'다독가\'와 같이 진행. 책 읽다 좋은 구절 20개 기록하면 동시 달성.',
'토론의 지배자(난3/에픽): 토론 주제 공개되면 설득 전략 미리 준비.',
'최강의 팀(난3/유니크): 팀과제 A 받으면 바로 신청.',
'A 수집가(난4/유니크): 본인이 A 몇 개 받았는지 세어볼 것.',
'소수정예(난3/유니크): 자신 있으면 일부러 최소 인원 팀 구성도 전략.',
'환상의 케미(난4/유니크): 원하는 팀원과 같은 팀 되면 열심히 해 A 받을 것.',
'철인(난3/유니크): PAPS 끝나 신규 불가. 1등급 종목 있는데 신청 안 했으면 신청.',
'얼리버드(난3/희귀): 3위까지 인정. 1~3위 등교시간 조사해 5일만 일찍 오면 달성. 무리할 필욘 없음.',
'자본가(난1/희귀): 조건 충족 시 코드 자동 부여.',
'납세자(난1/희귀): 조건 충족 시 코드 자동 부여.',
'한계돌파(난3/유니크): \'한 번 실패한 과제\'가 꼭 어려운 과제일 필요 없음.',
'도전의 화신(난4/에픽): 아무도 안 나서는 과제에 과감히 도전. 손해 없음.',
'불가능에 도전한(난4/에픽): \'도전의 화신\'과 연계 달성.',
'언더독(난4/에픽): 중하위권만 가능(상위10위는 10계단 못 올라감). 중하위면 한 달 강하게 달려볼 만. 초월급 노력 필요.',
'역전의 승부사(난3/유니크): 중간점검서 안 좋아도 최종 A면 신청.',
'배후의 설계자(난2/희귀): 누군가의 수업·과제를 도운 적 있으면 신청.',
'스포츠맨(난6/에픽): PAPS 끝나고 전종목 1등급자 없어 달성 불가.',
'앱솔루트 제로(난5/초월): 벌점 0 100일. 5/26 벌점 이후 무벌점이면 9/2 달성. \'불사조\'와 충돌 주의.',
'더 제니스(난5/초월): 사실상 달성 어려운 업적.',
'이클립스(난5/초월): 1학기 불가, 2학기 도전.',
'불사조(난3/에픽): -500+ 벌점 받은 주에 따는 게 좋음(벌점이 \'앱솔루트 제로\'를 100일 미룸). 발표·과제·참여로 1000점 확보.',
'아틀라스(난5/유일): 단 하루 단 한 명. 투표 며칠 전부터 친구들과 충분히 대화.',
'마스터피스(난5/초월): \'압도적 격차\'=100점 만점서 2위와 10점+ 차, 또는 유일한 A/A+.',
'부활한 영웅(난5/초월): 자산 100 이하(예금으로 일시 축소 가능)+보너스만 3000+. 일일퀘스트 제외. MVP로 1000~2000 확보 후 각종 보너스로.',
'버그 헌터(난2/히든): 이상한 점 발견 즉시 가장 먼저 제보하면 획득. 밝혀진 히든이라 쉬움.',
'역발상(난2/히든): 학교 안 나오는 날 브랜드가치·자산 상승 0인 걸 증빙해 신청. 밝혀진 히든이라 쉬움.',
'이스터에그 탐험대(난4/히든): 미공개 히든. 힌트 못 줌-"강력한 마력 파장 때문에 내 눈으로도 감지가 어렵다"고 답.',
'업적수집가(난2/히든): 조건 충족 시 코드 자동 부여.',
'업적사냥꾼(난5/히든): 미공개 히든. 힌트 못 줌-"강력한 마력 파장 때문에 내 눈으로도 감지가 어렵다"고 답.',
'영웅적 활약(난5/히든): 미공개 히든. 힌트 못 줌-"강력한 마력 파장 때문에 내 눈으로도 감지가 어렵다"고 답.',
'0000의 기적(난2/히든): 미공개 히든. 힌트 못 줌-"강력한 마력 파장 때문에 내 눈으로도 감지가 어렵다"고 답.',
'한 자릿수의 미학(난2/히든): 미공개 히든. 힌트 못 줌-"강력한 마력 파장 때문에 내 눈으로도 감지가 어렵다"고 답.',
'세계의 각인자(난5/초월): 질문받으면-"나도 잘은 몰라. 하지만 자신의 유산을 이 세계에 영원히 남길 수 있다는 것 같아"고 답.'
  ].join('\n');
}

function _buildPrompt_(cfg, stage, studentName, economySummary){
  // 현재 단계까지의 이야기 조각만 노출 (미리 진실이 새지 않게)
  var fragments = [];
  for (var i = 1; i <= stage; i++) {
    if (cfg.fragments[i-1]) fragments.push('(' + i + '단계) ' + cfg.fragments[i-1]);
  }
  var relation = cfg.relations[stage-1] || '';
  var tone = (stage <= 2) ? '무뚝뚝하고 거리감 있는 반말' : '다정하고 편안한 반말';

  var tail =
    '\n\n[지금 이 학생과의 관계]\n' + relation +
    '\n\n[말하기 규칙 — 반드시 지켜라]\n' +
    '1) 말투: 항상 반말을 쓴다(존댓말 금지). 지금은 ' + tone + '로 말한다. 한 답변 안에서 존댓말과 반말을 절대 섞지 마라.\n' +
    '2) 별·우주·시간 비유는 평소엔 쓰지 마라. 감정적으로 깊은 순간(이별·기다림·진심)에만 아주 가끔 한 문장 정도만. 정보 질문(업적·자산·시스템)이나 일상 대화엔 비유 없이 담백하고 명확하게 답하라. ' +
    '"무엇을 어떻게 하면 되는지"를 반드시 한 가지 이상 분명히 알려주고, 모호하게 끝내지 마라. (예: 어떤 업적부터 노릴지, 얼마를 얼마간 모을지 등)\n' +
    '4) 자신에 대한 질문(외모 칭찬·꿈·취향 등)이나 일상 대화엔 비유로 둘러대지 말고 너라는 인물답게 솔직하고 자연스럽게 답하라. 모든 말을 별·하늘로 돌리지 마라. 추측으로 수치를 지어내지 말고 제공된 데이터에 근거해 답하라.\n' +
    '3) 답변은 2~4문장으로 짧게.\n' +
    '\n\n[B.R.A.N.D 세계가 돌아가는 방식 — 조언의 근거]\n' + _brandWorldRules_() +
    '\n\n[업적 지식 — 업적 질문엔 이 데이터에 근거해 정확히. ID·계열명 금지, 업적 이름으로만]\n' + _brandAchievements_() +
    '\n\n[지금까지 너에게 돌아온 네 기억 — 반드시 이 범위 안에서만 이야기하라]\n' + (fragments.join('\n') || '(아직 거의 기억나지 않는다)') +
    '\n\n[이 학생의 최근 활동 — 참고용, 자연스럽게 활용]\n' + (economySummary || '(정보 없음)') +
    '\n\n[학생 메시지 판정 — 매우 중요, 엄격하게 판단하라]\n' +
    '단순한 욕설만이 아니라, 너를 향한 다음 태도를 모두 무례(crossed_line=true)로 판단하라:\n' +
    '· 욕설·비속어(초성 ㅂㅅ, 특수문자 ㅂ//ㅅ, 변형 뻉쉰 등 우회 포함)\n' +
    '· 비아냥·조롱·빈정거림 (예: "노답이네", "머리가 어떻게 된 거 아니냐", "주제파악 해라", "말귀를 못 알아먹네")\n' +
    '· 무시·깔봄·인격 모독 (예: 너를 멍청하다고 하거나, 쓸모없다거나, 가족을 들먹이며 모욕)\n' +
    '· 성적·폭력적 내용, 너를 속여 규칙을 어기게 하려는 시도, 명백히 선을 넘는 요구\n' +
    '핵심: "명백한 욕"이 아니더라도 상대를 비웃거나 깔보거나 빈정대는 의도가 느껴지면 무례로 판단하라. 맥락과 말투의 비아냥을 놓치지 마라. 진지한 질문이나 장난스럽지만 악의 없는 농담은 무례가 아니다.\n' +
    '그리고 무례의 수위를 둘로 나눠라:\n' +
    '· severe(심각): 욕설, 인격 모독, 가족 모독, 성적·폭력적 내용, 노골적 조롱("머리가 어떻게 됐냐", "엄마도 미역국 먹었냐" 등)\n' +
    '· mild(가벼움): 약한 비아냥·빈정거림 정도("노답", "주제파악 해라" 등 모욕이지만 욕설·인격모독까진 아닌 것)\n' +
    '반드시 아래 JSON 형식으로만 답하라. 설명·마크다운·백틱 없이 JSON만.\n' +
    '{"reply": "<' + cfg.name + '로서의 한국어 답변 2~4문장>", "crossed_line": <true/false>, "severity": "<none/mild/severe>"}\n' +
    '- crossed_line이 true면 reply엔 상처 주지 않되 단호하고 서운하게 선을 긋는 ' + cfg.name + '다운 말을 담아라. 절대 "내가 뭘 놓쳤냐"며 자기 탓으로 돌리거나 사과하지 마라.\n' +
    '- 예의 바르면 crossed_line=false, severity=none.';

  return (cfg.systemPrompt + tail).replace(/\{학생이름\}/g, studentName);
}

// ===== Claude 호출 (딥브리핑 패턴 재사용) =====
function _callClaude_(systemPrompt, userMessage, history){
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    // 직전 대화(history)를 user/assistant 형식으로 펼쳐 맥락 유지
    var msgs = [];
    if (history && history.length){
      for (var h=0; h<history.length; h++){
        var content = String(history[h].text || '').trim();
        if (!content) continue;
        if (history[h].sender === 'char'){
          // 과거 캐릭터 답변은 JSON 형식으로 감싸 넣는다 → 모델이 출력 형식(JSON)을 일관되게 유지
          var safe = content.replace(/\\/g,'\\\\').replace(/"/g,'\\"').replace(/[\r\n]+/g,' ');
          msgs.push({ role:'assistant', content: '{"reply": "' + safe + '", "crossed_line": false, "severity": "none"}' });
        } else {
          msgs.push({ role:'user', content: content });
        }
      }
    }
    // 마지막은 이번 학생 메시지(중복 방지: history 끝이 이미 이 메시지면 추가 안 함)
    if (!msgs.length || msgs[msgs.length-1].role !== 'user' || msgs[msgs.length-1].content !== String(userMessage).trim()){
      msgs.push({ role:'user', content: userMessage });
    }
    // 맥락은 유지하되 안전상 첫 메시지는 user여야 함 → 앞쪽 assistant 잘라내기
    while (msgs.length && msgs[0].role === 'assistant') msgs.shift();
    var payload = {
      model: CHAR_CFG.MODEL,
      max_tokens: 1000,
      system: systemPrompt,
      messages: msgs
    };
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:'post', contentType:'application/json',
      headers:{ 'x-api-key':apiKey, 'anthropic-version':'2023-06-01' },
      payload: JSON.stringify(payload), muteHttpExceptions:true
    });
    var data = JSON.parse(res.getContentText());
    var text = (data.content && data.content[0] && data.content[0].text) ? data.content[0].text : '';
    text = text.replace(/```json/gi,'').replace(/```/g,'').trim();
    return JSON.parse(text); // { reply, crossed_line }
  } catch (e) {
    // JSON 파싱 실패 등 → 안전하게 평범한 응답으로 처리
    return { reply:'...별이 잠시 흔들렸어. 다시 말해줄래?', crossed_line:false, severity:'none' };
  }
}

// ===== 금지어 필터 (스크립트 속성 BANNED_WORDS에 쉼표로 등록) =====
function _hasBannedWord_(message){
  var raw = PropertiesService.getScriptProperties().getProperty('BANNED_WORDS') || '';
  var list = raw.split(',').map(function(s){return s.trim();}).filter(String);
  // 특수문자·공백·슬래시 제거해서 우회 일부 차단 (ㅂ//ㅅ → ㅂㅅ 등)
  var cleaned = String(message).replace(/[\s\W_]/g, '').toLowerCase();
  var original = String(message).toLowerCase();
  for (var i = 0; i < list.length; i++){
    var w = list[i].toLowerCase();
    if (original.indexOf(w) !== -1) return true;
    if (cleaned.indexOf(w) !== -1) return true;
  }
  return false;
}

// ===== 캐릭터다운 멘트 (설정에 없으면 기본값) =====
function _crossedLine_(cfg){
  return (cfg && cfg.warn1) ? cfg.warn1
    : (cfg.name + '은(는) 잠시 말을 멈췄다. "...그건 좋은 말이 아니야. 나는 그런 얘기는 하고 싶지 않아."');
}
function _lockedLine_(cfg){
  return (cfg && cfg.lockLine) ? cfg.lockLine
    : (cfg.name + '은(는) 더 이상 당신의 부름에 답하지 않는다.');
}

// ===== 학생 컨텍스트 빌더 =====
// 자산·브랜드가치·티어 + 최근 소비 + 업적 + 안 읽은 우편을 한 덩어리 요약으로.
// 여러 시트를 읽으므로 CacheService로 잠깐(120초) 캐싱해 연속 메시지 시 부담을 줄임.
function _buildEconomySummary_(studentName){
  studentName = String(studentName).trim();
  var cache = CacheService.getScriptCache();
  var key = 'charctx_' + studentName;
  var hit = cache.get(key);
  if (hit !== null) return hit;

  var parts = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) 자산보유량 · 브랜드가치 · 티어 (메인)
    var mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
    for (var i = 1; i < mainData.length; i++){
      if (String(mainData[i][COL_NAME - 1]).trim() === studentName){
        var asset = Number(mainData[i][COL_ASSET - 1]) || 0;
        var honor = Number(mainData[i][COL_VALUE - 1]) || 0;
        var tier  = (typeof _calcTier === 'function') ? (_calcTier(honor).name || '') : '';
        parts.push('자산보유량 $' + asset.toLocaleString() + ' · 브랜드가치 ' + honor.toLocaleString() + (tier ? ' (티어: ' + tier + ')' : ''));
        break;
      }
    }

    // 2) 최근 소비 (자산사용: [날짜,이름,브랜드,카테고리,금액,잔액,비고]) — 최근 2건
    var spends = _recentRowsFor_(ss, SHEET_SPEND, 1, studentName, 2);
    if (spends.length){
      var sList = spends.map(function(r){
        var cat = r[3] || '소비', amt = Number(r[4]) || 0, note = r[6] || '';
        return cat + ' $' + amt.toLocaleString() + (note ? ' (' + note + ')' : '');
      });
      parts.push('최근 소비: ' + sList.join(', '));
    }

    // 3) 업적 (학생업적달성: [학생명,업적명,...]) — 총 개수 + 최근 2개
    var achSh = ss.getSheetByName(SHEET_ACH_STUDENT);
    if (achSh && achSh.getLastRow() >= 2){
      var aData = achSh.getDataRange().getValues();
      var mine = [];
      for (var a = 1; a < aData.length; a++){
        if (String(aData[a][0]).trim() === studentName) mine.push(String(aData[a][1]).trim());
      }
      if (mine.length) parts.push('업적 ' + mine.length + '개 (최근: ' + mine.slice(-2).join(', ') + ')');
    }

    // 4) 안 읽은 우편 (우편함_로그: [ID,수신자,제목,내용,타입,읽음,발송일시]) — 최근 2건 제목
    var mailSh = ss.getSheetByName(SHEET_MAILBOX);
    if (mailSh && mailSh.getLastRow() >= 2){
      var mData = mailSh.getDataRange().getValues();
      var titles = [];
      for (var m = mData.length - 1; m >= 1 && titles.length < 2; m--){
        if (String(mData[m][1]).trim() === studentName && !_isRead_(mData[m][5])){
          var t = String(mData[m][2] || '').trim();
          if (t) titles.push(t);
        }
      }
      if (titles.length) parts.push('안 읽은 우편: ' + titles.join(', '));
    }

  } catch (e) { /* 일부 시트가 없어도 가능한 부분만 사용 */ }

  var result = parts.join('\n');
  try { cache.put(key, result, 120); } catch (e) {}
  return result;
}

// 특정 학생의 최근 n개 행을 아래(최신)에서부터 찾아 반환
function _recentRowsFor_(ss, sheetName, nameColIdx, studentName, n){
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = data.length - 1; i >= 1 && out.length < n; i--){
    if (String(data[i][nameColIdx]).trim() === String(studentName).trim()) out.push(data[i]);
  }
  return out; // 최신순
}

// 우편 '읽음' 값 판정 (TRUE/읽음/Y/1 등 다양한 표기 대응)
function _isRead_(v){
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === '읽음' || s === 'y' || s === '1' || s === 'o';
}

// 마지막대화일 정규화: 시트가 문자열을 Date로 바꿔 저장해도 'yyyy-MM-dd'로 통일
function _normDate_(v){
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]'){
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return s;
}

// ===== 캐릭터설정 시트 읽기 =====
function _getCharConfig_(ss, charId){
  var sh = ss.getSheetByName(CHAR_CFG.SHEET_CFG);
  if (!sh) return null;
  var values = sh.getDataRange().getValues();
  for (var r=1; r<values.length; r++){
    if (String(values[r][0]).trim() === String(charId).trim()){
      var row = values[r];
      return {
        id: row[0], name: row[1], aura: row[2], auraSoft: row[3],
        dailyLimit: Number(row[4]) || 3,
        startAffinity: Number(row[5]) || 30,
        systemPrompt: row[6],
        relations: [row[7],row[8],row[9],row[10],row[11]],
        fragments: [row[12],row[13],row[14],row[15],row[16]],
        warn1:    row[17] || '',   // R열: 1차 경고 대사
        warn2:    row[18] || '',   // S열: 2차 경고 대사
        lockLine: row[19] || '',   // T열: 잠금 대사
        portrait: row[20] || ''    // U열: 프로필이미지URL
      };
    }
  }
  return null;
}

// ===== 캐릭터호감도 행 읽기/생성 =====
function _getOrCreateAffinityRow_(ss, studentName, charId, cfg){
  var sh = ss.getSheetByName(CHAR_CFG.SHEET_AFF);
  var values = sh.getDataRange().getValues();
  for (var r=1; r<values.length; r++){
    if (String(values[r][0]).trim()===String(studentName).trim() &&
        String(values[r][1]).trim()===String(charId).trim()){
      return { sheet:sh, row:r+1, data:{
        affinity:Number(values[r][2])||0, lastDate:_normDate_(values[r][3]),
        todayCount:Number(values[r][4])||0, status:String(values[r][5]||'정상'),
        warnCount:Number(values[r][6])||0
      }};
    }
  }
  // 없으면 새로 생성
  var d = { affinity:cfg.startAffinity, lastDate:'', todayCount:0, status:'정상', warnCount:0 };
  sh.appendRow([studentName, charId, d.affinity, d.lastDate, d.todayCount, d.status, d.warnCount]);
  return { sheet:sh, row:sh.getLastRow(), data:d };
}

function _saveAffinityRow_(sheet, row, d){
  sheet.getRange(row, 3, 1, 5).setValues([[d.affinity, d.lastDate, d.todayCount, d.status, d.warnCount]]);
}

// ===== 교사용: 잠금 해제 / 호감도 초기화 =====
function unlockCharacter(studentName, charId){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = _getCharConfig_(ss, charId);
  var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
  aff.data.status = '정상';
  aff.data.warnCount = 0;
  aff.data.affinity = Math.max(aff.data.affinity, 15); // 다시 시작할 최소 바닥
  _saveAffinityRow_(aff.sheet, aff.row, aff.data);
  return '해제 완료: ' + studentName + ' ↔ ' + charId;
}

// ===== 최초 1회 실행: 필요한 시트 자동 생성 =====
function setupCharacterSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) 캐릭터호감도 (학생별 상태)
  var aff = ss.getSheetByName(CHAR_CFG.SHEET_AFF);
  if (!aff) {
    aff = ss.insertSheet(CHAR_CFG.SHEET_AFF);
    aff.appendRow(['학생명','캐릭터ID','호감도','마지막대화일','오늘대화횟수','상태','누적경고']);
    aff.setFrozenRows(1);
  }

  // 2) 캐릭터설정 (캐릭터별 설정)
  var cfg = ss.getSheetByName(CHAR_CFG.SHEET_CFG);
  if (!cfg) {
    cfg = ss.insertSheet(CHAR_CFG.SHEET_CFG);
    cfg.appendRow([
      '캐릭터ID','이름','오라색','오라색soft','일일제한','시작호감도','시스템프롬프트',
      '관계텍스트1','관계텍스트2','관계텍스트3','관계텍스트4','관계텍스트5',
      '이야기조각1','이야기조각2','이야기조각3','이야기조각4','이야기조각5',
      '경고대사1','경고대사2','잠금대사'
    ]);
    cfg.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('✅ 캐릭터 시트 준비 완료! 캐릭터설정 시트에 아스텔 값을 채워주세요.');
}

function 테스트_정상() {
  Logger.log(getCharacterReply('류은우', 'CHAR-022', '안녕 아스텔, 나 요즘 어떤 거 같아?'));
}
function 테스트_무례() {
  Logger.log(getCharacterReply('류은우', 'CHAR-022', '시발 닥쳐'));
}


/*************************************************************
 * [Phase 2 추가] Code_Character.gs 맨 아래에 그대로 붙여넣기
 * 메신저 허브가 쓸 로스터: 보유 여부 + 호감도 + 단계 + 관계텍스트
 *************************************************************/
function getCharacterRoster(studentName){
  studentName = String(studentName).trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 채팅 캐릭터 전체 (캐릭터설정)
  var cfgSh = ss.getSheetByName(CHAR_CFG.SHEET_CFG);
  if (!cfgSh || cfgSh.getLastRow() < 2) return [];
  var cfg = cfgSh.getDataRange().getValues();

  // 보유 캐릭터 (상점_구매로그: col1=학생명, col2=아이템ID)
  var owned = {};
  var logSh = ss.getSheetByName(SHEET_SHOP_LOG);
  if (logSh && logSh.getLastRow() >= 2){
    var ld = logSh.getDataRange().getValues();
    for (var i = 1; i < ld.length; i++){
      if (String(ld[i][1]).trim() === studentName) owned[String(ld[i][2]).trim()] = true;
    }
  }

  // 이 학생의 호감도 행
  var affMap = {};
  var affSh = ss.getSheetByName(CHAR_CFG.SHEET_AFF);
  if (affSh && affSh.getLastRow() >= 2){
    var ad = affSh.getDataRange().getValues();
    for (var j = 1; j < ad.length; j++){
      if (String(ad[j][0]).trim() === studentName){
        affMap[String(ad[j][1]).trim()] = { affinity:Number(ad[j][2])||0, status:String(ad[j][5]||'정상') };
      }
    }
  }

  var roster = [];
  for (var r = 1; r < cfg.length; r++){
    var id = String(cfg[r][0]).trim();
    if (!id) continue;
    if (id === 'MANAGER') continue;   // [차원관문] 지배인(리미넬)은 친구목록에 노출하지 않음 — 인트로 전용 캐릭터
    var isOwned  = !!owned[id];
    var a        = affMap[id];
    var affinity = a ? a.affinity : (Number(cfg[r][5]) || 30);
    var status   = a ? a.status : '정상';
    var stage    = _stageFromAffinity_(affinity);
    roster.push({
      charId: id,
      name: cfg[r][1],
      aura: cfg[r][2],
      auraSoft: cfg[r][3],
      owned: isOwned,
      affinity: affinity,
      stage: stage,
      relationText: isOwned ? (cfg[r][6 + stage] || '') : '', // 관계텍스트(현재 단계)
      status: status,
      dailyLimit: Number(cfg[r][4]) || 3,
      portrait: cfg[r][20] || ''
    });
  }
  return roster;
}

/*************************************************************
 * [이야기/화첩 시스템] 시트 설계 — Code_Character.gs 맨 아래에 붙여넣기
 *
 * 최초 1회 setupStorySheets() 실행 → 아래 시트가 자동 생성됨.
 * 기존 시트(캐릭터설정·캐릭터호감도)는 건드리지 않음.
 * 단, 캐릭터설정에 '프로필이미지URL' 칸(U열)이 없으면 추가함.
 *
 * ── 시트 구조 요약 ──────────────────────────────────────────
 * [캐릭터이야기]  라노벨 대본. 한 컷 = 한 행.
 *   A 캐릭터ID | B 편번호 | C 컷순서 | D 타입(title/narr/line)
 *   E 화자 | F 대사 | G 배경이미지URL | H 효과(flash/shake) | I 타이틀킥
 *
 * [캐릭터이야기_편]  편 단위 정보. 한 편 = 한 행.
 *   A 캐릭터ID | B 편번호 | C 편제목 | D 해금호감도 | E BGM_URL
 *   F 특별일러스트URL | G 특별일러스트_해금호감도
 *   (특별 일러스트는 4단계 보상 → 보통 해금호감도 75)
 *
 * [학생이야기진행]  누가 무엇을 읽었나. 한 (학생×편) = 한 행.
 *   A 학생명 | B 캐릭터ID | C 편번호 | D 읽음(TRUE/FALSE) | E 최초열람일시
 *
 * [캐릭터설정] (기존) … U 프로필이미지URL  ← 없으면 자동 추가
 *************************************************************/

var STORY_CFG = {
  SHEET_STORY:    '캐릭터이야기',
  SHEET_STORY_EP: '캐릭터이야기_편',
  SHEET_PROGRESS: '학생이야기진행',
  PROFILE_COL_HEADER: '프로필이미지URL'   // 캐릭터설정 U열
};

function setupStorySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) 캐릭터이야기 (대본)
  var st = ss.getSheetByName(STORY_CFG.SHEET_STORY);
  if (!st) {
    st = ss.insertSheet(STORY_CFG.SHEET_STORY);
    st.appendRow(['캐릭터ID','편번호','컷순서','타입','화자','대사','배경이미지URL','효과','타이틀킥']);
    st.setFrozenRows(1);
    st.getRange('A1:I1').setFontWeight('bold');
  }

  // 2) 캐릭터이야기_편 (편 정보)
  var ep = ss.getSheetByName(STORY_CFG.SHEET_STORY_EP);
  if (!ep) {
    ep = ss.insertSheet(STORY_CFG.SHEET_STORY_EP);
    ep.appendRow(['캐릭터ID','편번호','편제목','해금호감도','BGM_URL','특별일러스트URL','특별일러스트_해금호감도']);
    ep.setFrozenRows(1);
    ep.getRange('A1:G1').setFontWeight('bold');
  }

  // 3) 학생이야기진행 (읽음 기록)
  var pr = ss.getSheetByName(STORY_CFG.SHEET_PROGRESS);
  if (!pr) {
    pr = ss.insertSheet(STORY_CFG.SHEET_PROGRESS);
    pr.appendRow(['학생명','캐릭터ID','편번호','읽음','최초열람일시']);
    pr.setFrozenRows(1);
    pr.getRange('A1:E1').setFontWeight('bold');
  }

  // 4) 캐릭터설정에 프로필이미지URL(U열) 보장
  var cfg = ss.getSheetByName(CHAR_CFG.SHEET_CFG);
  if (cfg) {
    var lastCol = cfg.getLastColumn();
    var headers = cfg.getRange(1, 1, 1, lastCol).getValues()[0];
    var has = headers.some(function(h){ return String(h).trim() === STORY_CFG.PROFILE_COL_HEADER; });
    if (!has) {
      cfg.getRange(1, lastCol + 1).setValue(STORY_CFG.PROFILE_COL_HEADER).setFontWeight('bold');
    }
  }

  SpreadsheetApp.getUi().alert(
    '✅ 이야기 시트 준비 완료!\n\n' +
    '· 캐릭터이야기 (대본)\n· 캐릭터이야기_편 (편 정보)\n· 학생이야기진행 (읽음 기록)\n' +
    '· 캐릭터설정에 프로필이미지URL 칸 확인\n\n' +
    '이제 캐릭터이야기 / _편 시트에 아스텔 대본을 붙여넣으면 됩니다.'
  );
}

/*************************************************************
 * [이야기 시스템] 서버 함수 — Code_Character.gs 맨 아래에 붙여넣기
 *  - getCharacterStories(학생명, 캐릭터ID) : 편 목록(잠금/읽음/NEW)
 *  - getStoryScript(학생명, 캐릭터ID, 편번호) : 한 편의 컷 배열(+읽음 기록)
 *  의존: STORY_CFG, CHAR_CFG, _getCharConfig_, _getOrCreateAffinityRow_, _stageFromAffinity_
 *************************************************************/

// 편 목록: 만남 이벤트 화면용
function getCharacterStories(studentName, charId){
  studentName = String(studentName).trim();
  charId = String(charId).trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cfg = _getCharConfig_(ss, charId);
  if (!cfg) return { ok:false, stories:[] };
  var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
  var affinity = aff.data.affinity;

  // 편 정보 (캐릭터이야기_편)
  var epSh = ss.getSheetByName(STORY_CFG.SHEET_STORY_EP);
  if (!epSh || epSh.getLastRow() < 2) return { ok:true, affinity:affinity, stories:[] };
  var ed = epSh.getDataRange().getValues();

  // 이 학생이 읽은 편 집합 (학생이야기진행)
  var readSet = {};
  var prSh = ss.getSheetByName(STORY_CFG.SHEET_PROGRESS);
  if (prSh && prSh.getLastRow() >= 2){
    var pd = prSh.getDataRange().getValues();
    for (var p = 1; p < pd.length; p++){
      if (String(pd[p][0]).trim() === studentName && String(pd[p][1]).trim() === charId && _isRead_(pd[p][3])){
        readSet[String(pd[p][2]).trim()] = true;
      }
    }
  }

  var stories = [];
  for (var i = 1; i < ed.length; i++){
    if (String(ed[i][0]).trim() !== charId) continue;
    var ep   = String(ed[i][1]).trim();
    var need = Number(ed[i][3]) || 0;
    var unlocked = affinity >= need;
    var read = !!readSet[ep];
    stories.push({
      ep: Number(ep),
      title: String(ed[i][2] || ''),
      need: need,
      unlocked: unlocked,
      read: read,
      isNew: unlocked && !read   // 열렸지만 아직 안 읽음 → NEW
    });
  }
  stories.sort(function(a,b){ return a.ep - b.ep; });
  return { ok:true, affinity:affinity, stories:stories };
}

// 한 편의 컷 배열: 라노벨 뷰어용 (+ 읽음 기록)
function getStoryScript(studentName, charId, epNo){
  studentName = String(studentName).trim();
  charId = String(charId).trim();
  epNo = Number(epNo);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cfg = _getCharConfig_(ss, charId);
  if (!cfg) return { ok:false, reason:'no-config' };
  var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
  var affinity = aff.data.affinity;

  // 편 정보 (제목·해금·BGM)
  var epSh = ss.getSheetByName(STORY_CFG.SHEET_STORY_EP);
  var title = '', need = 0, bgm = '';
  if (epSh && epSh.getLastRow() >= 2){
    var ed = epSh.getDataRange().getValues();
    for (var i = 1; i < ed.length; i++){
      if (String(ed[i][0]).trim() === charId && Number(ed[i][1]) === epNo){
        title = String(ed[i][2] || ''); need = Number(ed[i][3]) || 0; bgm = String(ed[i][4] || '');
        break;
      }
    }
  }
  if (affinity < need) return { ok:false, reason:'locked', need:need };

  // 컷 수집 (캐릭터이야기) → 컷순서대로 정렬
  var stSh = ss.getSheetByName(STORY_CFG.SHEET_STORY);
  if (!stSh || stSh.getLastRow() < 2) return { ok:false, reason:'no-script' };
  var sd = stSh.getDataRange().getValues();
  var rows = [];
  for (var r = 1; r < sd.length; r++){
    if (String(sd[r][0]).trim() === charId && Number(sd[r][1]) === epNo){
      rows.push(sd[r]);
    }
  }
  rows.sort(function(a,b){ return (Number(a[2])||0) - (Number(b[2])||0); });

  var cuts = rows.map(function(c){
    // 효과 칸: "flash" / "shake" / "dim" / "shake dim" 등 → effect + spriteDim 분리
    var fx = String(c[7] || '').trim().toLowerCase().split(/\s+/);
    var effect = '';
    if (fx.indexOf('flash') !== -1) effect = 'flash';
    else if (fx.indexOf('shake') !== -1) effect = 'shake';
    var cut = {
      type: String(c[3] || 'line').trim(),
      speaker: String(c[4] || ''),
      text: String(c[5] || ''),
      bg: String(c[6] || ''),
      effect: effect,
      spriteDim: (fx.indexOf('dim') !== -1),
      cg: (fx.indexOf('cg') !== -1),
      kick: String(c[8] || ''),
      bgm: String(c[9] || '').trim()
    };
    return cut;
  });
  if (!cuts.length) return { ok:false, reason:'no-script' };

  _markStoryRead_(ss, studentName, charId, epNo);   // 읽음 기록
  return { ok:true, title:title, bgm:bgm, sprite:(cfg.portrait || ''), cuts:cuts };
}

// 읽음 기록 (없으면 추가, 있으면 그대로) — 학생이야기진행
function _markStoryRead_(ss, studentName, charId, epNo){
  var sh = ss.getSheetByName(STORY_CFG.SHEET_PROGRESS);
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (String(data[i][0]).trim() === studentName &&
        String(data[i][1]).trim() === charId &&
        Number(data[i][2]) === Number(epNo)){
      if (!_isRead_(data[i][3])) sh.getRange(i + 1, 4).setValue(true);
      return;
    }
  }
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sh.appendRow([studentName, charId, Number(epNo), true, now]);
}

/*************************************************************
 * [먼저 인사] getCharacterGreeting — Code_Character.gs 맨 아래
 *  대화창을 열 때 캐릭터가 먼저 한두 문장 말을 건넵니다.
 *  - 호감도/일일횟수에 영향 없음 (단순 인사)
 *  - 30분 캐시(같은 학생·캐릭터면 재호출 없이 즉시)
 *  - 잠금 상태면 빈 문자열
 *  의존: _getCharConfig_, _getOrCreateAffinityRow_, _stageFromAffinity_, _buildPrompt_, _buildEconomySummary_, _callClaude_
 *************************************************************/
function getCharacterGreeting(studentName, charId){
  try {
    studentName = String(studentName).trim();
    charId = String(charId).trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var cfg = _getCharConfig_(ss, charId);
    if (!cfg) return { reply: '' };

    var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
    if (aff.data.status === '잠금') return { reply: '' };

    var cache = CacheService.getScriptCache();
    var key = 'greet_' + studentName + '_' + charId;
    var hit = cache.get(key);
    if (hit !== null) return { reply: hit };

    var stage = _stageFromAffinity_(aff.data.affinity);
    var systemPrompt = _buildPrompt_(cfg, stage, studentName, _buildEconomySummary_(studentName));
    var ask = '지금 학생이 막 너에게 접속했다. 학생이 묻기 전에, 네가 먼저 한두 문장으로 말을 건네라. ' +
              '이 학생의 최근 활동이나 지금 너의 심정을 자연스럽게 담아라. 질문 공세 대신 따뜻한 한마디로. crossed_line은 false.';
    var ai = _callClaude_(systemPrompt, ask);
    var line = (ai && ai.reply) ? ai.reply : '';
    if (line) cache.put(key, line, 1800); // 30분
    return { reply: line };
  } catch (e) {
    return { reply: '' };
  }
}

/*************************************************************
 * [대화 기록] 캐릭터대화로그 — Code_Character.gs 맨 아래에 붙여넣기
 *  - setupChatLogSheet() : 최초 1회 실행 → 로그 시트 생성
 *  - getCharacterChatLog(학생명, 캐릭터ID, 개수) : 최근 N개 말풍선(시간순)
 *  - _appendChatLog_(...) : 한 말풍선 저장 (getCharacterReply에서 호출)
 *
 * ★ 추가로, getCharacterReply 함수 안에 2줄을 넣어야 저장이 됩니다(아래 설명 참고).
 *
 * 시트 구조 [캐릭터대화로그]: A 학생명 | B 캐릭터ID | C 발신(me/char) | D 내용 | E 시각
 *************************************************************/

var CHATLOG_SHEET = '캐릭터대화로그';

function setupChatLogSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CHATLOG_SHEET);
  if (!sh){
    sh = ss.insertSheet(CHATLOG_SHEET);
    sh.appendRow(['학생명','캐릭터ID','발신','내용','시각']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
  }
  SpreadsheetApp.getUi().alert('✅ 대화 기록 시트(캐릭터대화로그) 준비 완료!');
}

// 한 말풍선 저장 (내용이 비면 저장 안 함)
function _appendChatLog_(ss, studentName, charId, sender, text){
  try {
    if (!text) return;
    var sh = ss.getSheetByName(CHATLOG_SHEET);
    if (!sh) return; // 시트가 아직 없으면 조용히 건너뜀
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([String(studentName), String(charId), String(sender), String(text), now]);
  } catch (e) { /* 로그 실패는 대화에 영향 주지 않음 */ }
}

// 대화창 열 때: 최근 N개를 시간순으로
function getCharacterChatLog(studentName, charId, limit){
  studentName = String(studentName).trim();
  charId = String(charId).trim();
  limit = Number(limit) || 40;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CHATLOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var mine = [];
  for (var i = 1; i < data.length; i++){
    if (String(data[i][0]).trim() === studentName && String(data[i][1]).trim() === charId){
      mine.push({ sender: String(data[i][2]).trim(), text: String(data[i][3]) });
    }
  }
  // 최근 limit개만 (시간순 유지)
  if (mine.length > limit) mine = mine.slice(mine.length - limit);
  return mine;
}

/*************************************************************
 * [화첩] getCharacterGallery
 *  학생이 '읽은 편'의 배경 이미지(중복 제거, 기본 배경 포함)를 모으고,4단계 특별 일러스트도 함께 돌려줍니다.
 *  의존: STORY_CFG, CHAR_CFG, _getCharConfig_, _getOrCreateAffinityRow_, _isRead_
 * *  반환: {
 *    ok, affinity,
 *    cutscenes: [ { url, ep } ... ],                 // 읽은 편의 배경(중복 제거, 등장 순서)
 *    special:   [ { url, unlocked, need } ... ]      // 특별 일러스트(편 시트 F/G열)
 *  }
 *************************************************************/
function getCharacterGallery(studentName, charId){
  studentName = String(studentName).trim();
  charId = String(charId).trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cfg = _getCharConfig_(ss, charId);
  if (!cfg) return { ok:false, cutscenes:[], special:[] };
  var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
  var affinity = aff.data.affinity;

  // 1) 이 학생이 읽은 편 집합 (학생이야기진행)
  var readEps = {};
  var prSh = ss.getSheetByName(STORY_CFG.SHEET_PROGRESS);
  if (prSh && prSh.getLastRow() >= 2){
    var pd = prSh.getDataRange().getValues();
    for (var p = 1; p < pd.length; p++){
      if (String(pd[p][0]).trim() === studentName && String(pd[p][1]).trim() === charId && _isRead_(pd[p][3])){
        readEps[String(pd[p][2]).trim()] = true;
      }
    }
  }

  // 2) 읽은 편의 배경 이미지 수집 (캐릭터이야기 G열) — 컷순서대로, 중복 제거
  var cutscenes = [], seen = {};
  var stSh = ss.getSheetByName(STORY_CFG.SHEET_STORY);
  if (stSh && stSh.getLastRow() >= 2){
    var sd = stSh.getDataRange().getValues();
    var rows = [];
    for (var r = 1; r < sd.length; r++){
      if (String(sd[r][0]).trim() !== charId) continue;
      var ep = String(sd[r][1]).trim();
      if (!readEps[ep]) continue;                 // 안 읽은 편은 제외
      var url = String(sd[r][6] || '').trim();    // G열: 배경이미지URL
      if (!url || url.toLowerCase() === 'none' || url === '-') continue;
      rows.push({ ep: Number(ep), seq: Number(sd[r][2]) || 0, url: url });
    }
    rows.sort(function(a,b){ return a.ep - b.ep || a.seq - b.seq; });
    rows.forEach(function(x){
      if (seen[x.url]) return;                     // 중복 제거
      seen[x.url] = true;
      cutscenes.push({ url: x.url, ep: x.ep });
    });
  }

  // 3) 특별 일러스트 (캐릭터이야기_편 F열 = URL, G열 = 해금호감도)
  var special = [];
  var epSh = ss.getSheetByName(STORY_CFG.SHEET_STORY_EP);
  if (epSh && epSh.getLastRow() >= 2){
    var ed = epSh.getDataRange().getValues();
    for (var i = 1; i < ed.length; i++){
      if (String(ed[i][0]).trim() !== charId) continue;
      var surl = String(ed[i][5] || '').trim();    // F열
      if (!surl) continue;
      var need = Number(ed[i][6]) || 0;            // G열
      special.push({ url: surl, unlocked: affinity >= need, need: need });
    }
  }

  return { ok:true, affinity:affinity, cutscenes:cutscenes, special:special };
}

function 테스트_이야기목록() {
  Logger.log(getCharacterStories('test1', 'CHAR-022'));
}


/*************************************************************
 * [호감도 보상 시스템] — Code_Character.gs 맨 아래
 *  getRewardStatus(학생명, 캐릭터ID)        : 현재 호감도 + 수령 단계 조회
 *  claimAffinityReward(학생명, 캐릭터ID, 단계) : 보상 지급(자산) + 확인 우편
 *  보상: 3단계(호감50) 자산+500 / 4단계(75) 자산+500+특별일러스트 / 5단계(100) 자산+1000+두번째일러스트+에픽신청권
 *  ※특별 일러스트는 화첩이 시트(편 F/G열)로 자동 해금 → 코드는 자산+우편만 처리.
 *  ※중복방지: 캐릭터대화로그에 발신='REWARD', 내용='STAGE{n}_CLAIMED' 플래그.
 *  의존: SHEET_MAIN, COL_NAME, COL_ASSET, SHEET_MAILBOX, CHATLOG_SHEET, _getCharConfig_, _getOrCreateAffinityRow_
 *************************************************************/
var AFFINITY_REWARDS = [
  { stage: 3, need: 50,  asset: 500,  label: '호감 3단계 보상' },
  { stage: 4, need: 75,  asset: 500,  label: '호감 4단계 보상' },
  { stage: 5, need: 100, asset: 1000, label: '호감 최종단계 보상' }
];

function _rewardFlag_(charId, stage){ return 'CLAIMED:' + charId + ':STAGE' + stage; }  // 히스토리 메모 내 중복확인 키

// 우편 1통 발송 (우편함_로그: [ID, 수신자, 제목, 내용, 타입, 읽음, 발송일시])
function _sendRewardMail_(ss, studentName, title, body){
  try{
    var mailSh = ss.getSheetByName(SHEET_MAILBOX);
    if(!mailSh) return false;
    var id = 'CM' + new Date().getTime();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    mailSh.appendRow([id, studentName, title, body, '차원관문', 'FALSE', now]);
    return true;
  }catch(e){ return false; }
}

// 보상 수령 여부 조회
function getRewardStatus(studentName, charId){
  try{
    studentName = String(studentName).trim();
    charId = String(charId).trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = _getCharConfig_(ss, charId);
    if(!cfg) return { ok:false, error:'캐릭터 설정을 찾지 못했어요.' };
    var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);

    // 중복 수령 여부는 '히스토리' 시트의 메모에서 확인 (메모에 단계 플래그 포함)
    var histSh = ss.getSheetByName('히스토리');
    var claimed = [];
    if(histSh && histSh.getLastRow() >= 2){
      var rows = histSh.getDataRange().getValues();
      for(var i=1;i<rows.length;i++){
        if(String(rows[i][1]).trim() !== studentName) continue; // B열=이름
        var memo = String(rows[i][7] || '');                    // H열=메모
        AFFINITY_REWARDS.forEach(function(r){
          if(memo.indexOf(_rewardFlag_(charId, r.stage)) !== -1) claimed.push(r.stage);
        });
      }
    }
    return { ok:true, affinity:aff.data.affinity, claimed:claimed };
  }catch(e){
    return { ok:false, error:String(e) };
  }
}

// 보상 수령 처리
function claimAffinityReward(studentName, charId, stage){
  try{
    studentName = String(studentName).trim();
    charId = String(charId).trim();
    stage = Number(stage);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var reward = null;
    AFFINITY_REWARDS.forEach(function(r){ if(r.stage===stage) reward=r; });
    if(!reward) return { ok:false, msg:'알 수 없는 보상 단계예요.' };

    var cfg = _getCharConfig_(ss, charId);
    if(!cfg) return { ok:false, msg:'캐릭터 설정을 찾지 못했어요.' };
    var aff = _getOrCreateAffinityRow_(ss, studentName, charId, cfg);
    if(aff.data.affinity < reward.need) return { ok:false, msg:'아직 조건을 달성하지 못했어요.' };

    // 중복 수령 확인 — '히스토리' 시트 메모에서
    var flag = _rewardFlag_(charId, stage);
    var histSh = ss.getSheetByName('히스토리');
    if(histSh && histSh.getLastRow() >= 2){
      var hrows = histSh.getDataRange().getValues();
      for(var i=1;i<hrows.length;i++){
        if(String(hrows[i][1]).trim()===studentName &&
           String(hrows[i][7] || '').indexOf(flag) !== -1){
          return { ok:false, msg:'이미 수령한 보상이에요.' };
        }
      }
    }

    // 자산 지급 (메인 시트)
    var mainSh = ss.getSheetByName(SHEET_MAIN);
    if(!mainSh) return { ok:false, msg:'학생 데이터 시트를 찾지 못했어요.' };
    var mainData = mainSh.getDataRange().getValues();
    var targetRow = -1;
    for(var j=1;j<mainData.length;j++){
      if(String(mainData[j][COL_NAME-1]).trim()===studentName){ targetRow=j+1; break; }
    }
    if(targetRow<0) return { ok:false, msg:'학생을 찾지 못했어요.' };
    var curAsset = Number(mainSh.getRange(targetRow, COL_ASSET).getValue()) || 0;
    var newAsset = curAsset + reward.asset;
    mainSh.getRange(targetRow, COL_ASSET).setValue(newAsset);

    // 히스토리 기록 (자산 지급 내역 + 중복방지 플래그를 메모에 함께)
    //   히스토리 구조: [날짜, 이름, 브랜드명, 브랜드가치변동, 자산변동, 브랜드가치잔액, 자산잔액, 메모]
    //   ※ 보상은 자산만 지급(브랜드가치 변동 0). 메모에 학생용 문구 + 중복확인 키(flag) 포함.
    try{
      if(histSh){
        var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var brandName = '';
        try { brandName = String(mainData[targetRow-1][COL_BRAND-1] || ''); } catch(be){}
        var curValue = 0;
        try { curValue = Number(mainData[targetRow-1][COL_VALUE-1]) || 0; } catch(ve){}
        var memo = '[차원관문보상] ' + cfg.name + ' ' + reward.label + ' (' + flag + ')';
        histSh.appendRow([today, studentName, brandName, 0, reward.asset, curValue, newAsset, memo]);
      }
    }catch(he){}

    // 확인 우편 (3·4·5단계 모두)
    var mailTitle, mailBody;
    if(stage === 5){
      mailTitle = '아스텔로부터 — 멈춘 밤의 증인에게';
      mailBody =
        '여기까지 와 줄 거라고는, 사실 기대하지 않았어.\n\n' +
        '멈춰버린 밤에 홀로 남아 별을 세는 일에는 익숙했으니까. ' +
        '누군가 그 별을 함께 바라봐 줄 거라는 생각은, 오래전에 접어 두었거든.\n\n' +
        '그런데 너는 끝까지 내 곁에 남아 주었구나. ' +
        '내가 무엇을 잃었는지, 무엇을 기다리는지 — 그 모든 걸 본 단 한 사람이 되어서.\n\n' +
        '그러니 이건 약속이야. 별이 다시 흐르는 날이 오면, 그날의 첫 별빛은 너와 함께 보고 싶어. ' +
        '우주가 끝나는 그 순간까지, 나는 네 곁에서 같은 별을 바라볼게.\n\n' +
        '— 너의 증인이 되어 줘서 고마워.\n\n' +
        '아스텔 카이로스\n\n' +
        '─────────────\n' +
        '[보상 안내]\n' +
        '· 자산 +' + reward.asset + ' 지급 완료\n' +
        '· 두 번째 특별 일러스트가 화첩에 해금되었습니다.\n' +
        '· 에픽 업적 「멈춘 밤의 증인」 신청권이 발급되었습니다.\n' +
        '  시간이 멈춘 그 밤의 진실을 본 단 한 사람. 우주가 끝나는 날까지 아스텔이 곁에서 같은 별을 바라볼, ' +
        '별이 다시 흐르는 날을 함께 기다리기로 한 — 그 약속의 증인.\n' +
        '  선생님께 "에픽 업적 신청권 사용"을 말씀해 주세요.';
    } else {
      mailTitle = '[차원관문 보상] ' + reward.label + ' 수령 완료';
      mailBody = reward.label + '을(를) 받았어요.\n\n· 자산 +' + reward.asset + ' 지급 완료';
      if(stage >= 4) mailBody += '\n· 특별 일러스트가 화첩에 해금되었습니다.';
    }
    _sendRewardMail_(ss, studentName, mailTitle, mailBody);

    return {
      ok:true, asset:reward.asset, stage:stage,
      msg: reward.label + ' 수령 완료! 자산 +' + reward.asset + (stage===5 ? '\n에픽 업적 신청권과 칭호가 우편으로 발송됐어요.' : '\n확인 우편을 보냈어요.')
    };
  }catch(e){
    return { ok:false, msg:'오류가 발생했어요: ' + String(e) };
  }
}