// ════════════════════════════════════════════════════════════════
// 업적검증도우미 시스템 (학생 전용 검증 페이지용 백엔드)
//
// ★ 사용법: 이 파일 전체를 복사해서 Code_Achievement.gs 맨 아래에 붙여넣으세요.
//
// 보안 원칙
//   - 도우미 전용 함수는 호출될 때마다 ① 비밀번호 ② 1인1역='업적검증도우미'
//     두 가지를 서버에서 반드시 재검증합니다.
//   - 도우미는 시트를 직접 보지 못하고, GAS가 걸러준 데이터만 받습니다.
//   - 도우미는 '추천'만 기록하고, 실제 업적 승인은 선생님이 합니다.
//
// 저장 위치
//   - 추천 결과는 '업적신청로그'(SHEET_ACH_LOG) 시트의 뒤쪽 3칸에만 기록됩니다.
//       F열(6): 도우미추천   ('승인추천' / '반려추천' / 빈칸)
//       G열(7): 추천메모      (도우미가 남긴 짧은 메모)
//       H열(8): 추천도우미명  (누가 추천했는지)
//   - 기존 A~E열(시각/이름/업적ID/증거/상태)은 전혀 건드리지 않습니다.
//   - 상태(E열)는 계속 '대기'로 유지되므로, 선생님의 기존 승인 흐름이
//     그대로 작동합니다.
// ════════════════════════════════════════════════════════════════


// ── [내부] 비밀번호 확인 (로그인_로그에 기록하지 않는 조용한 버전) ──
// 도우미 페이지는 데이터를 자주 불러오므로, 매번 로그인 기록이 쌓이지
// 않도록 verifyStudentPassword 와 별개로 만든 검증 전용 함수입니다.
function _checkStudentPasswordSilent_(studentName, password) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return false;
  const mainData = mainSheet.getDataRange().getValues();

  for (let i = 1; i < mainData.length; i++) {
    if (String(mainData[i][COL_NAME - 1]).trim() === String(studentName).trim()) {
      const correctPw = String(mainData[i][COL_PASSWORD - 1]).trim();
      const inputPw   = (password === null || password === undefined) ? null : String(password).trim();
      const masterPw  = _getMasterPassword();
      if (masterPw !== null && inputPw === masterPw) return true;  // 마스터키
      // 기존 verifyStudentPassword 와 동일한 통과 규칙
      if (inputPw !== null && correctPw && inputPw !== correctPw) return false;
      return true;
    }
  }
  return false;  // 학생을 찾지 못함
}


// ── [내부] 이 학생이 '업적검증도우미'인지 확인 (1인1역 시트 참조) ──
function _isVerifyHelper_(studentName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const jobSheet = ss.getSheetByName(SHEET_JOB);
  if (!jobSheet) return false;
  const jobData = jobSheet.getDataRange().getValues();
  for (let j = 1; j < jobData.length; j++) {
    if (String(jobData[j][0]).trim() === String(studentName).trim()) {
      const title = String(jobData[j][1] || '').trim();   // B열: 직업명
      return title.indexOf('업적검증도우미') !== -1;
    }
  }
  return false;
}


// ── [내부] 도우미 권한 통합 검사 (비번 + 역할) ──────────────────────
function _verifyHelperGuard_(studentName, password) {
  if (!_validateStudentName(studentName)) return { ok: false, msg: '유효하지 않은 이름입니다.' };
  if (!_validatePassword(password))       return { ok: false, msg: '유효하지 않은 비밀번호입니다.' };
  studentName = String(studentName).trim();
  if (!_checkStudentPasswordSilent_(studentName, password)) {
    return { ok: false, msg: '이름 또는 비밀번호가 일치하지 않습니다.' };
  }
  if (!_isVerifyHelper_(studentName)) {
    return { ok: false, msg: '업적검증도우미만 사용할 수 있는 페이지입니다.' };
  }
  return { ok: true, name: studentName };
}


// ── 도우미 로그인 (로그인 화면에서 호출) ────────────────────────────
function verifyHelperLogin(studentName, password) {
  const g = _verifyHelperGuard_(studentName, password);
  return { success: g.ok, msg: g.msg || '', helperName: g.ok ? g.name : '' };
}


// ── 도우미용: 대기 중인 업적 신청 + 검증 데이터 반환 ────────────────
// 각 신청 옆에 "신청자의 실제 숫자"(자산/브랜드가치/MVP/납세/기부 등)를
// 함께 내려보내 정량적 업적을 한눈에 판별할 수 있게 합니다.
// ════════════════════════════════════════════════════════════════
// 업적검증도우미 — 정량 근거 자동집계 버전 (getPendingForHelper 교체)
//
// ★ 사용법
//   기존 getPendingForHelper 함수를 통째로 지우고, 이 파일 내용 전체로
//   교체하세요. (_verifyHelperGuard_, _isVerifyHelper_, verifyHelperLogin,
//    helperRecommendAchievement 등 나머지 함수는 그대로 둡니다.)
//
//   _computeVerifyStats_ 가 신청 학생들의 모든 정량 근거를 한 번에 계산해
//   카드로 내려보냅니다. 어떤 근거를 보여줄지(업적별)는 HTML 쪽에서 고릅니다.
// ════════════════════════════════════════════════════════════════

function getPendingForHelper(studentName, password) {
  var g = _verifyHelperGuard_(studentName, password);
  if (!g.ok) return { success: false, msg: g.msg, pending: [] };

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet    = ss.getSheetByName(SHEET_ACH_LOG);
  var masterSheet = ss.getSheetByName(SHEET_ACH_MASTER);
  if (!logSheet) return { success: true, helperName: g.name, pending: [] };

  // 1) 업적마스터: ID → {명, 조건, 판단여부(가능/불가/자동)}
  //    업적마스터 컬럼: A=업적ID B=업적명 C=달성조건 D=업적도우미 판단가능 여부
  var achInfo = {};
  if (masterSheet) {
    var mData = masterSheet.getDataRange().getValues();
    for (var m = 1; m < mData.length; m++) {
      if (!mData[m][0]) continue;
      achInfo[String(mData[m][0]).trim()] = {
        name:      String(mData[m][1]).trim(),
        condition: String(mData[m][2]).trim(),
        judge:     String(mData[m][3] || '').trim()   // 가능 / 불가 / 자동
      };
    }
  }

  // 2) 대기 신청 행 수집 + 신청 학생 집합
  var logData = logSheet.getDataRange().getValues();
  var rows = [];
  var applicantSet = {};
  for (var i = 1; i < logData.length; i++) {
    if (String(logData[i][4]).trim() !== '대기') continue;   // E열 상태
    var nm = String(logData[i][1]).trim();
    applicantSet[nm] = true;
    rows.push({ rowIdx: i, name: nm });
  }

  // 3) 신청 학생들의 정량 근거 일괄 계산
  var statsMap = _computeVerifyStats_(applicantSet);

  // 4) 카드 조립
  var pending = [];
  rows.forEach(function(r) {
    var i = r.rowIdx;
    var achId = String(logData[i][2]).trim();
    var isSpecial = (achId === '특별보고');
    var info = achInfo[achId] || { name: '(미등록)', condition: '', judge: '' };
    pending.push({
      rowNumber:   i + 1,
      timestamp:   String(logData[i][0]),
      studentName: r.name,
      achId:       achId,
      achName:     isSpecial ? '특별보고' : info.name,
      condition:   isSpecial ? '(선생님이 직접 판단하는 특별보고)' : info.condition,
      judge:       isSpecial ? '불가' : (info.judge || ''),
      isSpecial:   isSpecial,
      proof:       String(logData[i][3]),
      stats:       statsMap[r.name] || _emptyStats_(),
      myRec:       String(logData[i][5] || '').trim(),   // F
      myMemo:      String(logData[i][6] || '').trim(),   // G
      recBy:       String(logData[i][7] || '').trim()    // H
    });
  });

  return { success: true, helperName: g.name, pending: pending, totalStudents: _countStudents_() };
}


// ── 빈 통계 객체 ────────────────────────────────────────────────
function _emptyStats_() {
  return {
    asset:0, honor:0, mvp:0, tax:0,
    depositPrincipal:0, loanBalance:0, netWorth:0, hasActiveLoan:false,
    donationTotal:0, donationMaxDay:0,
    spend7:0, spend30:0,
    snackLastDate:'', snackCount14:0,
    shopMarketTotal:0, shopMarketLastDate:'',
    p2pSent:0, p2pRecv:0, p2pPartners:0, p2pOver200:0, p2pMaxSpend:0,
    sellerRateAvg:0, sellerRateCnt:0,
    interestTotal:0, interestMaxOne:0, hasMaturedInterest:false,
    creditMax:0, creditLatest:0,
    auctionWinTotal:0, auctionWinMaxDay:0,
    maxPathsOneDay:0, mvpFromHistory:0,
    monthP2PSell:0, monthP2PBuy:0,
    achCount:0
  };
}

function _countStudents_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  if (!sh) return 0;
  return Math.max(0, sh.getLastRow() - 1);
}


// ── 정량 근거 일괄 계산 (신청 학생들만) ─────────────────────────
function _computeVerifyStats_(applicantSet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var S  = {};
  Object.keys(applicantSet).forEach(function(nm){ S[nm] = _emptyStats_(); });
  function has(nm){ return S.hasOwnProperty(nm); }

  function ymd(d){
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var now = new Date();
  var d7  = new Date(now); d7.setDate(d7.getDate()-7);
  var d14 = new Date(now); d14.setDate(d14.getDate()-14);
  var d30 = new Date(now); d30.setDate(d30.getDate()-30);
  var thisMonth = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM');

  // (A) 메인: 자산/브랜드가치/MVP/납세
  var main = ss.getSheetByName(SHEET_MAIN);
  if (main) {
    var md = main.getDataRange().getValues();
    for (var i=1;i<md.length;i++){
      var nm = String(md[i][COL_NAME-1]).trim();
      if (!has(nm)) continue;
      S[nm].asset = Number(md[i][COL_ASSET-1])||0;
      S[nm].honor = Number(md[i][COL_VALUE-1])||0;
      S[nm].mvp   = Number(md[i][COL_MVP-1])||0;
      S[nm].tax   = Number(md[i][COL_TAX-1])||0;
    }
  }

  // (B) 예금: 진행중 원금 + 이자
  var dep = ss.getSheetByName(SHEET_DEPOSIT_LOG);  // 학생별가입예금
  if (dep) {
    var dd = dep.getDataRange().getValues();
    for (var i=1;i<dd.length;i++){
      var nm = String(dd[i][1]).trim();      // B 학생명
      if (!has(nm)) continue;
      var principal = Number(dd[i][2])||0;   // C 원금
      var status    = String(dd[i][7]).trim(); // H 상태
      var interest  = Number(dd[i][8])||0;   // I 지급이자액
      if (status === '진행중') S[nm].depositPrincipal += principal;
      if (interest > 0) {
        S[nm].interestTotal += interest;
        if (interest > S[nm].interestMaxOne) S[nm].interestMaxOne = interest;
        if (status === '만기') S[nm].hasMaturedInterest = true;
      }
    }
  }

  // (C) 대출: 잔여원금 + 활성여부
  var loan = ss.getSheetByName(SHEET_LOAN_STATUS); // 대출현황
  if (loan) {
    var ld = loan.getDataRange().getValues();
    for (var i=1;i<ld.length;i++){
      var nm = String(ld[i][1]).trim();      // B 학생명
      if (!has(nm)) continue;
      var bal    = Number(ld[i][10])||0;     // K 잔여원금
      var status = String(ld[i][8]).trim();  // I 상환상태
      S[nm].loanBalance += bal;
      if (status !== '완료' && (bal > 0 || status === '연체' || status === '정상')) {
        S[nm].hasActiveLoan = true;
      }
    }
  }

  // 순자산 = 자산 + 진행중예금원금 − 대출잔액
  Object.keys(S).forEach(function(nm){
    S[nm].netWorth = S[nm].asset + S[nm].depositPrincipal - S[nm].loanBalance;
  });

  // (D) 자산사용: 기부/소비/간식/상점·거래소
  var spend = ss.getSheetByName(SHEET_SPEND);
  if (spend && spend.getLastRow() >= 2) {
    var sd = spend.getDataRange().getValues();
    var donDay = {};  // nm|date -> 합
    for (var i=1;i<sd.length;i++){
      var dt  = sd[i][0];                       // A 날짜
      var nm  = String(sd[i][1]).trim();        // B 이름
      if (!has(nm)) continue;
      var cat = String(sd[i][3]||'').trim();    // D 사용항목
      var amt = Number(sd[i][4])||0;            // E 금액
      var dd2 = (dt instanceof Date) ? dt : new Date(dt);

      // 기부
      if (cat === '기부') {
        S[nm].donationTotal += amt;
        var k = nm + '|' + ymd(dd2);
        donDay[k] = (donDay[k]||0) + amt;
      }
      // 소비(전체 자산사용 합) — 최근 7/30일
      if (dd2 instanceof Date && !isNaN(dd2.getTime())) {
        if (dd2 >= d7)  S[nm].spend7  += amt;
        if (dd2 >= d30) S[nm].spend30 += amt;
      }
      // 간식 (D열에 [간식] 포함)
      if (cat.indexOf('[간식]') !== -1) {
        if (dd2 >= d14) S[nm].snackCount14++;
        var sdate = ymd(dd2);
        if (sdate > S[nm].snackLastDate) S[nm].snackLastDate = sdate;
      }
      // 상점 + 물품거래소 (ECO-032/034)
      if (cat.indexOf('상점구매') !== -1 || cat.indexOf('[물품구매]') !== -1) {
        S[nm].shopMarketTotal += amt;
        var mdate = ymd(dd2);
        if (mdate > S[nm].shopMarketLastDate) S[nm].shopMarketLastDate = mdate;
      }
    }
    Object.keys(donDay).forEach(function(k){
      var nm = k.split('|')[0];
      if (has(nm) && donDay[k] > S[nm].donationMaxDay) S[nm].donationMaxDay = donDay[k];
    });
  }

  // (E) P2P거래로그: 송금/수령/상대수/200+/최대지출/판매자평점/이번달
  var p2p = ss.getSheetByName(SHEET_P2P);
  if (p2p && p2p.getLastRow() >= 2) {
    var pd = p2p.getDataRange().getValues();
    var partners = {}; // nm -> {상대명:true}
    Object.keys(S).forEach(function(nm){ partners[nm] = {}; });
    for (var i=1;i<pd.length;i++){
      var dt   = pd[i][1];                    // B 날짜
      var snd  = String(pd[i][2]).trim();     // C 보내는(구매자/지불)
      var rcv  = String(pd[i][3]).trim();     // D 받는(판매자/수취)
      var amt  = Number(pd[i][4])||0;         // E 금액
      var rate = Number(pd[i][9])||0;         // J 평점
      var dd3  = (dt instanceof Date) ? dt : new Date(dt);
      var ym   = (dd3 instanceof Date && !isNaN(dd3.getTime())) ? Utilities.formatDate(dd3,'Asia/Seoul','yyyy-MM') : '';

      if (has(snd)) {
        S[snd].p2pSent++;
        if (amt > S[snd].p2pMaxSpend) S[snd].p2pMaxSpend = amt;
        if (rcv) partners[snd][rcv] = true;
        if (amt >= 200) S[snd].p2pOver200++;
        if (ym === thisMonth) S[snd].monthP2PBuy += amt;   // 보낸 돈 = 구매
      }
      if (has(rcv)) {
        S[rcv].p2pRecv++;
        if (snd) partners[rcv][snd] = true;
        if (amt >= 200 && snd !== rcv) S[rcv].p2pOver200++;
        if (ym === thisMonth) S[rcv].monthP2PSell += amt;  // 받은 돈 = 판매
        if (rate > 0) {                                     // 판매자(받는쪽)가 평점 대상
          S[rcv].sellerRateAvg = (S[rcv].sellerRateAvg * S[rcv].sellerRateCnt + rate) / (S[rcv].sellerRateCnt + 1);
          S[rcv].sellerRateCnt++;
        }
      }
    }
    Object.keys(S).forEach(function(nm){
      S[nm].p2pPartners = Object.keys(partners[nm]).length;
      S[nm].sellerRateAvg = Math.round(S[nm].sellerRateAvg * 100) / 100;
    });
  }

  // (F) 신용점수이력: 최고/최신 총점
  var cr = ss.getSheetByName(SHEET_CREDIT_HISTORY);
  if (cr && cr.getLastRow() >= 2) {
    var cd = cr.getDataRange().getValues();
    var latestDate = {}; // nm -> Date
    for (var i=1;i<cd.length;i++){
      var nm = String(cd[i][1]).trim();   // B 학생명
      if (!has(nm)) continue;
      var score = Number(cd[i][6])||0;    // G 총점
      var dt    = cd[i][0];               // A 기준일
      var dd4   = (dt instanceof Date) ? dt : new Date(dt);
      if (score > S[nm].creditMax) S[nm].creditMax = score;
      if (!latestDate[nm] || (dd4 instanceof Date && dd4 >= latestDate[nm])) {
        latestDate[nm] = dd4;
        S[nm].creditLatest = score;
      }
    }
  }

  // (G) 히스토리: 경매낙찰/하루경로/MVP누적
  var hist = ss.getSheetByName(SHEET_HISTORY);
  if (hist && hist.getLastRow() >= 2) {
    var hd = hist.getDataRange().getValues();
    var auctionDay = {}; // nm|date -> 낙찰수
    var pathsDay   = {}; // nm|date -> {경로:true}
    for (var i=1;i<hd.length;i++){
      var dt  = hd[i][0];                      // A 날짜
      var nm  = String(hd[i][1]).trim();       // B 이름
      if (!has(nm)) continue;
      var gain= Number(hd[i][3])||0;           // D 당일지급점수
      var note= String(hd[i][7]||'').trim();   // H 비고
      var day = ymd((dt instanceof Date)?dt:new Date(dt));

      if (note.indexOf('[경매낙찰]') !== -1) {
        var ak = nm+'|'+day; auctionDay[ak] = (auctionDay[ak]||0)+1;
        S[nm].auctionWinTotal++;
      }
      if (note.indexOf('[MVP]') !== -1) S[nm].mvpFromHistory += gain;

      // 하루 획득 경로 (양(+)의 지급만, 경로=태그 대분류)
      if (gain > 0 && day) {
        var cat = _pathCategory_(note);
        if (cat) {
          var pk = nm+'|'+day;
          if (!pathsDay[pk]) pathsDay[pk] = {};
          pathsDay[pk][cat] = true;
        }
      }
    }
    Object.keys(auctionDay).forEach(function(k){
      var nm = k.split('|')[0];
      if (has(nm) && auctionDay[k] > S[nm].auctionWinMaxDay) S[nm].auctionWinMaxDay = auctionDay[k];
    });
    Object.keys(pathsDay).forEach(function(k){
      var nm = k.split('|')[0];
      var c = Object.keys(pathsDay[k]).length;
      if (has(nm) && c > S[nm].maxPathsOneDay) S[nm].maxPathsOneDay = c;
    });
  }

  // (H) 달성 업적 수
  var achSt = ss.getSheetByName(SHEET_ACH_STUDENT);
  if (achSt && achSt.getLastRow() >= 2) {
    var asd = achSt.getDataRange().getValues();
    for (var i=1;i<asd.length;i++){
      var nm = String(asd[i][0]).trim();
      if (has(nm)) S[nm].achCount++;
    }
  }

  return S;
}

// 히스토리 비고 → 획득 경로 대분류 (ECO-009 판정용)
function _pathCategory_(note) {
  if (!note) return '';
  if (note.indexOf('[MVP]') !== -1)            return 'MVP';
  if (note.indexOf('일일퀘스트') !== -1)        return '일일퀘스트';
  if (note.indexOf('업적') !== -1)             return '업적보상';
  if (note.indexOf('기부') !== -1)             return '기부보상';
  if (note.indexOf('[경매낙찰]') !== -1)        return '경매';
  if (note.indexOf('[수업]') !== -1)           return '수업';
  if (note.indexOf('[과제]') !== -1)           return '과제';
  if (note.indexOf('차원관문') !== -1)         return '차원관문';
  if (note.indexOf('보너스') !== -1)           return '보너스';
  if (note.indexOf('[생활]') !== -1)           return '생활';
  if (note.indexOf('[P2P') !== -1)             return '';   // P2P 이동은 경로로 보지 않음
  return '';
}


// ── 도우미용: 추천 기록(승인추천/반려추천/취소) ─────────────────────
// recommendation: '승인추천' | '반려추천' | '' (빈 문자열이면 추천 취소)
function helperRecommendAchievement(studentName, password, rowNumber, recommendation, memo) {
  const g = _verifyHelperGuard_(studentName, password);
  if (!g.ok) return { success: false, msg: g.msg };

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_ACH_LOG);
  if (!logSheet) return { success: false, msg: '업적신청로그 시트를 찾을 수 없습니다.' };

  rowNumber = Number(rowNumber);
  if (!(rowNumber >= 2)) return { success: false, msg: '유효하지 않은 신청입니다.' };

  // 선생님이 이미 처리(승인/반려)한 건은 추천 불가 → 새로고침 유도
  const status = String(logSheet.getRange(rowNumber, 5).getValue()).trim();
  if (status !== '대기') {
    return { success: false, msg: '이미 선생님이 처리한 신청입니다. 새로고침 해주세요.' };
  }

  let rec = String(recommendation || '').trim();
  if (rec !== '승인추천' && rec !== '반려추천' && rec !== '') {
    return { success: false, msg: '추천 값이 올바르지 않습니다.' };
  }
  const cleanMemo = String(memo || '').trim().slice(0, 200);

  logSheet.getRange(rowNumber, 6).setValue(rec);                  // F: 도우미추천
  logSheet.getRange(rowNumber, 7).setValue(rec ? cleanMemo : ''); // G: 추천메모
  logSheet.getRange(rowNumber, 8).setValue(rec ? g.name : '');    // H: 추천도우미명

  return {
    success: true,
    msg: rec ? (rec + ' 저장 완료') : '추천이 취소되었습니다.',
    myRec: rec, myMemo: rec ? cleanMemo : '', recBy: rec ? g.name : ''
  };
}