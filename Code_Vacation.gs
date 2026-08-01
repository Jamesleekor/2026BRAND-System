// ================================================================
// Code_Vacation.gs — 여름의 별자리 (방학 이벤트, 2026-07-27 ~ 08-23)
// ================================================================
// 이 파일 하나에 방학 이벤트 로직이 전부 들어 있습니다.
// 기존 Code.gs 에는 getStudentData 안에 2줄만 추가하면 됩니다. (안내문 참조)
// 방학 기간이 아니면 모든 함수가 즉시 종료되므로 평상시 성능에 영향 없음.
// ================================================================

// ── [설정] 여기 숫자만 바꾸면 규칙이 바뀝니다 ──────────────────────
const VAC_CFG = {
  START:        '2026-07-27',   // 방학 시작일
  END:          '2026-08-23',   // 방학 종료일
  MAKEUP_START: '2026-08-24',   // 보충 기간 시작 (개학 후 1주)
  MAKEUP_END:   '2026-08-28',   // 보충 기간 종료
  PROPOSAL_OPEN:'2026-08-03',   // 방학 시작 7일 후 제안소 공개
  PROPOSAL_MIN_CHARS: 20,       // 제안 본문 최소 글자 수
  PROPOSAL_LIST_LIMIT: 100,     // 공개 목록 최대 표시 수
  TOTAL_CELLS:  28,             // 별 개수 (= 방학 일수)
  DAILY_GOLD:   50,             // 매일 출석 골드
  DAILY_TOKEN:  1,              // 매일 출석 토큰
  STREAK_BONUS: { 3: 2, 7: 5, 14: 10, 21: 15 },  // 연속일수: 토큰
  COMPLETE_TOKEN: 20,           // 별자리 완성(28칸) 토큰
  RESTORE_TICKETS: 2,           // 학생당 복구권 개수
  MAKEUP_MAX_PER_DAY: 3,        // 보충 기간 하루 최대 칸 수
  // 학급 공동 진행도 단계 보상 (누적칸 합계 기준)
  MILESTONES: [
    { at: 100, type: '토큰', value: 3,   label: '전원 토큰 +3' },
    { at: 200, type: '골드', value: 200, label: '전원 골드 +200' },
    { at: 300, type: '토큰', value: 7,   label: '전원 토큰 +7' },
    { at: 400, type: '수동', value: 0,   label: '개학날 간식 1종 추가 (복지기금·교사 수동)' },
    { at: 500, type: '수동', value: 0,   label: '학급 공동 업적 「함께 본 여름 하늘」 (교사 수동)' }
  ],
  // 보물 칸 랜덤 보상 풀 (w = 가중치, 클수록 잘 나옴)
  TREASURE_POOL: [
    { label: '토큰 +3',        type: '토큰', value: 3,   w: 30 },
    { label: '토큰 +5',        type: '토큰', value: 5,   w: 15 },
    { label: '골드 +100',      type: '골드', value: 100, w: 35 },
    { label: '골드 +200',      type: '골드', value: 200, w: 15 },
    { label: '랜덤 상자 쿠폰',  type: '쿠폰', value: 1,   w: 5 }
  ],
  LIBERATION_DATE: '2026-08-15' // 히든 업적 「빛을 되찾은 날」 대상일
};

// ── [설정] 시트 이름 ──────────────────────────────────────────────
const SHEET_VAC_ATT   = '방학_출석';
const SHEET_VAC_HIST  = '방학_출석_히스토리';
const SHEET_VAC_MAP   = '방학_별자리';       // 칸별 유형·내용 (교사가 직접 편집)
const SHEET_VAC_PROP  = '방학_제안함';
const SHEET_VAC_PROP_REC = '방학_제안추천';
const SHEET_VAC_FINISH = '방학_완주자';
const SHEET_VAC_Q     = '방학_질문게시판';

// 방학_출석 열 번호 (1-indexed)
const VC = {
  NAME: 1,        // A 이름
  LAST_DATE: 2,   // B 마지막출석일
  CELLS: 3,       // C 누적칸수  ← 공동 진행도는 이 열의 합계
  STREAK: 4,      // D 연속일수
  TICKETS: 5,     // E 복구권잔여
  BEST: 6,        // F 최고연속 (연속 보너스 지급 기준)
  TOKENS: 7,      // G 토큰보유
  BONUS_LOG: 8,   // H 받은보너스기록 (예: "3,7,완주")
  TREASURE: 9,    // I 보물기록
  MAKEUP_DATE: 10,// J 보충일자
  MAKEUP_CNT: 11, // K 보충횟수(해당일)
  PREV_STREAK: 12,// L 직전연속 (복구용)
  RESTORE_DATE: 13,// M 복구가능일 (놓친 날짜, 다음날까지만 유효)
  LIB_NOTE: 14    // N 광복절소감 제출여부
};

// ════════════════════════════════════════════════════════════════
// 0. 최초 1회 실행: 시트 생성
// ════════════════════════════════════════════════════════════════
function setupVacationSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 방학_출석 — 메인 시트의 학생 명단으로 초기화
  if (!ss.getSheetByName(SHEET_VAC_ATT)) {
    const sh = ss.insertSheet(SHEET_VAC_ATT);
    sh.appendRow(['이름','마지막출석일','누적칸수','연속일수','복구권잔여',
                  '최고연속','토큰보유','받은보너스기록','보물기록',
                  '보충일자','보충횟수','직전연속','복구가능일','광복절소감']);
    // ★ 날짜가 들어가는 열(B,J,M)은 텍스트 서식으로 고정 —
    //   구글시트가 "2026-07-27" 문자열을 자동으로 날짜(Date) 타입으로
    //   바꿔버리면 다음 날 비교가 어긋나 연속일수가 올라가지 않는 문제가 생깁니다.
    sh.getRange(2, VC.LAST_DATE,   500, 1).setNumberFormat('@');
    sh.getRange(2, VC.MAKEUP_DATE, 500, 1).setNumberFormat('@');
    sh.getRange(2, VC.RESTORE_DATE,500, 1).setNumberFormat('@');
    const mainData = ss.getSheetByName(SHEET_MAIN).getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < mainData.length; i++) {
      const name = String(mainData[i][COL_NAME - 1]).trim();
      if (!name) continue;
      rows.push([name, '', 0, 0, VAC_CFG.RESTORE_TICKETS, 0, 0, '', '', '', 0, 0, '', '']);
    }
    if (rows.length) sh.getRange(2, 1, rows.length, 14).setValues(rows);
  }

  // ② 방학_출석_히스토리 — A열 날짜 / I열 타임스탬프 (기존 규칙 동일)
  if (!ss.getSheetByName(SHEET_VAC_HIST)) {
    const sh = ss.insertSheet(SHEET_VAC_HIST);
    sh.appendRow(['날짜','이름','유형','골드','토큰','칸번호','비고','예비','타임스탬프']);
  }

  // ③ 방학_별자리 — 칸별 내용. 생성 후 교사가 제목·내용을 직접 채우면 됨
  if (!ss.getSheetByName(SHEET_VAC_MAP)) {
    const sh = ss.insertSheet(SHEET_VAC_MAP);
    sh.appendRow(['칸번호','유형','제목','내용']);
    // 기본 배치: 기록 5(4,8,12,16,24) / 떡밥 4(6,13,20,26) / 보물 2(10,22) / 최종 1(28) / 나머지 일반
    const special = { 4:'기록', 8:'기록', 12:'기록', 16:'기록', 24:'기록',
                      6:'떡밥', 13:'떡밥', 20:'떡밥', 26:'떡밥',
                      10:'보물', 22:'보물', 28:'최종' };
    const rows = [];
    for (let n = 1; n <= VAC_CFG.TOTAL_CELLS; n++) {
      rows.push([n, special[n] || '일반', '', '']);
    }
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  // ④ 방학_제안함·의회 추천·완주자 기록
  // 기존 7열 제안함도 자동으로 새 구조로 확장됩니다.
  _vacEnsureProposalSheets_(ss);

  // ⑤ 방학_질문게시판 — 교사는 E열에 답변을 적고 F열 상태를 '답변완료'로 변경
  if (!ss.getSheetByName(SHEET_VAC_Q)) {
    const sh = ss.insertSheet(SHEET_VAC_Q);
    sh.appendRow(['등록일','이름','분류','질문','답변','상태','답변일','타임스탬프']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 500, 8).setWrap(true);
    sh.setColumnWidth(2, 90);
    sh.setColumnWidth(3, 120);
    sh.setColumnWidth(4, 360);
    sh.setColumnWidth(5, 360);
  }

  return '✅ 방학 시트 준비 완료! 제안함·의회 추천·완주자 기록 시트까지 최신 구조로 정리했습니다.';
}

// ════════════════════════════════════════════════════════════════
// 1. 로그인 훅 — getStudentData 에서 딱 이 함수 하나만 호출
//    방학 기간이 아니면 null 반환 (시트를 읽지도 않음 → 평상시 비용 0)
// ════════════════════════════════════════════════════════════════
function vacationOnLogin(studentName) {
  const today = _vacToday_(studentName);
  const inVacation = (today >= VAC_CFG.START && today <= VAC_CFG.END);
  const inMakeup   = (today >= VAC_CFG.MAKEUP_START && today <= VAC_CFG.MAKEUP_END);
  if (!inVacation && !inMakeup) return null;

  try {
    return inVacation
      ? _vacCheckIn_(studentName, today)
      : _vacMakeup_(studentName, today);
  } catch (e) {
    // 방학 로직 오류가 로그인 자체를 막으면 안 되므로 조용히 넘어감
    return { active: true, error: true };
  }
}

// ── 방학 중 출석 처리 (하루 1회) ─────────────────────────────────
function _vacCheckIn_(studentName, today) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_ATT);
  if (!sh) return null;

  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][VC.NAME - 1]).trim() === studentName) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return null;

  const row = data[rowIdx];
  const lastDate = _vacDateKey_(row[VC.LAST_DATE - 1]);
  const cells    = Number(row[VC.CELLS - 1]) || 0;

  // 이미 오늘 출석했으면 읽기만 하고 종료 (쓰기 없음 → 가벼움)
  if (lastDate === today) {
    return _vacStatusLight_(row, today, false);
  }
  // 별자리를 이미 완성했으면 더 열 칸이 없음
  if (cells >= VAC_CFG.TOTAL_CELLS) {
    return _vacStatusLight_(row, today, false);
  }

  // ── 여기부터 쓰기 작업: Lock 필수 ──
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Lock 획득 후 재확인 (동시 접속 대비)
    const cur = sh.getRange(rowIdx + 1, 1, 1, 14).getValues()[0];
    if (_vacDateKey_(cur[VC.LAST_DATE - 1]) === today) {
      return _vacStatusLight_(cur, today, false);
    }

    const yesterday = _vacAddDays_(today, -1);
    const prevStreak = Number(cur[VC.STREAK - 1]) || 0;
    let   streak, restoreDate = '', prevStreakSave = 0;

    if (_vacDateKey_(cur[VC.LAST_DATE - 1]) === yesterday || !cur[VC.LAST_DATE - 1]) {
      // 어제 출석했거나 첫 출석 → 연속 이어짐
      streak = prevStreak + 1;
    } else if (_vacDateKey_(cur[VC.LAST_DATE - 1]) === _vacAddDays_(today, -2)) {
      // 딱 하루 놓침 → 연속 리셋하되, 복구권으로 되살릴 수 있게 기록해 둠
      streak = 1;
      prevStreakSave = prevStreak;
      restoreDate = yesterday;
    } else {
      // 이틀 이상 놓침 → 복구 불가, 연속 리셋 (누적칸은 그대로!)
      streak = 1;
    }

    const newCells = Number(cur[VC.CELLS - 1]) + 1;
    let   tokens   = Number(cur[VC.TOKENS - 1]) + VAC_CFG.DAILY_TOKEN;
    let   bonusLog = String(cur[VC.BONUS_LOG - 1] || '');
    let   best     = Number(cur[VC.BEST - 1]) || 0;
    const messages = [];

    // 연속 보너스: '최고 기록 갱신' 방식 — 한 번 받은 구간은 재지급 안 됨
    if (streak > best) {
      for (const k in VAC_CFG.STREAK_BONUS) {
        const th = Number(k);
        if (streak >= th && best < th) {
          tokens += VAC_CFG.STREAK_BONUS[k];
          bonusLog = bonusLog ? bonusLog + ',' + th : String(th);
          messages.push(`🔥 ${th}일 연속! 토큰 +${VAC_CFG.STREAK_BONUS[k]}`);
        }
      }
      best = streak;
    }

    // 보물 칸이면 랜덤 보상
    let treasureLog = String(cur[VC.TREASURE - 1] || '');
    const cellType = _vacCellType_(ss, newCells);
    let extraGold = 0;
    if (cellType === '보물') {
      const prize = _vacRollTreasure_();
      if (prize.type === '토큰') tokens += prize.value;
      if (prize.type === '골드') extraGold = prize.value;
      treasureLog = treasureLog ? treasureLog + ` / ${newCells}번:${prize.label}` : `${newCells}번:${prize.label}`;
      messages.push(`🎁 보물 발견! ${prize.label}`);
    }

    // 별자리 완성
    if (newCells === VAC_CFG.TOTAL_CELLS && bonusLog.indexOf('완주') === -1) {
      tokens += VAC_CFG.COMPLETE_TOKEN;
      bonusLog += ',완주';
      messages.push(`🌌 별자리 완성! 토큰 +${VAC_CFG.COMPLETE_TOKEN} · 「28성의 개척자」 칭호와 별자리 의회 자격이 기록되었습니다`);
    }

    // 시트 반영 (한 번에)
    const newRow = cur.slice();
    newRow[VC.LAST_DATE - 1]   = today;
    newRow[VC.CELLS - 1]       = newCells;
    newRow[VC.STREAK - 1]      = streak;
    newRow[VC.BEST - 1]        = best;
    newRow[VC.TOKENS - 1]      = tokens;
    newRow[VC.BONUS_LOG - 1]   = bonusLog;
    newRow[VC.TREASURE - 1]    = treasureLog;
    newRow[VC.PREV_STREAK - 1] = prevStreakSave;
    newRow[VC.RESTORE_DATE - 1]= restoreDate;
    sh.getRange(rowIdx + 1, 1, 1, 14).setValues([newRow]);

    // 골드 지급 (메인 시트 자산만, BV 변동 없음, 세금 없음)
    const totalGold = VAC_CFG.DAILY_GOLD + extraGold;
    _vacGrantGold_(ss, studentName, totalGold, `[방학출석] ${newCells}번째 별` + (extraGold ? ' +보물' : ''));

    // 방학 히스토리 기록 (A열 날짜 / I열 타임스탬프)
    ss.getSheetByName(SHEET_VAC_HIST).appendRow([
      today, studentName, '출석', totalGold, tokens - Number(cur[VC.TOKENS - 1]),
      newCells, messages.join(' | '), '', _nowStr()
    ]);
    if (newCells === VAC_CFG.TOTAL_CELLS) _vacRecordCompletionPerk_(ss, studentName, today);

    return {
      active: true, mode: '방학',
      checkedToday: true, newCell: newCells,
      cells: newCells, streak: streak, tokens: tokens,
      tickets: Number(cur[VC.TICKETS - 1]),
      restoreAvailable: !!restoreDate && Number(cur[VC.TICKETS - 1]) > 0,
      gold: totalGold, messages: messages,
      isLiberationDay: (today === VAC_CFG.LIBERATION_DATE),
      libNoteDone: !!cur[VC.LIB_NOTE - 1]
    };
  } finally {
    lock.releaseLock();
  }
}

// ── 보충 기간 출석 처리 (하루 최대 3칸, 연속·복구 없음) ──────────
function _vacMakeup_(studentName, today) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_ATT);
  if (!sh) return null;

  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][VC.NAME - 1]).trim() === studentName) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return null;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cur = sh.getRange(rowIdx + 1, 1, 1, 14).getValues()[0];
    const cells = Number(cur[VC.CELLS - 1]) || 0;
    if (cells >= VAC_CFG.TOTAL_CELLS) return _vacStatusLight_(cur, today, true);

    const usedToday = (_vacDateKey_(cur[VC.MAKEUP_DATE - 1]) === today) ? Number(cur[VC.MAKEUP_CNT - 1]) || 0 : 0;
    if (usedToday >= VAC_CFG.MAKEUP_MAX_PER_DAY) return _vacStatusLight_(cur, today, true);

    const newCells = cells + 1;
    let tokens = Number(cur[VC.TOKENS - 1]) + VAC_CFG.DAILY_TOKEN;
    let bonusLog = String(cur[VC.BONUS_LOG - 1] || '');
    let treasureLog = String(cur[VC.TREASURE - 1] || '');
    const messages = [`⭐ 보충으로 ${newCells}번째 별을 켰어요 (오늘 ${usedToday + 1}/${VAC_CFG.MAKEUP_MAX_PER_DAY})`];

    let extraGold = 0;
    if (_vacCellType_(ss, newCells) === '보물') {
      const prize = _vacRollTreasure_();
      if (prize.type === '토큰') tokens += prize.value;
      if (prize.type === '골드') extraGold = prize.value;
      treasureLog = treasureLog ? treasureLog + ` / ${newCells}번:${prize.label}` : `${newCells}번:${prize.label}`;
      messages.push(`🎁 보물 발견! ${prize.label}`);
    }
    if (newCells === VAC_CFG.TOTAL_CELLS && bonusLog.indexOf('완주') === -1) {
      tokens += VAC_CFG.COMPLETE_TOKEN;
      bonusLog += ',완주';
      messages.push(`🌌 별자리 완성! 토큰 +${VAC_CFG.COMPLETE_TOKEN} · 별자리 의회 자격 획득`);
    }

    const newRow = cur.slice();
    newRow[VC.CELLS - 1]       = newCells;
    newRow[VC.TOKENS - 1]      = tokens;
    newRow[VC.BONUS_LOG - 1]   = bonusLog;
    newRow[VC.TREASURE - 1]    = treasureLog;
    newRow[VC.MAKEUP_DATE - 1] = today;
    newRow[VC.MAKEUP_CNT - 1]  = usedToday + 1;
    sh.getRange(rowIdx + 1, 1, 1, 14).setValues([newRow]);

    const totalGold = VAC_CFG.DAILY_GOLD + extraGold;
    _vacGrantGold_(ss, studentName, totalGold, `[방학보충] ${newCells}번째 별`);
    ss.getSheetByName(SHEET_VAC_HIST).appendRow([
      today, studentName, '보충', totalGold, tokens - Number(cur[VC.TOKENS - 1]),
      newCells, messages.join(' | '), '', _nowStr()
    ]);
    if (newCells === VAC_CFG.TOTAL_CELLS) _vacRecordCompletionPerk_(ss, studentName, today);

    return {
      active: true, mode: '보충', checkedToday: true, newCell: newCells,
      cells: newCells, streak: 0, tokens: tokens,
      tickets: Number(cur[VC.TICKETS - 1]),
      restoreAvailable: false, gold: totalGold, messages: messages,
      makeupLeft: VAC_CFG.MAKEUP_MAX_PER_DAY - usedToday - 1
    };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// 2. 복구권 사용 (학생이 버튼으로 직접 선택)
//    조건: 어제 딱 하루 놓쳤고, 오늘 출석하면서 복구가능일이 기록된 경우
// ════════════════════════════════════════════════════════════════
function useVacationRestore(studentName) {
  studentName = String(studentName).trim();
  const today = _vacToday_(studentName);
  if (today < VAC_CFG.START || today > VAC_CFG.END) return { success: false, msg: '방학 기간에만 사용할 수 있어요.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_ATT);
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][VC.NAME - 1]).trim() === studentName) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return { success: false, msg: '학생을 찾을 수 없어요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cur = sh.getRange(rowIdx + 1, 1, 1, 14).getValues()[0];
    const restoreDate = _vacDateKey_(cur[VC.RESTORE_DATE - 1]);
    const tickets = Number(cur[VC.TICKETS - 1]) || 0;

    if (!restoreDate)                          return { success: false, msg: '지금은 복구할 별이 없어요.' };
    if (restoreDate !== _vacAddDays_(today, -1)) return { success: false, msg: '복구는 놓친 다음 날까지만 가능해요.' };
    if (tickets <= 0)                          return { success: false, msg: '복구권을 모두 사용했어요.' };
    if (Number(cur[VC.CELLS - 1]) >= VAC_CFG.TOTAL_CELLS) return { success: false, msg: '이미 별자리를 완성했어요!' };

    const newCells = Number(cur[VC.CELLS - 1]) + 1;
    // 연속 복원: 놓치기 전 연속 + 어제(복구) + 오늘(이미 출석) = 직전연속 + 2
    const restoredStreak = (Number(cur[VC.PREV_STREAK - 1]) || 0) + 2;
    let tokens = Number(cur[VC.TOKENS - 1]) + VAC_CFG.DAILY_TOKEN;
    let bonusLog = String(cur[VC.BONUS_LOG - 1] || '');
    let best = Number(cur[VC.BEST - 1]) || 0;
    const messages = [`🌠 ${restoreDate}의 별을 되살렸어요 (연속 ${restoredStreak}일 복원)`];

    if (restoredStreak > best) {
      for (const k in VAC_CFG.STREAK_BONUS) {
        const th = Number(k);
        if (restoredStreak >= th && best < th) {
          tokens += VAC_CFG.STREAK_BONUS[k];
          bonusLog = bonusLog ? bonusLog + ',' + th : String(th);
          messages.push(`🔥 ${th}일 연속! 토큰 +${VAC_CFG.STREAK_BONUS[k]}`);
        }
      }
      best = restoredStreak;
    }
    if (newCells === VAC_CFG.TOTAL_CELLS && bonusLog.indexOf('완주') === -1) {
      tokens += VAC_CFG.COMPLETE_TOKEN;
      bonusLog += ',완주';
      messages.push(`🌌 별자리 완성! 토큰 +${VAC_CFG.COMPLETE_TOKEN} · 별자리 의회 자격 획득`);
    }

    const newRow = cur.slice();
    newRow[VC.CELLS - 1]        = newCells;
    newRow[VC.STREAK - 1]       = restoredStreak;
    newRow[VC.TICKETS - 1]      = tickets - 1;
    newRow[VC.BEST - 1]         = best;
    newRow[VC.TOKENS - 1]       = tokens;
    newRow[VC.BONUS_LOG - 1]    = bonusLog;
    newRow[VC.PREV_STREAK - 1]  = 0;
    newRow[VC.RESTORE_DATE - 1] = '';
    sh.getRange(rowIdx + 1, 1, 1, 14).setValues([newRow]);

    _vacGrantGold_(ss, studentName, VAC_CFG.DAILY_GOLD, `[방학복구] ${restoreDate}의 별`);
    ss.getSheetByName(SHEET_VAC_HIST).appendRow([
      today, studentName, '복구', VAC_CFG.DAILY_GOLD, tokens - Number(cur[VC.TOKENS - 1]),
      newCells, messages.join(' | '), '', _nowStr()
    ]);
    if (newCells === VAC_CFG.TOTAL_CELLS) _vacRecordCompletionPerk_(ss, studentName, today);

    return { success: true, cells: newCells, streak: restoredStreak,
             tokens: tokens, tickets: tickets - 1, messages: messages };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// 3. 별자리 화면 데이터 (버튼을 눌렀을 때만 호출 — 로그인 경로와 분리)
//    열린 칸의 내용만 보내고, 안 연 칸은 유형조차 숨김
// ════════════════════════════════════════════════════════════════
function getVacationStatus(studentName) {
  studentName = String(studentName).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_ATT);
  if (!sh) return { success: false };

  const data = sh.getDataRange().getValues();
  let myRow = null, classTotal = 0;
  for (let i = 1; i < data.length; i++) {
    classTotal += Number(data[i][VC.CELLS - 1]) || 0;   // ← 공동 진행도 = 단순 합계
    if (String(data[i][VC.NAME - 1]).trim() === studentName) myRow = data[i];
  }
  if (!myRow) return { success: false };

  const myCells = Number(myRow[VC.CELLS - 1]) || 0;
  const today = _vacToday_(studentName);
  const councilEligible = myCells >= VAC_CFG.TOTAL_CELLS;
  if (councilEligible) _vacRecordCompletionPerk_(ss, studentName, today);

  // 열린 칸의 내용만 전달 (스포일러 방지)
  const mapSheet = ss.getSheetByName(SHEET_VAC_MAP);
  const mapData  = mapSheet ? mapSheet.getDataRange().getValues() : [];
  const openedCells = [];
  for (let i = 1; i < mapData.length; i++) {
    const n = Number(mapData[i][0]);
    if (n >= 1 && n <= myCells) {
      openedCells.push({ no: n, type: mapData[i][1], title: mapData[i][2], content: mapData[i][3] });
    }
  }

  // 공동 진행도 단계 현황
  const milestones = VAC_CFG.MILESTONES.map(function(m) {
    return { at: m.at, label: m.label, reached: classTotal >= m.at };
  });

  return {
    success: true,
    totalCells: VAC_CFG.TOTAL_CELLS,
    myCells: myCells,
    streak: Number(myRow[VC.STREAK - 1]) || 0,
    best: Number(myRow[VC.BEST - 1]) || 0,
    tickets: Number(myRow[VC.TICKETS - 1]) || 0,
    tokens: Number(myRow[VC.TOKENS - 1]) || 0,
    openedCells: openedCells,
    classTotal: classTotal,          // 숫자와 게이지만 표시할 것 (개인별 노출 금지)
    classMax: (data.length - 1) * VAC_CFG.TOTAL_CELLS,
    milestones: milestones,
    proposalOpen: today >= VAC_CFG.PROPOSAL_OPEN && today <= VAC_CFG.MAKEUP_END,
    proposalOpenDate: VAC_CFG.PROPOSAL_OPEN,
    proposalMinChars: VAC_CFG.PROPOSAL_MIN_CHARS,
    councilEligible: councilEligible,
    completionTitle: councilEligible ? '28성의 개척자' : '',
    completionBadge: councilEligible ? '✦28' : ''
  };
}

// ════════════════════════════════════════════════════════════════
// 4. 광복절(8/15) 한 줄 소감 — 히든 업적 「빛을 되찾은 날」 조건
// ════════════════════════════════════════════════════════════════
function submitLiberationNote(studentName, note) {
  studentName = String(studentName).trim();
  note = _sanitizeString(note);
  if (!note) return { success: false, msg: '내용을 입력해주세요.' };
  if (note.length > 200) return { success: false, msg: '200자 이내로 적어주세요.' };

  const today = _vacToday_(studentName);
  if (today !== VAC_CFG.LIBERATION_DATE) {
    return { success: false, msg: '오늘은 제출할 수 없어요.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_ATT);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][VC.NAME - 1]).trim() === studentName) {
      if (data[i][VC.LIB_NOTE - 1]) return { success: false, msg: '이미 제출했어요.' };
      sh.getRange(i + 1, VC.LIB_NOTE).setValue('제출');
      ss.getSheetByName(SHEET_VAC_HIST).appendRow([
        today, studentName, '광복절소감', 0, 0, '', note, '', _nowStr()
      ]);
      return { success: true, msg: '소중한 한 줄이 기록되었어요. ✨' };
    }
  }
  return { success: false, msg: '학생을 찾을 수 없어요.' };
}

// ════════════════════════════════════════════════════════════════
// 5. 시즌2 공개 제안소
//    8/3부터 전원 공개 · 20자 이상 · 반복 제출 가능 · 작성자 익명
//    28별 완주자는 「별자리 의회」 구성원으로 다른 제안을 추천할 수 있음
// ════════════════════════════════════════════════════════════════
function submitSeasonProposal(studentName, category, title, content) {
  studentName = String(studentName).trim();
  category = _vacProposalCategory_(category);
  title    = _sanitizeString(title);
  content  = _sanitizeString(content);

  if (!title || !content) return { success: false, msg: '제목과 내용을 모두 적어주세요.' };
  if (title.length > 50 || content.length > 1000) {
    return { success: false, msg: '너무 길어요! 제목 50자, 내용 1000자 이내로 적어주세요.' };
  }
  if (content.length < VAC_CFG.PROPOSAL_MIN_CHARS) {
    return { success: false, msg: `제안 내용은 ${VAC_CFG.PROPOSAL_MIN_CHARS}자 이상 적어주세요. 현재 ${content.length}자입니다.` };
  }

  const today = _vacToday_(studentName);
  if (today < VAC_CFG.PROPOSAL_OPEN) {
    return { success: false, msg: `시즌2 공개 제안소는 ${VAC_CFG.PROPOSAL_OPEN}부터 열립니다.` };
  }
  if (today > VAC_CFG.MAKEUP_END) {
    return { success: false, msg: '여름 시즌2 제안 접수 기간이 종료되었습니다.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  if (!att || !_vacStudentExists_(att, studentName)) {
    return { success: false, msg: '학생 정보를 확인할 수 없어요.' };
  }
  _vacEnsureProposalSheets_(ss);
  const sh = ss.getSheetByName(SHEET_VAC_PROP);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = sh.getDataRange().getValues();
    let hasPrevious = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() === studentName && String(data[i][4] || '').trim()) {
        hasPrevious = true;
        break;
      }
    }

    const proposalId = 'P-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const achievementMark = hasPrevious ? '' : '시즌2의 설계자 조건충족';
    sh.appendRow([today, studentName, category, title, content, '새 제안', _nowStr(), proposalId, achievementMark, 0]);

    if (!hasPrevious) {
      const hist = ss.getSheetByName(SHEET_VAC_HIST);
      if (hist) hist.appendRow([
        today, studentName, '업적조건충족', 0, 0, '', '시즌2의 설계자 — 20자 이상 첫 제안 제출', '', _nowStr()
      ]);
    }

    return {
      success: true,
      firstValid: !hasPrevious,
      msg: !hasPrevious
        ? '제안이 공개 제안소에 등록되었습니다. 업적 「시즌2의 설계자」 달성 조건도 충족했습니다. 🛠️'
        : '새 제안이 공개 제안소에 등록되었습니다. 다른 아이디어도 계속 제출할 수 있어요. 💡'
    };
  } finally {
    lock.releaseLock();
  }
}

function getSeasonProposals(studentName) {
  studentName = String(studentName).trim();
  const today = _vacToday_(studentName);
  const isOpen = today >= VAC_CFG.PROPOSAL_OPEN && today <= VAC_CFG.MAKEUP_END;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  if (!att || !_vacStudentExists_(att, studentName)) {
    return { success: false, msg: '학생 정보를 확인할 수 없어요.', proposals: [] };
  }

  const studentRow = _vacFindStudentRow_(att, studentName);
  const myCells = studentRow ? Number(studentRow[VC.CELLS - 1]) || 0 : 0;
  const councilEligible = myCells >= VAC_CFG.TOTAL_CELLS;
  if (councilEligible) _vacRecordCompletionPerk_(ss, studentName, today);

  if (!isOpen) {
    return {
      success: true, open: false, openDate: VAC_CFG.PROPOSAL_OPEN,
      minChars: VAC_CFG.PROPOSAL_MIN_CHARS, proposals: [], councilEligible: councilEligible,
      completionTitle: councilEligible ? '28성의 개척자' : '', completionBadge: councilEligible ? '✦28' : ''
    };
  }

  _vacEnsureProposalSheets_(ss);
  const sh = ss.getSheetByName(SHEET_VAC_PROP);
  const recSh = ss.getSheetByName(SHEET_VAC_PROP_REC);
  const data = sh.getDataRange().getValues();
  const recData = recSh.getDataRange().getValues();
  const recCounts = {};
  const myRecs = {};

  for (let i = 1; i < recData.length; i++) {
    const pid = String(recData[i][1] || '').trim();
    const who = String(recData[i][2] || '').trim();
    if (!pid || !who) continue;
    recCounts[pid] = (recCounts[pid] || 0) + 1;
    if (who === studentName) myRecs[pid] = true;
  }

  const proposals = [];
  for (let i = data.length - 1; i >= 1 && proposals.length < VAC_CFG.PROPOSAL_LIST_LIMIT; i--) {
    const owner = String(data[i][1] || '').trim();
    const content = String(data[i][4] || '').trim();
    if (!content) continue;
    const proposalId = String(data[i][7] || '').trim() || ('P-LEGACY-' + (i + 1));
    const status = _vacProposalStatus_(data[i][5]);
    proposals.push({
      id: proposalId,
      date: _vacDateKey_(data[i][0]),
      category: String(data[i][2] || '기타'),
      title: String(data[i][3] || ''),
      content: content,
      status: status.label,
      statusKey: status.key,
      recommendationCount: recCounts[proposalId] || Number(data[i][9]) || 0,
      isMine: owner === studentName,
      canRecommend: councilEligible && owner !== studentName,
      recommendedByMe: !!myRecs[proposalId]
    });
  }

  return {
    success: true, open: true, openDate: VAC_CFG.PROPOSAL_OPEN,
    minChars: VAC_CFG.PROPOSAL_MIN_CHARS, limit: VAC_CFG.PROPOSAL_LIST_LIMIT,
    proposals: proposals, councilEligible: councilEligible,
    completionTitle: councilEligible ? '28성의 개척자' : '', completionBadge: councilEligible ? '✦28' : ''
  };
}

function toggleSeasonProposalRecommendation(studentName, proposalId) {
  studentName = String(studentName).trim();
  proposalId = String(proposalId || '').trim();
  const today = _vacToday_(studentName);
  if (today < VAC_CFG.PROPOSAL_OPEN || today > VAC_CFG.MAKEUP_END) {
    return { success: false, msg: '지금은 별자리 의회 추천 기간이 아닙니다.' };
  }
  if (!proposalId) return { success: false, msg: '제안을 찾을 수 없어요.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  const studentRow = att ? _vacFindStudentRow_(att, studentName) : null;
  if (!studentRow || (Number(studentRow[VC.CELLS - 1]) || 0) < VAC_CFG.TOTAL_CELLS) {
    return { success: false, msg: '28개의 별을 모두 밝힌 별자리 의회 구성원만 추천할 수 있어요.' };
  }

  _vacEnsureProposalSheets_(ss);
  const sh = ss.getSheetByName(SHEET_VAC_PROP);
  const recSh = ss.getSheetByName(SHEET_VAC_PROP_REC);
  const data = sh.getDataRange().getValues();
  let proposalRow = -1, owner = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][7] || '').trim() === proposalId) {
      proposalRow = i + 1;
      owner = String(data[i][1] || '').trim();
      break;
    }
  }
  if (proposalRow < 0) return { success: false, msg: '제안을 찾을 수 없어요. 목록을 새로고침해주세요.' };
  if (owner === studentName) return { success: false, msg: '자신이 제출한 제안은 직접 추천할 수 없어요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const recData = recSh.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < recData.length; i++) {
      if (String(recData[i][1] || '').trim() === proposalId && String(recData[i][2] || '').trim() === studentName) {
        existingRow = i + 1;
        break;
      }
    }

    let recommended;
    if (existingRow > 0) {
      recSh.deleteRow(existingRow);
      recommended = false;
    } else {
      recSh.appendRow([today, proposalId, studentName, _nowStr()]);
      recommended = true;
    }

    const latest = recSh.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < latest.length; i++) {
      if (String(latest[i][1] || '').trim() === proposalId) count++;
    }
    sh.getRange(proposalRow, 10).setValue(count);

    return {
      success: true, recommended: recommended, count: count,
      msg: recommended ? '별자리 의회 추천을 보냈습니다. ✦' : '추천을 취소했습니다.'
    };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// 6. 여름 방학 질문게시판 — 정규 방학 기간에만 운영
//    모든 학생의 질문과 답변을 익명 공개하고, 본인 글만 isMine으로 표시
// ════════════════════════════════════════════════════════════════
function submitVacationQuestion(studentName, category, question) {
  studentName = String(studentName).trim();
  category = _sanitizeString(category) || '기타';
  question = _sanitizeString(question);

  const today = _vacToday_(studentName);
  if (today < VAC_CFG.START || today > VAC_CFG.END) {
    return { success: false, msg: '질문게시판은 여름 방학 기간에만 운영됩니다.' };
  }
  if (!question) return { success: false, msg: '질문 내용을 적어주세요.' };
  if (question.length > 500) return { success: false, msg: '질문은 500자 이내로 적어주세요.' };
  if (category.length > 30) category = category.slice(0, 30);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  const qSheet = ss.getSheetByName(SHEET_VAC_Q);
  if (!att || !qSheet) return { success: false, msg: '질문게시판이 아직 준비되지 않았어요.' };

  // 로그인 학생 명단 확인
  const names = att.getRange(2, VC.NAME, Math.max(att.getLastRow() - 1, 1), 1).getValues();
  const exists = names.some(function(r) { return String(r[0]).trim() === studentName; });
  if (!exists) return { success: false, msg: '학생 정보를 확인할 수 없어요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    qSheet.appendRow([today, studentName, category, question, '', '답변대기', '', _nowStr()]);
  } finally {
    lock.releaseLock();
  }
  return { success: true, msg: '질문이 등록되었습니다. 교사의 답변은 이 게시판에서 확인할 수 있어요. 💬' };
}

function getMyVacationQuestions(studentName) {
  studentName = String(studentName).trim();
  const today = _vacToday_(studentName);
  if (today < VAC_CFG.START || today > VAC_CFG.END) {
    return { success: false, msg: '질문게시판 운영 기간이 아닙니다.', questions: [] };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_Q);
  if (!sh) return { success: false, msg: '질문게시판이 아직 준비되지 않았어요.', questions: [] };

  const data = sh.getDataRange().getValues();
  const questions = [];
  // 최신 질문 최대 50개를 공개한다. 작성자 이름 자체는 응답에 포함하지 않는다.
  for (let i = data.length - 1; i >= 1 && questions.length < 50; i--) {
    const owner = String(data[i][1] || '').trim();
    const question = String(data[i][3] || '').trim();
    if (!question) continue;
    const answer = String(data[i][4] || '').trim();
    const status = String(data[i][5] || '').trim();
    questions.push({
      date: _vacDateKey_(data[i][0]),
      category: String(data[i][2] || '기타'),
      question: question,
      answer: answer,
      answered: !!answer || status === '답변완료',
      answerDate: _vacDateKey_(data[i][6]),
      isMine: owner === studentName
    });
  }
  return { success: true, questions: questions, limit: 50 };
}


// ════════════════════════════════════════════════════════════════
// 7. 학급 공동 진행도 단계 보상 (하루 1회 트리거)
//    이미 지급한 단계는 PropertiesService 에 기록해 중복 지급 방지
// ════════════════════════════════════════════════════════════════
function checkVacationClassMilestones() {
  const today = _vacToday_(null);
  if (today < VAC_CFG.START || today > VAC_CFG.MAKEUP_END) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VAC_ATT);
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  let classTotal = 0;
  for (let i = 1; i < data.length; i++) classTotal += Number(data[i][VC.CELLS - 1]) || 0;

  const props = PropertiesService.getScriptProperties();
  const granted = (props.getProperty('VAC_MILESTONE_GRANTED') || '').split(',').filter(String);

  VAC_CFG.MILESTONES.forEach(function(m) {
    if (classTotal < m.at) return;
    if (granted.indexOf(String(m.at)) !== -1) return;  // 이미 지급함

    if (m.type === '토큰') {
      // 전원 토큰 지급
      for (let i = 1; i < data.length; i++) {
        const cur = Number(sh.getRange(i + 1, VC.TOKENS).getValue()) || 0;
        sh.getRange(i + 1, VC.TOKENS).setValue(cur + m.value);
      }
    } else if (m.type === '골드') {
      // 전원 골드 지급 (자산만, BV 변동·세금 없음)
      for (let i = 1; i < data.length; i++) {
        const name = String(data[i][VC.NAME - 1]).trim();
        if (name) _vacGrantGold_(ss, name, m.value, `[공동목표] 누적 ${m.at}회 달성`);
      }
    }
    // type '수동' 은 기록만 하고 교사가 직접 처리

    ss.getSheetByName(SHEET_VAC_HIST).appendRow([
      today, '(학급전체)', '공동목표달성', m.type === '골드' ? m.value : 0,
      m.type === '토큰' ? m.value : 0, '', `누적 ${m.at}회 — ${m.label}`, '', _nowStr()
    ]);
    granted.push(String(m.at));
  });

  props.setProperty('VAC_MILESTONE_GRANTED', granted.join(','));

  // 하루 1회 랭킹 갱신 (로그인 경로에서는 생략했으므로 여기서 보정)
  updateRankings();
}

// 트리거 설치/제거 (스프레드시트 메뉴 없이 에디터에서 1회 실행)
function setupVacationTrigger() {
  removeVacationTrigger();
  ScriptApp.newTrigger('checkVacationClassMilestones')
    .timeBased().everyDays(1).atHour(21).create();   // 매일 밤 9시
  return '✅ 공동 목표 확인 트리거 설치 완료 (매일 21시)';
}
function removeVacationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkVacationClassMilestones') ScriptApp.deleteTrigger(t);
  });
}

// ── 공개 제안소 시트 구조 보정 ─────────────────────────────────
function _vacEnsureProposalSheets_(ss) {
  let sh = ss.getSheetByName(SHEET_VAC_PROP);
  if (!sh) sh = ss.insertSheet(SHEET_VAC_PROP);
  const headers = ['제출일','이름','분류','제목','내용','검토상태','타임스탬프','제안ID','업적처리','의회추천수'];
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), headers.length).setWrap(true);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 420);
  sh.setColumnWidth(6, 130);

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
    let changed = false;
    for (let i = 0; i < rows.length; i++) {
      if (!String(rows[i][7] || '').trim() && String(rows[i][4] || '').trim()) {
        rows[i][7] = 'P-LEGACY-' + String(i + 2).padStart(3, '0');
        changed = true;
      }
      if (!String(rows[i][5] || '').trim() && String(rows[i][4] || '').trim()) {
        rows[i][5] = '새 제안';
        changed = true;
      }
      if (rows[i][9] === '') {
        rows[i][9] = 0;
        changed = true;
      }
    }
    if (changed) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  let rec = ss.getSheetByName(SHEET_VAC_PROP_REC);
  if (!rec) {
    rec = ss.insertSheet(SHEET_VAC_PROP_REC);
    rec.appendRow(['추천일','제안ID','추천자','타임스탬프']);
  }
  rec.setFrozenRows(1);
  rec.getRange(1, 1, Math.max(rec.getMaxRows(), 2), 4).setWrap(true);

  let finish = ss.getSheetByName(SHEET_VAC_FINISH);
  if (!finish) {
    finish = ss.insertSheet(SHEET_VAC_FINISH);
    finish.appendRow(['완주일','이름','한정칭호','프로필표식','별자리의회','타임스탬프']);
  }
  finish.setFrozenRows(1);
  finish.getRange(1, 1, Math.max(finish.getMaxRows(), 2), 6).setWrap(true);
  return sh;
}

function _vacRecordCompletionPerk_(ss, studentName, completionDate) {
  _vacEnsureProposalSheets_(ss);
  const sh = ss.getSheetByName(SHEET_VAC_FINISH);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === studentName) return;
  }
  sh.appendRow([completionDate, studentName, '28성의 개척자', '✦28', '초대', _nowStr()]);
}

function _vacStudentExists_(attSheet, studentName) {
  return !!_vacFindStudentRow_(attSheet, studentName);
}

function _vacFindStudentRow_(attSheet, studentName) {
  if (!attSheet || attSheet.getLastRow() < 2) return null;
  const data = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, 14).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][VC.NAME - 1] || '').trim() === studentName) return data[i];
  }
  return null;
}

function _vacProposalCategory_(category) {
  const allowed = ['새 제도','이벤트','길드 운영','직업·성장','웹앱 개선','기타'];
  category = _sanitizeString(category);
  return allowed.indexOf(category) >= 0 ? category : '기타';
}

function _vacProposalStatus_(raw) {
  const v = String(raw || '').replace(/\s/g, '');
  if (v.indexOf('검토중') !== -1) return { key:'review', label:'🔍 검토 중' };
  if (v.indexOf('시험운영후보') !== -1 || v.indexOf('시험운영') !== -1) return { key:'candidate', label:'🧪 시험 운영 후보' };
  if (v.indexOf('시즌2반영예정') !== -1 || v.indexOf('반영예정') !== -1 || v.indexOf('채택') !== -1) return { key:'accepted', label:'✅ 시즌2 반영 예정' };
  if (v.indexOf('보관') !== -1 || v.indexOf('보류') !== -1) return { key:'archive', label:'📦 보관' };
  return { key:'new', label:'💡 새 제안' };
}

// 교사용 1회 실행: 기존 28별 완주자를 완주자 시트에 일괄 기록
function syncVacationCompletionPerks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const att = ss.getSheetByName(SHEET_VAC_ATT);
  if (!att) return '방학_출석 시트를 찾을 수 없습니다.';
  _vacEnsureProposalSheets_(ss);
  const data = att.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][VC.NAME - 1] || '').trim();
    if (!name || (Number(data[i][VC.CELLS - 1]) || 0) < VAC_CFG.TOTAL_CELLS) continue;
    const before = ss.getSheetByName(SHEET_VAC_FINISH).getLastRow();
    _vacRecordCompletionPerk_(ss, name, _vacDateKey_(data[i][VC.LAST_DATE - 1]) || _vacToday_(name));
    if (ss.getSheetByName(SHEET_VAC_FINISH).getLastRow() > before) count++;
  }
  return `✅ 완주자 ${count}명의 별자리 의회·칭호 기록을 추가했습니다.`;
}

// ════════════════════════════════════════════════════════════════
// 내부 도우미 함수들
// ════════════════════════════════════════════════════════════════

// ── 오늘 날짜 (테스트 모드면 지정된 학생에게만 가짜 날짜를 돌려줌) ──
// 평상시에는 _todayStr() 과 완전히 동일하게 동작합니다.
function _vacToday_(studentName) {
  try {
    const props = PropertiesService.getScriptProperties();
    const testName = props.getProperty('VAC_TEST_NAME');
    if (!testName) return _todayStr();                       // 테스트 모드 꺼짐
    if (studentName !== null && String(studentName).trim() !== testName) {
      return _todayStr();                                    // 다른 학생은 진짜 날짜
    }
    const testDate = props.getProperty('VAC_TEST_DATE');
    return testDate || _todayStr();
  } catch (e) {
    return _todayStr();
  }
}

// 골드 지급 — applyAssetOnly 관례를 따름 (BV 불변, 세금 없음, 히스토리 9열)
function _vacGrantGold_(ss, studentName, amount, note) {
  const main = ss.getSheetByName(SHEET_MAIN);
  const data = main.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NAME - 1]).trim() === studentName) {
      const curValue = Number(data[i][COL_VALUE - 1]) || 0;
      const curAsset = Number(data[i][COL_ASSET - 1]) || 0;
      main.getRange(i + 1, COL_ASSET).setValue(curAsset + amount);
      ss.getSheetByName(SHEET_HISTORY).appendRow([
        _vacToday_(studentName), studentName, data[i][COL_BRAND - 1],
        0, amount, curValue, curAsset + amount, note, _nowStr()
      ]);
      return;
    }
  }
}

// 칸 유형 조회 (방학_별자리 시트 기준)
function _vacCellType_(ss, cellNo) {
  const sh = ss.getSheetByName(SHEET_VAC_MAP);
  if (!sh) return '일반';
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === cellNo) return String(data[i][1] || '일반');
  }
  return '일반';
}

// 보물 가중치 랜덤 뽑기
function _vacRollTreasure_() {
  const pool = VAC_CFG.TREASURE_POOL;
  const totalW = pool.reduce(function(s, p) { return s + p.w; }, 0);
  let r = Math.random() * totalW;
  for (let i = 0; i < pool.length; i++) {
    r -= pool[i].w;
    if (r <= 0) return pool[i];
  }
  return pool[0];
}

// 날짜 셀 값을 항상 'yyyy-MM-dd' 문자열로 되돌림.
// ★ 구글시트가 "2026-07-27" 같은 문자열을 자동으로 날짜(Date) 타입으로
//   바꿔버리는 경우가 있어, 그대로 String()만 씌우면 비교가 항상 어긋납니다.
//   (예: "Mon Jul 27 2026 00:00:00 GMT+0900" 같은 문자열이 되어 "2026-07-27"와 다른 값이 됨)
//   → 저장된 값이 Date 객체이든 문자열이든 항상 같은 형식으로 맞춰줍니다.
function _vacDateKey_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v || '').trim();
}

// 날짜 문자열 계산 (yyyy-MM-dd 에 days 만큼 더하기)
function _vacAddDays_(dateStr, days) {
  const parts = dateStr.split('-');
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// 쓰기 없이 현재 상태만 돌려주는 가벼운 응답
function _vacStatusLight_(row, today, isMakeup) {
  const tickets = Number(row[VC.TICKETS - 1]) || 0;
  const restoreDate = _vacDateKey_(row[VC.RESTORE_DATE - 1] || '');
  return {
    active: true,
    mode: isMakeup ? '보충' : '방학',
    checkedToday: _vacDateKey_(row[VC.LAST_DATE - 1]) === today || isMakeup,
    cells: Number(row[VC.CELLS - 1]) || 0,
    streak: Number(row[VC.STREAK - 1]) || 0,
    tokens: Number(row[VC.TOKENS - 1]) || 0,
    tickets: tickets,
    restoreAvailable: !isMakeup && !!restoreDate &&
                      restoreDate === _vacAddDays_(today, -1) && tickets > 0,
    isLiberationDay: (!isMakeup && today === VAC_CFG.LIBERATION_DATE),
    libNoteDone: !!row[VC.LIB_NOTE - 1],
    messages: []
  };
}
