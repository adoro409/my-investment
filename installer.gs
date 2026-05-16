// =====================================================
//  투자 포트폴리오 트래커 v2.0 — 구글 시트 설치 스크립트
//  소유자: 나 / 상미 / 다인
//  계좌유형: 해외(USD) / ISA(KRW) / IRP(KRW)
//  실행: install() 함수 선택 후 ▶ 실행
// =====================================================

// ── 색상 테마 (다크 모드) ──
const C = {
  BG:      '#1a1a2e',
  HEADER:  '#16213e',
  INPUT:   '#0f3460',
  CALC:    '#0d0d1a',
  SECTION: '#080818',
  PROFIT:  '#3ecf8e',
  LOSS:    '#ff6b6b',
  ACCENT:  '#4f8ef7',
  TEXT:    '#e8e8e8',
  TEXT2:   '#a0a0b0',
  TEXT3:   '#444466',
};

// ── 시트명 ──
const SN = {
  SET: '⚙ 설정',
  TRD: '📋 거래내역',
  DEP: '💰 입출금',
  DIV: '🎁 배당',
  SNP: '📅 월말스냅샷',
  STK: '📈 종목현황',
  ACC: '🏦 계좌현황',
  MON: '📉 월별현황',
  DSH: '🎯 대시보드',
  PRC: '💹 가격캐시',  // GOOGLEFINANCE 영구 캐시 (절대 삭제 안 함)
};

const MAX_ROWS = 500;

// ── 스프레드시트 접근 헬퍼 ──
// 웹 앱 URL 접근 시 getActiveSpreadsheet()가 null을 반환하는 경우 대비
// install() 실행 시 SS_ID를 저장하고 이후 openById()로 접근
function _ss() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}


// =====================================================
//  1. MAIN INSTALL
// =====================================================
function install() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 스프레드시트 ID 저장 (웹 앱 접근 시 openById로 사용)
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  const resp = ui.alert(
    '📊 포트폴리오 트래커 설치',
    '같은 이름의 시트가 있으면 초기화됩니다.\n계속 진행할까요?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    _makeSettings(ss);
    _makeTrades(ss);
    _makeDeposits(ss);
    _makeDividends(ss);
    _makeSnapshot(ss);
    _makeStocks(ss);
    _makeAccounts(ss);
    _makeMonthly(ss);
    _makeDashboard(ss);
    _makePriceCache(ss);
    _setupTriggers(ss);
    _createMenu();
    _reorderSheets(ss);
    ss.setActiveSheet(ss.getSheetByName(SN.DSH));

    ui.alert(
      '✅ 설치 완료!',
      '9개 시트가 생성됐습니다.\n\n' +
      '시작 방법:\n' +
      '1. ⚙설정 시트에서 계좌 정보 확인/수정\n' +
      '2. 📋거래내역 시트에 거래 내역 입력\n' +
      '3. 상단 메뉴 [📊 포트폴리오 → 전체 새로고침] 실행\n\n' +
      '웹 UI를 배포하면 앱으로도 사용할 수 있습니다.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('오류 발생', e.message + '\n\n' + e.stack, ui.ButtonSet.OK);
  }
}

function _reorderSheets(ss) {
  const order = [SN.DSH, SN.ACC, SN.STK, SN.MON, SN.TRD, SN.DEP, SN.DIV, SN.SNP, SN.SET];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
}


// =====================================================
//  2. 헬퍼 함수
// =====================================================

function _getSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear();
    sh.clearFormats();
    sh.clearConditionalFormatRules();
    sh.clearNotes();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  } else {
    sh = ss.insertSheet(name);
  }
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())
    .setBackground(C.BG).setFontColor(C.TEXT).setFontFamily('Malgun Gothic').setFontSize(10);
  sh.setHiddenGridlines(true);
  return sh;
}

function _section(sh, row, text, cols) {
  const r = sh.getRange(row, 1, 1, cols || 10);
  r.merge().setValue(text)
   .setBackground(C.SECTION).setFontColor(C.ACCENT)
   .setFontWeight('bold').setFontSize(11).setVerticalAlignment('middle');
  sh.setRowHeight(row, 30);
}

function _headers(sh, row, labels) {
  sh.getRange(row, 1, 1, labels.length).setValues([labels])
   .setBackground(C.HEADER).setFontColor(C.TEXT2)
   .setFontWeight('bold').setFontSize(10).setVerticalAlignment('middle')
   .setBorder(false, false, true, false, false, false, C.ACCENT, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(row, 28);
}

function _inputStyle(range) {
  range.setBackground(C.INPUT).setFontColor(C.TEXT).setFontSize(10);
}

function _calcStyle(range) {
  range.setBackground(C.CALC).setFontColor(C.TEXT2).setFontSize(10);
}

function _dropdown(sh, r1, c1, rows, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).build();
  sh.getRange(r1, c1, rows, 1).setDataValidation(rule);
}

function _dropdownRange(sh, r1, c1, rows, sourceA1) {
  const source = SpreadsheetApp.getActiveSpreadsheet().getRange(sourceA1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true).setAllowInvalid(false).build();
  sh.getRange(r1, c1, rows, 1).setDataValidation(rule);
}

function _colWidths(sh, widths) {
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function _addPnlFormat(sh, rangeA1) {
  const rules = sh.getConditionalFormatRules();
  const range = sh.getRange(rangeA1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(C.PROFIT).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor(C.LOSS).setRanges([range]).build()
  );
  sh.setConditionalFormatRules(rules);
}


// =====================================================
//  3. ⚙ 설정 시트
// =====================================================
function _makeSettings(ss) {
  const sh = _getSheet(ss, SN.SET);
  sh.setTabColor(C.ACCENT);

  // 소유자 목록
  _section(sh, 1, '👤  소유자 목록', 5);
  _headers(sh, 2, ['소유자ID', '이름', '표시명', '색상', '메모']);
  const owners = [
    ['P01', '나',   '나',   '파랑', '기본 소유자'],
    ['P02', '상미', '상미', '주황', '배우자'],
    ['P03', '다인', '다인', '보라', '자녀'],
  ];
  _inputStyle(sh.getRange(3, 1, owners.length, 5));
  sh.getRange(3, 1, owners.length, 5).setValues(owners);

  // 계좌 목록
  _section(sh, 8, '🏦  계좌 목록', 8);
  _headers(sh, 9, ['계좌ID', '소유자', '계좌명', '증권사', '계좌유형', '기준통화', '메모']);
  const accounts = [
    ['A01', '나',   '토스_본계좌', '토스증권', '해외', 'USD', '미국주식 전용'],
    ['A02', '나',   '키움_ISA',   '키움증권', 'ISA',  'KRW', '비과세 한도 2천만원'],
    ['A03', '나',   '키움_IRP',   '키움증권', 'IRP',  'KRW', '퇴직연금 세액공제'],
    ['A04', '상미', '상미_ISA',   '',         'ISA',  'KRW', ''],
  ];
  _inputStyle(sh.getRange(10, 1, accounts.length, 7));
  sh.getRange(10, 1, accounts.length, 7).setValues(accounts);
  // setAllowInvalid(true): 드롭다운 안내는 하되 엄격 검증 OFF
  // (행 15~16 환율 섹션 헤더와 범위 충돌 방지)
  const _settingsDrop = (c, vals) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(vals, true).setAllowInvalid(true).build();
    sh.getRange(10, c, 20, 1).setDataValidation(rule);
  };
  _settingsDrop(2, ['나', '상미', '다인']);
  _settingsDrop(5, ['해외', 'ISA', 'IRP']);
  _settingsDrop(6, ['USD', 'KRW']);

  // 실시간 환율
  _section(sh, 15, '💱  실시간 환율  (GOOGLEFINANCE 자동)', 5);
  _headers(sh, 16, ['통화쌍', '현재값', '설명', '기준시각']);
  sh.getRange(17, 1).setValue('USD/KRW').setBackground(C.INPUT).setFontColor(C.TEXT);
  sh.getRange(17, 2).setFormula('=GOOGLEFINANCE("CURRENCY:USDKRW")')
    .setBackground(C.INPUT).setFontColor(C.ACCENT).setFontWeight('bold').setFontSize(14)
    .setNumberFormat('#,##0.00');
  sh.getRange(17, 3).setValue('GOOGLEFINANCE 실시간 자동갱신').setBackground(C.CALC).setFontColor(C.TEXT2);
  sh.getRange(17, 4).setFormula('=NOW()').setBackground(C.CALC).setFontColor(C.TEXT2)
    .setNumberFormat('YYYY-MM-DD HH:mm');
  sh.setRowHeight(17, 36);

  _colWidths(sh, [80, 100, 160, 120, 80, 80, 200]);
  sh.setRowHeight(17, 36);

  sh.getRange(10, 1).setNote('계좌 추가: 이 행 아래에 행 삽입 후 입력하세요. 계좌ID 형식: A05, A06...');
}


// =====================================================
//  4. 📋 거래내역 시트
// =====================================================
function _makeTrades(ss) {
  const sh = _getSheet(ss, SN.TRD);
  sh.setTabColor(C.ACCENT);

  _headers(sh, 1, ['날짜','소유자','계좌ID','계좌명','통화','종목코드','종목명','구분','수량','단가','수수료','환율','원화환산단가','메모']);

  // 입력 컬럼 스타일
  [1,2,3,6,7,8,9,10,11,12,14].forEach(c => _inputStyle(sh.getRange(2, c, MAX_ROWS, 1)));
  // 자동계산 컬럼 스타일
  [4,5,13].forEach(c => _calcStyle(sh.getRange(2, c, MAX_ROWS, 1)));

  // 자동계산 수식 (ARRAYFORMULA)
  sh.getRange(2,4).setFormula("=ARRAYFORMULA(IF(C2:C=\"\",\"\",IFERROR(VLOOKUP(C2:C,'⚙ 설정'!$A$10:$G$50,3,0),\"\")))");
  sh.getRange(2,5).setFormula("=ARRAYFORMULA(IF(C2:C=\"\",\"\",IFERROR(VLOOKUP(C2:C,'⚙ 설정'!$A$10:$G$50,6,0),\"\")))");
  sh.getRange(2,13).setFormula("=ARRAYFORMULA(IF(J2:J=\"\",\"\",IF(E2:E=\"KRW\",J2:J,IFERROR(J2:J*L2:L,\"\"))))");

  // 드롭다운
  _dropdownRange(sh, 2, 2, MAX_ROWS, "'⚙ 설정'!$B$3:$B$5");
  _dropdownRange(sh, 2, 3, MAX_ROWS, "'⚙ 설정'!$A$10:$A$50");
  _dropdown(sh, 2, 8, MAX_ROWS, ['매수','매도']);

  // 숫자 포맷
  sh.getRange(2,1,MAX_ROWS,1).setNumberFormat('YYYY-MM-DD');
  sh.getRange(2,9,MAX_ROWS,1).setNumberFormat('#,##0.####');
  sh.getRange(2,10,MAX_ROWS,1).setNumberFormat('#,##0.####');
  sh.getRange(2,11,MAX_ROWS,1).setNumberFormat('#,##0.##');
  sh.getRange(2,12,MAX_ROWS,1).setNumberFormat('#,##0.00');
  sh.getRange(2,13,MAX_ROWS,1).setNumberFormat('#,##0.##');

  _colWidths(sh, [100,70,70,130,60,90,160,60,70,90,80,90,110,180]);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  sh.getRange(1,12).setNote('계좌ID 입력 시 자동입력됩니다.\nUSD: GOOGLEFINANCE 실시간환율\nKRW: 1 자동입력\n거래당시 환율이 다르면 수정하세요.');
  sh.getRange(1,13).setNote('🔒 자동계산 (수정금지)\n단가 × 환율 (KRW계좌는 단가 그대로)');
}


// =====================================================
//  5. 💰 입출금 시트
// =====================================================
function _makeDeposits(ss) {
  const sh = _getSheet(ss, SN.DEP);
  sh.setTabColor(C.PROFIT);

  _headers(sh, 1, ['날짜','소유자','계좌ID','계좌명','구분','금액(₩)','환율','달러환산($)','메모']);

  [1,2,3,5,6,7,9].forEach(c => _inputStyle(sh.getRange(2, c, MAX_ROWS, 1)));
  [4,8].forEach(c => _calcStyle(sh.getRange(2, c, MAX_ROWS, 1)));

  sh.getRange(2,4).setFormula("=ARRAYFORMULA(IF(C2:C=\"\",\"\",IFERROR(VLOOKUP(C2:C,'⚙ 설정'!$A$10:$G$50,3,0),\"\")))");
  sh.getRange(2,8).setFormula("=ARRAYFORMULA(IF(F2:F=\"\",\"\",IF(IFERROR(VLOOKUP(C2:C,'⚙ 설정'!$A$10:$G$50,6,0),\"KRW\")=\"USD\",IFERROR(F2:F/G2:G,\"\"),\"\")))");

  _dropdownRange(sh, 2, 2, MAX_ROWS, "'⚙ 설정'!$B$3:$B$5");
  _dropdownRange(sh, 2, 3, MAX_ROWS, "'⚙ 설정'!$A$10:$A$50");
  _dropdown(sh, 2, 5, MAX_ROWS, ['입금','출금']);

  sh.getRange(2,1,MAX_ROWS,1).setNumberFormat('YYYY-MM-DD');
  sh.getRange(2,6,MAX_ROWS,1).setNumberFormat('#,##0');
  sh.getRange(2,7,MAX_ROWS,1).setNumberFormat('#,##0.00');
  sh.getRange(2,8,MAX_ROWS,1).setNumberFormat('#,##0.##');

  _colWidths(sh, [100,70,70,130,70,130,90,110,200]);
  sh.setFrozenRows(1);
}


// =====================================================
//  6. 🎁 배당 시트
// =====================================================
function _makeDividends(ss) {
  const sh = _getSheet(ss, SN.DIV);
  sh.setTabColor(C.PROFIT);

  _headers(sh, 1, ['날짜','소유자','계좌ID','계좌명','통화','종목코드','종목명','배당금','환율','원화환산(₩)','연도','메모']);
  //  A      B       C      D      E      F       G      H      I       J         K     L

  [1,2,3,6,7,8,9,12].forEach(c => _inputStyle(sh.getRange(2, c, MAX_ROWS, 1)));
  [4,5,10,11].forEach(c => _calcStyle(sh.getRange(2, c, MAX_ROWS, 1)));

  sh.getRange(2,4).setFormula("=ARRAYFORMULA(IF(C2:C=\"\",\"\",IFERROR(VLOOKUP(C2:C,'⚙ 설정'!$A$10:$G$50,3,0),\"\")))");
  sh.getRange(2,5).setFormula("=ARRAYFORMULA(IF(C2:C=\"\",\"\",IFERROR(VLOOKUP(C2:C,'⚙ 설정'!$A$10:$G$50,6,0),\"\")))");
  sh.getRange(2,10).setFormula("=ARRAYFORMULA(IF(H2:H=\"\",\"\",IF(E2:E=\"KRW\",H2:H,IFERROR(H2:H*I2:I,\"\"))))");
  // 연도 자동 추출
  sh.getRange(2,11).setFormula("=ARRAYFORMULA(IF(A2:A=\"\",\"\",YEAR(A2:A)))");

  _dropdownRange(sh, 2, 2, MAX_ROWS, "'⚙ 설정'!$B$3:$B$5");
  _dropdownRange(sh, 2, 3, MAX_ROWS, "'⚙ 설정'!$A$10:$A$50");

  sh.getRange(2,1,MAX_ROWS,1).setNumberFormat('YYYY-MM-DD');
  sh.getRange(2,8,MAX_ROWS,1).setNumberFormat('#,##0.####');
  sh.getRange(2,9,MAX_ROWS,1).setNumberFormat('#,##0.00');
  sh.getRange(2,10,MAX_ROWS,1).setNumberFormat('#,##0');
  sh.getRange(2,11,MAX_ROWS,1).setNumberFormat('0');

  _colWidths(sh, [100,70,70,130,60,90,160,90,90,110,60,180]);
  sh.setFrozenRows(1);
  _addPnlFormat(sh, 'H2:H' + (MAX_ROWS+1));
  _addPnlFormat(sh, 'J2:J' + (MAX_ROWS+1));

  sh.getRange(1,11).setNote('🔒 자동계산 — 날짜에서 연도 추출.\n웹UI 배당차트에서 연도별/전체기간 필터에 사용됩니다.');
}


// =====================================================
//  7. 📅 월말스냅샷 시트
// =====================================================
function _makeSnapshot(ss) {
  const sh = _getSheet(ss, SN.SNP);
  sh.setTabColor('#a064f0');

  sh.getRange(1,1,1,9).merge()
    .setValue('📅 매월 말일, 증권사 앱에서 계좌 총액 확인 후 입력하세요. (월 1회 · 약 2분) → 월별현황 차트 자동 업데이트')
    .setBackground('#0a0a20').setFontColor(C.TEXT2).setFontSize(10).setWrap(true);
  sh.setRowHeight(1, 36);

  _headers(sh, 2, ['년월','소유자','계좌ID','계좌명','통화','계좌총액','환율','원화환산(₩)','메모']);

  [1,2,3,6,7,9].forEach(c => _inputStyle(sh.getRange(3, c, MAX_ROWS, 1)));
  [4,5,8].forEach(c => _calcStyle(sh.getRange(3, c, MAX_ROWS, 1)));

  sh.getRange(3,4).setFormula("=ARRAYFORMULA(IF(C3:C=\"\",\"\",IFERROR(VLOOKUP(C3:C,'⚙ 설정'!$A$10:$G$50,3,0),\"\")))");
  sh.getRange(3,5).setFormula("=ARRAYFORMULA(IF(C3:C=\"\",\"\",IFERROR(VLOOKUP(C3:C,'⚙ 설정'!$A$10:$G$50,6,0),\"\")))");
  sh.getRange(3,8).setFormula("=ARRAYFORMULA(IF(F3:F=\"\",\"\",IF(E3:E=\"KRW\",F3:F,IFERROR(F3:F*G3:G,\"\"))))");

  _dropdownRange(sh, 3, 2, MAX_ROWS, "'⚙ 설정'!$B$3:$B$5");
  _dropdownRange(sh, 3, 3, MAX_ROWS, "'⚙ 설정'!$A$10:$A$50");

  sh.getRange(3,1,MAX_ROWS,1).setNumberFormat('YYYY-MM');
  sh.getRange(3,6,MAX_ROWS,1).setNumberFormat('#,##0.####');
  sh.getRange(3,7,MAX_ROWS,1).setNumberFormat('#,##0.00');
  sh.getRange(3,8,MAX_ROWS,1).setNumberFormat('#,##0');

  _colWidths(sh, [80,70,70,130,60,130,90,130,200]);
  sh.setFrozenRows(2);
}


// =====================================================
//  8. 계산 시트들 (GAS 새로고침)
// =====================================================
function _makeStocks(ss) {
  const sh = _getSheet(ss, SN.STK);
  sh.setTabColor(C.ACCENT);
  sh.getRange(1,1,1,17).merge()
    .setValue('📈 종목현황 — 🔄 자동계산 시트. [포트폴리오 → 전체 새로고침] 후 업데이트됩니다.')
    .setBackground(C.SECTION).setFontColor(C.TEXT2).setFontSize(10);
  sh.setRowHeight(1, 30);
  _headers(sh, 2, ['소유자','계좌명','유형','통화','종목코드','종목명','보유수량','평균매입가','현재가','평가금','투자금','평가금(₩)','투자금(₩)','수익금(₩)','수익률(%)','비중(%)','배당누계(₩)']);
  _colWidths(sh, [70,130,60,60,90,160,80,90,90,100,100,110,110,110,90,70,110]);
  sh.setFrozenRows(2);
  // setFrozenColumns 제거: 1행 병합셀(A1:Q1)과 충돌하여 오류 발생
  _addPnlFormat(sh, 'N3:N1000');
  _addPnlFormat(sh, 'O3:O1000');
}

function _makeAccounts(ss) {
  const sh = _getSheet(ss, SN.ACC);
  sh.setTabColor(C.PROFIT);
  sh.getRange(1,1,1,13).merge()
    .setValue('🏦 계좌현황 — 🔄 자동계산 시트. [포트폴리오 → 전체 새로고침] 후 업데이트됩니다.')
    .setBackground(C.SECTION).setFontColor(C.TEXT2).setFontSize(10);
  sh.setRowHeight(1, 30);
  _headers(sh, 2, ['소유자','계좌ID','계좌명','증권사','유형','종목수','투자금(₩)','평가금(₩)','수익금(₩)','수익률(%)','비중(%)','입금누계(₩)','배당누계(₩)']);
  _colWidths(sh, [70,70,130,100,60,60,120,120,120,90,70,120,110]);
  sh.setFrozenRows(2);
  _addPnlFormat(sh, 'I3:I500');
  _addPnlFormat(sh, 'J3:J500');
}

function _makeMonthly(ss) {
  const sh = _getSheet(ss, SN.MON);
  sh.setTabColor('#a064f0');
  sh.getRange(1,1,1,5).merge()
    .setValue('📉 월별현황 — 🔄 월말스냅샷 데이터 기반 자동계산. 매월 스냅샷 입력 후 새로고침하세요.')
    .setBackground(C.SECTION).setFontColor(C.TEXT2).setFontSize(10);
  sh.setRowHeight(1, 30);
  _headers(sh, 2, ['년월','총 평가금(₩)','누적입금(₩)','월 수익(₩)','누적 수익률(%)']);
  _colWidths(sh, [80,130,130,120,110]);
  sh.setFrozenRows(2);
  _addPnlFormat(sh, 'D3:D500');
  _addPnlFormat(sh, 'E3:E500');
}

function _makeDashboard(ss) {
  const sh = _getSheet(ss, SN.DSH);
  sh.setTabColor(C.LOSS);
  sh.getRange(1,1,1,8).merge()
    .setValue('🎯  투자 포트폴리오 대시보드')
    .setBackground(C.SECTION).setFontColor(C.ACCENT).setFontWeight('bold').setFontSize(14);
  sh.setRowHeight(1, 44);

  sh.getRange(2,1,1,6).setValues([['총 평가금(₩)','총 투자금(₩)','수익금(₩)','수익률(%)','배당누계(₩)','실시간 환율']])
    .setBackground(C.HEADER).setFontColor(C.TEXT2).setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(2, 28);

  sh.getRange(3,1,1,6).setBackground(C.CALC).setFontColor(C.TEXT).setFontSize(16)
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(3,6).setFormula("='⚙ 설정'!B17").setFontColor(C.ACCENT).setNumberFormat('#,##0.00');
  sh.setRowHeight(3, 54);

  sh.getRange(5,1,1,8).merge()
    .setValue('⬆ 상단 메뉴 [📊 포트폴리오 → 전체 새로고침]을 실행하면 최신 데이터로 업데이트됩니다.')
    .setBackground(C.BG).setFontColor(C.TEXT3).setFontSize(10);

  _colWidths(sh, [140,140,130,110,130,130,130,130]);
}


// =====================================================
//  9. 새로고침 함수 (GAS 계산)
// =====================================================

function refreshAll() {
  refreshStocks();
  refreshAccounts();
  refreshMonthly();
  refreshDashboard();
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 새로고침 완료!', '📊 포트폴리오', 3);
}

// =====================================================
//  💹 가격캐시 + 타임 트리거 방식
//
//  구조:
//  💹 가격캐시 시트
//    A: 종목코드  B: 통화  C: GF코드
//    D: =GOOGLEFINANCE(...)  ← Google이 자동 업데이트
//    E: 캐시 현재가 (트리거가 D→E 복사, 정적 값)
//    F: 마지막 업데이트 시각
//
//  흐름:
//  30분마다 트리거 → cachePricesFromSheet() → D열 읽어 E열 저장
//  웹앱 "시세 새로고침" → E열(정적값) 읽기 → 종목현황 업데이트
// =====================================================

function _makePriceCache(ss) {
  if (!ss) ss = _ss();
  let sh = ss.getSheetByName(SN.PRC);
  if (sh) return sh;
  sh = ss.insertSheet(SN.PRC);
  sh.getRange(1,1,1,6).setValues([['종목코드','통화','GF코드','GOOGLEFINANCE현재가','캐시현재가','업데이트시각']]);
  sh.setColumnWidths(1, 6, 110);
  sh.hideSheet();
  SpreadsheetApp.flush();
  return sh;
}

// 종목현황에 있는 종목을 가격캐시에 등록 (신규만, 기존 수식 절대 건드리지 않음)
function _ensurePriceCacheRows(ss, holdings) {
  const sh = _makePriceCache(ss);
  const existing = new Set();
  if (sh.getLastRow() > 1) {
    sh.getRange(2,1,sh.getLastRow()-1,1).getValues()
      .forEach(r => { if (r[0]) existing.add(String(r[0]).trim()); });
  }
  let added = false;
  holdings.forEach(h => {
    if (existing.has(h.code)) return;
    const gfCode = h.curr === 'USD' ? h.code : 'KRX:' + h.code;
    const nr = sh.getLastRow() + 1;
    sh.getRange(nr,1,1,3).setValues([[h.code, h.curr, gfCode]]);
    sh.getRange(nr,4).setFormula('=IFERROR(GOOGLEFINANCE("' + gfCode + '"),"")');
    existing.add(h.code);
    added = true;
    Logger.log('가격캐시 등록: ' + h.code + ' (' + gfCode + ')');
  });
  if (added) SpreadsheetApp.flush();
}

// ★ 30분 트리거 함수: D열(GOOGLEFINANCE) → E열(정적 캐시) 복사
function cachePricesFromSheet() {
  const ss = _ss();
  const sh = ss.getSheetByName(SN.PRC);
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  const now  = new Date();
  let cnt = 0;
  rows.forEach((r,i) => {
    const price = parseFloat(r[3]); // D열: GOOGLEFINANCE 계산값
    if (r[0] && price > 0) {
      sh.getRange(i+2,5).setValue(price); // E열 캐시
      sh.getRange(i+2,6).setValue(now);   // F열 시각
      cnt++;
    }
  });
  SpreadsheetApp.flush();
  Logger.log('cachePricesFromSheet 완료: ' + cnt + '개 / ' + now);
}

// 30분 타임 트리거 설치 (GAS 편집기에서 최초 1회 실행)
function setupPriceTrigger() {
  // 기존 동일 트리거 삭제
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'cachePricesFromSheet') ScriptApp.deleteTrigger(t);
  });
  // 30분마다 자동 실행
  ScriptApp.newTrigger('cachePricesFromSheet').timeBased().everyMinutes(30).create();
  Logger.log('✅ 트리거 설치 완료 — 30분마다 cachePricesFromSheet 자동 실행');
  // 즉시 1회 실행 (현재 GOOGLEFINANCE 값 캐시)
  cachePricesFromSheet();
}

// E열(캐시)에서 현재가 맵 반환, holdings의 신규 종목은 자동 등록
function _getPriceMap(ss, holdings) {
  if (holdings && holdings.length) _ensurePriceCacheRows(ss, holdings);
  const sh = ss.getSheetByName(SN.PRC);
  if (!sh || sh.getLastRow() < 2) return {};
  const priceMap = {};
  sh.getRange(2,1,sh.getLastRow()-1,5).getValues().forEach(r => {
    const code  = String(r[0]).trim();
    const price = parseFloat(r[4]); // E열: 캐시된 정적 값
    if (code && price > 0) priceMap[code] = price;
  });
  Logger.log('_getPriceMap: ' + JSON.stringify(priceMap));
  return priceMap;
}

// 종목현황 현재가 업데이트 (캐시에서 읽기)
function refreshStockPricesOnly() {
  const ss    = _ss();
  const stkSh = ss.getSheetByName(SN.STK);
  const setSh = ss.getSheetByName(SN.SET);
  if (!stkSh || stkSh.getLastRow() < 3) return {ok:false, msg:'종목 없음'};

  const fxRate   = parseFloat(setSh.getRange(17,2).getValue()) || 1450;
  const dataRows = stkSh.getRange(3,1,stkSh.getLastRow()-2,17).getValues()
    .map((r,i) => ({r,i})).filter(({r}) => r[4]);
  if (!dataRows.length) return {ok:false, msg:'종목 없음'};

  // 캐시에 없는 신규 종목 자동 등록
  const holdings = dataRows.map(o => ({code:o.r[4], curr:o.r[3]}));
  const priceMap = _getPriceMap(ss, holdings);

  // 평가금 합산 (비중 계산용)
  const totalEval = dataRows.reduce((s,obj) => {
    const p = priceMap[obj.r[4]];
    if (!p) return s;
    const ev = obj.r[3]==='USD' ? parseFloat(obj.r[6])*p*fxRate : parseFloat(obj.r[6])*p;
    return s + ev;
  }, 0);

  // 종목현황 업데이트
  dataRows.forEach(obj => {
    const row = obj.i + 3;
    const curr = obj.r[3], qty = parseFloat(obj.r[6])||0;
    const avgOrig = parseFloat(obj.r[7])||0;
    const price   = priceMap[obj.r[4]];
    const investKrw = parseFloat(obj.r[12]) || Math.round(qty*avgOrig*(curr==='USD'?fxRate:1));
    if (!price) return;
    const evalOrig = qty * price;
    const evalKrw  = Math.round(curr==='USD' ? evalOrig*fxRate : evalOrig);
    const pnlKrw   = evalKrw - investKrw;
    const pnlPct   = investKrw > 0 ? pnlKrw/investKrw*100 : 0;
    const weight   = totalEval > 0 ? evalKrw/totalEval*100 : 0;
    stkSh.getRange(row,9).setValue(price).setNumberFormat('#,##0.####');
    stkSh.getRange(row,10).setValue(evalOrig).setNumberFormat('#,##0.##');
    stkSh.getRange(row,12).setValue(evalKrw).setNumberFormat('#,##0');
    stkSh.getRange(row,14).setValue(pnlKrw).setNumberFormat('#,##0');
    stkSh.getRange(row,15).setValue(pnlPct).setNumberFormat('0.00');
    stkSh.getRange(row,16).setValue(weight).setNumberFormat('0.0');
  });
  SpreadsheetApp.flush();
  return {ok:true, count:dataRows.length};
}

function refreshStocks() {
  const ss = _ss();
  const trdSh = ss.getSheetByName(SN.TRD);
  const divSh = ss.getSheetByName(SN.DIV);
  const stkSh = ss.getSheetByName(SN.STK);
  const setSh = ss.getSheetByName(SN.SET);
  if (!trdSh || !stkSh) return;

  const fxRate = parseFloat(setSh.getRange(17, 2).getValue()) || 1450;
  const trdData = trdSh.getLastRow() > 1
    ? trdSh.getRange(2, 1, trdSh.getLastRow()-1, 14).getValues().filter(r => r[0] && r[5])
    : [];
  const divData = divSh && divSh.getLastRow() > 1
    ? divSh.getRange(2, 1, divSh.getLastRow()-1, 11).getValues().filter(r => r[0])
    : [];

  // 계좌 정보 맵
  const acctRows = setSh.getRange(10, 1, 30, 7).getValues().filter(r => r[0]);
  const acctMap = {};
  acctRows.forEach(r => acctMap[r[0]] = {name: r[2], type: r[4], curr: r[5]});

  // 종목별 집계
  const hMap = {};
  trdData.forEach(r => {
    const [date, owner, acctId,, , code, name, type, qty, price, fee, rate] = r;
    const key = `${owner}||${acctId}||${code}`;
    const ac  = acctMap[acctId] || {name: acctId, type:'', curr:'KRW'};
    const fx  = ac.curr === 'USD' ? (parseFloat(rate) || fxRate) : 1;
    const q   = parseFloat(qty) || 0;
    const p   = parseFloat(price) || 0;
    const f   = parseFloat(fee) || 0;
    if (!hMap[key]) hMap[key] = {owner, acctId, acctName: ac.name, acctType: ac.type, curr: ac.curr, code, name, buyQty:0, buyAmtKrw:0, sellQty:0};
    const h = hMap[key];
    if (type === '매수') { h.buyQty += q; h.buyAmtKrw += (q * p + f) * fx; }
    else if (type === '매도') { h.sellQty += q; }
  });

  // 배당 집계
  const divMap = {};
  divData.forEach(r => {
    const key = `${r[1]}||${r[2]}||${r[5]}`;
    divMap[key] = (divMap[key] || 0) + (parseFloat(r[9]) || 0);
  });

  // 보유 종목 (수량 > 0)
  const holdings = Object.values(hMap).filter(h => (h.buyQty - h.sellQty) > 0.00001);
  holdings.sort((a,b) => (a.owner+a.acctName+a.code).localeCompare(b.owner+b.acctName+b.code));

  // 기존 데이터 삭제
  if (stkSh.getLastRow() >= 3) stkSh.getRange(3,1,stkSh.getLastRow()-2,17).clear();

  if (holdings.length === 0) {
    stkSh.getRange(3,1).setValue('거래내역 입력 후 새로고침하세요.').setFontColor(C.TEXT3);
    return;
  }

  // ── 가격캐시(GOOGLEFINANCE)에서 현재가 조회 ──
  // 신규 종목은 가격캐시에 자동 등록 후 GOOGLEFINANCE 수식 세팅
  const priceMap = _getPriceMap(ss, holdings);
  holdings.forEach(h => {
    h.currentPrice = priceMap[h.code] || null;
  });

  // ── 파생값 계산 ──
  holdings.forEach(h => {
    h.netQty    = h.buyQty - h.sellQty;
    h.avgKrw    = h.buyQty > 0 ? h.buyAmtKrw / h.buyQty : 0;
    h.avgOrig   = h.curr === 'USD' ? h.avgKrw / fxRate : h.avgKrw;
    h.investKrw = Math.round(h.netQty * h.avgKrw);
    if (h.currentPrice) {
      const evalOrig = h.netQty * h.currentPrice;
      h.evalKrw = Math.round(h.curr === 'USD' ? evalOrig * fxRate : evalOrig);
      h.pnlKrw  = h.evalKrw - h.investKrw;
      h.pnlPct  = h.investKrw > 0 ? h.pnlKrw / h.investKrw * 100 : 0;
    } else {
      h.evalKrw = ''; h.pnlKrw = ''; h.pnlPct = '';
    }
  });

  // 비중 계산 (전체 평가금 합산 후)
  const totalEval = holdings.reduce((s,h) => s + (h.evalKrw || 0), 0);
  holdings.forEach(h => {
    h.weight = (totalEval > 0 && h.evalKrw) ? h.evalKrw / totalEval * 100 : '';
  });

  // ── 시트 쓰기 (수식 없이 모두 정적 값) ──
  holdings.forEach((h, i) => {
    const row = i + 3;
    const key = `${h.owner}||${h.acctId}||${h.code}`;
    const div = divMap[key] || 0;
    const investOrig = h.netQty * h.avgOrig;
    const evalOrig   = h.currentPrice ? h.netQty * h.currentPrice : '';

    // A~Q 한 번에 쓰기
    stkSh.getRange(row, 1, 1, 17).setValues([[
      h.owner,        // A: 소유자
      h.acctName,     // B: 계좌명
      h.acctType,     // C: 유형
      h.curr,         // D: 통화
      h.code,         // E: 종목코드
      h.name,         // F: 종목명
      h.netQty,       // G: 수량
      h.avgOrig,      // H: 평단가(원화통화)
      h.currentPrice || '', // I: 현재가 (API직접조회)
      evalOrig,       // J: 현재가×수량
      investOrig,     // K: 평단가×수량
      h.evalKrw,      // L: 평가금(₩)
      h.investKrw,    // M: 투자금(₩)
      h.pnlKrw,       // N: 수익금(₩)
      h.pnlPct,       // O: 수익률%
      h.weight,       // P: 비중%
      div             // Q: 배당금
    ]]);

    // 배경/폰트색 일괄 적용
    stkSh.getRange(row,1,1,6).setBackground(C.BG).setFontColor(C.TEXT);
    stkSh.getRange(row,7,1,2).setBackground(C.BG).setFontColor(C.TEXT);
    stkSh.getRange(row,9).setBackground(C.CALC).setFontColor(C.ACCENT).setNumberFormat('#,##0.####');
    stkSh.getRange(row,10,1,8).setBackground(C.CALC).setFontColor(C.TEXT);
    stkSh.getRange(row,17).setFontColor(C.PROFIT);

    // 숫자 포맷
    stkSh.getRange(row,7).setNumberFormat('#,##0.####');
    stkSh.getRange(row,8).setNumberFormat('#,##0.####');
    stkSh.getRange(row,10,1,2).setNumberFormat('#,##0.##');
    stkSh.getRange(row,12,1,2).setNumberFormat('#,##0');
    stkSh.getRange(row,14).setNumberFormat('#,##0');
    stkSh.getRange(row,15).setNumberFormat('0.00');
    stkSh.getRange(row,16).setNumberFormat('0.0');
    stkSh.getRange(row,17).setNumberFormat('#,##0');
    stkSh.setRowHeight(row, 24);
  });

  stkSh.clearConditionalFormatRules();
  _addPnlFormat(stkSh, `N3:N${holdings.length+2}`);
  _addPnlFormat(stkSh, `O3:O${holdings.length+2}`);
  SpreadsheetApp.flush();
}

// 🏦 계좌현황 새로고침
function refreshAccounts() {
  const ss    = _ss();
  const setSh = ss.getSheetByName(SN.SET);
  const stkSh = ss.getSheetByName(SN.STK);
  const depSh = ss.getSheetByName(SN.DEP);
  const divSh = ss.getSheetByName(SN.DIV);
  const accSh = ss.getSheetByName(SN.ACC);
  if (!accSh) return;

  const acctRows = setSh.getRange(10, 1, 30, 7).getValues().filter(r => r[0]);
  const stats = {};
  acctRows.forEach(a => stats[a[0]] = {owner:a[1],name:a[2],broker:a[3],type:a[4],investKrw:0,evalKrw:0,stockCnt:0,deposit:0,div:0});

  // 종목현황 집계 (GAS가 채운 값 읽기)
  if (stkSh && stkSh.getLastRow() >= 3) {
    stkSh.getRange(3,1,stkSh.getLastRow()-2,17).getValues().filter(r=>r[0]).forEach(r => {
      const ac = acctRows.find(a => a[2] === r[1]);
      if (!ac) return;
      const id = ac[0];
      if (!stats[id]) return;
      stats[id].evalKrw   += parseFloat(r[11]) || 0;
      stats[id].investKrw += parseFloat(r[12]) || 0;
      stats[id].stockCnt++;
    });
  }

  // 입출금 누계
  if (depSh && depSh.getLastRow() > 1) {
    depSh.getRange(2,1,depSh.getLastRow()-1,9).getValues().filter(r=>r[0]).forEach(r => {
      if (!stats[r[2]]) return;
      stats[r[2]].deposit += (r[4]==='입금'?1:-1) * (parseFloat(r[5])||0);
    });
  }

  // 배당 누계
  if (divSh && divSh.getLastRow() > 1) {
    divSh.getRange(2,1,divSh.getLastRow()-1,11).getValues().filter(r=>r[0]).forEach(r => {
      if (!stats[r[2]]) return;
      stats[r[2]].div += parseFloat(r[9]) || 0;
    });
  }

  const totalEval = Object.values(stats).reduce((s,a) => s + a.evalKrw, 0);

  if (accSh.getLastRow() >= 3) accSh.getRange(3,1,accSh.getLastRow()-2,13).clear();

  acctRows.forEach((a, i) => {
    const id = a[0];
    const st = stats[id] || {};
    const row = i + 3;
    const profit = (st.evalKrw||0) - (st.investKrw||0);
    const rate   = st.investKrw > 0 ? profit/st.investKrw*100 : 0;
    const weight = totalEval > 0 ? (st.evalKrw||0)/totalEval*100 : 0;

    accSh.getRange(row,1,1,13).setValues([[
      st.owner||'', id, st.name||'', st.broker||'', st.type||'', st.stockCnt||0,
      Math.round(st.investKrw||0), Math.round(st.evalKrw||0), Math.round(profit),
      +rate.toFixed(2), +weight.toFixed(1),
      Math.round(st.deposit||0), Math.round(st.div||0)
    ]]).setBackground(C.BG).setFontColor(C.TEXT);

    accSh.getRange(row,9).setFontColor(profit>=0?C.PROFIT:C.LOSS);
    accSh.getRange(row,10).setFontColor(rate>=0?C.PROFIT:C.LOSS);
    accSh.getRange(row,7,1,3).setNumberFormat('#,##0');
    accSh.getRange(row,10).setNumberFormat('0.00');
    accSh.getRange(row,11).setNumberFormat('0.0');
    accSh.getRange(row,12,1,2).setNumberFormat('#,##0');
    accSh.setRowHeight(row, 24);
  });
}

// 📉 월별현황 새로고침
function refreshMonthly() {
  const ss    = _ss();
  const snpSh = ss.getSheetByName(SN.SNP);
  const depSh = ss.getSheetByName(SN.DEP);
  const monSh = ss.getSheetByName(SN.MON);
  if (!monSh || !snpSh || snpSh.getLastRow() < 3) return;

  const snpData = snpSh.getRange(3,1,snpSh.getLastRow()-2,9).getValues().filter(r=>r[0]);
  const evalMap = {};
  snpData.forEach(r => {
    const ym = _toYM(r[0]);
    if (ym) evalMap[ym] = (evalMap[ym]||0) + (parseFloat(r[7])||0);
  });

  const depData = depSh && depSh.getLastRow()>1
    ? depSh.getRange(2,1,depSh.getLastRow()-1,9).getValues().filter(r=>r[0]) : [];
  const depMap = {};
  depData.forEach(r => {
    const ym = _toYM(r[0]);
    if (ym) depMap[ym] = (depMap[ym]||0) + (r[4]==='입금'?1:-1)*(parseFloat(r[5])||0);
  });

  const months = Object.keys(evalMap).sort();
  let cumDep = 0;
  const result = months.map((ym, i) => {
    cumDep += depMap[ym] || 0;
    const ev   = evalMap[ym];
    const prev = i > 0 ? evalMap[months[i-1]] : cumDep;
    const mPro = ev - prev - (depMap[ym]||0);
    const rate = cumDep > 0 ? (ev - cumDep)/cumDep*100 : 0;
    return [ym, Math.round(ev), Math.round(cumDep), Math.round(mPro), +rate.toFixed(2)];
  });

  if (monSh.getLastRow() >= 3) monSh.getRange(3,1,monSh.getLastRow()-2,5).clear();
  if (!result.length) return;

  monSh.getRange(3,1,result.length,5).setValues(result).setBackground(C.BG).setFontColor(C.TEXT);
  monSh.getRange(3,1,result.length,1).setNumberFormat('YYYY-MM');
  monSh.getRange(3,2,result.length,3).setNumberFormat('#,##0');
  monSh.getRange(3,5,result.length,1).setNumberFormat('0.00');
  result.forEach((r,i) => {
    const row = i+3;
    monSh.getRange(row,4).setFontColor(r[3]>=0?C.PROFIT:C.LOSS);
    monSh.getRange(row,5).setFontColor(r[4]>=0?C.PROFIT:C.LOSS);
    monSh.setRowHeight(row, 24);
  });
}

// 🎯 대시보드 새로고침
function refreshDashboard() {
  const ss    = _ss();
  const accSh = ss.getSheetByName(SN.ACC);
  const dshSh = ss.getSheetByName(SN.DSH);
  if (!dshSh) return;

  let totalEval=0, totalInvest=0, totalDiv=0;
  if (accSh && accSh.getLastRow()>=3) {
    accSh.getRange(3,1,accSh.getLastRow()-2,13).getValues().filter(r=>r[0]).forEach(r => {
      totalInvest += parseFloat(r[6])||0;
      totalEval   += parseFloat(r[7])||0;
      totalDiv    += parseFloat(r[12])||0;
    });
  }

  const profit = totalEval - totalInvest;
  const rate   = totalInvest>0 ? profit/totalInvest*100 : 0;

  const vals = [
    _won(totalEval), _won(totalInvest),
    (profit>=0?'+':'')+_won(profit),
    (rate>=0?'+':'')+rate.toFixed(2)+'%',
    _won(totalDiv), null
  ];
  vals.forEach((v, i) => {
    if (v === null) return;
    dshSh.getRange(3, i+1).setValue(v);
  });
  dshSh.getRange(3,3).setFontColor(profit>=0?C.PROFIT:C.LOSS);
  dshSh.getRange(3,4).setFontColor(rate>=0?C.PROFIT:C.LOSS);

  // 계좌별 요약
  const setSh = ss.getSheetByName(SN.SET);
  if (dshSh.getLastRow()>6) dshSh.getRange(6,1,dshSh.getLastRow()-5,8).clear();
  _section(dshSh, 6, '🏦  계좌별 현황', 8);
  _headers(dshSh, 7, ['소유자','계좌명','유형','투자금(₩)','평가금(₩)','수익금(₩)','수익률(%)','비중(%)']);
  if (accSh && accSh.getLastRow()>=3) {
    accSh.getRange(3,1,accSh.getLastRow()-2,13).getValues().filter(r=>r[0]).forEach((r,i) => {
      const row = i+8;
      const p = parseFloat(r[8])||0, rt = parseFloat(r[9])||0;
      dshSh.getRange(row,1,1,8).setValues([[r[0],r[2],r[4],r[6],r[7],r[8],r[9],r[10]]])
        .setBackground(C.BG).setFontColor(C.TEXT).setFontSize(10);
      dshSh.getRange(row,6).setFontColor(p>=0?C.PROFIT:C.LOSS);
      dshSh.getRange(row,7).setFontColor(rt>=0?C.PROFIT:C.LOSS);
      dshSh.getRange(row,4,1,3).setNumberFormat('#,##0');
      dshSh.getRange(row,7).setNumberFormat('0.00');
      dshSh.getRange(row,8).setNumberFormat('0.0');
      dshSh.setRowHeight(row, 24);
    });
  }
}


// =====================================================
//  10. 유틸리티
// =====================================================
function _won(n) { return Math.round(n).toLocaleString('ko-KR')+'원'; }
function _toYM(v) {
  if (!v) return null;
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v).substring(0,7);
}


// =====================================================
//  11. onEdit 트리거 — 환율 자동입력
// =====================================================
function onEdit(e) {
  if (!e) return;
  const sh   = e.range.getSheet();
  const name = sh.getName();
  const col  = e.range.getColumn();
  const row  = e.range.getRow();
  // 시트별 설정: 계좌ID 컬럼 → 환율 컬럼
  const cfg = {[SN.TRD]:{ac:3,fx:12,start:2},[SN.DEP]:{ac:3,fx:7,start:2},[SN.DIV]:{ac:3,fx:9,start:2},[SN.SNP]:{ac:3,fx:7,start:3}};
  const c = cfg[name];
  if (!c || col !== c.ac || row < c.start) return;
  const acctId = e.range.getValue();
  if (!acctId) return;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const setSh = ss.getSheetByName(SN.SET);
  if (!setSh) return;
  const acctData = setSh.getRange(10,1,30,7).getValues();
  const acct = acctData.find(r => r[0] === acctId);
  if (!acct) return;
  const curr   = acct[5];
  const fxCell = sh.getRange(row, c.fx);
  if (curr === 'KRW') { fxCell.setValue(1); }
  else { fxCell.setValue(parseFloat(setSh.getRange(17,2).getValue()) || 1); }
}

function _setupTriggers(ss) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['onEdit','onOpen'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
}


// =====================================================
//  12. 커스텀 메뉴
// =====================================================
function onOpen() { _createMenu(); }

function _createMenu() {
  SpreadsheetApp.getUi().createMenu('📊 포트폴리오')
    .addItem('🔄 전체 새로고침', 'refreshAll')
    .addSeparator()
    .addItem('📈 종목현황만', 'refreshStocks')
    .addItem('🏦 계좌현황만', 'refreshAccounts')
    .addItem('📉 월별현황만', 'refreshMonthly')
    .addItem('🎯 대시보드만', 'refreshDashboard')
    .addSeparator()
    .addItem('🌐 웹 UI URL 확인', 'openWebUI')
    .addToUi();
}

function openWebUI() {
  const url = ScriptApp.getService().getUrl();
  const ui  = SpreadsheetApp.getUi();
  if (url) ui.alert('🌐 웹 UI URL', url + '\n\n브라우저 즐겨찾기에 추가하세요.', ui.ButtonSet.OK);
  else ui.alert('웹 UI 미배포', '[배포 → 새 배포 → 웹앱]으로 배포하세요.', ui.ButtonSet.OK);
}


// =====================================================
//  13. 웹앱 진입점 & 서버 함수
// =====================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('투자 포트폴리오 트래커')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheetData() {
  const ss = _ss();
  const sh = ss.getSheetByName(SN.SET);
  // P01, P02... 형식만 소유자로 인식 (섹션 헤더/계좌 데이터 혼입 방지)
  const owners   = sh.getRange(3,1,20,5).getValues().filter(r=>String(r[0]).match(/^P\d+/));
  // A01, A02... 형식만 계좌로 인식
  const accounts = sh.getRange(10,1,40,7).getValues().filter(r=>String(r[0]).match(/^A\d+/));
  return {
    owners,
    accounts,
    fxRate: parseFloat(sh.getRange(17,2).getValue()) || 0,
  };
}

// Date 객체 → ISO 문자열 변환 (google.script.run 직렬화 대응)
function _serializeRow(row) {
  return row.map(function(v) {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
    return v;
  });
}

function _readSheet(name, startRow, cols) {
  const sh = _ss().getSheetByName(name);
  if (!sh || sh.getLastRow() < startRow) return [];
  return sh.getRange(startRow,1,sh.getLastRow()-startRow+1,cols)
    .getValues()
    .filter(r => r[0] instanceof Date ? true : !!r[0])
    .map(_serializeRow);
}

function getTradesData()    { return _readSheet(SN.TRD, 2, 14); }
function getDepositsData()  { return _readSheet(SN.DEP, 2, 9);  }
function getDividendsData() { return _readSheet(SN.DIV, 2, 12); }
function getSnapshotData()  { return _readSheet(SN.SNP, 3, 9);  }
function getStocksData()    { return _readSheet(SN.STK, 3, 17).filter(r=>r[4]); } // r[4]=종목코드

// 시세만 빠르게 읽기 (GOOGLEFINANCE 폴링용)
// 반환: [{code, price, evalKrw, pnlKrw, pnlPct, weight}]
function getStockPrices() {
  const sh = _ss().getSheetByName(SN.STK);
  if (!sh || sh.getLastRow() < 3) return [];
  const rows = sh.getRange(3, 1, sh.getLastRow()-2, 17).getValues()
    .filter(r => r[4]); // r[4]=종목코드
  return rows.map(r => ({
    code:    r[4],
    price:   r[8],   // I열: GOOGLEFINANCE 현재가
    evalKrw: r[11],  // L열: 평가금(₩)
    pnlKrw:  r[13],  // N열: 수익금(₩)
    pnlPct:  r[14],  // O열: 수익률%
    weight:  r[15],  // P열: 비중%
  }));
}
function getAccountsData()  { return _readSheet(SN.ACC, 3, 13); }
function getMonthlyData()   { return _readSheet(SN.MON, 3, 5);  }

function saveTrade(d) {
  const sh = _ss().getSheetByName(SN.TRD);
  sh.appendRow([new Date(d.date),d.owner,d.acctId,'','',d.code,d.name,d.type,
    +d.qty,+d.price,+(d.fee||0),+(d.rate||1),'',d.memo||'']);
  return {ok:true};
}

function saveDeposit(d) {
  const sh = _ss().getSheetByName(SN.DEP);
  sh.appendRow([new Date(d.date),d.owner,d.acctId,'',d.type,+d.amount,+(d.rate||1),'',d.memo||'']);
  return {ok:true};
}

function saveDividend(d) {
  const sh = _ss().getSheetByName(SN.DIV);
  // 간편 월배당 입력(acctId 없음): ARRAYFORMULA 컬럼(D,E,J,K)은 빈 값으로 appendRow
  // → J: ARRAYFORMULA가 E="" 이면 H*I = amount*1 = amount로 자동 계산
  // → K: ARRAYFORMULA가 A에서 연도 자동 추출
  sh.appendRow([
    new Date(d.date),
    d.owner,
    d.acctId || '',   // C: 계좌ID (없으면 빈 값)
    '',               // D: 계좌명 (ARRAYFORMULA)
    '',               // E: 통화   (ARRAYFORMULA — 없으면 J에서 H*I로 계산)
    d.code  || '',    // F: 종목코드
    d.name  || '월배당합계',  // G: 종목명
    +d.amount,        // H: 배당금
    +(d.rate||1),     // I: 환율
    '',               // J: 원화환산 (ARRAYFORMULA)
    '',               // K: 연도    (ARRAYFORMULA)
    d.memo  || '',    // L: 메모
  ]);
  return {ok:true};
}

function saveSnapshot(d) {
  const sh = _ss().getSheetByName(SN.SNP);
  sh.appendRow([d.ym,d.owner,d.acctId,'','',+d.total,+(d.rate||1),'',d.memo||'']);
  return {ok:true};
}

function saveOwner(d) {
  const sh = _ss().getSheetByName(SN.SET);
  const vals = sh.getRange(3, 1, 10, 1).getValues();
  let nextRow = 3;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0]) nextRow = 3 + i + 1; else break;
  }
  const newId = 'P' + String(nextRow - 2).padStart(2, '0');
  sh.getRange(nextRow, 1, 1, 5).setValues([[newId, d.name, d.name, d.color||'파랑', d.memo||'']]);
  return {ok:true, id:newId};
}

function saveAccount(d) {
  const sh = _ss().getSheetByName(SN.SET);
  const vals = sh.getRange(10, 1, 30, 1).getValues();
  let nextRow = 10;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0]) nextRow = 10 + i + 1; else break;
  }
  const newId = 'A' + String(nextRow - 9).padStart(2, '0');
  sh.getRange(nextRow, 1, 1, 7).setValues([[newId, d.owner, d.name, d.broker||'', d.type, d.currency, d.memo||'']]);
  return {ok:true, id:newId};
}

// ── 과거 데이터 일괄 저장 ──
// rows: [{ym:'2024-01', totalAsset:75000000, deposit:500000, owner:'나'}]
// 각 row → 📅 월말스냅샷 1건 + (deposit>0 이면) 💰 입출금 1건(입금)
function saveBulkHistory(rows) {
  if (!rows || rows.length === 0) return {ok: false, msg: '입력 데이터 없음'};
  const ss = _ss();
  const shSnp = ss.getSheetByName(SN.SNP);
  const shDep = ss.getSheetByName(SN.DEP);
  if (!shSnp) return {ok: false, msg: '시트 없음: ' + SN.SNP};
  if (!shDep) return {ok: false, msg: '시트 없음: ' + SN.DEP};

  let snapCount = 0, depCount = 0;

  rows.forEach(function(r) {
    const ym     = String(r.ym || '');          // 'YYYY-MM'
    const asset  = parseFloat(r.totalAsset) || 0;
    const dep    = parseFloat(r.deposit)    || 0;
    const owner  = r.owner || '나';
    const memo   = '과거 일괄입력';

    if (!ym || ym.length < 7) return;           // 유효하지 않은 행 skip

    // 📅 월말스냅샷 append
    // 컬럼 순서: 날짜, 소유자, 계좌ID, 계좌명(auto), 통화(auto), 총평가금, 환율, 원화환산(auto), 메모
    shSnp.appendRow([ym, owner, '', '', '', asset, 1, '', memo]);
    snapCount++;

    // 💰 입출금: deposit > 0 이면 '입금' 행 추가
    // 컬럼 순서: 날짜, 소유자, 계좌ID, 계좌명(ARRAYFORMULA), 입금/출금, 금액, 환율, 원화환산(ARRAYFORMULA), 메모
    if (dep > 0) {
      const dateObj = new Date(ym + '-01');
      shDep.appendRow([dateObj, owner, '', '', '입금', dep, 1, '', memo]);
      depCount++;
    }
  });

  return {ok: true, snapCount: snapCount, depCount: depCount, total: rows.length};
}

// ── 가격캐시 진단 (GAS 편집기에서 직접 실행) ──
function testPriceCache() {
  const ss = _ss();
  const sh = ss.getSheetByName(SN.PRC);
  if (!sh) { Logger.log('가격캐시 시트 없음 — install() 또는 _makePriceCache() 실행 필요'); return; }
  const rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow()-1, 4).getValues()
    : [];
  Logger.log('가격캐시 행 수: ' + rows.length);
  rows.forEach(r => Logger.log(r[0] + ' (' + r[1] + ') GF=' + r[2] + ' → ' + r[3]));
  if (rows.length === 0) Logger.log('→ refreshStocks() 실행 후 1~2분 기다리세요');
}

// ── 진단용 함수 (Apps Script 편집기에서 직접 실행) ──
function testDeposits() {
  const rows = getDepositsData();
  Logger.log('rows count: ' + rows.length);
  if (rows.length > 0) {
    Logger.log('first row: ' + JSON.stringify(rows[0]));
    Logger.log('last  row: ' + JSON.stringify(rows[rows.length - 1]));
  }
}

function testDividends() {
  const rows = getDividendsData();
  Logger.log('dividends count: ' + rows.length);
  if (rows.length > 0) Logger.log('first row: ' + JSON.stringify(rows[0]));
}

function testAll() {
  testDeposits();
  testDividends();
  const sd = getSheetData();
  Logger.log('owners: ' + sd.owners.length + ', fxRate: ' + sd.fxRate);
}

// ── 웹 앱에서 직접 호출 가능한 진단 함수 ──
function debugDeposits() {
  try {
    const ss = _ss();
    if (!ss) return { ok: false, msg: 'ss is null — SS_ID 미설정 또는 getActiveSpreadsheet null' };
    const allSheets = ss.getSheets().map(s => s.getName());
    const sh = ss.getSheetByName(SN.DEP);
    if (!sh) {
      return { ok: false, msg: '시트 없음: ' + SN.DEP, allSheets: allSheets, ssId: ss.getId() };
    }
    const lastRow = sh.getLastRow();
    let firstRow = null, rawA2 = null;
    if (lastRow >= 2) {
      firstRow = sh.getRange(2, 1, 1, 9).getValues()[0].map(String);
      rawA2 = sh.getRange(2, 1).getDisplayValue();
    }
    const rows = sh.getRange(2, 1, Math.max(1, lastRow - 1), 9).getValues();
    const filtered = rows.filter(r => r[0]);
    return {
      ok: true,
      ssId: ss.getId(),
      sheetName: sh.getName(),
      lastRow: lastRow,
      totalRows: rows.length,
      filteredRows: filtered.length,
      firstRow: firstRow,
      rawA2: rawA2,
      allSheets: allSheets
    };
  } catch(e) {
    return { ok: false, msg: e.message, stack: e.stack };
  }
}

// 시트 이름 목록만 반환 (가장 단순한 진단)
function getSheetNames() {
  try {
    const ss = _ss();
    if (!ss) return ['ERROR: ss is null'];
    return ss.getSheets().map(s => s.getName());
  } catch(e) {
    return ['ERROR: ' + e.message];
  }
}
