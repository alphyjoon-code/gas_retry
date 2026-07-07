/**
 * 🏠 네비게이션 시트 생성 (용역명에 하이퍼링크 통합 버전)
 */
function createNavigationDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  const NAV_SHEET_NAME = "🏠 네비게이션";

  // ── 고정 시트 설정 ──────────────────────────────────────────────────────────
  const FIXED_SHEETS = [
    "PQ백데이터",
    "최근 30일 차트",
    "통계_대시보드",
    "전기간 예가율 수주 시뮬레이션"
  ];
  const SHEET_DESCRIPTIONS = {
    "PQ백데이터":               "전 입찰 통계, 면적 X 예가율, 업체수 X 예가율, 구간별 예가율 등",
    "최근 30일 차트":           "최근 30일 간의 실예가율 및 판단예가율",
    "통계_대시보드":            "구간별 경쟁강도와 가치 평가",
    "전기간 예가율 수주 시뮬레이션": "한 줄로 쐈을 때, 얼마나 수주했을지 추론"
  };

  const FIXED_EXCLUDES   = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "업체매핑"];
  const EXCLUDE_KEYWORDS = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "실적"];

  // ── 개별 용역 시트에서 읽어올 셀 주소 ──────────────────────────────────────
  const CELL_MAP = {
    용역명:   "B1",
    면적:     "B6",   // B5  → B6
    기초가격: "B11",  // B9  → B11
    예가율:   "B13",  // B11 → B13
    개찰일:   "B14",  // B12 → B14
    업체수:   "B15",  // B13 → B15
    순위1위:  "B16",  // B14 → B16
    낙찰업체: "B17",  // B15 → B17
    연번:     "B18",  // B16 → B18
  };

  // ── 네비게이션 시트 준비 ────────────────────────────────────────────────────
  let navSheet = ss.getSheetByName(NAV_SHEET_NAME);
  if (!navSheet) {
    navSheet = ss.insertSheet(NAV_SHEET_NAME, 0);
  } else {
    navSheet.clear();
    navSheet.clearFormats();
  }

  // ── 헤더 ────────────────────────────────────────────────────────────────────
  const headers = [
    "번호", "용역명", "개찰일", "면적", "기초가격",
    "실예가율", "1순위", "낙찰업체", "정우 순위", "정우종합 순위", "업체수"
  ];
  const TOTAL_COLS = headers.length; // 11

  // ── 데이터 수집 ─────────────────────────────────────────────────────────────
  let fixedRowsMeta = [];
  let dynamicRows   = [];

  allSheets.forEach((sheet) => {
    const name = sheet.getName();

    // 고정 시트 처리
    if (FIXED_SHEETS.includes(name)) {
      const gid  = sheet.getSheetId();
      const link = `=HYPERLINK("#gid=${gid}", "🔗 ${name}")`;
      const desc = SHEET_DESCRIPTIONS[name] || "고정 메뉴";
      fixedRowsMeta.push({
        sortKey: FIXED_SHEETS.indexOf(name),
        rowData: ["-", link, "-", desc, "-", "-", "-", "-", "-", "-", "-"]
      });
      return;
    }

    // 제외 시트 필터
    const isExcluded = FIXED_EXCLUDES.includes(name) ||
                       EXCLUDE_KEYWORDS.some(kw => name.includes(kw));
    if (isExcluded) return;

    // 개별 용역 시트 데이터 수집
    try {
      const gid  = sheet.getSheetId();
      const link = `=HYPERLINK("#gid=${gid}", "${name}")`;

      const cellAddresses = Object.values(CELL_MAP);
      const ranges = sheet.getRangeList(cellAddresses).getRanges();

      const cellValues = {};
      Object.keys(CELL_MAP).forEach((key, idx) => {
        cellValues[key] = ranges[idx].getValue();
      });

      // 연번 파싱
      const fullSerial = cellValues["연번"].toString().trim();
      let serialNum = 0;
      if (fullSerial.includes("-")) {
        const parts = fullSerial.split("-");
        serialNum = parseInt(parts[parts.length - 1], 10) || 0;
      }

      // 정우 / 정우종합 순위
      const lastRowInSheet = sheet.getLastRow();
      let jungwooRank      = "-";
      let jungwooTotalRank = "-";

      if (lastRowInSheet > 1) {
        const dataValues = sheet.getRange(2, 5, lastRowInSheet - 1, 3).getValues();
        for (let i = 0; i < dataValues.length; i++) {
          const rank        = dataValues[i][0];
          const companyName = dataValues[i][2].toString().trim();
          if (companyName === "정우")     jungwooRank      = rank;
          if (companyName === "정우종합") jungwooTotalRank = rank;
        }
      }

      dynamicRows.push([
        serialNum,
        link,
        cellValues["개찰일"],
        cellValues["면적"],
        cellValues["기초가격"],
        cellValues["예가율"],
        cellValues["순위1위"],
        cellValues["낙찰업체"],
        jungwooRank,
        jungwooTotalRank,
        cellValues["업체수"],
      ]);

    } catch (e) {
      console.log(name + " 시트 데이터 수집 스킵: " + e.message);
    }
  });

  // ── 정렬 ────────────────────────────────────────────────────────────────────
  fixedRowsMeta.sort((a, b) => a.sortKey - b.sortKey);
  dynamicRows.sort((a, b) => b[0] - a[0]);

  // ── finalRows 조립 ──────────────────────────────────────────────────────────
  const finalRows = [headers];
  fixedRowsMeta.forEach(({ rowData }) => finalRows.push(rowData));
  dynamicRows.forEach(row => {
    row[0] = row[0] > 0 ? Utilities.formatString("%03d", row[0]) : "-";
    finalRows.push(row);
  });

  const totalRows = finalRows.length;

  // ── 데이터 일괄 쓰기 ────────────────────────────────────────────────────────
  const dataRange = navSheet.getRange(1, 1, totalRows, TOTAL_COLS);
  dataRange.setValues(finalRows);

  // ── 기본 서식 ───────────────────────────────────────────────────────────────
  dataRange
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setVerticalAlignment("middle");
  navSheet.setRowHeights(1, totalRows, 35);

  // 헤더 행
  navSheet.getRange(1, 1, 1, TOTAL_COLS)
    .setBackground("#1F4E79")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 고정 시트 행 서식
  const fixedCount = fixedRowsMeta.length;
  if (fixedCount > 0) {
    navSheet.getRange(2, 1, fixedCount, TOTAL_COLS)
      .setBackground("#F8F9FA")
      .setFontWeight("normal");
    navSheet.getRange(2, 2, fixedCount, 1)
      .setHorizontalAlignment("left");
  }

  // 날짜 / 숫자 / 정렬 서식 (동적 데이터 영역)
  const dynStart = fixedCount + 2;
  const dynCount = totalRows - fixedCount - 1;
  if (dynCount > 0) {
    navSheet.getRange(dynStart, 3, dynCount, 1)
      .setHorizontalAlignment("center")
      .setNumberFormat("yyyy-mm-dd");
    navSheet.getRange(dynStart, 5, dynCount, 1)
      .setNumberFormat("#,##0");
    navSheet.getRange(dynStart, 6, dynCount, 1)
      .setNumberFormat("0.0000%")
      .setFontWeight("bold");
    navSheet.getRange(dynStart, 7, dynCount, 4)
      .setHorizontalAlignment("center");
    navSheet.getRange(dynStart, 11, dynCount, 1)
      .setHorizontalAlignment("center");
  }

  // ── 하이라이트: 정우/정우종합 순위 1~3위 행 ──────────────────────────────
  if (dynCount > 0) {
    const rankValues = navSheet
      .getRange(dynStart, 9, dynCount, 2)
      .getValues();

    const highlightRanges = [];
    for (let i = 0; i < rankValues.length; i++) {
      const rankI = parseInt(rankValues[i][0], 10);
      const rankJ = parseInt(rankValues[i][1], 10);
      const isHighRank = (rankI >= 1 && rankI <= 3) || (rankJ >= 1 && rankJ <= 3);
      if (isHighRank) {
        highlightRanges.push(`A${dynStart + i}:K${dynStart + i}`);
      }
    }

    if (highlightRanges.length > 0) {
      const hl = navSheet.getRangeList(highlightRanges);
      hl.setBackground("#E8F0FE");
      hl.setFontWeight("bold");
    }
  }

  // ── 테두리 ──────────────────────────────────────────────────────────────────
  dataRange.setBorder(
    true, true, true, true, true, true,
    "#DDDDDD", SpreadsheetApp.BorderStyle.SOLID
  );

  // ── 열 너비 ─────────────────────────────────────────────────────────────────
  const colWidths = [45, 400, 100, 70, 100, 100, 100, 100, 80, 100, 70];
  colWidths.forEach((w, idx) => navSheet.setColumnWidth(idx + 1, w));

  // ── 1행 고정 ────────────────────────────────────────────────────────────────
  navSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("네비게이션 및 순위 업데이트 완료");
}