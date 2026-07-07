/**
 * [수정본] 1순위 탈환 가능성 분석 (분류코드 1 추가)
 * 로직: 실예가율(B13) <= 가상 판단예가율 < 기존 1순위 판단예가율(J2)
 */

const COL = {
  COUNT:   0,
  NAMES:   1,
  TOTAL:   2,
  CAT_1:   3,
  CAT_2:   4,
  CAT_5:   5,
  CAT_10:  6,
  CAT_30:  7,
};

const CATEGORY_COL_MAP = {
  1:  COL.CAT_1,
  2:  COL.CAT_2,
  5:  COL.CAT_5,
  10: COL.CAT_10,
  30: COL.CAT_30,
};

function updateSimulationWithRevenue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("전기간 예가율 수주 시뮬레이션");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // ── 1. 헤더 초기화 ──────────────────────────────────────────────────────────
  const headers = [
    "1순위 횟수", "해당 용역", "기대수익(합계)",
    "분류: 1", "분류: 2", "분류: 5", "분류: 10", "분류: 30"
  ];
  const HEADER_COUNT = headers.length;

  sheet.getRange(2, 2, lastRow - 1, HEADER_COUNT).clearContent();

  const headerRange = sheet.getRange(1, 2, 1, HEADER_COUNT);
  headerRange.setValues([headers]);
  sheet.getRange("A1").copyTo(headerRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // ── 2. 데이터 수집 ──────────────────────────────────────────────────────────
  const FIXED_EXCLUDES   = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "통계_대시보드", "누적현황", "전기간 예가율 수주 시뮬레이션", "업체매핑"];
  const EXCLUDE_KEYWORDS = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "exc", "실적"];

  const aValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  const projectsToAnalyze = [];

  ss.getSheets().forEach(s => {
    const sName = s.getName();
    if (
      FIXED_EXCLUDES.includes(sName) ||
      EXCLUDE_KEYWORDS.some(kw => sName.includes(kw))
    ) return;

    try {
      // ★ B8→B10, B9→B11, B11→B13 으로 수정
      const ranges     = s.getRangeList(["B10", "B11", "B13", "J2"]).getRanges();
      const code       = parseInt(ranges[0].getValue());   // B10: 분류코드
      const price      = Number(ranges[1].getValue()) || 0; // B11: 기초가격
      const actualRate = ranges[2].getValue();              // B13: 실예가율
      const winnerRate = ranges[3].getValue();              // J2:  1순위 판단예가율

      if (actualRate && winnerRate && !isNaN(actualRate) && !isNaN(winnerRate)) {
        projectsToAnalyze.push({
          name:  sName,
          min:   actualRate,
          max:   winnerRate,
          price: price,
          code:  code,
        });
      }
    } catch (e) {
      console.log(sName + " 데이터 읽기 실패: " + e.message);
    }
  });

  // ── 3. 누적 계산 ─────────────────────────────────────────────────────────────
  const resultData = Array.from({ length: aValues.length }, () =>
    new Array(HEADER_COUNT).fill(0).map((v, i) => i === COL.NAMES ? "" : 0)
  );

  aValues.forEach((myRate, i) => {
    const matchedNames = [];

    projectsToAnalyze.forEach(p => {
      if (myRate >= p.min && myRate < p.max) {
        resultData[i][COL.COUNT] += 1;
        resultData[i][COL.TOTAL] += p.price;
        matchedNames.push(p.name);

        const colIdx = CATEGORY_COL_MAP[p.code];
        if (colIdx !== undefined) {
          resultData[i][colIdx] += p.price;
        }
      }
    });

    resultData[i][COL.NAMES] = matchedNames.join(", ");
  });

  // ── 4. 결과 출력 ─────────────────────────────────────────────────────────────
  sheet.getRange(2, 2, resultData.length, HEADER_COUNT).setValues(resultData);

  const MONEY_COL_START = 2 + COL.TOTAL;
  const MONEY_COL_COUNT = HEADER_COUNT - COL.TOTAL;
  sheet.getRange(2, MONEY_COL_START, lastRow - 1, MONEY_COL_COUNT)
    .setNumberFormat('#,##0"원"')
    .setHorizontalAlignment("right");

  sheet.getRange(1, 1, lastRow, HEADER_COUNT + 1)
    .setBorder(true, true, true, true, true, true, "#bcbcbc", SpreadsheetApp.BorderStyle.SOLID);

  highlightMaxPoints(sheet, lastRow, HEADER_COUNT + 1);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "전기간 1순위 시뮬레이션이 실행되었습니다.",
    "✅ 실행 완료",
    5
  );
}


/**
 * 최대 1순위 지점 행 전체 강조
 */
function highlightMaxPoints(sheet, lastRow, totalCols) {
  const dataRowCount = lastRow - 1;
  const counts = sheet.getRange(2, 2, dataRowCount, 1).getValues();

  let maxCount = 0;
  counts.forEach(r => { if (Number(r[0]) > maxCount) maxCount = Number(r[0]); });

  const backgrounds = counts.map(r =>
    new Array(totalCols).fill(Number(r[0]) === maxCount && maxCount > 0 ? "#d9ead3" : null)
  );

  sheet.getRange(2, 1, dataRowCount, totalCols).setBackgrounds(backgrounds);
}