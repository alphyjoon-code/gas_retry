/**
 * [정밀 분석본] 기간별 시뮬레이션 (전기간 버전과 구조 통일)
 */

const PERIOD_COL = {
  COUNT:  0,
  NAMES:  1,
  TOTAL:  2,
  CAT_1:  3,
  CAT_2:  4,
  CAT_5:  5,
  CAT_10: 6,
  CAT_30: 7,
};

const PERIOD_CATEGORY_COL_MAP = {
  1:  PERIOD_COL.CAT_1,
  2:  PERIOD_COL.CAT_2,
  5:  PERIOD_COL.CAT_5,
  10: PERIOD_COL.CAT_10,
  30: PERIOD_COL.CAT_30,
};

const PERIOD_HEADER_COUNT = Object.keys(PERIOD_COL).length; // 8

function updateSimulationByPeriod() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── 1. 기간 입력 ─────────────────────────────────────────────────────────────
  const startRes = ui.prompt("기간 설정 (1/2)", "시작일 (예: 2026-02-15)", ui.ButtonSet.OK_CANCEL);
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  const startText = startRes.getResponseText().trim();

  const endRes = ui.prompt("기간 설정 (2/2)", "종료일 (예: 2026-02-28)", ui.ButtonSet.OK_CANCEL);
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  const endText = endRes.getResponseText().trim();

  const startDate = _parseLocalDate(startText, 0, 0, 0);
  const endDate   = _parseLocalDate(endText, 23, 59, 59);
  if (!startDate || !endDate) {
    ui.alert("날짜 형식을 확인해주세요. (예: 2026-02-15)");
    return;
  }

  // ── 2. 시트 준비 ─────────────────────────────────────────────────────────────
  const masterSheet = ss.getSheetByName("전기간 예가율 수주 시뮬레이션");
  if (!masterSheet) { ui.alert("'전기간 예가율 수주 시뮬레이션' 시트가 없습니다."); return; }

  const dateLabel   = "[" + startText.slice(5).replace("-", "") +
                      "-" + endText.slice(5).replace("-", "") + "]";
  const newSheetName = dateLabel + " 시뮬레이션";

  let targetSheet = ss.getSheetByName(newSheetName);
  if (!targetSheet) {
    targetSheet = masterSheet.copyTo(ss).setName(newSheetName);
  } else {
    const existLastRow = targetSheet.getLastRow();
    if (existLastRow >= 2) {
      targetSheet.getRange(2, 2, existLastRow - 1, PERIOD_HEADER_COUNT).clearContent();
    }
  }

  // ── 3. 헤더 전체 갱신 ────────────────────────────────────────────────────────
  const headers = [
    "1순위 횟수", "해당 용역", "기대수익(합계)",
    "분류: 1", "분류: 2", "분류: 5", "분류: 10", "분류: 30"
  ];
  targetSheet.getRange(1, 2, 1, PERIOD_HEADER_COUNT).setValues([headers]);

  // ── 4. 프로젝트 데이터 수집 ──────────────────────────────────────────────────
  const EXCLUDE_KEYWORDS = [
    "시뮬레이션", "차트", "백데이터", "대시보드", "업체매핑",
    "🏠", "RawData", "통계", "현황", "Form", "GT", "exc", "실적"
  ];

  const periodProjects = [];
  let processedCount   = 0;

  ss.getSheets().forEach(s => {
    const sName = s.getName();
    if (EXCLUDE_KEYWORDS.some(kw => sName.includes(kw))) return;

    try {
      // ★ B8→B10, B9→B11, B11→B13, B12→B14 으로 수정
      const ranges     = s.getRangeList(["B10", "B11", "B13", "B14", "J2"]).getRanges();
      const code       = parseInt(ranges[0].getValue());         // B10: 분류코드
      const price      = Number(ranges[1].getValue()) || 0;      // B11: 기초가격
      const actualRate = parseFloat(ranges[2].getValue());       // B13: 실예가율
      const dateVal    = ranges[3].getValue();                   // B14: 개찰일
      const winnerRate = parseFloat(ranges[4].getValue());       // J2:  1순위 판단예가율

      const dateStr = dateVal instanceof Date
        ? Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd")
        : String(dateVal).trim();
      const itemDate = _parseLocalDate(dateStr, 12, 0, 0);

      processedCount++;

      if (itemDate && itemDate >= startDate && itemDate <= endDate) {
        if (!isNaN(actualRate) && !isNaN(winnerRate)) {
          periodProjects.push({
            name:  sName,
            min:   actualRate,
            max:   winnerRate,
            price: price,
            code:  code,
          });
        }
      }
    } catch (e) {
      console.log(sName + " 스킵: " + e.message);
    }
  });

  // ── 5. 누적 계산 ─────────────────────────────────────────────────────────────
  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return;

  const aValues = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  const resultData = Array.from({ length: aValues.length }, () =>
    new Array(PERIOD_HEADER_COUNT).fill(0).map((v, i) => i === PERIOD_COL.NAMES ? "" : 0)
  );

  aValues.forEach((myRate, i) => {
    const rate = parseFloat(myRate);
    if (isNaN(rate)) return;

    const matchedNames = [];
    periodProjects.forEach(p => {
      if (rate >= p.min && rate < p.max) {
        resultData[i][PERIOD_COL.COUNT] += 1;
        resultData[i][PERIOD_COL.TOTAL] += p.price;
        matchedNames.push(p.name);

        const colIdx = PERIOD_CATEGORY_COL_MAP[p.code];
        if (colIdx !== undefined) {
          resultData[i][colIdx] += p.price;
        }
      }
    });

    resultData[i][PERIOD_COL.NAMES] = matchedNames.join(", ");
  });

  // ── 6. 결과 출력 ─────────────────────────────────────────────────────────────
  targetSheet.getRange(2, 2, resultData.length, PERIOD_HEADER_COUNT).setValues(resultData);

  const MONEY_COL_START = 2 + PERIOD_COL.TOTAL;
  const MONEY_COL_COUNT = PERIOD_HEADER_COUNT - PERIOD_COL.TOTAL;
  targetSheet.getRange(2, MONEY_COL_START, lastRow - 1, MONEY_COL_COUNT)
    .setNumberFormat('#,##0"원"')
    .setHorizontalAlignment("right");

  targetSheet.getRange(1, 1, lastRow, PERIOD_HEADER_COUNT + 1)
    .setBorder(true, true, true, true, true, true, "#bcbcbc", SpreadsheetApp.BorderStyle.SOLID);

  highlightMaxPointsSafe(targetSheet, lastRow);

  ui.alert(
    "기간별 분석 완료\n" +
    "- 날짜 확인 시트: " + processedCount + "개\n" +
    "- 기간 내 용역: "   + periodProjects.length + "건"
  );
}

/**
 * 날짜 문자열(yyyy-MM-dd)을 KST 기준 Date로 파싱
 */
function _parseLocalDate(str, h, m, s) {
  const parts = str.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!parts) return null;
  const d = new Date(
    parseInt(parts[1]),
    parseInt(parts[2]) - 1,
    parseInt(parts[3]),
    h, m, s
  );
  return isNaN(d.getTime()) ? null : d;
}

/**
 * 최대 1순위 지점 행 전체 강조
 */
function highlightMaxPointsSafe(sheet, lastRow) {
  const counts = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let max = 0;
  counts.forEach(r => { if (Number(r[0]) > max) max = Number(r[0]); });
  if (max === 0) return;

  const totalCols = PERIOD_HEADER_COUNT + 1;
  const bg = counts.map(r =>
    new Array(totalCols).fill(Number(r[0]) === max ? "#d9ead3" : null)
  );
  sheet.getRange(2, 1, lastRow - 1, totalCols).setBackgrounds(bg);
}