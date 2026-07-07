/**
 * 경쟁강도 취합 — 각 용역 시트 R·S열(판단예가율 구간, 경쟁, 구 U·V열) → 분석용_RawData
 */
function collectBiddingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("분석용_RawData");

  if (!logSheet) {
    logSheet = ss.insertSheet("분석용_RawData");
    logSheet.appendRow(["개찰일", "용역명", "분류코드", "실예가율", "판단예가율_구간", "경쟁강도(업체수)", "실예가_포함여부"]);
  }

  const existingData = logSheet.getDataRange().getValues();
  const processedBids = new Set();
  existingData.forEach(row => {
    if (row[1]) {
      processedBids.add(row[1].toString().trim());
    }
  });

  const excludeKeywords   = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "exc", "실적", "업체매핑"];
  const excludeExactNames = ["GT", "Form", "🏠 네비게이션", "분석용_RawData"];

  const allSheets = ss.getSheets();

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();

    if (excludeExactNames.includes(sheetName)) return;

    const hasExcludeKeyword = excludeKeywords.some(keyword => sheetName.includes(keyword));
    if (hasExcludeKeyword) return;

    if (sheetName.match(/^\d{4}-\d{2}-\d{2}$/)) return;

    if (processedBids.has(sheetName)) return;

    try {
      const bidDate      = sheet.getRange("B14").getValue();
      const actualRate   = sheet.getRange("B13").getValue();
      const categoryCode = sheet.getRange("B10").getValue();

      if (bidDate === "" || actualRate === "") return;

      // R열(판단예가율 구간, 구 U열) 기준 마지막 행 — L열은 추정PQ이므로 사용하지 않음
      const uValues = sheet.getRange("R:R").getValues();
      let lastRow = 0;
      for (let i = uValues.length - 1; i >= 0; i--) {
        if (uValues[i][0] !== "") {
          lastRow = i + 1;
          break;
        }
      }

      if (lastRow < 2) return;

      const intervalData = sheet.getRange("R2:S" + lastRow).getValues();
      const rowsToAdd = [];

      intervalData.forEach(row => {
        const intervalText = row[0];
        const intensity    = row[1];
        if (!intervalText) return;

        const isWin = checkRateInInterval(actualRate, intervalText);

        rowsToAdd.push([
          bidDate,
          sheetName,
          categoryCode,
          actualRate,
          intervalText,
          intensity === "" ? 0 : intensity,
          isWin ? "YES" : "",
        ]);
      });

      if (rowsToAdd.length > 0) {
        logSheet.getRange(logSheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
      }
    } catch (e) {
      console.log(sheetName + " 오류: " + e.message);
    }
  });

  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 7).sort({ column: 1, ascending: false });
  }

  SpreadsheetApp.getUi().alert("데이터 수집이 완료되었습니다!");
}

/**
 * 실예가율이 구간 문자열 안에 포함되는지 판별
 */
function checkRateInInterval(rate, intervalText) {
  const numbers = intervalText.match(/[\d.]+/g);
  if (numbers && numbers.length === 2) {
    const upper = parseFloat(numbers[0]) / 100;
    const lower = parseFloat(numbers[1]) / 100;
    const rateVal = typeof rate === "string"
      ? parseFloat(rate.replace("%", "")) / 100
      : rate;
    return rateVal <= upper && rateVal > lower;
  }
  return false;
}
