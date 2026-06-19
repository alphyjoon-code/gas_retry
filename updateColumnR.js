/**
 * [점추정 복원판] O(추정예가율) 계산
 * - 2026-06-19 롤백: no-op 폐기 스텁을 걷어내고, "PQ 범위추정" 도입 이전의
 *   단순 산식을 복원함.
 *
 *   O(추정예가율) = I(투찰금액) / (B11(기초가격) × N(추정낙찰하한율))
 *
 * - N열이 비어 있거나(참조표 매칭 실패 등) 0이면 해당 행은 공란 처리한다.
 */
function updateColumnR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const basePrice = parseFloat(sheet.getRange("B11").getValue()); // 기초가격

  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;

  sheet.getRange("O1").setValue("추정예가율").setHorizontalAlignment("center");

  if (!basePrice || isNaN(basePrice) || basePrice <= 0) {
    SpreadsheetApp.getUi().alert("B11 셀의 기초가격이 유효하지 않습니다.");
    return;
  }

  const iValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();  // I 투찰금액
  const nValues = sheet.getRange(2, 14, lastRow - 1, 1).getValues(); // N 추정낙찰하한율

  const results = [];
  for (let i = 0; i < iValues.length; i++) {
    const bidAmt = Number(iValues[i][0]);
    const n = Number(nValues[i][0]);

    if (nValues[i][0] === "" || !isFinite(bidAmt) || !isFinite(n) || n === 0) {
      results.push([""]);
      continue;
    }

    results.push([bidAmt / (basePrice * n)]);
  }

  sheet.getRange(2, 15, results.length, 1).setValues(results).setNumberFormat("0.000%");
  ss.toast("O열(추정예가율) 계산 완료", "✅ 완료");
}
