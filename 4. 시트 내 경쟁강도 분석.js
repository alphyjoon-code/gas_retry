/**
 * [기능 4] 경쟁강도 분석 — J열(판단예가율) 구간별 업체 수 집계
 * 출력: R(판단예가율 구간, 구 U열), S(경쟁, 구 V열)
 * L~Q(구 O~T)는 입찰 분석(추정PQ·가점·N·O·P·Q) 전용 — 절대 clear 하지 않음
 * (K~M 삭제로 구 U~X열이 R~U열로 3열씩 이동)
 */
const INTENSITY_COL = {
  INTERVAL: 18, // R: 판단예가율 구간 (구 U열)
  COUNT:    19, // S: 경쟁 (구 V열)
  WORK1:    20, // T: 작업용 (구 W열, 구구 Q열)
  WORK2:    21, // U: 작업용 (구 X열, 구구 R열)
};

function analyzeIntensityToColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  const lastRow = activeSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("분석할 데이터가 없습니다.");
    return;
  }

  const data = activeSheet.getRange(2, 10, lastRow - 1, 1).getValues(); // J열: 판단예가율

  const buckets = [];
  const step = 0.07;

  for (let val = 101.5; val > 98.8; val -= step) {
    const upper = parseFloat(val.toFixed(2));
    let lower = parseFloat((val - step).toFixed(2));
    if (lower < 98.8) lower = 98.8;

    buckets.push({
      label: upper.toFixed(2) + "% ~ " + lower.toFixed(2) + "%",
      upper: upper,
      lower: lower,
      count: 0,
    });
    if (lower <= 98.0) break;
  }

  data.forEach(row => {
    const ratio = row[0];
    if (typeof ratio === "number" && !isNaN(ratio)) {
      const displayVal = ratio < 2 ? ratio * 100 : ratio;
      for (let i = 0; i < buckets.length; i++) {
        if (displayVal <= buckets[i].upper && displayVal > buckets[i].lower) {
          buckets[i].count++;
          break;
        }
      }
    }
  });

  const results = [["판단예가율 구간", "경쟁"]];
  buckets.forEach(b => {
    results.push([b.label, b.count]);
  });

  // R~U만 초기화 (L~Q 입찰 분석 데이터 보호, 구 U~X → R~U)
  activeSheet.getRange("R:U").clear();

  const resultRange = activeSheet.getRange(1, INTENSITY_COL.INTERVAL, results.length, 2);
  resultRange.setValues(results);

  const counts = buckets.map(b => b.count).filter(c => c > 0);
  const uniqueSortedCounts = [...new Set(counts)].sort((a, b) => b - a);

  const top1 = uniqueSortedCounts[0];
  const top2 = uniqueSortedCounts[1];
  const top3 = uniqueSortedCounts[2];

  for (let i = 0; i < buckets.length; i++) {
    const rowIdx = i + 2;
    const currentCount = buckets[i].count;
    const targetRange = activeSheet.getRange(rowIdx, INTENSITY_COL.INTERVAL, 1, 2);

    if (currentCount > 0) {
      if (currentCount === top1) {
        targetRange.setBackground("#ff0000").setFontColor("white").setFontWeight("bold");
      } else if (currentCount === top2) {
        targetRange.setBackground("#ff7b7b").setFontColor("black").setFontWeight("bold");
      } else if (currentCount === top3) {
        targetRange.setBackground("#ffc7c7").setFontColor("black").setFontWeight("bold");
      }
    }
  }

  activeSheet.getRange(1, INTENSITY_COL.INTERVAL, 1, 2)
    .setBackground("#45818e")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  activeSheet.getRange(2, INTENSITY_COL.COUNT, buckets.length, 1)
    .setHorizontalAlignment("center");

  activeSheet.getRange(1, INTENSITY_COL.INTERVAL, results.length, 2)
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  activeSheet.setColumnWidth(INTENSITY_COL.INTERVAL, 130);
  activeSheet.setColumnWidth(INTENSITY_COL.COUNT, 50);
  activeSheet.setColumnWidth(INTENSITY_COL.WORK1, 100);
  activeSheet.setColumnWidth(INTENSITY_COL.WORK2, 100);

  SpreadsheetApp.getUi().alert("경쟁강도 출력완료! (R·S열, 구 U·V열)");
}
