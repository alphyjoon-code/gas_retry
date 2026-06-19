/**
 * [점추정 복원판] N(추정낙찰하한율) 계산
 * - 2026-06-19 롤백: "PQ 범위추정" 엔진(updatePQRangeColumns/buildReverseRefMap_/
 *   calcPqRange_)을 폐기하고, 그 이전의 정방향 조회(L 추정PQ점수 → N 낙찰하한율)
 *   방식으로 되돌림.
 * - L열(추정PQ점수)이 비어 있으면 100으로 간주해 조회한다(추정PQ점수를 아직
 *   입력하지 않은 경우 만점으로 가정하는 기존 의도된 동작).
 * - 참조 파일은 동일: "PQ점수별 낙찰하한율_반올림까지 고려한 값"
 *   (시트 "낙찰하한율 정리표", A열=PQ원점수 90.00~100.00 0.01단위,
 *   B~E열=분류코드 2/5/10/30별 낙찰하한율)
 */
function updateColumnQ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const classCode = parseInt(sheet.getRange("B10").getValue(), 10); // 분류코드

  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;

  sheet.getRange("N1").setValue("추정낙찰하한율").setHorizontalAlignment("center");

  const refMap = getReferenceLimitMap_(classCode);
  if (Object.keys(refMap).length === 0) {
    SpreadsheetApp.getUi().alert("분류코드 " + classCode + "에 대한 낙찰하한율 참조표를 찾을 수 없습니다.");
    return;
  }

  // L열(추정PQ점수, 12번째 열)
  const lValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues();

  const results = [];
  for (let i = 0; i < lValues.length; i++) {
    let score = lValues[i][0];
    if (score === "" || score === null || isNaN(score)) {
      score = 100; // 추정PQ점수 미입력 시 100으로 간주(의도된 동작)
    }
    score = Math.max(90, Math.min(100, Number(score)));
    const key = (Math.round(score * 100) / 100).toFixed(2);
    const rate = refMap[key];
    results.push([rate !== undefined ? rate : ""]);
  }

  sheet.getRange(2, 14, results.length, 1).setValues(results).setNumberFormat("0.000%");
  ss.toast("N열(추정낙찰하한율) 계산 완료", "✅ 완료");
}

// ═══════════════════════════════════════════════════════
//  [점추정] 낙찰하한율 참조표 → 정방향(PQ점수→낙찰하한율) 맵
//  파일: "PQ점수별 낙찰하한율_반올림까지 고려한 값" (시트 "낙찰하한율 정리표")
//  A열=PQ원점수, B~E열=분류코드 2/5/10/30별 낙찰하한율
// ═══════════════════════════════════════════════════════
var PQ_FORWARD_REF_FILE_ID = "1qWRw6Ttl3GkX1YNBhcMqV1yYfadOv3ukEKDF9LjsOio";

function getReferenceLimitMap_(classCode) {
  const SHEET_NAME     = "낙찰하한율 정리표";
  const DATA_START_ROW = 3;
  const COL_MAP        = { "2": 1, "5": 2, "10": 3, "30": 4 }; // A=PQ원점수, B~E

  const codeStr = String(parseInt(classCode, 10));
  const codeKey = (codeStr === "1") ? "2" : codeStr; // 분류코드 1 = 코드 2와 동일 산식
  const colIdx  = COL_MAP[codeKey];

  const map = {};
  if (!colIdx) return map;

  try {
    const refSS    = SpreadsheetApp.openById(PQ_FORWARD_REF_FILE_ID);
    const refSheet = refSS.getSheetByName(SHEET_NAME);
    const lastRow  = refSheet.getLastRow();
    const vals     = refSheet
      .getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 5)
      .getValues();

    vals.forEach(row => {
      const pqScore = parseFloat(row[0]);
      if (isNaN(pqScore) || pqScore < 90 || pqScore > 100) return;
      const rate = parseFloat(row[colIdx]);
      if (isNaN(rate)) return;
      map[pqScore.toFixed(2)] = rate;
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert("PQ 참조 테이블 로드 실패: " + e.message);
  }

  return map;
}
