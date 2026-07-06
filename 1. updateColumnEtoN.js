/**
 * [2026-06-23 열구조 개편판] 데이터 계산 + 서식 복구 (E~L 담당)
 * - 신규 열 구조(claude.md 9절 기준): K=가점(fillBonusScores 담당),
 *   L=PQ 하한/필요PQ점수(K열이 아직 채워지기 전이라 이 함수는 계산하지 않고
 *   비워두며, calcPQScore_를 이용한 실제 계산은 updatePQTierColumns.js가 담당),
 *   M~Q=PQ 티어별 추정예가율(updatePQTierColumns.js, 건드리지 않음).
 * - 종전(~2026-06-22)에는 이 함수가 K열에 필요PQ점수를 직접 채웠으나,
 *   K가 가점 전용 열로 재배치되면서 그 계산 결과는 L열로 이동함.
 *   K는 이 함수가 쓰지 않고(가점은 fillBonusScores가 별도로 채움),
 *   다만 fillBonusScores 실행 전 K에 남아있는 값(나라장터 원본 비고
 *   텍스트 등)을 한 번 더 가산점 입력값으로 동행시켜 L 계산에 사용한다.
 * - N(추정낙찰하한율)/O(추정예가율)/P(PQ Gap)/Q(특이사항)를 채우던
 *   updateColumnQ/R/S/T.js는 2026-06-23 폐기됨(M~Q가 PQ 티어 열로
 *   재정의되어 더 이상 유효하지 않음).
 * - 비적격 수동 플래그("예가초과"/"적격점수미달")는 E열(순위)에 기록
 */

// ─────────────────────────────────────────────
// [설정] 제외 시트 목록 (최상단 상수로 분리)
// ─────────────────────────────────────────────
const FIXED_EXCLUDES   = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "업체매핑"];
const EXCLUDE_KEYWORDS = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "실적"];

// ─────────────────────────────────────────────
// [유틸] 숫자 파싱 (공통)
// ─────────────────────────────────────────────
const cleanNum = (val) => {
  if (typeof val === "number") return val;
  if (!val) return 0;
  return Number(String(val).replace(/[^0-9.-]+/g, ""));
};

// ─────────────────────────────────────────────
// [유틸] 실제 업체 데이터의 마지막 행 (공통)
// sheet.getLastRow()는 시트 전체(R·S열 경쟁강도 참고표 등 다른 영역 포함) 기준이라
// 실제 업체 데이터보다 훨씬 아래까지 잡힐 수 있다. 이를 그대로 쓰면 L/N/P 등을
// 계산하는 함수들이 업체가 없는 빈 행까지 값을 채워버리는 문제가 생긴다.
// F열(사업자등록번호) 기준으로 실제 마지막 데이터 행을 찾는다.
// ─────────────────────────────────────────────
function getRealDataLastRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return lastRow;
  const colF = sheet.getRange(2, 6, lastRow - 1, 1).getValues(); // F열: 사업자등록번호
  for (let i = colF.length - 1; i >= 0; i--) {
    if (colF[i][0] !== "" && colF[i][0] !== null) return i + 2; // 헤더(1행) 제외 절대 행번호
  }
  return 1; // 실제 데이터 없음
}


// ═══════════════════════════════════════════════════════
//  메인 함수
// ═══════════════════════════════════════════════════════
function updateBiddingResults() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // ── 제외 시트 필터 ──
  const isExcluded = FIXED_EXCLUDES.includes(sheetName) ||
                     EXCLUDE_KEYWORDS.some(kw => sheetName.includes(kw));
  if (isExcluded) {
    SpreadsheetApp.getUi().alert(`이 시트('${sheetName}')는 자동채우기 실행 대상이 아닙니다.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // ── 기준값 읽기 (B열) ──
  const [lowerLimitRate, classCode, basePrice, scheduledPrice, actualRate] =
    sheet.getRange("B9:B13").getValues().map(r => cleanNum(r[0]));
  //   B9=낙찰하한율  B10=분류코드  B11=기초금액  B12=예정가격  B13=실예가율
  const thresholdAmt = scheduledPrice * lowerLimitRate;

  // H1 헤더 (K1/L1 헤더는 updatePQTierColumns()가 설정하므로 여기서는 건드리지 않음)
  sheet.getRange("H1").setValue("투찰률").setHorizontalAlignment("center");

  // ── 데이터 로드 (E~L열, 8열) ──
  const DATA_START_COL = 5;   // E
  const DATA_COL_COUNT = 8;   // E~L (K=가점, L=PQ하한/필요PQ점수)
  const rawRange = sheet.getRange(2, DATA_START_COL, lastRow - 1, DATA_COL_COUNT);
  const rawData  = rawRange.getValues();

  // K열(11번째) 원본 비고: 나라장터 원본 데이터는 비적격(예가초과/적격점수미달) 업체에는
  // 순위(E)를 부여하지 않고 비워두는 대신, 비고란에 판정 결과가 들어있는 경우가 많음.
  // 최초 입력(신규 붙여넣기) 시점에는 E열만으로 비적격 여부를 알 수 없으므로 분류 보조용으로 읽어둔다.
  // (이후 fillBonusScores 단계에서 K열은 가점 값으로 덮어써짐)
  const remarkRawAll = sheet.getRange(2, 11, lastRow - 1, 1).getValues();

  // 유효 행 필터 + 투찰률(H열=index 3) 재계산
  // row[8]에 K열 원본 비고를 임시로 동행시켜 분류(정렬 포함) 후 기록 직전에 제거한다.
  // (row[6]=K, row[7]=L은 이제 모두 실제 데이터 칸이므로 임시 슬롯으로 쓰지 않음)
  //
  // 주의: rawData는 E~L(8열)이므로
  //   row[0]=E, row[1]=F, row[2]=G, row[3]=H, row[4]=I,
  //   row[5]=J, row[6]=K(가점), row[7]=L(PQ하한, 이 함수가 계산)
  let data = rawData
    .map((row, i) => { row[6] = row[6] || 0; row[8] = remarkRawAll[i][0]; return row; })  // K열(row[6]) 가점 기본값 0, row[8]=임시 비고 슬롯
    .filter(r => r[1] !== "" && r[4] !== "")
    .map(row => {
      row[4] = cleanNum(row[4]);                                      // I열: 입찰금액
      row[3] = scheduledPrice > 0 ? row[4] / scheduledPrice : 0;     // H열: 투찰률
      return row;
    });

  if (data.length === 0) return;

  // ── 그룹 분류 (1순위 / 정상 / 하한미달 / 제외) ──
  const rank1Row = data.find(r => r[0] == 1);
  if (!rank1Row) {
    SpreadsheetApp.getUi().alert("1순위 업체를 찾을 수 없습니다.");
    return;
  }

  const positiveGroup    = [];
  const negativeGroup    = [];
  const disqualifiedGroup = [];

  data.forEach(row => {
    if (row[0] == 1) return;
    // E열에 이미 비적격 플래그가 있으면 그것을 우선 사용(재실행 시 유지).
    // E열이 비어있다면(나라장터 원본이 비적격 업체에 순위를 주지 않는 경우) K열 원본 비고로 판정.
    const eFlag = String(row[0] ?? "").trim();
    const flagSource = eFlag || String(row[8] ?? "").trim();  // row[8]=임시 슬롯(K열 원본 비고)
    if (flagSource.includes("초과")) {
      row[0] = "예가초과";      disqualifiedGroup.push(row);
    } else if (flagSource.includes("적격점") || (flagSource.includes("미달") && !flagSource.includes("하한"))) {
      row[0] = "적격점수미달"; disqualifiedGroup.push(row);
    } else if (row[4] < thresholdAmt) {
      negativeGroup.push(row);
    } else {
      positiveGroup.push(row);
    }
  });

  positiveGroup.sort((a, b) => a[4] - b[4]);
  negativeGroup.sort((a, b) => b[4] - a[4]);

  // ── 계산 ──
  calculateRankValues(
    rank1Row, positiveGroup, negativeGroup, disqualifiedGroup,
    actualRate, thresholdAmt, lowerLimitRate, basePrice, scheduledPrice, classCode
  );

  // ── 시트 반영 ──
  rawRange.clearContent();
  // E~L만 기록 (row[8]의 임시 비고 슬롯은 제거)
  const finalData   = [rank1Row, ...positiveGroup, ...negativeGroup, ...disqualifiedGroup]
    .map(r => r.slice(0, DATA_COL_COUNT));  // E~L 8열만
  const outputRange = sheet.getRange(2, DATA_START_COL, finalData.length, DATA_COL_COUNT);
  outputRange.setValues(finalData);

  // ── 서식 ──
  applyStyles(sheet, outputRange, finalData);
  ss.toast(`'${sheetName}' 데이터 분석 및 서식 적용 완료!`);
}

// ═══════════════════════════════════════════════════════
//  계산 함수
//  row index 기준 (E=0, F=1, G=2, H=3, I=4, J=5, K=6(가점), L=7(PQ하한))
// ═══════════════════════════════════════════════════════
function calculateRankValues(
  rank1Row, positiveGroup, negativeGroup, disqualifiedGroup,
  actualRate, thresholdAmt, lowerLimitRate, basePrice, scheduledPrice, classCode
) {
  const calcJ = (amt) =>
    (lowerLimitRate > 0 && basePrice > 0) ? amt / lowerLimitRate / basePrice : 0;
  // L(PQ하한/필요PQ점수)은 이 시점엔 K열(가점)이 아직 fillBonusScores로 채워지기
  // 전이라 여기서 계산해도 stale한 값이 되고, 곧바로 updatePQTierColumns()가
  // 덮어쓴다. 중복 계산을 피하기 위해 여기서는 비워두고 L 계산은
  // updatePQTierColumns()에 전적으로 위임한다.

  // ── 1순위 ──
  rank1Row[5]        = calcJ(rank1Row[4]);
  rank1Row[7]        = "";

  // ── 정상 그룹 ──
  positiveGroup.forEach((r, i) => {
    r[0]      = i + 2;
    r[5]      = calcJ(r[4]);
    r[7]      = "";
  });

  // ── 하한미달 그룹 ──
  negativeGroup.forEach((r, i) => {
    r[0]       = -(i + 1);
    r[5]       = calcJ(r[4]);
    r[7]       = "";
  });

  // ── 제외 그룹 ──
  // r[0]은 분류 단계(updateBiddingResults)에서 이미 "예가초과" 또는 "적격점수미달"로 지정됨
  disqualifiedGroup.forEach(r => {
    r[5] = calcJ(r[4]);
    r[7] = "";
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  L열(PQ 하한/필요PQ점수) 최종 계산 함수 (2024.12.3 훈령 기준)
//  (※ 2026-06-23 이전에는 이 결과를 K열에 기록했으나, 현재는 L열에 기록함.
//     이 함수 자체의 산식/로직은 변경하지 않음 — 결과 기록 위치만 바뀜)
//  (※ 이 파일의 calculateRankValues()는 더 이상 이 함수를 호출하지 않음 —
//     실제 호출/기록은 updatePQTierColumns.js가 전담)
//
//  로직:
//  1. 투찰률(소수 형태)을 소수점 다섯째자리에서 반올림
//  2. 분류코드별로 입찰가격 평점 계산
//  3. L = (적격통과점수 - 입찰가격평점 - 가산점) / divisor
//  4. 최종값을 소수점 셋째자리에서 반올림
//
//  분류코드별 기준:
//  - 30 (30억 이상): 적격 85점, 입찰가격배점 30, divisor 0.7
//  - 10 (10억~30억): 적격 90점, 입찰가격배점 30, divisor 0.5
//  - 5 (5억~10억): 적격 95점, 입찰가격배점 50, divisor 0.5
//  - 2 (2.3억~5억): 적격 95점, 입찰가격배점 70, divisor 0.3
//  - 1 (≤2.3억): 적격 95점, 입찰가격배점 70, divisor 0.3
// ═══════════════════════════════════════════════════════════════════════════════
function calcPQScore_(hPercent, classCode, bonusScore = 0) {
  const code = parseInt(classCode, 10);

  // ─────────────────────────────────────────────────────────────────────────────
  // 1단계: 투찰률 반올림 (소수점 다섯째자리)
  // hPercent는 0~100 범위의 숫자 (예: 86.8166597820946)
  // ─────────────────────────────────────────────────────────────────────────────
  let hDecimal = hPercent / 100;  // 0.868166597820946
  let hRounded = Math.round(hDecimal * 10000) / 10000;  // 0.8682
  let hForCalc = hRounded * 100;  // 86.82

  // ─────────────────────────────────────────────────────────────────────────────
  // 2단계: 분류코드별 파라미터
  // ─────────────────────────────────────────────────────────────────────────────
  let passScore, bidScoreMax, divisor, hasThreshold = false, threshold, thresholdScore;

  if (code === 30) {
    passScore = 85;
    bidScoreMax = 30;
    divisor = 0.7;
  } else if (code === 10) {
    passScore = 90;
    bidScoreMax = 30;
    divisor = 0.7;
    hasThreshold = true;
    threshold = 89.25;
    thresholdScore = 27;
  } else if (code === 5) {
    passScore = 95;
    bidScoreMax = 50;
    divisor = 0.5;
    hasThreshold = true;
    threshold = 90.5;
    thresholdScore = 45;
  } else if (code === 2 || code === 1) {
    passScore = 95;
    bidScoreMax = 70;
    divisor = 0.3;
    hasThreshold = true;
    threshold = 88.25;
    thresholdScore = 85;
  } else {
    return ""; // 분류코드 불명
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // 3단계: 입찰가격 평점 계산
  // ─────────────────────────────────────────────────────────────────────────────
  let bidScore;
  const diff = Math.abs(88 - hForCalc);

  if (code === 30 || code === 10) {
    // 평점 = 30 - |88 - H|
    bidScore = bidScoreMax - diff;
  } else if (code === 5) {
    // 평점 = 50 - 2×|88 - H|
    bidScore = bidScoreMax - 2 * diff;
  } else if (code === 2 || code === 1) {
    // 평점 = 70 - 4×|88 - H|
    bidScore = bidScoreMax - 4 * diff;
  }

  // 임계값 처리
  if (hasThreshold && hForCalc >= threshold) {
    bidScore = thresholdScore;
  }

  // 최저점 2점
  bidScore = Math.max(bidScore, 2);

  // ─────────────────────────────────────────────────────────────────────────────
  // 4단계: 필요PQ점수(PQ 하한) 계산
  // L = (적격통과점수 - 입찰가격평점 - 가산점) / divisor
  // ─────────────────────────────────────────────────────────────────────────────
  const requiredPerformance = passScore - bidScore - bonusScore;
  const pqScore = requiredPerformance / divisor;

  // ─────────────────────────────────────────────────────────────────────────────
  // 5단계: PQ점수 반올림 (소수점 셋째자리)
  // ─────────────────────────────────────────────────────────────────────────────
  const pqRounded = Math.round(pqScore * 100) / 100;

  return pqRounded;
}

// ═══════════════════════════════════════════════════════
//  가점 divisor — 분류코드별(5→0.5, 10/30→0.7, 그외(1,2)→0.3)
//  L(PQ하한/필요PQ점수) 공식의 divisor와 동일 기준.
//  (2026-06-23: 이 divisor를 쓰던 PQ Gap 계산(updateColumnS.js)은 폐기됨.
//   현재는 calcPQScore_ 내부에서만 동일 기준을 자체적으로 적용하고 있고,
//   이 함수 자체를 외부에서 호출하는 곳은 없음 — 참고용으로 유지)
// ═══════════════════════════════════════════════════════
function getPqBonusDivisor_(classCode) {
  const code = parseInt(classCode, 10);
  if (code === 5) return 0.5;
  if (code === 10 || code === 30) return 0.7;
  return 0.3; // 1, 2, 그외
}

// ═══════════════════════════════════════════════════════
//  서식 적용 (E~L)
//  K,L 숫자서식은 fillBonusScores 실행 전이라 의미 없는 값(가점 미기입/L 공백)
//  단계라 여기서 지정하지 않음 — K,L 숫자서식은 updatePQTierColumns.js/
//  formatAnalysisColumns()가 담당. M~Q는 updatePQTierColumns.js가 관리.
// ═══════════════════════════════════════════════════════
function applyStyles(sheet, range, data) {
  const rowCount  = data.length;
  const startRow  = range.getRow();

  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(startRow, 8,  rowCount, 1).setNumberFormat("0.000%");  // H 투찰률
  sheet.getRange(startRow, 9,  rowCount, 1).setNumberFormat("#,##0");   // I 투찰금액
  sheet.getRange(startRow, 10, rowCount, 1).setNumberFormat("0.000%");  // J 판단예가율

  sheet.getRange(startRow, 5,  rowCount, 2).setHorizontalAlignment("center"); // E,F
  sheet.getRange(startRow, 7,  rowCount, 1).setHorizontalAlignment("left");   // G
  sheet.getRange(startRow, 8,  rowCount, 3).setHorizontalAlignment("right");  // H,I,J (K,L 정렬은 formatAnalysisColumns()가 담당)

  const backgrounds = data.map(r => {
    let color = null;
    if      (r[0] == 1)              color = "#D9EAD3";
    else if (String(r[2]).includes("정우")) color = "#CFE2F3";
    else if (r[0] == -1)             color = "#FCE5CD";
    else if (typeof r[0] === "string" && (r[0].includes("초과") || r[0].includes("미달"))) color = "#F4CCCC";
    return Array(r.length).fill(color);  // r.length = DATA_COL_COUNT(8)에 맞춰 동적 생성 — range 폭과 불일치 시 setBackgrounds가 에러를 던지므로 고정값 대신 사용
  });
  range.setBackgrounds(backgrounds);
}
