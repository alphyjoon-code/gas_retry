/**
 * [점추정 복원판] 데이터 계산 + 서식 복구 (E~K 담당)
 * - 2026-06-19 롤백: "PQ 범위추정"(K=PQ하한/L=PQ상한 등) 도입 이전의
 *   점추정 체계로 되돌림. K열(필요PQ점수, 분류코드별 4갈래 공식)을 다시
 *   이 함수가 직접 채운다. L(추정PQ점수)/N(추정낙찰하한율)/O(추정예가율)/
 *   P(PQ Gap)/Q(특이사항)는 각각 updateColumnQ/R/S/T.js가 채움.
 * - K~M(예가율차이/금액차이/비고) 삭제는 그대로 유지(이건 범위추정 redesign과
 *   무관한 별개의 이전 마이그레이션이라 되돌리지 않음): E~K(7열)
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

  // H1 헤더
  sheet.getRange("H1").setValue("투찰률").setHorizontalAlignment("center");
  sheet.getRange("K1").setValue("필요PQ점수").setHorizontalAlignment("center");

  // ── 데이터 로드 (E~K열, 7열) ──
  const DATA_START_COL = 5;   // E
  const DATA_COL_COUNT = 7;   // E~K (K=필요PQ점수, 이 함수가 직접 계산)
  const rawRange = sheet.getRange(2, DATA_START_COL, lastRow - 1, DATA_COL_COUNT);
  const rawData  = rawRange.getValues();

  // M열(13번째) 원본 비고: 나라장터 원본 데이터는 비적격(예가초과/적격점수미달) 업체에는
  // 순위(E)를 부여하지 않고 비워두는 대신, 비고란에 판정 결과가 들어있는 경우가 많음.
  // 최초 입력(신규 붙여넣기) 시점에는 E열만으로 비적격 여부를 알 수 없으므로 분류 보조용으로 읽어둔다.
  // (이후 fillBonusScores 단계에서 M열은 가점 값으로 덮어써짐)
  const remarkRawAll = sheet.getRange(2, 13, lastRow - 1, 1).getValues();

  // 유효 행 필터 + 투찰률(H열=index 3) 재계산
  // row[7]에 M열 원본 비고를 임시로 동행시켜 분류(정렬 포함) 후 기록 직전에 제거한다.
  // (row[6]=K열은 이제 실제 데이터 칸이므로 더 이상 임시 슬롯으로 쓰지 않음)
  let data = rawData
    .map((row, i) => { row[7] = remarkRawAll[i][0]; return row; })
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
    // E열이 비어있다면(나라장터 원본이 비적격 업체에 순위를 주지 않는 경우) M열 원본 비고로 판정.
    const eFlag = String(row[0] ?? "").trim();
    const flagSource = eFlag || String(row[7] ?? "").trim();
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
  // row[7]에 분류용으로 임시 동행시켰던 M열 원본 비고를 제거하고 E~K 7열만 기록
  const finalData   = [rank1Row, ...positiveGroup, ...negativeGroup, ...disqualifiedGroup]
    .map(r => r.slice(0, DATA_COL_COUNT));
  const outputRange = sheet.getRange(2, DATA_START_COL, finalData.length, DATA_COL_COUNT);
  outputRange.setValues(finalData);

  // ── 서식 ──
  applyStyles(sheet, outputRange, finalData);
  ss.toast(`'${sheetName}' 데이터 분석 및 서식 적용 완료!`);
}


// ═══════════════════════════════════════════════════════
//  계산 함수
//  row index 기준 (E=0, F=1, G=2, H=3, I=4, J=5, K=6) — 점추정 7열 구조
// ═══════════════════════════════════════════════════════
function calculateRankValues(
  rank1Row, positiveGroup, negativeGroup, disqualifiedGroup,
  actualRate, thresholdAmt, lowerLimitRate, basePrice, scheduledPrice, classCode
) {
  const calcJ = (amt) =>
    (lowerLimitRate > 0 && basePrice > 0) ? amt / lowerLimitRate / basePrice : 0;
  const calcK = (r) => calcPQScore_(r[3] * 100, classCode); // H(투찰률)×100 기준 필요PQ점수

  // ── 1순위 ──
  rank1Row[5]        = calcJ(rank1Row[4]);
  rank1Row[6]        = calcK(rank1Row);

  // ── 정상 그룹 ──
  positiveGroup.forEach((r, i) => {
    r[0]      = i + 2;
    r[5]      = calcJ(r[4]);
    r[6]      = calcK(r);
  });

  // ── 하한미달 그룹 ──
  negativeGroup.forEach((r, i) => {
    r[0]       = -(i + 1);
    r[5]       = calcJ(r[4]);
    r[6]       = calcK(r);
  });

  // ── 제외 그룹 ──
  // r[0]은 분류 단계(updateBiddingResults)에서 이미 "예가초과" 또는 "적격점수미달"로 지정됨
  disqualifiedGroup.forEach(r => {
    r[5] = calcJ(r[4]);
    r[6] = calcK(r);
  });
}

// ═══════════════════════════════════════════════════════
//  K열(필요PQ점수) 점추정 공식 — 분류코드별 4갈래 (마스터문서 3-3 기준 복원)
//  H = 투찰률×100. 분류1은 분류2와 동일 산식 적용.
//  최종값은 항상 90 이상으로 floor.
// ═══════════════════════════════════════════════════════
function calcPQScore_(hPercent, classCode) {
  const clamp = (x, lo, hi) => Math.min(Math.max(x, lo), hi);
  const code = parseInt(classCode, 10);
  const codeKey = (code === 1) ? 2 : code;

  let raw;
  if (codeKey === 2) {
    raw = clamp((95 - 70 + 4 * (88 - hPercent)) / 0.3, 2, 100);
  } else if (codeKey === 5) {
    raw = (hPercent >= 90.5) ? 45 : clamp((95 - 50 + 2 * (88 - hPercent)) / 0.5, 2, 100);
  } else if (codeKey === 10) {
    raw = (hPercent >= 89.25) ? 65 : clamp((90 - 30 + (88 - hPercent)) / 0.7, 2, 100);
  } else if (codeKey === 30) {
    raw = clamp((85 - 30 + (88 - hPercent)) / 0.7, 2, 100);
  } else {
    return ""; // 분류코드 불명
  }
  return Math.round(Math.max(raw, 90) * 100) / 100;
}

// ═══════════════════════════════════════════════════════
//  가점 divisor — 분류코드별(5→0.5, 10/30→0.7, 그외(1,2)→0.3)
//  K(필요PQ점수) 공식의 divisor와 동일 기준. PQ Gap(P열) 계산에도 쓰임.
// ═══════════════════════════════════════════════════════
function getPqBonusDivisor_(classCode) {
  const code = parseInt(classCode, 10);
  if (code === 5) return 0.5;
  if (code === 10 || code === 30) return 0.7;
  return 0.3; // 1, 2, 그외
}


// ═══════════════════════════════════════════════════════
//  서식 적용 (E~K)
//  L~Q는 updateColumnQ/R/S/T.js, formatAnalysisColumns()가 관리
// ═══════════════════════════════════════════════════════
function applyStyles(sheet, range, data) {
  const rowCount  = data.length;
  const startRow  = range.getRow();

  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(startRow, 8,  rowCount, 1).setNumberFormat("0.000%");  // H 투찰률
  sheet.getRange(startRow, 9,  rowCount, 1).setNumberFormat("#,##0");   // I 투찰금액
  sheet.getRange(startRow, 10, rowCount, 1).setNumberFormat("0.000%");  // J 판단예가율
  sheet.getRange(startRow, 11, rowCount, 1).setNumberFormat("0.00");    // K 필요PQ점수

  sheet.getRange(startRow, 5,  rowCount, 2).setHorizontalAlignment("center"); // E,F
  sheet.getRange(startRow, 7,  rowCount, 1).setHorizontalAlignment("left");   // G
  sheet.getRange(startRow, 8,  rowCount, 4).setHorizontalAlignment("right");  // H,I,J,K

  const backgrounds = data.map(r => {
    let color = null;
    if      (r[0] == 1)              color = "#D9EAD3";
    else if (String(r[2]).includes("정우")) color = "#CFE2F3";
    else if (r[0] == -1)             color = "#FCE5CD";
    else if (typeof r[0] === "string" && (r[0].includes("초과") || r[0].includes("미달"))) color = "#F4CCCC";
    return Array(7).fill(color);
  });
  range.setBackgrounds(backgrounds);
}