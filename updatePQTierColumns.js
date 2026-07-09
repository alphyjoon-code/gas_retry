/**
 * [신규] K~Q열 PQ 티어별 추정예가율 계산
 *
 * 열 구조 (신규 기준):
 *   K열 (11) : 가점
 *   L열 (12) : PQ 하한 (필요PQ점수, 기존 calcPQScore_ 로직)
 *   M열 (13) : 추정예가율 @ PQ 100점
 *   N열 (14) : 추정예가율 @ PQ 98.2점
 *   O열 (15) : 추정예가율 @ PQ 96.4점
 *   P열 (16) : 추정예가율 @ PQ 94.6점
 *   Q열 (17) : 추정예가율 @ PQ 92.8점
 *
 * 전제:
 *   - E~J열(순위/사업자번호/업체명/투찰률/투찰금액/판단예가율)은 이미 채워진 상태
 *   - K열(가점)은 fillBonusScores가 사전에 채운 상태
 *   - B10: 분류코드, B11: 기초가격
 *
 * 추정예가율 공식:
 *   추정예가율 = 투찰금액(I열) ÷ (기초가격(B11) × 낙찰하한율(해당PQ점수))
 *   ※ 낙찰하한율 조회 시 L열(calcPQScore_)과 동일하게 K열 가점을 반영:
 *     유효PQ = 티어PQ + (가점 ÷ divisor) → 해당 유효PQ 이하 최고 티어의 낙찰하한율 사용
 *
 * 조건부 서식:
 *   티어 PQ점수 < L열(PQ하한)인 경우 해당 셀 배경을 회색(#CCCCCC)으로 표시
 *   (해당 PQ점수로는 적격 통과 불가 → 현실적으로 불가능한 시나리오)
 */

// ─────────────────────────────────────────────────────────────────────────────
// [하드코딩] PQ 티어별 낙찰하한율 테이블
// 출처: "PQ점수별 낙찰하한율_반올림까지 고려한 값" (2024.12.3 훈령 기준)
// 훈령 개정 시 이 테이블만 업데이트할 것
// 단위: 소수 형태 (예: 0.8674500)
// 분류코드 1은 코드 2와 동일 산식 적용
// ─────────────────────────────────────────────────────────────────────────────
const PQ_TIER_LIMIT_TABLE = {
  //          코드2(=1)    코드5       코드10      코드30
  "100.00": { 2: 0.8674500, 5: 0.8549500, 10: 0.779950, 30: 0.729950 },
  "98.20":  { 2: 0.8688500, 5: 0.8594500, 10: 0.792550, 30: 0.742550 },
  "96.40":  { 2: 0.8701500, 5: 0.8637500, 10: 0.804450, 30: 0.754450 },
  "94.60":  { 2: 0.8714500, 5: 0.8680500, 10: 0.816350, 30: 0.766350 },
  "92.80":  { 2: 0.8727500, 5: 0.8722500, 10: 0.828250, 30: 0.778250 },
};

// 티어 순서 (M~Q열 순서와 일치)
const PQ_TIERS = ["100.00", "98.20", "96.40", "94.60", "92.80"];

// 조건부 서식용 색상
const COLOR_IMPOSSIBLE = "#CCCCCC"; // PQ하한 미달 티어 (불가능한 시나리오)
const COLOR_DEFAULT    = null;       // 기본 배경 (투명)


// ─────────────────────────────────────────────────────────────────────────────
// M~Q 추정예가율 계산 (L열 calcPQScore_와 동일하게 가점 반영)
// L = (적격통과점수 - 입찰가격평점 - 가점) / divisor 이므로,
// 가점은 PQ divisor 만큼의 PQ 점수와 동등 → 유효PQ = 티어PQ + 가점/divisor
// ─────────────────────────────────────────────────────────────────────────────
function getTierLimitRate_(tierScore, codeKey, bonus, classCode) {
  const divisor = getPqBonusDivisor_(classCode);
  const effectivePq = tierScore + (bonus / divisor);

  // PQ_TIERS는 내림차순(100→92.8). effectivePq 이하 최고 티어의 낙찰하한율 사용
  let selectedTier = PQ_TIERS[PQ_TIERS.length - 1];
  for (const tier of PQ_TIERS) {
    if (parseFloat(tier) <= effectivePq) {
      selectedTier = tier;
      break;
    }
  }
  return PQ_TIER_LIMIT_TABLE[selectedTier][codeKey];
}

function calcEstRateAtTier_(bidAmt, basePrice, tierScore, classCode, codeKey, bonus) {
  const limitRate = getTierLimitRate_(tierScore, codeKey, bonus, classCode);
  if (!(bidAmt > 0 && basePrice > 0 && limitRate > 0)) return "";
  return bidAmt / (basePrice * limitRate);
}


// ─────────────────────────────────────────────────────────────────────────────
// 메인 함수
// ─────────────────────────────────────────────────────────────────────────────
function updatePQTierColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // ── 기준값 읽기 ──
  const classCodeRaw = sheet.getRange("B10").getValue();
  const basePrice    = parseFloat(sheet.getRange("B11").getValue()); // 기초가격
  const classCode    = parseInt(classCodeRaw, 10);

  if (!basePrice || basePrice <= 0) {
    SpreadsheetApp.getUi().alert("B11(기초가격)이 올바르지 않습니다.");
    return;
  }

  // 분류코드 정규화: 코드 1 → 코드 2와 동일
  const codeKey = (classCode === 1) ? 2 : classCode;
  if (![2, 5, 10, 30].includes(codeKey)) {
    SpreadsheetApp.getUi().alert("지원하지 않는 분류코드입니다: " + classCode);
    return;
  }

  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;
  const dataRows = lastRow - 1;

  // ── 헤더 설정 ──
  sheet.getRange("K1").setValue("가점").setHorizontalAlignment("center");
  sheet.getRange("L1").setValue("필요PQ점수").setHorizontalAlignment("center");
  sheet.getRange("M1").setValue("추정예가율@100").setHorizontalAlignment("center");
  sheet.getRange("N1").setValue("추정예가율@98.2").setHorizontalAlignment("center");
  sheet.getRange("O1").setValue("추정예가율@96.4").setHorizontalAlignment("center");
  sheet.getRange("P1").setValue("추정예가율@94.6").setHorizontalAlignment("center");
  sheet.getRange("Q1").setValue("추정예가율@92.8").setHorizontalAlignment("center");

  // ── 기존 데이터 읽기 ──
  // H열(8): 투찰률, I열(9): 투찰금액, K열(11): 가점
  const hValues = sheet.getRange(2, 8,  dataRows, 1).getValues(); // 투찰률
  const iValues = sheet.getRange(2, 9,  dataRows, 1).getValues(); // 투찰금액
  const kValues = sheet.getRange(2, 11, dataRows, 1).getValues(); // 가점 (fillBonusScores가 채운 값)

  // ── 계산 ──
  const lResults = []; // PQ 하한
  const mResults = []; // 티어별 추정예가율 (5개 열)
  const bgColors = []; // 조건부 서식용 배경색 (5개 열)

  for (let i = 0; i < dataRows; i++) {
    const hPercent = hValues[i][0]; // 투찰률 (소수 형태, 예: 0.86781)
    const bidAmt   = parseFloat(iValues[i][0]) || 0; // 투찰금액
    const bonus    = parseFloat(kValues[i][0]) || 0;  // 가점

    // hPercent가 소수 형태(0~1)로 저장된 경우 ×100 변환
    // (applyStyles에서 H열을 "0.000%" 서식으로 저장하므로 실제 값은 소수 형태)
    const hPct = (hPercent < 1 && hPercent > 0) ? hPercent * 100 : hPercent;

    // L열: PQ 하한 계산
    const pqMin = (hPct > 0 && bidAmt > 0)
      ? calcPQScore_(hPct, classCode, bonus)
      : "";
    lResults.push([pqMin]);

    // M~Q열: 티어별 추정예가율 계산
    const tierRates = [];
    const tierColors = [];

    for (const tier of PQ_TIERS) {
      const tierScore = parseFloat(tier);
      const estRate = calcEstRateAtTier_(bidAmt, basePrice, tierScore, classCode, codeKey, bonus);
      tierRates.push(estRate);

      // 조건부 서식: 이 티어 PQ점수 < PQ하한이면 회색
      const isImpossible = (pqMin !== "" && tierScore < pqMin);
      tierColors.push(isImpossible ? COLOR_IMPOSSIBLE : COLOR_DEFAULT);
    }

    mResults.push(tierRates);
    bgColors.push(tierColors);
  }

  // ── 시트 반영 ──
  // L열: PQ 하한
  const lRange = sheet.getRange(2, 12, dataRows, 1);
  lRange.setValues(lResults).setNumberFormat("0.00").setHorizontalAlignment("right");

  // M~Q열: 티어별 추정예가율
  const tierRange = sheet.getRange(2, 13, dataRows, 5);
  tierRange.setValues(mResults).setNumberFormat("0.000%").setHorizontalAlignment("right");

  // 조건부 서식 (배경색)
  tierRange.setBackgrounds(bgColors);

  // 테두리
  sheet.getRange(2, 12, dataRows, 6).setBorder(
    true, true, true, true, true, true,
    "black", SpreadsheetApp.BorderStyle.SOLID
  );

  ss.toast("K~Q열(가점/PQ하한/티어별 추정예가율) 계산 완료", "✅ 완료");
}
