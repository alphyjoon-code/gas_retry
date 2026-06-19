/**
 * Q열(특이사항/추론) 자동 계산 스크립트 [점추정 복원판]
 * - 2026-06-19 롤백: 범위 기반(N~O 예가율범위 이탈 여부) 재설계를 폐기하고,
 *   PQ 범위추정 도입 이전의 우선순위 캐스케이드로 되돌림.
 *
 * ※ 주의: 이 캐스케이드는 git 이력이나 원본 코드 백업이 없는 상태에서
 *   마스터문서(PQ입찰분석시스템_마스터문서.md, 2026-06-16 작성, 두 차례의
 *   재설계 이전 버전을 근거로 한 메모 수준 설명)를 바탕으로 재구성한 것이다.
 *   문구·우선순위가 100% 원본과 동일하다는 보장은 없으니, 실데이터로
 *   한 번 검증해 보는 것을 권장한다.
 *
 * 우선순위 캐스케이드:
 *   1) E열에 "예가초과"/"적격점수미달" 문자 플래그가 있으면 그대로 표시
 *   2) E열(순위)이 음수면 "낙찰하한선 미달"
 *   3) P열(PQ Gap)이 음수면(가점을 더해도 필요PQ점수 미달) "부적격추정"
 *   4) 그 외에는 판단예가율(J)의 위치로 3구간 분류
 *      - 99.5% ≤ J ≤ 100.3% → "중간틈새(C)"
 *      - J < 100.8%          → "공격하단(A)"
 *      - J ≥ 100.8%          → "수세상단(B)"
 */
function updateColumnT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;

  // 데이터 영역 로드: E(5)열부터 Q(17)열까지 (총 13개 열)
  // 인덱스 기준: E(0,순위), G(2,업체명), J(5,판단예가율), P(11,PQ Gap)
  const range = sheet.getRange(2, 5, lastRow - 1, 13);
  const values = range.getValues();
  const results = [];
  const etBackgrounds = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const rank = row[0];                  // E열: 순위 (또는 "예가초과"/"적격점수미달")
    const company = String(row[2] ?? ""); // G열: 업체명/단축명
    const jVal = Number(row[5]);          // J열: 판단예가율
    const pRaw = row[11];                 // P열: PQ Gap

    let inference = "";

    if (typeof rank === "string" && rank.includes("초과")) {
      inference = "예가초과";
    } else if (typeof rank === "string" && rank.includes("적격점")) {
      inference = "적격점수미달";
    } else if (typeof rank === "number" && rank < 0) {
      inference = "낙찰하한선 미달";
    } else if (pRaw !== "" && isFinite(Number(pRaw)) && Number(pRaw) < 0) {
      inference = "부적격추정";
    } else if (jVal >= 0.995 && jVal <= 1.003) {
      inference = "중간틈새(C)";
    } else if (jVal < 1.008) {
      inference = "공격하단(A)";
    } else {
      inference = "수세상단(B)";
    }

    results.push([inference]);

    // 분류 그룹 시각 구분용 배경색 (Q열만 적용)
    let groupBg = null;
    if      (inference === "공격하단(A)")   groupBg = "#FCE5CD"; // 연주황
    else if (inference === "수세상단(B)")   groupBg = "#CFE2F3"; // 연파랑
    else if (inference === "중간틈새(C)")   groupBg = "#D9EAD3"; // 연초록
    else if (inference === "부적격추정")    groupBg = "#F4CCCC"; // 연빨강

    // 행 하이라이트(1순위, -1순위, 자사)를 E~Q 전체에 적용
    // 우선순위: 1순위 > -1순위 > 자사 > 예가초과/미달 > 기타(그대로)
    let rowBg = null;
    if (rank == 1) rowBg = "#D9EAD3";
    else if (rank == -1) rowBg = "#FCE5CD";
    else if (company.includes("정우")) rowBg = "#CFE2F3";
    else if (typeof rank === "string" && (rank.includes("초과") || rank.includes("미달"))) rowBg = "#F4CCCC";

    const rowColors = new Array(13).fill(rowBg);
    // Q열은 분류 그룹색이 있으면 그룹색을 우선 적용(행 하이라이트보다 우선)
    rowColors[12] = groupBg ?? rowBg;
    etBackgrounds.push(rowColors);
  }

  // Q열(17번째 열)에 결과값을 일괄 기입
  const tRange = sheet.getRange(2, 17, results.length, 1);
  tRange.setValues(results);

  // E~Q(5~17열) 전체 행 배경색 동기화 (Q열은 그룹색 포함)
  sheet.getRange(2, 5, results.length, 13).setBackgrounds(etBackgrounds);

  ss.toast("Q열 특이사항 추론(점추정 복원: 낙찰하한선미달/부적격추정/A·B·C) 업데이트 완료", "✅ 완료");
}

// 기존 짧은 함수명 호환용
function updateColumn() {
  return updateColumnT();
}
