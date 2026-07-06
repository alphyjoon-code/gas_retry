/**
 * [포맷팅] 분석 결과 열(K~Q) 테두리 + 정렬 + 숫자서식 적용 [2026-06-23 열구조 개편판]
 * - 신규 열 구조(claude.md 9절): K=가점, L=PQ하한(필요PQ점수), M~Q=PQ 티어별 추정예가율.
 * - L~Q는 updatePQTierColumns()가 자체적으로 테두리/숫자서식/조건부 배경색까지
 *   이미 적용하므로, 이 함수는 그 범위와 겹치되 값 자체는 건드리지 않는(서식만
 *   재적용하는) 선이라 안전함. 이 함수가 유일하게 책임지는 건 K(updatePQTierColumns가
 *   서식을 적용하지 않는 열)의 테두리/정렬/숫자서식이다.
 * - 테두리: K~Q 전체
 * - 정렬: K~Q 전체 중앙(Q도 더 이상 텍스트 특이사항이 아니라 수치형 예가율이므로
 *   기존 "Q는 텍스트라 좌측" 예외를 제거함)
 * - 숫자서식: K,L은 0.00(가점/PQ점수 단위), M~Q는 0.000%(추정예가율 단위)
 */
function formatAnalysisColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // sheet.getLastRow()는 R·S열 경쟁강도 참고표 때문에 부풀려질 수 있어
  // F열(사업자등록번호) 기준 실제 데이터 마지막 행을 사용한다.
  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;

  const dataRows = lastRow - 1; // 2행부터 마지막 행까지

  // 테두리: K~Q 열(11~17번째)
  const analysisCols = sheet.getRange(2, 11, dataRows, 7); // K열~Q열
  analysisCols.setBorder(
    true, true, true, true, true, true,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );

  // 정렬: K~Q 전체 중앙(모두 수치형)
  sheet.getRange(2, 11, dataRows, 7).setHorizontalAlignment("center"); // K~Q

  // 숫자서식 통일
  sheet.getRange(2, 11, dataRows, 1).setNumberFormat("0.00");   // K 가점
  sheet.getRange(2, 12, dataRows, 1).setNumberFormat("0.00");   // L PQ하한(필요PQ점수)
  sheet.getRange(2, 13, dataRows, 5).setNumberFormat("0.000%"); // M~Q 티어별 추정예가율
}
