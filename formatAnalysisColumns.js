/**
 * [포맷팅] 분석 결과 열(K~Q) 테두리 + 정렬 + 숫자서식 적용 [점추정 복원판]
 * - 2026-06-19 롤백: K/L/N/O/P를 "PQ 범위추정" 출력 기준 서식에서
 *   점추정(필요PQ점수/추정PQ점수/추정낙찰하한율/추정예가율/PQ Gap) 기준으로 복원.
 * - 테두리: K~Q 전체
 * - 정렬: K~P(수치형) 중앙 / Q(특이사항, 텍스트) 좌측
 * - 숫자서식: K,L,M,P는 0.00(PQ 점수 단위) / N,O는 0.000%(예가율 단위)
 *   M(가점)은 외부 가산점 파일 값을 그대로 받아써서 소수 자릿수가
 *   들쭉날쭉(0.8 / 0.80 / 0 등)했던 것을 통일
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

  // 정렬: K~P(수치형) 중앙, Q(특이사항, 텍스트) 좌측
  sheet.getRange(2, 11, dataRows, 6).setHorizontalAlignment("center"); // K,L,M,N,O,P
  sheet.getRange(2, 17, dataRows, 1).setHorizontalAlignment("left");   // Q

  // 숫자서식 통일
  sheet.getRange(2, 11, dataRows, 1).setNumberFormat("0.00");   // K 필요PQ점수
  sheet.getRange(2, 12, dataRows, 1).setNumberFormat("0.00");   // L 추정PQ점수
  sheet.getRange(2, 13, dataRows, 1).setNumberFormat("0.00");   // M 가점
  sheet.getRange(2, 14, dataRows, 1).setNumberFormat("0.000%"); // N 추정낙찰하한율
  sheet.getRange(2, 15, dataRows, 1).setNumberFormat("0.000%"); // O 추정예가율
  sheet.getRange(2, 16, dataRows, 1).setNumberFormat("0.00");   // P PQ Gap(점수 단위)

  // Q열(17번째) 너비 자동조정
  sheet.autoResizeColumn(17);
  // 최소 너비 보장 (자동조정 후 너무 좁으면 200px로 강제)
  if (sheet.getColumnWidth(17) < 200) {
    sheet.setColumnWidth(17, 200);
  }
}
