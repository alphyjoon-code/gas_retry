/**
 * 용역시트 개찰결과 분석을 한 번에 실행 [2026-06-23 열구조 개편판]
 * K~Q 열 구조 개편(claude.md 9절)에 맞춰 파이프라인 교체:
 * - 폐기: updateColumnQ/R/S/T (N~Q가 PQ 티어 열로 재정의되어 더 이상 유효하지 않음)
 * - 추가: updatePQTierColumns (L~Q 일괄 계산, updatePQTierColumns.js)
 * 실행 순서(의존성 기준):
 * 1) E~L 계산/정렬 (L=PQ하한/필요PQ점수 포함)
 * 2) G(업체명 단축) 정리
 * 3) K(가점) 채우기
 * 4) L(PQ하한)/M~Q(PQ 티어별 추정예가율) 계산
 * 5) K~Q 서식(테두리/정렬/숫자서식) 적용
 */
function runAllBidAnalysis() {
  updateBiddingResults(); // E~L 계산/정렬

  convertCompanyNamesToShort(true); // G열 단축명 (K열 매칭 전)

  // K열 가점: B1에서 식별자를 자동 추출 (없으면 fillBonusScores가 프롬프트로 보완)
  const identifier = extractBonusIdentifierFromB1_();
  fillBonusScores(identifier);

  updatePQTierColumns(); // L(PQ하한)/M~Q(PQ 티어별 추정예가율)

  formatAnalysisColumns(); // K~Q 테두리/정렬/숫자서식

  SpreadsheetApp.getActiveSpreadsheet().toast("전체 분석(EL→G→K→L~Q→서식) 완료!", "✅ 완료");
}

/**
 * B1 텍스트에서 (B007) 같은 식별자를 추출.
 * 예: "26-D-... (B007)" -> "B007"
 */
function extractBonusIdentifierFromB1_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const b1 = (sheet.getRange("B1").getValue() ?? "").toString();
  const m = b1.match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : "";
}

