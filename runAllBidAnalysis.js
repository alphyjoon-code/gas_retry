/**
 * 용역시트 개찰결과 분석을 한 번에 실행 [점추정 복원판]
 * 2026-06-19 롤백: "PQ 범위추정" 파이프라인(updatePQRangeColumns 단일 호출)을
 * 폐기하고, 그 이전의 점추정 단계별 파이프라인으로 되돌림.
 * 실행 순서(의존성 기준):
 * 1) E~K 계산/정렬 (K=필요PQ점수 포함)
 * 2) G(업체명 단축) 정리
 * 3) M(가점) 채우기
 * 4) N(추정낙찰하한율) 계산 — L(추정PQ점수) 기준 정방향 조회
 * 5) O(추정예가율) 계산 — N 기준
 * 6) P(PQ Gap) 계산 — K/L/M 기준
 * 7) Q(특이사항) 판정 — A/B/C 캐스케이드
 * 8) K~Q 서식(테두리/정렬/숫자서식) 적용
 */
function runAllBidAnalysis() {
  updateBiddingResults(); // E~K 계산/정렬

  convertCompanyNamesToShort(true); // G열 단축명 (M열 매칭 전)

  // M열 가점: B1에서 식별자를 자동 추출 (없으면 fillBonusScores가 프롬프트로 보완)
  const identifier = extractBonusIdentifierFromB1_();
  fillBonusScores(identifier);

  updateColumnQ();  // N(추정낙찰하한율)
  updateColumnR();  // O(추정예가율)
  updateColumnS();  // P(PQ Gap)
  updateColumnT();  // Q(특이사항)

  formatAnalysisColumns(); // K~Q 테두리/정렬/숫자서식

  SpreadsheetApp.getActiveSpreadsheet().toast("전체 분석(EK→G→M→N→O→P→Q→서식) 완료!", "✅ 완료");
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

