/**
 * [점추정 복원판] P열(PQ Gap) 계산
 * - 2026-06-19 롤백: 범위 기반(N~O 예가율범위 이탈도) 재설계를 폐기하고,
 *   "PQ 범위추정" 도입 이전의 점추정 산식으로 되돌림.
 *
 *   P = ROUND( MIN(L(추정PQ점수) + ROUND(M(가점)/divisor, 2), 100) - K(필요PQ점수), 2 )
 *
 *   divisor는 분류코드별 가점 환산 기준(getPqBonusDivisor_, 1.updateColumnEtoN.gs.js
 *   에 정의: 분류5→0.5, 분류10/30→0.7, 그외(1,2)→0.3)와 동일하다.
 *
 * - L(추정PQ점수)이 비어 있으면 100으로 간주해 계산하고, 그 100을 L셀에도
 *   그대로 채워 넣는다(추정PQ점수 미입력 시 만점으로 보는 기존 의도된 동작).
 */
function updateColumnS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const classCode = parseInt(sheet.getRange("B10").getValue(), 10); // 분류코드
  const divisor = getPqBonusDivisor_(classCode);

  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;

  const kValues = sheet.getRange(2, 11, lastRow - 1, 1).getValues(); // K 필요PQ점수
  const lValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues(); // L 추정PQ점수
  const mValues = sheet.getRange(2, 13, lastRow - 1, 1).getValues(); // M 가점

  const lFix = []; // L열 빈칸 자동보정(100) 기록용
  const results = [];

  for (let i = 0; i < kValues.length; i++) {
    const k = kValues[i][0];
    let l = lValues[i][0];

    if (l === "" || l === null) {
      l = 100; // 추정PQ점수 미입력 시 100으로 간주(의도된 동작)
      lFix.push([100]);
    } else {
      lFix.push([l]);
      l = Number(l);
    }

    const m = Number(mValues[i][0]) || 0;

    if (k === "" || !isFinite(Number(k))) {
      results.push([""]);
      continue;
    }

    const bonus = divisor > 0 ? Math.round((m / divisor) * 100) / 100 : 0;
    const gap = Math.round((Math.min(l + bonus, 100) - Number(k)) * 100) / 100;
    results.push([gap]);
  }

  // L열 빈칸 자동보정(100) 반영
  sheet.getRange(2, 12, lFix.length, 1).setValues(lFix);

  sheet.getRange(2, 16, results.length, 1).setValues(results).setNumberFormat("0.00");
  ss.toast("P열(PQ Gap) 계산 완료", "✅ 완료");
}

// 기존 오타 함수명 호환용
function updateCoulumnS() {
  return updateColumnS();
}
