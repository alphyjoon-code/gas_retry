/**
 * [1회성 마이그레이션] K~M열(예가율차이/금액차이/비고) 삭제
 * - 대상: FIXED_EXCLUDES/EXCLUDE_KEYWORDS(1. updateColumnEtoN.gs.js 정의)에 해당하지 않는
 *         모든 용역별 시트 중, 실제로 구 레이아웃(K~M 보유)인 시트만 처리
 * - 안전장치: M1 셀 값이 "비고"인 시트만 삭제 진행 → 이미 마이그레이션된 시트는 자동 스킵(재실행 안전)
 * - Sheets 기본 동작으로 N열 이후가 자동으로 3열씩 좌측 이동 (N→K, O→L, P→M, Q→N, R→O, S→P, T→Q, U→R, V→S, W→T, X→U)
 * - GAS 코드의 하드코딩 열 참조는 별도로 이미 새 레이아웃 기준으로 수정되어 있음
 *   (updateColumnEtoN/P/Q/R/S/T, formatAnalysisColumns, estimatePQScores, 경쟁강도 분석 2종, 계약누계 발행)
 *
 * ⚠ 실행 전 필수: 반드시 운영 스프레드시트의 "사본"에서 먼저 실행하여 전체 파이프라인을 검증할 것.
 *    (사본에서 검증 완료 후에만 운영 파일에 적용)
 */
function migrateDeleteKLMColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "K~M열 삭제 마이그레이션 실행",
    "현재 스프레드시트의 모든 대상 시트에서 K~M열(예가율차이/금액차이/비고)을 영구 삭제합니다.\n" +
      "이 작업은 되돌리기 어려우니 사본 파일에서 실행 중인지 다시 확인하세요.\n\n계속하시겠습니까?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const sheets = ss.getSheets();
  const results = { migrated: [], skipped: [], excluded: [], errors: [] };

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    const isExcluded = FIXED_EXCLUDES.includes(sheetName) ||
                       EXCLUDE_KEYWORDS.some(kw => sheetName.includes(kw));
    if (isExcluded) {
      results.excluded.push(sheetName);
      return;
    }

    try {
      const lastCol = sheet.getLastColumn();
      if (lastCol < 13) {                         // M열(13번째)까지 데이터/서식이 없으면 대상 아님
        results.skipped.push(sheetName);
        return;
      }

      const m1 = String(sheet.getRange("M1").getValue()).trim();
      if (m1 !== "비고") {
        results.skipped.push(sheetName);          // 이미 마이그레이션됨 또는 구조가 다른 시트로 판단
        return;
      }

      sheet.deleteColumns(11, 3);                 // K(11), L(12), M(13) 삭제
      results.migrated.push(sheetName);

    } catch (e) {
      results.errors.push(sheetName + ": " + e.message);
    }
  });

  const summary =
    `✅ 삭제 완료: ${results.migrated.length}개\n` +
    `⏭ 스킵(대상 아님/이미 처리됨): ${results.skipped.length}개\n` +
    `🚫 제외 시트: ${results.excluded.length}개` +
    (results.errors.length ? `\n❌ 오류: ${results.errors.length}개\n` + results.errors.join("\n") : "");

  Logger.log(summary);
  Logger.log("삭제된 시트 목록: " + results.migrated.join(", "));
  if (results.errors.length) Logger.log("오류 상세: " + results.errors.join(" | "));

  ui.alert("K~M열 삭제 마이그레이션 결과", summary, ui.ButtonSet.OK);
}


/**
 * [점검용] 실제 삭제 없이 마이그레이션 대상/스킵 시트 목록만 미리 확인
 * 사본에서 migrateDeleteKLMColumns() 실행 전에 먼저 돌려서 대상 범위를 검증하는 용도
 */
function previewMigrateDeleteKLMColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheets = ss.getSheets();

  const willMigrate = [];
  const willSkip = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const isExcluded = FIXED_EXCLUDES.includes(sheetName) ||
                       EXCLUDE_KEYWORDS.some(kw => sheetName.includes(kw));
    if (isExcluded) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 13) { willSkip.push(sheetName); return; }

    const m1 = String(sheet.getRange("M1").getValue()).trim();
    if (m1 === "비고") willMigrate.push(sheetName);
    else willSkip.push(sheetName);
  });

  Logger.log("삭제 대상 (" + willMigrate.length + "개): " + willMigrate.join(", "));
  Logger.log("스킵 대상 (" + willSkip.length + "개): " + willSkip.join(", "));

  ui.alert(
    "마이그레이션 미리보기",
    `삭제 대상: ${willMigrate.length}개 시트\n스킵: ${willSkip.length}개 시트\n\n(상세 목록은 실행 로그(Ctrl+Enter 후 보기 ▸ 로그)에서 확인)`,
    ui.ButtonSet.OK
  );
}
