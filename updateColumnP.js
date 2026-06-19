const BONUS_FILE_ID = "1mzFM-_6GmXtRA3NN7IDN2r01vSk1rHghU2jGbG4cbv8";

function fillBonusScores(identifierArg) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  // ── 제외 시트 필터 ──
  const isExcluded = FIXED_EXCLUDES.includes(sheetName) ||
                     EXCLUDE_KEYWORDS.some(kw => sheetName.includes(kw));
  if (isExcluded) {
    ui.alert(`이 시트('${sheetName}')는 실행 대상이 아닙니다.`);
    return;
  }

  // ── 식별자 결정(인자 우선, 없으면 프롬프트) ──
  let identifier = (identifierArg ?? "").toString().trim();
  if (!identifier) {
    const response = ui.prompt(
      "가산점 식별자 입력",
      "가산점 시트의 식별자를 입력하세요.\n(예: B007, 강설-육-32)",
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    identifier = response.getResponseText().trim();
  }
  if (!identifier) {
    ui.alert('식별자가 입력되지 않아 취소합니다.');
    return;
  }

  // ── 가산점 파일 열기 ──
  let bonusSS;
  try {
    bonusSS = SpreadsheetApp.openById(BONUS_FILE_ID);
  } catch (e) {
    ui.alert('가산점 파일을 열 수 없습니다.\n파일 ID 또는 접근 권한을 확인해 주세요.');
    return;
  }

  // ── 식별자 포함 시트 찾기 (대소문자 무관) ──
  const identifierUpper = identifier.toUpperCase();
  const bonusSheets = bonusSS.getSheets();
  const bonusSheet = bonusSheets.find(s => s.getName().toUpperCase().includes(identifierUpper));
  if (!bonusSheet) {
    ui.alert(`'${identifier}'을(를) 포함한 시트를 가산점 파일에서 찾을 수 없습니다.`);
    return;
  }

  const bonusSheetName = bonusSheet.getName();

  // ── 가산점 결과 시트 읽기 ──
  // 구조: A(번호) / B(단축명) / C(가산점) / D(보유항목)
  const bonusLastRow = bonusSheet.getLastRow();
  if (bonusLastRow < 2) {
    ui.alert(`'${bonusSheetName}' 시트에 데이터가 없습니다.`);
    return;
  }

  const bonusData = bonusSheet.getRange(2, 1, bonusLastRow - 1, 3).getValues();
  // index: 0=A(번호), 1=B(단축명), 2=C(가산점)

  // 단축명 → 점수 맵
  const nameToScore = {};
  bonusData.forEach(row => {
    const shortName = String(row[1]).trim(); // B열: 단축명
    const score     = row[2];               // C열: 가산점
    if (shortName) nameToScore[shortName] = score;
  });

  // ── 입찰결과 시트 데이터 로드 (G열: 단축명) ──
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const targetData = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  // G열(index 0): 단축명

  // ── P열에 가점 기입 ──
  const pValues = targetData.map(row => {
    const name = String(row[0]).trim();
    if (name && nameToScore[name] !== undefined) {
      return [nameToScore[name]];
    } else {
      return [""];
    }
  });

  sheet.getRange(2, 13, pValues.length, 1).setValues(pValues); // M열(13번째, 구 P열 16번째)

  const matched = pValues.filter(v => v[0] !== "").length;
  ss.toast(`'${bonusSheetName}' 기준으로 가산점 ${matched}건 기입 완료!`, '✅ 완료');
}