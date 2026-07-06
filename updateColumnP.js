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

  const nameToScore = {};
  let sourceLabel = "";

  if (bonusSheet) {
    // ── 식별자별 가산점 시트 읽기 ──
    // 구조: A(번호) / B(단축명) / C(가산점) / D(보유항목)
    const bonusSheetName = bonusSheet.getName();
    const bonusLastRow = bonusSheet.getLastRow();
    if (bonusLastRow < 2) {
      ui.alert(`'${bonusSheetName}' 시트에 데이터가 없습니다.`);
      return;
    }

    const bonusData = bonusSheet.getRange(2, 1, bonusLastRow - 1, 3).getValues();
    // index: 0=A(번호), 1=B(단축명), 2=C(가산점)
    bonusData.forEach(row => {
      const shortName = String(row[1]).trim(); // B열: 단축명
      const score     = row[2];               // C열: 가산점
      if (shortName) nameToScore[shortName] = score;
    });

    sourceLabel = `'${bonusSheetName}'`;

  } else {
    // ── Fallback: 식별자별 시트가 없으면 전체업체 시트("조사대상기업가산점")에서 가져옴 ──
    // 투찰 전 공개되는 참여업체 목록 기준으로 사전에 식별자별 시트를 만들어두는
    // 워크플로우 도입 이전에 입찰한 과거 용역은 식별자별 시트가 없음 — 더 이상
    // alert 후 중단하지 않고 전체업체 가산점 데이터에서 대체 입력한다.
    // 가점 기준(국방부/조달청) 자동 식별 필드는 아직 없으므로 사용자에게 직접 확인한다.
    const agencyResponse = ui.alert(
      '가산점 기준 선택',
      `'${identifier}'을(를) 포함한 식별자별 가산점 시트를 찾을 수 없습니다.\n전체업체 시트("조사대상기업가산점")를 기준으로 대체 입력합니다.\n\n이 용역은 [국방부] 가점 기준인가요?\n\n(예 = 국방부 기준 / 아니오 = 조달청 기준 / 취소 = 가점 입력 중단)`,
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (agencyResponse === ui.Button.CANCEL) {
      ui.alert('가점 기준을 선택하지 않아 취소합니다.');
      return;
    }
    const isMilitary = (agencyResponse === ui.Button.YES);

    const fallbackSheet = bonusSS.getSheetByName('조사대상기업가산점');
    if (!fallbackSheet) {
      ui.alert(`'${identifier}'을(를) 포함한 시트도, 전체업체 시트("조사대상기업가산점")도 가산점 파일에서 찾을 수 없습니다.`);
      return;
    }

    const fbLastRow = fallbackSheet.getLastRow();
    if (fbLastRow < 2) {
      ui.alert(`'조사대상기업가산점' 시트에 데이터가 없습니다.`);
      return;
    }

    // 구조: A연번/B업체명/C사업자등록번호/D단축명/E~K(일반 카테고리)/L(조달청)/M(국방부)
    // L/M 값은 이미 최종 가산점이므로 다른 카테고리와 합산하지 않고 그대로 사용한다.
    const fbData = fallbackSheet.getRange(2, 1, fbLastRow - 1, 13).getValues();
    fbData.forEach(row => {
      const shortName = String(row[3]).trim(); // D열: 단축명
      const score     = isMilitary ? row[12] : row[11]; // M(국방부) 또는 L(조달청)
      if (shortName && score !== "" && score !== null) nameToScore[shortName] = score;
    });

    sourceLabel = `전체업체 기준(${isMilitary ? '국방부' : '조달청'})`;
  }

  // ── 입찰결과 시트 데이터 로드 (G열: 단축명) ──
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const targetData = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  // G열(index 0): 단축명

  // ── K열에 가점 기입 ── (2026-06-23 열구조 개편: 가점은 K열 담당)
  const pValues = targetData.map(row => {
    const name = String(row[0]).trim();
    if (name && nameToScore[name] !== undefined) {
      return [nameToScore[name]];
    } else {
      return [""];
    }
  });

  sheet.getRange(2, 11, pValues.length, 1).setValues(pValues); // K열(11번째, 구 M열 13번째)

  const matched = pValues.filter(v => v[0] !== "").length;
  ss.toast(`${sourceLabel} 기준으로 가산점 ${matched}건 기입 완료!`, '✅ 완료');
}