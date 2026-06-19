function insertNewFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 제외할 시트 - 완전일치
  const excludeExactNames = ["GT", "Form", "🏠 네비게이션", "업체매핑"];
  
  // 제외할 시트 - 시트명에 포함된 키워드
  const excludeKeywords = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "통계", "통합", "실적"];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // 완전일치 제외
    if (excludeExactNames.includes(sheetName)) {
      Logger.log(`스킵(완전일치): ${sheetName}`);
      return;
    }
    
    // 키워드 포함 제외
    if (excludeKeywords.some(keyword => sheetName.includes(keyword))) {
      Logger.log(`스킵(키워드): ${sheetName}`);
      return;
    }

    // ── 1단계: 행5 위치에 "사업 면적" 삽입 ──────────────────
    // A~D열의 행5부터 아래 데이터를 한 칸 아래로 이동
    const lastRow1 = getLastRowInRange(sheet, 5);
    if (lastRow1 >= 5) {
      sheet.getRange(5, 1, lastRow1 - 5 + 1, 4)
           .copyTo(sheet.getRange(6, 1, lastRow1 - 5 + 1, 4),
                   SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    }
    sheet.getRange(5, 1).setValue('사업 면적');
    sheet.getRange(5, 2).setValue('');

    // ── 2단계: 행7 위치에 "인정 기준" 삽입 ──────────────────
    // 1단계 후: 행6=유사면적, 행7=추정가격(기존) → 행7 앞에 삽입
    const lastRow2 = getLastRowInRange(sheet, 7);
    if (lastRow2 >= 7) {
      sheet.getRange(7, 1, lastRow2 - 7 + 1, 4)
           .copyTo(sheet.getRange(8, 1, lastRow2 - 7 + 1, 4),
                   SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    }
    sheet.getRange(7, 1).setValue('인정 기준');
    sheet.getRange(7, 2).setValue('');

    Logger.log(`완료: ${sheetName}`);
  });

  SpreadsheetApp.getUi().alert('모든 시트 업데이트 완료!');
}

// A열 기준으로 해당 행부터 아래 마지막 데이터 행 반환
function getLastRowInRange(sheet, startRow) {
  const maxRow = sheet.getLastRow();
  if (maxRow < startRow) return startRow;
  
  const data = sheet.getRange(startRow, 1, maxRow - startRow + 1, 1).getValues();
  let last = startRow;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== '') last = startRow + i;
  }
  return last;
}