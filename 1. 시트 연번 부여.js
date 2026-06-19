/**
 * [표준 필터 적용] 시트 배열 순서에 따라 B18 셀에 연번 부여
 */
function updateSheetSerialNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // --- 사용자 설정 영역 ---
  const year = "2026";             // 연도 설정
  const targetCell = "B18";        // 번호를 입력할 셀 위치 (행5,7 삽입으로 B16→B18)

  // [표준 필터링 규칙]
  const fixedExcludes = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "업체매핑"];
  const excludeKeywords = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "실적"];
  // -----------------------

  let serialCount = 1; // 연번 시작 값

  sheets.forEach((sheet) => {
    const sheetName = sheet.getName();
    const sheetType = sheet.getType();

    // 1. 표준 필터링 검사
    const isExcluded = fixedExcludes.includes(sheetName) || 
                       excludeKeywords.some(kw => sheetName.includes(kw));

    // 2. 제외 대상이 아니고, 실제 데이터가 있는 그리드 시트인 경우만 처리
    if (!isExcluded && sheetType === SpreadsheetApp.SheetType.GRID) {
      try {
        // 3자리 숫자 형식으로 변환 (예: 001, 002...)
        let formattedIndex = Utilities.formatString('%03d', serialCount);
        let serialNumber = year + "-" + formattedIndex;
        
        // 해당 시트의 B18 셀에 값 입력
        sheet.getRange(targetCell).setValue(serialNumber);
        
        // 다음 번호를 위해 카운트 증가
        serialCount++;
      } catch (e) {
        console.log(`${sheetName} 시트 처리 중 오류 발생: ${e.message}`);
      }
    }
  });
  
  ss.toast("표준 필터에 따라 총 " + (serialCount - 1) + "개 시트에 연번 부여를 완료했습니다.", "알림");
}