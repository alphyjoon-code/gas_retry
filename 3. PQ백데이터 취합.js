/**
 * 예가율 표시 방식을 표준 퍼센트(1.0 = 100%) 방식으로 수집합니다.
 * 키워드 기반 필터링을 통해 시트 제외 로직을 효율화했습니다.
 */
function collectJungwooData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const TARGET_NAME = "PQ백데이터";
  let targetSheet = ss.getSheetByName(TARGET_NAME);
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet(TARGET_NAME);
  }

  // --- 1. 제외 설정 ---
  const excludeKeywords   = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "실적", "업체매핑"];
  const excludeExactNames = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "누적현황"];

  // --- 2. 헤더 생성 (비어있을 경우) ---
  if (targetSheet.getLastRow() === 0) {
    const headers = ["순번", "연번", "용역명", "개찰일", "기초가격", "실예가율", "정우 판단예가", "정우종합 판단예가", "업체수", "유사면적", "분류코드"];
    targetSheet.appendRow(headers);
    targetSheet.getRange(1, 1, 1, headers.length)
               .setBackground("#f3f3f3")
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
  }

  // --- 3. 중복 체크용 기존 연번(B열) 가져오기 ---
  const lastRow = targetSheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) {
    existingIds = targetSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().map(String);
  }

  // --- 4. 시트 순회 및 데이터 수집 ---
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();

    const isExcluded = excludeExactNames.includes(sheetName) || 
                       excludeKeywords.some(kw => sheetName.includes(kw));
    if (isExcluded) return;

    // 데이터 추출 전 기초 검증
    const actualRateRaw = sheet.getRange("B13").getValue();  // B11 → B13
    const firstCompany  = sheet.getRange("G2").getValue();
    if (actualRateRaw === 0 || firstCompany === "") return;

    // 연번 체크
    const serialNum = String(sheet.getRange("B18").getValue()).trim();  // B16 → B18
    if (serialNum === "" || existingIds.includes(serialNum)) return;

    // --- 데이터 추출 영역 ---
    const projectTitle = sheet.getRange("B1").getValue();
    const area         = sheet.getRange("B6").getValue();   // B5  → B6
    const classCode    = sheet.getRange("B10").getValue();  // B8  → B10
    const basePrice    = sheet.getRange("B11").getValue();  // B9  → B11
    const openDate     = sheet.getRange("B14").getValue();  // B12 → B14
    const companyCount = sheet.getRange("B15").getValue();  // B13 → B15

    const data = sheet.getDataRange().getValues();
    let jungwooRate = ""; 
    let jonghapRate = ""; 

    // G열에서 업체 찾기 및 J열 예가 추출
    for (let i = 1; i < data.length; i++) {
      const companyName = String(data[i][6]).trim(); // G열 (Index 6)
      if (companyName === "정우") {
        jungwooRate = data[i][9]; // J열 (Index 9)
      } else if (companyName === "정우종합") {
        jonghapRate = data[i][9]; // J열 (Index 9)
      }
    }

    // 데이터 저장 실행
    if (jungwooRate !== "" || jonghapRate !== "") {
      const nextRow = targetSheet.getLastRow() + 1;
      const newIdx  = nextRow - 1;

      const rowData = [
        newIdx, serialNum, projectTitle, openDate, basePrice, 
        actualRateRaw, jungwooRate, jonghapRate, companyCount, area, classCode
      ];

      targetSheet.appendRow(rowData);

      // --- 서식 지정 ---
      targetSheet.getRange(nextRow, 4).setNumberFormat("yyyy-mm-dd"); // 개찰일
      targetSheet.getRange(nextRow, 5).setNumberFormat("#,##0");       // 기초가격
      targetSheet.getRange(nextRow, 6, 1, 3).setNumberFormat("0.000%"); // 실/판단예가율
      
      console.log(`수집 성공: [${serialNum}] ${projectTitle}`);
    }
  });
  
  SpreadsheetApp.getUi().alert("전체 PQ 백데이터 수집이 완료되었습니다!");
}