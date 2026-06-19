/**
 * [기능 2] PQ 일일 현황 발행 (1순위 투찰률 및 예가율 열 추가 버전)
 */
function createLatestPQStatusSheetWithProfessionalSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const gtSheet = ss.getSheetByName("GT");
  
  // 외부 타겟 파일 ID (현황판 전송용)
  const targetFileId = "1mdXYtzi2yWBViAqFvVyzmpaT5EqGF5HWCFmC4KD49Y8";
  let targetSS = SpreadsheetApp.openById(targetFileId);

  // GT 시트의 매핑 데이터 로드 (업체명 약칭 처리)
  const mappingData = gtSheet.getRange("F2:G" + gtSheet.getLastRow()).getValues();
  const nameMap = {};
  mappingData.forEach(row => { 
    if (row[0] && row[1]) nameMap[row[0].toString().trim()] = row[1].toString().trim(); 
  });

  // --- [설정] 시트 제외 로직 (키워드 및 고유 명칭) ---
  const excludeKeywords = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "(exc)", "분석용", "실적", "업체매핑"];
  const excludeExactNames = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "누적현황", "통합", "PQ백데이터", "PQ백데이터(분석용)", "통계_대시보드", "최근 30일 실예가"];
  // -----------------------------------------------

  const allStatusData = [];
  let latestDate = new Date(0);

  // 1. 데이터 수집 단계
  sheets.forEach(sheet => {
    const name = sheet.getName();
    
    const isExcluded = excludeExactNames.includes(name) || 
                       excludeKeywords.some(kw => name.includes(kw)) ||
                       name.match(/^\d{4}-\d{2}-\d{2}$/);
    
    if (isExcluded) return;

    if (!sheet.getRange("G2").getValue()) return;

    const fullSerial = sheet.getRange("B18").getValue().toString().trim();   // B16 → B18
    const serialNumOnly = fullSerial.includes("-") ? parseInt(fullSerial.split("-")[1], 10) : 0;
    
    const serviceName = sheet.getRange("B1").getValue();
    const openDateRaw = sheet.getRange("B14").getValue();   // B12 → B14
    const client      = sheet.getRange("B2").getValue();
    const basePrice   = sheet.getRange("B11").getValue();   // B9  → B11
    const classCode   = sheet.getRange("B10").getValue();   // B8  → B10
    const area        = sheet.getRange("B6").getValue();    // B5  → B6
    const actualRate  = sheet.getRange("B13").getValue();   // B11 → B13
    
    // 1순위 업체 및 투찰 정보
    const firstPlaceShort    = nameMap[sheet.getRange("B16").getValue().toString().trim()] || sheet.getRange("B16").getValue(); // B14 → B16
    const firstPlaceRate     = sheet.getRange("J2").getValue();
    const firstPlaceBidRate  = sheet.getRange("H2").getValue();

    // 정우/정우종합 데이터 추출 (E~J열 순회)
    let jungwooData = { "정우": { rate: "-", rank: "-" }, "정우종합": { rate: "-", rank: "-" } };
    const dataRows = sheet.getRange("E2:J" + sheet.getLastRow()).getValues();
    dataRows.forEach(row => {
      const cFullName = row[2].toString().trim();
      const short = nameMap[cFullName] || cFullName;
      if (short === "정우") { 
        jungwooData["정우"].rank = row[0]; 
        jungwooData["정우"].rate = row[5]; 
      }
      if (short === "정우종합") { 
        jungwooData["정우종합"].rank = row[0]; 
        jungwooData["정우종합"].rate = row[5]; 
      }
    });

    if (serviceName && openDateRaw) {
      const openDate = new Date(openDateRaw);
      allStatusData.push({
        fullSerial: fullSerial,
        serialIdx: serialNumOnly,
        classCode: classCode,
        date: openDate,
        // 누적현황용 데이터 구조
        data: [
          serviceName, openDateRaw, client, basePrice, classCode, area, actualRate, 
          sheet.getRange("B15").getValue(),  // B13 → B15 (업체수)
          firstPlaceShort, 
          firstPlaceRate, 
          sheet.getRange("B17").getValue()   // B15 → B17 (낙찰업체)
        ],
        // 상단 요약용 데이터 구조
        summary: [
          serviceName, openDateRaw, client, actualRate, 
          jungwooData["정우"].rate, jungwooData["정우"].rank, 
          jungwooData["정우종합"].rate, jungwooData["정우종합"].rank, 
          firstPlaceShort, firstPlaceRate, firstPlaceBidRate
        ]
      });
      if (openDate > latestDate) latestDate = openDate;
    }
  });

  // 2. 타겟 시트 준비
  const finalTargetDateStr = Utilities.formatDate(latestDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const titleDateStr = Utilities.formatDate(latestDate, Session.getScriptTimeZone(), "yyyyMMdd");
  
  let targetSheet = targetSS.getSheetByName(finalTargetDateStr) || targetSS.insertSheet(finalTargetDateStr);
  targetSheet.clear();
  targetSS.setActiveSheet(targetSheet);
  targetSS.moveActiveSheet(1);

  // 테이블 공통 스타일 함수
  const applyTableStyle = (range, hasData) => {
    if (!hasData) return;
    range.setFontSize(11).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);
  };

  // --- 1단계. 상단 당일 투찰결과 요약 ---
  targetSheet.getRange("B1").setValue(titleDateStr + " 투찰결과").setFontSize(16).setFontWeight("bold");
  const summaryLabels = [["연번", "용역명", "개찰일", "발주처", "실예가율", "정우(예가율)", "정우(순위)", "정우종합(예가율)", "정우종합(순위)", "1순위 업체", "1순위(예가율)", "1순위(투찰률)"]];
  const todayData = allStatusData.filter(d => Utilities.formatDate(d.date, Session.getScriptTimeZone(), "yyyy-MM-dd") === finalTargetDateStr).sort((a, b) => a.serialIdx - b.serialIdx);
  const summaryRows = todayData.map((d, i) => [i + 1, ...d.summary]);

  if (summaryRows.length > 0) {
    targetSheet.getRange(2, 1, 1, summaryLabels[0].length).setValues(summaryLabels).setBackground("#2c3e50").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    const r = targetSheet.getRange(3, 1, summaryRows.length, summaryLabels[0].length);
    r.setValues(summaryRows);
    applyTableStyle(r, true);
    targetSheet.getRange(3, 1, summaryRows.length, 1).setHorizontalAlignment("center");
    [5, 6, 8, 11, 12].forEach(col => {
      targetSheet.getRange(3, col, summaryRows.length, 1).setNumberFormat("0.0000%");
    });
  }

  // --- 2단계. 중단 전체 누적현황 ---
  let currentRow = (summaryRows.length || 0) + 5;
  targetSheet.getRange(currentRow, 2).setValue("2026 PQ 누적입찰현황").setFontSize(16).setFontWeight("bold");
  const fullHeader = [["연번", "용역명", "개찰일", "발주처", "기초가격", "분류코드", "유사면적", "실예가율", "업체수", "1순위", "1순위(예가율)", "낙찰업체"]];
  const fullRows = allStatusData.sort((a, b) => b.serialIdx - a.serialIdx).map(d => [d.serialIdx, ...d.data]);

  targetSheet.getRange(currentRow + 1, 1, 1, fullHeader[0].length).setValues(fullHeader).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  const fRange = targetSheet.getRange(currentRow + 2, 1, fullRows.length, fullHeader[0].length);
  fRange.setValues(fullRows);
  applyTableStyle(fRange, true);
  targetSheet.getRange(currentRow + 2, 1, fullRows.length, 1).setHorizontalAlignment("center");
  targetSheet.getRange(currentRow + 2, 5, fullRows.length, 1).setNumberFormat("#,##0");
  targetSheet.getRange(currentRow + 2, 8, fullRows.length, 1).setNumberFormat("0.0000%"); 
  targetSheet.getRange(currentRow + 2, 11, fullRows.length, 1).setNumberFormat("0.0000%");

  // --- 3단계. 하단 분류코드별 현황 (2, 5, 10) ---
  const classCodes = [2, 5, 10];
  currentRow += fullRows.length + 5;

  classCodes.forEach(code => {
    targetSheet.getRange(currentRow, 2).setValue(`2026 PQ ${code} 입찰현황`).setFontSize(16).setFontWeight("bold");
    const filteredData = allStatusData.filter(d => Number(d.classCode) === code).sort((a, b) => a.serialIdx - b.serialIdx);
    const rows = filteredData.map(d => [d.serialIdx, ...d.data]);

    targetSheet.getRange(currentRow + 1, 1, 1, fullHeader[0].length).setValues(fullHeader).setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    
    if (rows.length > 0) {
      const r = targetSheet.getRange(currentRow + 2, 1, rows.length, fullHeader[0].length);
      r.setValues(rows);
      applyTableStyle(r, true);
      targetSheet.getRange(currentRow + 2, 1, rows.length, 1).setHorizontalAlignment("center");
      targetSheet.getRange(currentRow + 2, 5, rows.length, 1).setNumberFormat("#,##0");
      targetSheet.getRange(currentRow + 2, 8, rows.length, 1).setNumberFormat("0.0000%");
      targetSheet.getRange(currentRow + 2, 11, rows.length, 1).setNumberFormat("0.0000%");
      currentRow += rows.length + 5;
    } else {
      targetSheet.getRange(currentRow + 2, 2).setValue("해당 데이터 없음").setFontStyle("italic");
      currentRow += 4;
    }
  });

  // --- 디자인 마무리 (열 너비 조정) ---
  targetSheet.setColumnWidth(1, 60); 
  for (let i = 2; i <= fullHeader[0].length; i++) {
    targetSheet.autoResizeColumn(i);
    let autoWidth = targetSheet.getColumnWidth(i);
    targetSheet.setColumnWidth(i, Math.max(autoWidth + 25, 100));
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("PQ 일일현황이 발행되었습니다.", "알림", 5);
}