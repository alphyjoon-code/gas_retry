function analyzeSizeData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // --- 1. 기준 정보 수집 (공고일, 기준면적, 용역명) ---
  var dateValue = activeSheet.getRange('B19').getValue();
  var targetDate = parseSafeDate(dateValue); // 안전하게 날짜 객체로 변환
  
  if (!targetDate || isNaN(targetDate.getTime())) {
    ui.alert('오류', 'B19 셀에 올바른 날짜 형식이 입력되어 있지 않습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var areaValue = activeSheet.getRange('B6').getValue();
  var inputArea = parseFloat(areaValue);
  
  if (isNaN(inputArea)) {
    ui.alert('오류', 'B6 셀에 올바른 숫자가 입력되어 있지 않습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var newSheetName = activeSheet.getRange('B1').getValue();
  if (!newSheetName || newSheetName.toString().trim() === '') {
    newSheetName = '분석결과_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  } else {
    newSheetName = newSheetName.toString().trim();
  }
  
  // --- 2. 조회 기간(5년) 산정 및 '정수(YYYYMMDD)' 변환 ---
  // 종료일 = 공고일 (예: 20260702)
  var endDateInt = toDateInt(targetDate); 
  
  // 시작일 = 공고일 - 5년 + 1일 (예: 20210703)
  var startDate = new Date(targetDate.getTime());
  startDate.setFullYear(startDate.getFullYear() - 5);
  startDate.setDate(startDate.getDate() + 1); 
  var startDateInt = toDateInt(startDate);
  
  // --- 3. 데이터 조회 ---
  var sourceSheet = ss.getSheetByName('경쟁업체실적5000이상');
  if (!sourceSheet) {
    ui.alert('오류', "'경쟁업체실적5000이상' 시트를 찾을 수 없습니다.", ui.ButtonSet.OK);
    return;
  }
  
  var data = sourceSheet.getDataRange().getValues();
  var companyData = {}; 
  
  // --- 4. 데이터 필터링 (기획 조건과 정확히 일치하게 검증) ---
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var projectName = row[0]; // A열: 용역명
    var expectedCompletionDate = parseSafeDate(row[4]); // E열: 예상준공일
    var areaData = row[5]; // F열: 발주면적
    var company = row[6]; // G열: 낙찰업체
    
    // 유효한 날짜가 아니면 패스
    if (!expectedCompletionDate || isNaN(expectedCompletionDate.getTime())) continue;
    
    // 비교 대상의 준공일도 정수(YYYYMMDD)로 변환
    var compDateInt = toDateInt(expectedCompletionDate);
    
    // 면적 데이터 검증
    if (areaData === '' || areaData === '-' || isNaN(parseFloat(areaData))) continue;
    var area = parseFloat(areaData);
    
    // [가장 중요한 부분] 기획한 대로 숫자(면적 및 정수화된 날짜) 크기 비교
    // 예: 면적이 6035 이상이고, 20210703 <= 준공일 <= 20260702 일 것
    if (area >= inputArea && compDateInt >= startDateInt && compDateInt <= endDateInt) {
      if (!companyData[company]) {
        companyData[company] = [];
      }
      companyData[company].push(projectName);
    }
  }
  
  // --- 5. 결과 가공 및 출력 ---
  var outputData = [['수행업체', '건수', '해당 용역명(검증용)']];
  for (var comp in companyData) {
    if (comp.toString().trim() !== '') {
      var count = companyData[comp].length;
      var projectList = companyData[comp].join('\n');
      outputData.push([comp, count, projectList]);
    }
  }
  
  var headers = outputData.slice(0, 1);
  var rows = outputData.slice(1);
  rows.sort(function(a, b) {
    return b[1] - a[1];
  });
  
  var finalData = headers.concat(rows);
  
  if (finalData.length <= 1) {
    ui.alert('알림', '조건에 맞는 실적 데이터가 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var targetUrl = 'https://docs.google.com/spreadsheets/d/1iDAqTQ_zEAFUiVOvJDiMEtSwdRoeQfM-ifhE_bU5HI4/edit?gid=0#gid=0';
  var targetSs;
  try {
    targetSs = SpreadsheetApp.openByUrl(targetUrl);
  } catch (e) {
    ui.alert('오류', '출력할 대상 파일을 열 수 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var finalSheetName = newSheetName;
  if (targetSs.getSheetByName(finalSheetName)) {
    var timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
    finalSheetName = finalSheetName + '_' + timeStamp;
  }
  
  var newSheet = targetSs.insertSheet(finalSheetName);
  var dataRange = newSheet.getRange(1, 1, finalData.length, 3);
  dataRange.setValues(finalData);
  
  dataRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
  dataRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  if (rows.length > 0) {
    var cColumnDataRange = newSheet.getRange(2, 3, rows.length, 1);
    cColumnDataRange.setHorizontalAlignment('left').setWrap(true);
  }
  
  var headerRange = newSheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  newSheet.setColumnWidth(1, 150); 
  newSheet.setColumnWidth(2, 80);  
  newSheet.setColumnWidth(3, 500); 
  newSheet.setFrozenRows(1);
  
  ui.alert('완료', '분석이 완료되었습니다.\n[' + finalSheetName + '] 시트의 C열에서 조건에 부합하는 용역명을 검증해 보세요.', ui.ButtonSet.OK);
}


/**
 * 보조 함수 1: Date 객체를 YYYYMMDD 형태의 정수로 변환 (예: 20260702)
 */
function toDateInt(dateObj) {
  var y = dateObj.getFullYear();
  var m = dateObj.getMonth() + 1;
  var d = dateObj.getDate();
  return y * 10000 + m * 100 + d; 
}

/**
 * 보조 함수 2: 셀의 값이 문자열(Text) 형태인 경우에도 안전하게 타임존 오차 없이 Date 객체로 파싱
 */
function parseSafeDate(cellValue) {
  if (cellValue instanceof Date) {
    return cellValue;
  }
  if (typeof cellValue === 'string') {
    var parts = cellValue.split('-'); // "2026-07-02" 형식일 경우
    if (parts.length === 3) {
      return new Date(parts[0], parts[1] - 1, parts[2]); // 로컬 시간으로 강제 할당
    }
  }
  return new Date(cellValue);
}