function analyzeSizeData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  
  // --- 1. 기준 정보 수집 (공고일, 기준면적, 용역명) ---
  var dateValue = activeSheet.getRange('B19').getValue();
  var targetDate = parseSafeDate(dateValue); 
  
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
  var endDateInt = toDateInt(targetDate); 
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
  
  // --- 4. 데이터 필터링 ---
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var projectName = row[0]; 
    var expectedCompletionDate = parseSafeDate(row[4]); 
    var areaData = row[5]; 
    var company = row[6]; 
    
    if (!expectedCompletionDate || isNaN(expectedCompletionDate.getTime())) continue;
    
    var compDateInt = toDateInt(expectedCompletionDate);
    
    if (areaData === '' || areaData === '-' || isNaN(parseFloat(areaData))) continue;
    var area = parseFloat(areaData);
    
    if (area >= inputArea && compDateInt >= startDateInt && compDateInt <= endDateInt) {
      if (!companyData[company]) {
        companyData[company] = [];
      }
      companyData[company].push(projectName);
    }
  }
  
  // --- 5. 결과 가공 및 출력 ---
  var outputData = [['수행업체', '건수', '해당 용역명']];
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
  
  // --- 6. [수정됨] 완료 알림을 Toast 메시지로 변경 ---
  // ss.toast('보여줄 내용', '제목', 노출시간(초))
  ss.toast('[' + finalSheetName + '] 시트 생성 완료!', '실적 분석 완료 🚀', 5);
}

/**
 * 보조 함수 1: Date 객체를 YYYYMMDD 형태의 정수로 변환
 */
function toDateInt(dateObj) {
  var y = dateObj.getFullYear();
  var m = dateObj.getMonth() + 1;
  var d = dateObj.getDate();
  return y * 10000 + m * 100 + d; 
}

/**
 * 보조 함수 2: 안전하게 날짜 객체로 파싱
 */
function parseSafeDate(cellValue) {
  if (cellValue instanceof Date) {
    return cellValue;
  }
  if (typeof cellValue === 'string') {
    var parts = cellValue.split('-'); 
    if (parts.length === 3) {
      return new Date(parts[0], parts[1] - 1, parts[2]); 
    }
  }
  return new Date(cellValue);
}