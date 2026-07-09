function analyzeSizeData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 현재 활성화된(보고 있는) 시트 가져오기
  var activeSheet = ss.getActiveSheet();
  
  // 1. 시트에서 공고일, 면적, 용역명(새 시트 이름) 읽어오기
  var dateValue = activeSheet.getRange('B19').getValue(); // B19: 공고일
  var targetDate = new Date(dateValue);
  
  if (isNaN(targetDate.getTime())) {
    ui.alert('오류', 'B19 셀에 올바른 날짜 형식이 입력되어 있지 않습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var areaValue = activeSheet.getRange('B6').getValue(); // B6: 유사 면적
  var inputArea = parseFloat(areaValue);
  
  if (isNaN(inputArea)) {
    ui.alert('오류', 'B6 셀에 올바른 숫자가 입력되어 있지 않습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var newSheetName = activeSheet.getRange('B1').getValue(); // B1: 용역명
  // 용역명이 비어있을 경우 대비
  if (!newSheetName || newSheetName.toString().trim() === '') {
    newSheetName = '분석결과_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  } else {
    newSheetName = newSheetName.toString().trim();
  }
  
  // 2. 공고일 기준 5년 이내 기간 산정
  var endDate = new Date(targetDate.getTime());
  var startDate = new Date(targetDate.getTime());
  startDate.setFullYear(startDate.getFullYear() - 5);
  startDate.setDate(startDate.getDate() + 1); // 5년 전 다음 날
  
  // 3. 같은 파일 내의 '경쟁업체실적5000이상' 시트에서 데이터 가져오기
  var sourceSheet = ss.getSheetByName('경쟁업체실적5000이상');
  
  if (!sourceSheet) {
    ui.alert('오류', "'경쟁업체실적5000이상' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.", ui.ButtonSet.OK);
    return;
  }
  
  var data = sourceSheet.getDataRange().getValues();
  var companyCounts = {};
  
  // 4. 데이터 필터링 및 업체별 건수 집계
  // E열(인덱스 4): 예상준공일, F열(인덱스 5): 발주면적, G열(인덱스 6): 낙찰업체
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var expectedCompletionDate = new Date(row[4]);
    var areaData = row[5];
    var company = row[6];
    
    if (areaData === '' || areaData === '-' || isNaN(parseFloat(areaData))) {
      continue;
    }
    var area = parseFloat(areaData);
    
    if (area >= inputArea && 
        expectedCompletionDate >= startDate && 
        expectedCompletionDate <= endDate) {
      
      if (companyCounts[company]) {
        companyCounts[company]++;
      } else {
        companyCounts[company] = 1;
      }
    }
  }
  
  // 5. 결과 데이터 배열 구성
  var outputData = [['수행업체', '건수']];
  for (var comp in companyCounts) {
    if (comp.toString().trim() !== '') {
      outputData.push([comp, companyCounts[comp]]);
    }
  }
  
  // 건수 기준 내림차순 정렬
  var headers = outputData.slice(0, 1);
  var rows = outputData.slice(1);
  rows.sort(function(a, b) {
    return b[1] - a[1];
  });
  
  var finalData = headers.concat(rows);
  
  // 조건에 맞는 데이터가 없으면 종료
  if (finalData.length <= 1) {
    ui.alert('알림', '조건에 맞는 실적 데이터가 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  // 6. 지정된 새 스프레드시트 파일 열기 및 시트 생성
  var targetUrl = 'https://docs.google.com/spreadsheets/d/1iDAqTQ_zEAFUiVOvJDiMEtSwdRoeQfM-ifhE_bU5HI4/edit?gid=0#gid=0';
  var targetSs;
  
  try {
    targetSs = SpreadsheetApp.openByUrl(targetUrl);
  } catch (e) {
    ui.alert('오류', '출력할 대상 파일을 열 수 없습니다. 권한이 없거나 URL이 잘못되었습니다.', ui.ButtonSet.OK);
    return;
  }
  
  // 새 파일에 동일한 이름의 시트가 이미 있는 경우 에러 방지를 위해 시간 덧붙이기
  var finalSheetName = newSheetName;
  if (targetSs.getSheetByName(finalSheetName)) {
    var timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
    finalSheetName = finalSheetName + '_' + timeStamp;
  }
  
  var newSheet = targetSs.insertSheet(finalSheetName);
  
  // 7. [수정됨] 시트에 데이터 쓰기 및 시인성(가독성) 대폭 개선
  var dataRange = newSheet.getRange(1, 1, finalData.length, 2);
  dataRange.setValues(finalData);
  
  // 가운데 정렬 및 테두리 적용
  dataRange.setHorizontalAlignment('center');
  dataRange.setVerticalAlignment('middle');
  dataRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // 헤더(1행) 디자인
  var headerRange = newSheet.getRange(1, 1, 1, 2);
  headerRange.setBackground('#4285F4'); // 구글 블루 색상
  headerRange.setFontColor('#FFFFFF');  // 흰색 글씨
  headerRange.setFontWeight('bold');
  
  // 열 너비 수동 조정 (자동 조정보다 여유롭게)
  newSheet.setColumnWidth(1, 200); // 수행업체 열
  newSheet.setColumnWidth(2, 100); // 건수 열
  
  // 1행 틀 고정 (스크롤해도 헤더 유지)
  newSheet.setFrozenRows(1);
  
  ui.alert('완료', '분석이 완료되어 대상 파일에 깔끔하게 정리된 [' + finalSheetName + '] 시트가 생성되었습니다.', ui.ButtonSet.OK);
}