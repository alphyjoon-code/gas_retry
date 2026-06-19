function moveCompetitionMetricsSafely() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    
    // 1. '용역'이 포함된 시트만 대상으로 함
    if (sheetName.indexOf('용역') === -1) return;
    
    // 2. O1, P1 헤더가 있을 때만 실행 (안전장치)
    var valO = sheet.getRange("O1").getValue();
    var valP = sheet.getRange("P1").getValue();
    
    if (valO === '판단예가율 구간' && valP === '경쟁') {
      
      // 3. 6개 열 삽입 (O열 위치) - 이때 서식이 N열로부터 복제됨
      sheet.insertColumnsBefore(15, 6);
      
      // 4. 삽입된 O:T 영역(15~20번 열)의 일반 서식(배경색, 테두리 등) 즉시 초기화
      var newRange = sheet.getRange(1, 15, sheet.getMaxRows(), 6);
      newRange.clearFormat();
      newRange.setBackground(null);
      
      // 5. [중요] 늘어난 '조건부 서식' 범위를 원래대로 복구
      // N열에서 O~T열까지 강제로 확장된 조건부 서식 규칙들을 찾아 범위를 재수정합니다.
      var rules = sheet.getConditionalFormatRules();
      var adjustedRules = [];
      
      for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var oldRanges = rule.getRanges();
        var fixedRanges = [];
        
        oldRanges.forEach(function(range) {
          var startCol = range.getColumn();
          var lastCol = range.getLastColumn();
          
          // 규칙 범위가 삽입된 열(15~20)을 포함하여 확장된 경우
          if (startCol < 15 && lastCol >= 15) {
            // 왼쪽(N열까지)과 오른쪽(U열부터)으로 범위를 쪼개서 다시 설정
            fixedRanges.push(sheet.getRange(range.getRow(), startCol, range.getNumRows(), 14 - startCol + 1));
            if (lastCol > 20) {
              fixedRanges.push(sheet.getRange(range.getRow(), 21, range.getNumRows(), lastCol - 20));
            }
          } else {
            fixedRanges.push(range);
          }
        });
        
        // 수정된 범위로 규칙 업데이트
        adjustedRules.push(rule.copy().setRanges(fixedRanges).build());
      }
      sheet.setConditionalFormatRules(adjustedRules);
      
      console.log(sheetName + " 시트: 지표 이동 및 서식 복구 완료");
    }
  });
}