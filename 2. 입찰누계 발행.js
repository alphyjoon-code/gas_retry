/**
 * [기능 3] PQ 입찰 누계 리포트 갱신 (퍼센트 오류 완벽 수정본)
 */
function updateMasterDashboard() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gtSheet = ss.getSheetByName("GT");
    if (!gtSheet) { SpreadsheetApp.getUi().alert("'GT' 시트를 찾을 수 없습니다."); return; }
  
    const targetFileId = "1mdXYtzi2yWBViAqFvVyzmpaT5EqGF5HWCFmC4KD49Y8";
    const targetSS = SpreadsheetApp.openById(targetFileId);
    
    let dash = targetSS.getSheetByName("1순위누계") || targetSS.insertSheet("1순위누계");
    dash.clear();
    targetSS.setActiveSheet(dash);
    targetSS.moveActiveSheet(1);
  
    // --- [설정] 시트 제외 로직 (키워드 및 고유 명칭) ---
    const excludeKeywords = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "(exc)", "분석용", "실적", "업체매핑"];
    const excludeExactNames = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "누적현황", "통합", "PQ백데이터", "PQ백데이터(분석용)", "통계_대시보드", "최근 30일 실예가"];
    // -----------------------------------------------
  
    // 1. 업체명 매핑 데이터 수집
    const nameMap = {};
    const mappingRows = gtSheet.getRange("F2:G" + gtSheet.getLastRow()).getValues();
    mappingRows.forEach(r => {
      if (r[0] && r[1]) nameMap[r[0].toString().trim()] = r[1].toString().trim();
    });
  
    const allData = [];
    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
  
      // --- [제외 로직 적용] ---
      const isExcluded = excludeExactNames.includes(name) || 
                         excludeKeywords.some(kw => name.includes(kw)) ||
                         name.match(/^\d{4}-\d{2}-\d{2}$/);
      
      if (isExcluded) return;
      // -----------------------
  
      const price  = Number(sheet.getRange("B11").getValue()) || 0;  // B9  → B11
      const codeVal = sheet.getRange("B10").getValue();               // B8  → B10
      const code   = codeVal ? codeVal.toString().trim() : "미분류";
      const firstVal = sheet.getRange("B16").getValue();              // B14 → B16
      
      if (price > 100 && firstVal) {
        const firstFull  = firstVal.toString().trim();
        const firstShort = nameMap[firstFull] || firstFull;
        allData.push({ price: price, code: code, first: firstShort });
      }
    });
  
    if (allData.length === 0) return;
  
    // 2. 리포트 생성
    dash.getRange("B1").setValue("2026 PQ입찰 1순위 누계 현황").setFontSize(22).setFontWeight("bold");
    let startRow = 3;
    const targetCodes = ["전체", "1", "2", "5", "10", "30"];
  
    targetCodes.forEach(targetCode => {
      let filtered = targetCode === "전체" ? allData : allData.filter(d => d.code === targetCode);
      if (filtered.length === 0) return;
  
      let sectionTotalSum = filtered.reduce((acc, curr) => acc + curr.price, 0);
      let jwSum = filtered.filter(d => d.first === "정우" || d.first === "정우종합").reduce((acc, curr) => acc + curr.price, 0);
      
      let winMap = {};
      filtered.forEach(d => { 
        if (!winMap[d.first]) winMap[d.first] = { sum: 0, count: 0 };
        winMap[d.first].sum += d.price;
        winMap[d.first].count += 1;
      });
  
      let sortedRank = Object.keys(winMap).map(name => [name, winMap[name]]).sort((a, b) => b[1].sum - a[1].sum);
  
      // 가. 입찰 1순위 누계
      dash.getRange(startRow, 2).setValue("■ [" + targetCode + "] 입찰 1순위 누계").setFontWeight("bold").setFontSize(14);
      const statsHeader = [["총 발주 규모", "자사(정우계열) 수주액", "자사 시장 점유율"]];
      const statsVal = [[sectionTotalSum, jwSum, sectionTotalSum > 0 ? jwSum / sectionTotalSum : 0]];
      
      dash.getRange(startRow + 1, 2, 1, 3).setValues(statsHeader).setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
      dash.getRange(startRow + 2, 2, 1, 3).setValues(statsVal).setBorder(true,true,true,true,true,true).setHorizontalAlignment("center");
      dash.getRange(startRow + 2, 2, 1, 2).setNumberFormat("#,##0\"원\"");
      
      dash.getRange(startRow + 2, 4).setNumberFormat("0.0%").setBackground("#fff2cc").setFontWeight("bold");
  
      // 나. 업체별 누적 수주 현황
      let displayRank = targetCode === "전체" ? sortedRank : sortedRank.slice(0, 5);
      let rankHeader = [["순위", "업체명", "수주 건수", "수주액 누계", "점유율"]];
      let rankRows = displayRank.map((v, i) => [
        i + 1, v[0], v[1].count, v[1].sum, sectionTotalSum > 0 ? v[1].sum / sectionTotalSum : 0
      ]);
  
      dash.getRange(startRow + 1, 6, 1, 5).setValues(rankHeader).setBackground("#34495e").setFontColor("white").setHorizontalAlignment("center");
      let rDataRange = dash.getRange(startRow + 2, 6, rankRows.length, 5);
      rDataRange.setValues(rankRows).setBorder(true,true,true,true,true,true).setFontSize(10);
      
      dash.getRange(startRow + 2, 6, rankRows.length, 2).setHorizontalAlignment("center"); 
      dash.getRange(startRow + 2, 8, rankRows.length, 1).setHorizontalAlignment("center").setNumberFormat("#,##0\"건\""); 
      dash.getRange(startRow + 2, 9, rankRows.length, 1).setNumberFormat("#,##0\"원\"").setHorizontalAlignment("right");
      
      dash.getRange(startRow + 2, 10, rankRows.length, 1).setNumberFormat("0.0%").setHorizontalAlignment("right");
  
      startRow += (rankRows.length + 5); 
    });
  
    // 너비 설정
    dash.setColumnWidth(2, 200); dash.setColumnWidth(3, 200); dash.setColumnWidth(4, 130);
    dash.setColumnWidth(6, 45); dash.setColumnWidth(7, 160); dash.setColumnWidth(8, 85); dash.setColumnWidth(9, 200); dash.setColumnWidth(10, 90);
  
    SpreadsheetApp.getActiveSpreadsheet().toast("PQ입찰 1순위 누계가 발행되었습니다.", "알림", 5);
  }