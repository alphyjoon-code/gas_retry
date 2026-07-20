/**
 * 용역파일의 데이터를 결과 파일(ID: 191f4xeEmXslCS6...WI)로 전송
 */
function updateResultFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  
  // 1. 용역파일(Source) 데이터 추출
  const serviceName     = sourceSheet.getRange("B1").getValue();   // B1:  용역명
  const estPrice        = sourceSheet.getRange("B12").getValue();   // B12: 예정가격 (B10→B12)
  const realRate        = sourceSheet.getRange("B13").getValue();   // B13: 실예가율  (B11→B13)
  const totalCompanies  = sourceSheet.getRange("B15").getValue();   // B15: 업체수    (B13→B15)
  
  // 순위 테이블(E:J) 데이터 (순위, 업체명, 투찰금액, 판단예가율)
  const rankingData = sourceSheet.getRange("E2:J100").getValues(); 
  
  let dataMap = {
    "1st":    { name: "", price: 0 },
    "정우":   { rank: "", price: 0 },
    "정우종합": { rank: "", price: 0 }
  };
  
  rankingData.forEach(row => {
    const rank      = row[0];  // E열: 순위
    const coName    = row[2];  // G열: 업체명
    const bidPrice  = row[4];  // I열: 투찰금액
    const judgeRate = row[5];  // J열: 판단예가율
    
    if (rank === 1) {
      dataMap["1st"] = { name: coName, price: bidPrice };
    }
    if (coName === "정우") {
      dataMap["정우"] = { rank: rank, price: bidPrice };
    }
    if (coName === "정우종합") {
      dataMap["정우종합"] = { rank: rank, price: bidPrice };
    }
  });

  // 2. 결과파일(Target) 연결
  const targetId    = "191f4xeEmXslCS6n3FLpagvKx3grCg_q3bRsmXC6VpWI";
  const targetSs    = SpreadsheetApp.openById(targetId);
  const targetSheet = targetSs.getSheetByName("통합"); 
  const targetData  = targetSheet.getDataRange().getValues();
  
  let updateCount = 0;
  
  // 3. 결과파일 순회하며 업데이트
  for (let i = 0; i < targetData.length; i++) {
    const targetServiceName = targetData[i][2]; // C열: 용역명
    const targetCompany     = targetData[i][1]; // B열: 업체명
    
    if (targetServiceName && targetServiceName.includes(serviceName)) {
      const rowIdx = i + 1;
      
      // 공통 정보 업데이트
      targetSheet.getRange(rowIdx, 18).setValue(estPrice);               // R열: 예정가격
      targetSheet.getRange(rowIdx, 21).setValue(totalCompanies);         // U열: 업체수
      targetSheet.getRange(rowIdx, 30).setValue(dataMap["1st"].name);    // AD열: 1순위 업체명
      targetSheet.getRange(rowIdx, 31).setValue(dataMap["1st"].price);   // AE열: 1위가격

      // 자사(정우/정우종합)별 개별 정보 업데이트
      if (targetCompany === "정우" && dataMap["정우"].rank !== "") {
        targetSheet.getRange(rowIdx, 14).setValue(dataMap["정우"].price); // N열: 투찰가격
        targetSheet.getRange(rowIdx, 20).setValue(dataMap["정우"].rank);  // T열: 결과(순위)
        updateCount++;
      } else if (targetCompany === "정우종합" && dataMap["정우종합"].rank !== "") {
        targetSheet.getRange(rowIdx, 14).setValue(dataMap["정우종합"].price); // N열
        targetSheet.getRange(rowIdx, 20).setValue(dataMap["정우종합"].rank);  // T열: 결과(순위)
        updateCount++;
      }
    }
  }
  
  if (updateCount > 0) {
    SpreadsheetApp.getUi().alert("'" + serviceName + "'\n총 " + updateCount + "건의 행이 업데이트되었습니다.");
  } else {
    SpreadsheetApp.getUi().alert("일치하는 용역명 또는 업체를 찾을 수 없습니다.");
  }
}