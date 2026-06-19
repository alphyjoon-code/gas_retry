/**
 * 성능 최적화 + 발주처/업체수 동기화 + Toast 알림 버전
 */
function syncProjectInfoWithToast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheets = ss.getSheets();
  
  // --- [설정] 제외 로직 ---
  const excludeKeywords   = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "exc", "실적", "업체매핑"];
  const excludeExactNames = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "누적현황", "통합", "PQ백데이터", "PQ백데이터(분석용)", "통계_대시보드", "최근 30일 실예가"];
  
  const targetId = "1aG6Sj3r9bpDJS6HV8Wa3g1hbfh4hb63PDDhkM7SQUWE";
  const targetSs = SpreadsheetApp.openById(targetId);
  const targetSheet = targetSs.getSheetByName("복수예가");
  
  if (!targetSheet) {
    ss.toast("❌ '복수예가' 시트를 찾을 수 없습니다.", "오류 발생");
    return;
  }

  // 데이터 로드
  const targetData = targetSheet.getDataRange().getValues();
  const lastRow = targetData.length;
  let cColumnData = targetSheet.getRange(1, 3, lastRow, 1).getValues(); 
  let xColumnData = targetSheet.getRange(1, 24, lastRow, 1).getValues(); 

  // 정규화 맵 생성
  const targetMap = new Map();
  for (let i = 1; i < lastRow; i++) {
    const rawName = targetData[i][1];
    if (rawName) targetMap.set(normalizeName(rawName), i); 
  }

  let updatedCount = 0;
  let failCount = 0;

  // 시트 순회 및 매칭
  sourceSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const isExcluded = excludeExactNames.includes(sheetName) || 
                       excludeKeywords.some(kw => sheetName.includes(kw)) ||
                       sheetName.match(/^\d{4}-\d{2}-\d{2}$/);
    if (isExcluded) return;

    const sourceProjectName = sheet.getRange("B1").getValue().toString();
    const sourceClient      = sheet.getRange("B2").getValue();
    const companyCount      = sheet.getRange("B15").getValue();  // B13 → B15

    if (sourceProjectName) {
      const targetIndex = targetMap.get(normalizeName(sourceProjectName));
      if (targetIndex !== undefined) {
        cColumnData[targetIndex][0] = sourceClient;
        xColumnData[targetIndex][0] = companyCount;
        updatedCount++;
      } else {
        failCount++;
      }
    }
  });

  // 데이터 일괄 반영
  targetSheet.getRange(1, 3, lastRow, 1).setValues(cColumnData);
  targetSheet.getRange(1, 24, lastRow, 1).setValues(xColumnData);

  // --- Toast 알림 ---
  const statusMsg = `✅ ${updatedCount}건 업데이트 완료` + (failCount > 0 ? ` (미매칭 ${failCount}건)` : "");
  ss.toast(statusMsg, "데이터 동기화 완료", 5);
}

/**
 * 이름 정규화 함수
 */
function normalizeName(name) {
  if (!name) return "";
  return name.toString().replace(/[\s\(\)\[\]\-_]/g, "").trim();
}