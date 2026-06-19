/**
 * [기능 1] 업체명 단축 정리 (현재 활성화된 시트만 대상)
 * - 1차: 업체명으로 매핑
 * - 2차: 업체명 매핑 실패 시 사업자등록번호(숫자만 추출)로 매핑
 * @param {boolean} [silent=false] true면 alert 없이 조용히 종료 (runAllBidAnalysis용)
 */
function convertCompanyNamesToShort(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mappingSheet = ss.getSheetByName("업체매핑");

  // UI 알림 헬퍼 (편집기 테스트 실행 시 getUi() 오류 방지)
  function showAlert(msg) {
    if (silent) return;
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (e) {
      Logger.log("[알림] " + msg);
    }
  }

  if (!mappingSheet) {
    showAlert("'업체매핑' 시트가 없습니다.");
    return;
  }

  // 1. 매핑 데이터 가져오기 (업체매핑 시트 기준: A=업체명, B=사업자등록번호, C=단축명)
  const lastMappingRow = mappingSheet.getLastRow();
  if (lastMappingRow < 2) {
    showAlert("업체매핑 시트에 매핑 데이터가 없습니다.");
    return;
  }
  const mappingData = mappingSheet.getRange("A2:C" + lastMappingRow).getValues();

  // 업체명 → 단축명 맵
  const nameMap = {};
  // 사업자등록번호(숫자만) → 단축명 맵
  const bizMap = {};

  mappingData.forEach(row => {
    const companyName = row[0] ? row[0].toString().trim() : "";
    const bizNo       = row[1] ? row[1].toString().replace(/[^0-9]/g, "").trim() : "";
    const shortName   = row[2] ? row[2].toString().trim() : "";

    if (companyName && shortName) nameMap[companyName] = shortName;
    if (bizNo && shortName)       bizMap[bizNo]        = shortName;
  });

  // 2. 현재 활성화된 시트 가져오기 및 표준 필터 적용
  const currentSheet = ss.getActiveSheet();
  const sheetName = currentSheet.getName();

  // [표준 필터링 로직]
  const fixedExcludes   = ["GT", "Form", "🏠 네비게이션", "분석용_RawData", "업체매핑"];
  const excludeKeywords = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "분석용", "실적"];

  const isExcluded = fixedExcludes.includes(sheetName) ||
                     excludeKeywords.some(kw => sheetName.includes(kw));

  if (isExcluded) {
    showAlert("이 시트('" + sheetName + "')는 변환 대상이 아닙니다.");
    return;
  }

  // 3. 업체명 변환 작업 실행 (7열: G열 기준)
  const lastRow = currentSheet.getLastRow();
  if (lastRow < 2) {
    showAlert("변환할 데이터가 없습니다.");
    return;
  }

  // G열(7열)의 데이터만 정확히 가져옵니다.
  const range  = currentSheet.getRange(2, 7, lastRow - 1, 1);
  const values = range.getValues();
  let changeCount = 0;

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) continue; // 빈 셀은 건너뛰기

    const cellValue = values[i][0].toString().trim();

    // 1차: 업체명으로 매핑
    if (nameMap[cellValue]) {
      values[i][0] = nameMap[cellValue];
      changeCount++;
      continue;
    }

    // 2차: 사업자등록번호(숫자만 추출)로 매핑
    const digitsOnly = cellValue.replace(/[^0-9]/g, "");
    if (digitsOnly && bizMap[digitsOnly]) {
      values[i][0] = bizMap[digitsOnly];
      changeCount++;
    }
  }

  // 4. 변경사항 반영 (실제 변경된 데이터가 있을 때만 시트에 반영)
  if (changeCount > 0) {
    range.setValues(values);
    ss.toast("'" + sheetName + "' 시트에서 " + changeCount + "건의 업체명을 단축 정리했습니다.", "완료");
  } else {
    ss.toast("변환할 업체명이 없습니다.", "알림");
  }
}