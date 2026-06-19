/**
 * 업체별 개별 시트 갱신(순위 포함) + 도트플롯용 통합 로데이터 생성 통합 스크립트
 */
function extractAndPrepareDotPlot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // [수정 1] GT 시트 → 업체매핑 시트 C열(단축명)으로 변경
  const mappingSheet = ss.getSheetByName("업체매핑");
  if (!mappingSheet) {
    SpreadsheetApp.getUi().alert("'업체매핑' 시트가 없습니다.");
    return;
  }

  // [설정] 외부 저장소 파일 ID
  const targetSpreadsheetId = "1zyEiDQVvUHFW0a-3OqloNeDAu9-i2Pvm7978z9jNqNc";
  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(targetSpreadsheetId);
  } catch (e) {
    SpreadsheetApp.getUi().alert("대상 파일(외부 저장소)을 열 수 없습니다. ID와 권한을 확인하세요.");
    return;
  }

  // 1. 업체 리스트 수집 (업체매핑 시트 C열 단축명)
  const lastRowMapping = mappingSheet.getLastRow();
  const companyNames = mappingSheet
    .getRange(2, 3, lastRowMapping - 1, 1)
    .getValues()
    .flat()
    .filter(n => n.toString().trim() !== "");

  const masterData = {};
  companyNames.forEach(name => { masterData[name] = []; });

  // 2. 전수 조사 및 데이터 수집
  // [표준 필터링 규칙] 시트 추가/변경 시 여기만 수정
  const FIXED_EXCLUDES = [
    "GT", "Form", "🏠 네비게이션", "분석용_RawData", "업체매핑", "누적현황"
  ];
  const EXCLUDE_KEYWORDS = [
    "시뮬레이션", "차트", "백데이터", "실예가", "분석용", "실적", "exc", "통계"
  ];

  const shouldExclude = (name) =>
    FIXED_EXCLUDES.includes(name) ||
    EXCLUDE_KEYWORDS.some(keyword => name.includes(keyword));

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (shouldExclude(sheetName)) return;

    // G2가 비어있으면 업체 데이터 없는 시트로 간주, 스킵
    const firstVendor = sheet.getRange("G2").getValue().toString().trim();
    if (firstVendor === "") return;

    // [수정 2] 실제 셀 위치에 맞게 commonInfo 수정
    const commonInfo = [
      sheet.getRange("B14").getValue(), // 1) 개찰일
      sheet.getRange("B1").getValue(),  // 2) 용역명
      sheet.getRange("B2").getValue(),  // 3) 발주처
      sheet.getRange("B4").getValue(),  // 4) 채널
      sheet.getRange("B10").getValue(), // 5) 분류코드
      sheet.getRange("B15").getValue(), // 6) 업체수
      sheet.getRange("B6").getValue(),  // 7) 유사면적
      sheet.getRange("B11").getValue(), // 8) 기초가격
      sheet.getRange("B13").getValue()  // 9) 실예가율 (% 그대로 읽음)
    ];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // E~O열 (인덱스 0~10): E=순위, F=사업자번호, G=업체명, H=투찰률, I=투찰금액,
    // J=판단예가율, K=필요PQ점수, L=추정PQ점수, M=가점, N=추정낙찰하한율, O=추정예가율
    // (K/L/N/O는 용역시트에서 updateColumnQ()/updateColumnR() 등이 이미 계산해 둔
    //  값을 그대로 읽음 — 여기서 다시 계산하지 않음)
    const rows = sheet.getRange(2, 5, lastRow - 1, 11).getValues();

    rows.forEach(row => {
      const nameInRow = row[2] ? row[2].toString().trim() : ""; // G열 = 인덱스 2
      if (nameInRow !== "" && masterData.hasOwnProperty(nameInRow)) {
        const biddingAmount = row[4]; // I열 투찰금액 (인덱스 4)

        masterData[nameInRow].push([
          ...commonInfo,
          row[5],        // 10) 판단예가율 (J열 = 인덱스 5)
          biddingAmount, // 11) 투찰금액   (I열 = 인덱스 4)
          row[0],        // 12) 순위       (E열 = 인덱스 0)
          row[6],        // 13) 필요PQ점수     (K열 = 인덱스 6)
          row[7],        // 14) 추정PQ점수     (L열 = 인덱스 7)
          row[9],        // 15) 추정낙찰하한율 (N열 = 인덱스 9)
          row[10]        // 16) 추정예가율     (O열 = 인덱스 10)
        ]);
      }
    });
  });

  // 3. 업체별 개별 시트 갱신
  const headers = [
    "개찰일", "용역명", "발주처", "채널", "분류코드",
    "업체수", "유사면적", "기초가격", "실예가율",
    "판단예가율", "투찰금액", "순위",
    "필요PQ점수", "추정PQ점수", "추정낙찰하한율", "추정예가율"
  ];

  for (const company in masterData) {
    if (masterData[company].length === 0) continue;

    const safeName = company.toString().substring(0, 30).replace(/[\[\]\?\*\/\\\:]/g, "");
    let companySheet = targetSS.getSheetByName(safeName);

    if (!companySheet) {
      companySheet = targetSS.insertSheet(safeName);
    } else {
      companySheet.clear();
    }

    companySheet.appendRow(headers);
    const dataRows = masterData[company].length;
    const dataRange = companySheet.getRange(2, 1, dataRows, headers.length);
    dataRange.setValues(masterData[company]);
    dataRange.sort({ column: 1, ascending: true });

    const fullRange = companySheet.getRange(1, 1, dataRows + 1, headers.length);
    fullRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    fullRange.setHorizontalAlignment("center");

    companySheet.getRange(2, 1, dataRows, 3).setHorizontalAlignment("left");
    companySheet.getRange(2, 6, dataRows, 11).setHorizontalAlignment("right"); // 업체수~추정예가율(6~16열)
    companySheet.getRange(1, 1, 1, headers.length).setBackground("#f3f3f3").setFontWeight("bold");
    companySheet.getRange(2, 1, dataRows, 1).setNumberFormat("yyyy-mm-dd");
    companySheet.getRange(2, 7, dataRows, 2).setNumberFormat("#,##0");
    companySheet.getRange(2, 11, dataRows, 1).setNumberFormat("#,##0");
    companySheet.getRange(2, 9, dataRows, 2).setNumberFormat("0.0000%");
    companySheet.getRange(2, 13, dataRows, 2).setNumberFormat("0.00");      // 필요PQ점수/추정PQ점수
    companySheet.getRange(2, 15, dataRows, 2).setNumberFormat("0.0000%");  // 추정낙찰하한율/추정예가율
    companySheet.setColumnWidth(2, 250);
    companySheet.setColumnWidth(3, 150);
    companySheet.setColumnWidth(12, 50);
  }

  // 4. 도트플롯용 통합 로데이터 시트 생성
  // 데이터가 하나도 없으면 중단 (시트 날아가는 것 방지)
  const hasData = Object.values(masterData).some(arr => arr.length > 0);
  if (!hasData) {
    SpreadsheetApp.getUi().alert(
      "수집된 데이터가 없습니다.\n" +
      "업체매핑 C열의 단축명과 용역 시트 G열의 업체명이 일치하는지 확인하세요.\n" +
      "Apps Script > 실행 로그에서 상세 내용을 확인할 수 있습니다."
    );
    return;
  }

  let dotDataSheet = targetSS.getSheetByName("도트플롯_로데이터");
  if (!dotDataSheet) {
    dotDataSheet = targetSS.insertSheet("도트플롯_로데이터");
  } else {
    dotDataSheet.clear();
  }

  const dotTable = [];
  let maxObservations = 0;

  for (const company in masterData) {
    if (masterData[company].length === 0) continue;
    // 판단예가율 = masterData 배열의 인덱스 9 (commonInfo 9개 + 판단예가율)
    const rates = masterData[company].map(item => item[9]);
    maxObservations = Math.max(maxObservations, rates.length);
    dotTable.push([company, ...rates]);
  }

  if (dotTable.length > 0) {
    const dotHeader = ["업체명"];
    for (let i = 1; i <= maxObservations; i++) dotHeader.push("투찰" + i);

    dotDataSheet
      .getRange(1, 1, 1, dotHeader.length)
      .setValues([dotHeader])
      .setBackground("#f3f3f3")
      .setFontWeight("bold");

    const finalTable = dotTable.map(row => {
      const padding = new Array(maxObservations + 1 - row.length).fill("");
      return row.concat(padding);
    });

    dotDataSheet
      .getRange(2, 1, finalTable.length, dotHeader.length)
      .setValues(finalTable)
      .setNumberFormat("0.0000%");

    dotDataSheet.setFrozenColumns(1);
  }

  // 5. 시트 정렬
  sortTargetSheets(targetSS);

  SpreadsheetApp.getUi().alert("업체별 시트와 정렬, 도트플롯 데이터 갱신이 모두 완료되었습니다.");
}

/**
 * 외부 스프레드시트의 시트들을 가나다순으로 정렬하는 함수
 */
function sortTargetSheets(spreadsheet) {
  const FIXED_SHEETS = ["도트플롯_로데이터", "전체투찰차트"];

  // 실제 존재하는 고정 시트만 추림
  const existingFixed = FIXED_SHEETS.filter(name => spreadsheet.getSheetByName(name) !== null);

  const sheetNames = spreadsheet.getSheets()
    .map(s => s.getName())
    .filter(name => !FIXED_SHEETS.includes(name));

  sheetNames.sort();

  // 업체 시트들을 고정 시트 수 이후 순서로 배치
  sheetNames.forEach((name, index) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(existingFixed.length + index + 1);
  });

  // 고정 시트들을 맨 앞으로 이동
  existingFixed.forEach((fixedName, i) => {
    const fixedSheet = spreadsheet.getSheetByName(fixedName);
    if (!fixedSheet) return;
    spreadsheet.setActiveSheet(fixedSheet);
    spreadsheet.moveActiveSheet(i + 1);
  });
}

// 2026-06-19 롤백: K/L/N/O는 용역시트의 updateColumnQ()/updateColumnR() 등
// 점추정 파이프라인이 이미 계산해 둔 값을 그대로 읽기만 합니다. 참조표
// 조회 로직(getReferenceLimitMap_)은 updateColumnQ.js에 있습니다.