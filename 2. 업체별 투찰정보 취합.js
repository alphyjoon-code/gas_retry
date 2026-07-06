/**
 * 업체별 개별 시트 갱신(순위 포함) + 도트플롯용 통합 로데이터 생성 통합 스크립트
 * 트리거 체이닝 방식으로 5분 제한 우회
 * 
 * 진입점: extractAndPrepareDotPlot()
 * 흐름: 1단계(수집) → _temp 덤프 → 2단계(업체시트, 10개씩) → 3단계(도트플롯) → 완료로그
 */

// ── 설정 ───────────────────────────────────────────────────────────────────
const TARGET_SS_ID = "1zyEiDQVvUHFW0a-3OqloNeDAu9-i2Pvm7978z9jNqNc";
const BATCH_SIZE   = 10;   // 2단계에서 한 번에 처리할 업체 수
const TEMP_SHEET   = "_temp_masterData";
const LOG_SHEET    = "_완료로그";

const HEADERS = [
  "개찰일","용역명","발주처","채널","분류코드",
  "업체수","유사면적","기초가격","실예가율",
  "판단예가율","투찰금액","순위",
  "가점","필요PQ점수",
  "추정예가율@100","추정예가율@98.2","추정예가율@96.4","추정예가율@94.6","추정예가율@92.8"
];

// ── 유틸: 기존 트리거 정리 ─────────────────────────────────────────────────
function _deleteTriggers(fnName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fnName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ── 유틸: 다음 단계 트리거 등록 ───────────────────────────────────────────
function _registerTrigger(fnName, delayMinutes) {
  ScriptApp.newTrigger(fnName)
    .timeBased()
    .after(delayMinutes * 60 * 1000)
    .create();
}

// ══════════════════════════════════════════════════════════════════════════
// 1단계: 데이터 수집 → _temp 덤프 → writeCompanySheets 트리거 등록
// ══════════════════════════════════════════════════════════════════════════
function extractAndPrepareDotPlot() {
  // 혹시 남아있을 이전 트리거 정리
  _deleteTriggers("writeCompanySheets");
  _deleteTriggers("writeDotPlotSheet");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mappingSheet = ss.getSheetByName("업체매핑");
  if (!mappingSheet) {
    SpreadsheetApp.getUi().alert("'업체매핑' 시트가 없습니다.");
    return;
  }

  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(TARGET_SS_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert("대상 파일(외부 저장소)을 열 수 없습니다. ID와 권한을 확인하세요.");
    return;
  }

  // 업체 리스트 수집
  const lastRowMapping = mappingSheet.getLastRow();
  const companyNames = mappingSheet
    .getRange(2, 3, lastRowMapping - 1, 1)
    .getValues().flat()
    .filter(n => n.toString().trim() !== "");

  const masterData = {};
  companyNames.forEach(name => { masterData[name] = []; });

  // [전역 충돌 방지] 다른 .gs 파일과 이름이 겹치는 설정값은 함수 내부 지역변수로 선언
  const FIXED_EXCLUDES    = ["GT","Form","🏠 네비게이션","분석용_RawData","업체매핑","누적현황"];
  const EXCLUDE_KEYWORDS  = ["시뮬레이션","차트","백데이터","실예가","분석용","실적","exc","통계"];
  const shouldExclude = n =>
    FIXED_EXCLUDES.includes(n) || EXCLUDE_KEYWORDS.some(k => n.includes(k));

  // 용역 시트 전수 순회
  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (shouldExclude(sheetName)) return;

    const firstVendor = sheet.getRange("G2").getValue().toString().trim();
    if (firstVendor === "") return;

    // [최적화] B1:B15 한 번에 읽기
    const bCol = sheet.getRange("B1:B15").getValues();
    const commonInfo = [
      bCol[13][0], // B14 개찰일
      bCol[0][0],  // B1  용역명
      bCol[1][0],  // B2  발주처
      bCol[3][0],  // B4  채널
      bCol[9][0],  // B10 분류코드
      bCol[14][0], // B15 업체수
      bCol[5][0],  // B6  유사면적
      bCol[10][0], // B11 기초가격
      bCol[12][0]  // B13 실예가율
    ];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const rows = sheet.getRange(2, 5, lastRow - 1, 13).getValues();
    rows.forEach(row => {
      const nameInRow = row[2] ? row[2].toString().trim() : "";
      if (nameInRow !== "" && masterData.hasOwnProperty(nameInRow)) {
        masterData[nameInRow].push([
          ...commonInfo,
          row[5],   // 판단예가율
          row[4],   // 투찰금액
          row[0],   // 순위
          row[6],   // 가점
          row[7],   // 필요PQ점수
          row[8],   // 추정예가율@100
          row[9],   // 추정예가율@98.2
          row[10],  // 추정예가율@96.4
          row[11],  // 추정예가율@94.6
          row[12]   // 추정예가율@92.8
        ]);
      }
    });
  });

  // 데이터 유무 확인
  const hasData = Object.values(masterData).some(arr => arr.length > 0);
  if (!hasData) {
    SpreadsheetApp.getUi().alert(
      "수집된 데이터가 없습니다.\n업체매핑 C열 단축명과 용역 시트 G열 업체명이 일치하는지 확인하세요."
    );
    return;
  }

  // _temp 시트에 덤프 (타겟SS)
  // 구조: 1열=업체명, 2열=행인덱스(0부터), 3열~=데이터 19컬럼
  let tempSheet = targetSS.getSheetByName(TEMP_SHEET);
  if (tempSheet) targetSS.deleteSheet(tempSheet);
  tempSheet = targetSS.insertSheet(TEMP_SHEET);

  const tempRows = [];
  for (const company in masterData) {
    masterData[company].forEach((dataRow, idx) => {
      tempRows.push([company, idx, ...dataRow]);
    });
  }

  if (tempRows.length > 0) {
    tempSheet.getRange(1, 1, tempRows.length, tempRows[0].length)
      .setValues(tempRows);
  }

  // 업체 목록을 PropertiesService에 저장 (2단계에서 순서 참조용)
  PropertiesService.getScriptProperties()
    .setProperty("companyList", JSON.stringify(companyNames))
    .setProperty("companyOffset", "0");

  console.log(`[1단계 완료] 업체 ${companyNames.length}개, 행 ${tempRows.length}개 덤프 완료. writeCompanySheets 트리거 등록.`);

  // 2단계 트리거 등록
  _registerTrigger("writeCompanySheets", 1);
}


// ══════════════════════════════════════════════════════════════════════════
// 2단계: 업체 시트 BATCH_SIZE개씩 처리 → 자기 체이닝 or 3단계로
// ══════════════════════════════════════════════════════════════════════════
function writeCompanySheets() {
  _deleteTriggers("writeCompanySheets");

  const props = PropertiesService.getScriptProperties();
  const companyNames = JSON.parse(props.getProperty("companyList"));
  const offset       = parseInt(props.getProperty("companyOffset"), 10);

  const targetSS  = SpreadsheetApp.openById(TARGET_SS_ID);
  const tempSheet = targetSS.getSheetByName(TEMP_SHEET);
  if (!tempSheet) {
    console.log("[2단계 오류] _temp 시트가 없습니다.");
    return;
  }

  // _temp 전체 읽기 (한 번만)
  const tempData = tempSheet.getDataRange().getValues();

  // 이번 배치 대상
  const batch = companyNames.slice(offset, offset + BATCH_SIZE);

  batch.forEach(company => {
    // 해당 업체 행만 필터
    const rows = tempData
      .filter(r => r[0] === company)
      .map(r => r.slice(2)); // 업체명·인덱스 제거, 데이터 19컬럼만

    if (rows.length === 0) return;

    const safeName = company.toString().substring(0, 30).replace(/[\[\]\?\*\/\\\:]/g, "");
    let sheet = targetSS.getSheetByName(safeName);
    if (!sheet) {
      sheet = targetSS.insertSheet(safeName);
    } else {
      sheet.clear();
    }

    // 헤더 + 데이터 한 번에 쓰기
    const allRows = [HEADERS, ...rows];
    sheet.getRange(1, 1, allRows.length, HEADERS.length).setValues(allRows);

    // 날짜순 정렬
    if (rows.length > 1) {
      sheet.getRange(2, 1, rows.length, HEADERS.length).sort({ column: 1, ascending: true });
    }

    // 서식
    const dataRows = rows.length;
    sheet.getRange(1, 1, dataRows + 1, HEADERS.length)
      .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(1, 1, 1, HEADERS.length).setBackground("#f3f3f3").setFontWeight("bold");
    sheet.getRange(2, 1, dataRows, HEADERS.length).setHorizontalAlignment("center");
    sheet.getRange(2, 1, dataRows, 3).setHorizontalAlignment("left");
    sheet.getRange(2, 6, dataRows, 14).setHorizontalAlignment("right");
    sheet.getRange(2, 1, dataRows, 1).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(2, 7, dataRows, 2).setNumberFormat("#,##0");
    sheet.getRange(2, 9, dataRows, 2).setNumberFormat("0.0000%");
    sheet.getRange(2, 11, dataRows, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 13, dataRows, 2).setNumberFormat("0.00");
    sheet.getRange(2, 15, dataRows, 5).setNumberFormat("0.000%");

    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 60);
    sheet.setColumnWidth(6, 60);
    sheet.setColumnWidth(7, 60);
    sheet.setColumnWidth(12, 60);
    sheet.setColumnWidth(13, 60);
    sheet.setColumnWidth(14, 130);
    sheet.setColumnWidth(15, 130);
    sheet.setColumnWidth(16, 130);
    sheet.setColumnWidth(17, 130);
    sheet.setColumnWidth(18, 130);
    sheet.setColumnWidth(19, 130);
  });

  const nextOffset = offset + BATCH_SIZE;
  console.log(`[2단계] 업체 ${offset}~${Math.min(nextOffset, companyNames.length) - 1} 처리 완료.`);

  if (nextOffset < companyNames.length) {
    // 아직 남은 업체 있음 → 자기 자신을 다시 트리거
    props.setProperty("companyOffset", String(nextOffset));
    _registerTrigger("writeCompanySheets", 1);
  } else {
    // 모든 업체 완료 → 3단계 트리거
    console.log("[2단계 완료] 모든 업체 처리. writeDotPlotSheet 트리거 등록.");
    _registerTrigger("writeDotPlotSheet", 1);
  }
}


// ══════════════════════════════════════════════════════════════════════════
// 3단계: 도트플롯 시트 생성 + 시트 정렬 + _temp 삭제 + 완료로그
// ══════════════════════════════════════════════════════════════════════════
function writeDotPlotSheet() {
  _deleteTriggers("writeDotPlotSheet");

  const targetSS  = SpreadsheetApp.openById(TARGET_SS_ID);
  const tempSheet = targetSS.getSheetByName(TEMP_SHEET);
  if (!tempSheet) {
    console.log("[3단계 오류] _temp 시트가 없습니다.");
    return;
  }

  const props        = PropertiesService.getScriptProperties();
  const companyNames = JSON.parse(props.getProperty("companyList"));
  const tempData     = tempSheet.getDataRange().getValues();

  // 도트플롯 데이터 구성
  // 판단예가율 = 데이터 열의 인덱스 9 (commonInfo 9개 기준, _temp에서는 col 2+9=11)
  const dotTable         = [];
  let   maxObservations  = 0;

  companyNames.forEach(company => {
    const rates = tempData
      .filter(r => r[0] === company)
      .map(r => r[11]); // col 0=업체명, col 1=idx, col 2~20=데이터, col 11=판단예가율(2+9)
    if (rates.length === 0) return;
    maxObservations = Math.max(maxObservations, rates.length);
    dotTable.push([company, ...rates]);
  });

  if (dotTable.length > 0) {
    let dotSheet = targetSS.getSheetByName("도트플롯_로데이터");
    if (!dotSheet) {
      dotSheet = targetSS.insertSheet("도트플롯_로데이터");
    } else {
      dotSheet.clear();
    }

    const dotHeader = ["업체명", ...Array.from({ length: maxObservations }, (_, i) => `투찰${i + 1}`)];
    const finalTable = dotTable.map(row => {
      const padding = new Array(maxObservations + 1 - row.length).fill("");
      return row.concat(padding);
    });

    dotSheet.getRange(1, 1, 1, dotHeader.length)
      .setValues([dotHeader]).setBackground("#f3f3f3").setFontWeight("bold");
    dotSheet.getRange(2, 1, finalTable.length, dotHeader.length)
      .setValues(finalTable).setNumberFormat("0.0000%");
    dotSheet.setFrozenColumns(1);
  }

  // 시트 정렬
  sortTargetSheets(targetSS);

  // _temp 삭제
  targetSS.deleteSheet(tempSheet);

  // PropertiesService 정리
  props.deleteProperty("companyList");
  props.deleteProperty("companyOffset");

  // 완료 로그 시트
  let logSheet = targetSS.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = targetSS.insertSheet(LOG_SHEET);
    logSheet.getRange(1, 1, 1, 3).setValues([["완료시각", "업체수", "비고"]]);
  }
  logSheet.appendRow([new Date(), companyNames.length, "정상 완료"]);

  console.log(`[3단계 완료] 도트플롯 생성, _temp 삭제, 완료로그 기록.`);
}


// ══════════════════════════════════════════════════════════════════════════
// 시트 정렬 (기존과 동일)
// ══════════════════════════════════════════════════════════════════════════
function sortTargetSheets(spreadsheet) {
  const FIXED_SHEETS   = ["도트플롯_로데이터", "전체투찰차트"];
  const existingFixed  = FIXED_SHEETS.filter(name => spreadsheet.getSheetByName(name) !== null);

  const sheetNames = spreadsheet.getSheets()
    .map(s => s.getName())
    .filter(name => !FIXED_SHEETS.includes(name) && name !== TEMP_SHEET && name !== LOG_SHEET);

  sheetNames.sort();

  sheetNames.forEach((name, index) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(existingFixed.length + index + 1);
  });

  existingFixed.forEach((fixedName, i) => {
    const fixedSheet = spreadsheet.getSheetByName(fixedName);
    if (!fixedSheet) return;
    spreadsheet.setActiveSheet(fixedSheet);
    spreadsheet.moveActiveSheet(i + 1);
  });
}