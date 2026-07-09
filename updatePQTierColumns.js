// ─────────────────────────────────────────────────────────────────────────────
// [완전 수정본] 외부 시트(PQ점수별 낙찰하한율) 연동 및 정밀 추정예가율 역산
// ─────────────────────────────────────────────────────────────────────────────

const PQ_TIERS = ["100.00", "98.20", "96.40", "94.60", "92.80"];
const COLOR_IMPOSSIBLE = "#CCCCCC";
const COLOR_DEFAULT    = null;

// 제공해주신 'PQ점수별 낙찰하한율' 스프레드시트의 ID
const PQ_TABLE_SPREADSHEET_ID = "1qWRw6Ttl3GkX1YNBhcMqV1yYfadOv3ukEKDF9LjsOio";
let pqLimitTableCache = null;

/**
 * 1. 지정된 링크의 'PQ점수별 낙찰하한율' 시트 데이터를 읽어와 메모리에 캐싱합니다.
 */
function getPqLimitRateFromTable_(targetPq, codeKey) {
  if (!pqLimitTableCache) {
    pqLimitTableCache = {};
    const ss = SpreadsheetApp.openById(PQ_TABLE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("낙찰하한율정리");
    if (!sheet) {
      throw new Error("링크된 파일에서 '낙찰하한율정리' 시트를 찾을 수 없습니다.");
    }
    const data = sheet.getDataRange().getValues();

    // 데이터 파싱 (스크린샷 기준: A열(0)=PQ점수, B열(1)=코드2, C열(2)=코드5, D열(3)=코드10, E열(4)=코드30)
    for (let r = 0; r < data.length; r++) {
      const pqRaw = data[r][0];
      if (pqRaw === "" || isNaN(pqRaw)) continue;
      
      const pq = parseFloat(pqRaw);
      const key = pq.toFixed(2); // "98.60" 형태로 키 생성
      
      pqLimitTableCache[key] = {
        2: parseFloat(data[r][1]) / 100,  // 표의 85.84500 등을 0.85845로 변환
        5: parseFloat(data[r][2]) / 100,
        10: parseFloat(data[r][3]) / 100,
        30: parseFloat(data[r][4]) / 100
      };
    }
  }

  // PQ점수는 100점을 초과할 수 없으므로 상한선 적용 (적격심사 일반 룰)
  if (targetPq > 100) targetPq = 100.00;

  const lookupKey = targetPq.toFixed(2);
  const rowData = pqLimitTableCache[lookupKey];

  if (!rowData) {
    return null; // 매핑 테이블에 해당 점수가 없는 경우
  }

  return rowData[codeKey];
}

/**
 * 2. 메인 업데이트 함수
 */
function updatePQTierColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const classCodeRaw = sheet.getRange("B10").getValue();
  const basePrice    = parseFloat(sheet.getRange("B11").getValue()); // 기초가격
  const classCode    = parseInt(classCodeRaw, 10);

  if (!basePrice || basePrice <= 0) {
    SpreadsheetApp.getUi().alert("B11(기초가격)이 올바르지 않습니다.");
    return;
  }

  const codeKey = (classCode === 1) ? 2 : classCode;
  const lastRow = getRealDataLastRow_(sheet);
  if (lastRow < 2) return;
  const dataRows = lastRow - 1;

  // 기존 데이터 범위 읽기 (투찰률, 투찰금액, 가점)
  const hValues = sheet.getRange(2, 8, dataRows, 1).getValues(); 
  const iValues = sheet.getRange(2, 9, dataRows, 1).getValues(); 
  const kValues = sheet.getRange(2, 11, dataRows, 1).getValues(); 

  // divisor 가져오기 (문서 기준 5->0.5, 10/30->0.7, 1/2->0.3)
  const divisor = getPqBonusDivisor_(classCode);

  const lResults = [];
  const mResults = [];
  const bgColors = [];

  for (let i = 0; i < dataRows; i++) {
    const hPct   = (hValues[i][0] < 1 && hValues[i][0] > 0) ? hValues[i][0] * 100 : hValues[i][0];
    const bidAmt = parseFloat(iValues[i][0]) || 0;
    const bonus  = parseFloat(kValues[i][0]) || 0;

    // L열: 필요 PQ 하한 점수 계산
    const pqMin = (hPct > 0 && bidAmt > 0) ? calcPQScore_(hPct, classCode, bonus) : "";
    lResults.push([pqMin]);

    const tierRates  = [];
    const tierColors = [];
    const pMin       = parseFloat(pqMin);

    for (const tier of PQ_TIERS) {
      const basePq = parseFloat(tier); // 가정하는 PQ원점 (예: 98.20)
      
      // [STEP 1] 환산 점수 = 원점 + (가점 / divisor)
      let convertedPq = basePq + (bonus / divisor);
      
      // [STEP 2] 소수점 이하 셋째자리에서 반올림하여 둘째자리까지 산출
      convertedPq = Math.round(convertedPq * 100) / 100;
      
      // [STEP 3] 외부 링크의 'PQ점수별 낙찰하한율' 표에서 정확한 하한율 조회
      const limitRate = getPqLimitRateFromTable_(convertedPq, codeKey);
      
      // [STEP 4] 역산 산식: 추정예가율 = 투찰금액 / (기초가격 * 낙찰하한율)
      let estRate = "";
      if (bidAmt > 0 && basePrice > 0 && limitRate && limitRate > 0) {
        estRate = bidAmt / (basePrice * limitRate);
      }
      
      tierRates.push(estRate);

      // 조건부 서식: 가정하는 원점(basePq)이 적격통과를 위한 하한점(pMin)에 미달하면 회색 표시
      const isImpossible = (!isNaN(pMin) && basePq < pMin);
      tierColors.push(isImpossible ? COLOR_IMPOSSIBLE : COLOR_DEFAULT);
    }
    
    mResults.push(tierRates);
    bgColors.push(tierColors);
  }

  // 3. 계산 결과를 시트에 반영
  sheet.getRange(2, 12, dataRows, 1).setValues(lResults).setNumberFormat("0.00");
  
  const tierRange = sheet.getRange(2, 13, dataRows, 5);
  tierRange.setValues(mResults).setNumberFormat("0.000%");
  tierRange.setBackgrounds(bgColors);

  ss.toast("외부 낙찰하한율 매핑 표 기반 추정예가율 계산 완료", "✅ 완료");
}