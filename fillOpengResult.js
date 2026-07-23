/**
 * [나라장터 API 연동] 개찰결과 원본을 E~I열에 채워넣는 함수
 * ------------------------------------------------------------
 * 역할 범위(중요):
 *   이 함수는 "나라장터 웹화면에서 수동으로 복붙하던 작업"만 대체한다.
 *   즉 E~I열에 원본 데이터를 그대로 꽂아넣는 것까지만 하고,
 *   그 이후의 정렬/그룹분류(1순위·정상·하한미달·제외)/재계산/서식은
 *   전부 기존 updateBiddingResults() 함수가 담당한다.
 *   이 함수는 updateBiddingResults()의 로직을 절대 대체하거나
 *   중복 구현하지 않는다.
 *
 * 시트 구조:
 *   B20  : 입찰공고번호 (bidNtceNo) - 사용자가 미리 입력해두는 셀
 *   E1:I1: 헤더 (순위 / 사업자등록번호 / 업체명 / 투찰률 / 투찰금액)
 *   E2~  : 데이터 시작 행 (원본을 여기서부터 채움)
 *
 * 사용 오퍼레이션: getOpengResultListInfoOpengCompt (개찰결과 개찰완료 목록 조회)
 *   - 45개사 규모 공고에서 전체 데이터를 정상 반환하는 것을 확인함.
 *
 * 매핑 (원본 그대로, 재계산 없음):
 *   E열(순위)     <- opengRank (숫자)
 *   F열(사업자등록번호) <- prcbdrBizno
 *   G열(업체명)    <- prcbdrNm
 *   H열(투찰률)    <- bidprcrt / 100 (예: 86.747 -> 0.86747, 셀서식 0.000%와 맞춤)
 *   I열(투찰금액)  <- bidprcAmt (숫자)
 *
 * 실행 흐름:
 *   1. B20에서 입찰공고번호를 읽는다.
 *   2. 입찰공고번호 정제 (하이픈'-' 분리 후 공고번호 및 차수 추출)
 *   3. API 호출
 *      - 차수 명시 시(-01 등): 해당 차수로 1회 즉시 호출 (속도 최적화)
 *      - 차수 미지정 시: 000 -> 001 -> 002 -> 003 순차 탐색 후 데이터 발견 즉시 종료
 *   4. E2:I(기존 마지막 데이터 행)을 지운다 (이전 공고 잔여 데이터 방지).
 *   5. opengRank 순으로 정렬 후 E2부터 원본 그대로 채운다.
 *   6. updateBiddingResults()를 이어서 호출해 정렬/분류/재계산/서식을 맡긴다.
 *
 * 사전 준비: 스크립트 속성에 G2B_SERVICE_KEY 등록 완료 상태여야 함.
 */

const G2B_BASE_URL = 'https://apis.data.go.kr/1230000/as/ScsbidInfoService';

const OPENG_SHEET_LAYOUT = {
  BID_NTCE_NO_CELL: 'B20',
  DATA_START_ROW: 2,
  COL_RANK: 5,       // E
  COL_BIZNO: 6,      // F
  COL_CORP_NAME: 7,  // G
  COL_BID_RATE: 8,   // H
  COL_BID_AMOUNT: 9, // I
};

function getServiceKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('G2B_SERVICE_KEY');
  if (!key) {
    throw new Error('G2B_SERVICE_KEY가 설정되지 않았습니다. 스크립트 속성에 먼저 등록하세요.');
  }
  return key;
}

/**
 * 개찰결과 개찰완료 목록을 조회한다.
 */
function getOpengCompleteList_(bidNtceNo, bidNtceOrd, bidClsfcNo, rbidNo) {
  const serviceKey = getServiceKey_();

  const queryParts = [
    'ServiceKey=' + serviceKey,
    'type=json',
    'bidNtceNo=' + encodeURIComponent(bidNtceNo),
    'bidNtceOrd=' + encodeURIComponent(bidNtceOrd || '000'),
    'bidClsfcNo=' + encodeURIComponent(bidClsfcNo || '0'),
    'rbidNo=' + encodeURIComponent(rbidNo || '000'),
    'pageNo=1',
    'numOfRows=100',
  ];

  const url = G2B_BASE_URL + '/getOpengResultListInfoOpengCompt?' + queryParts.join('&');

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const text = response.getContentText();

  if (statusCode !== 200) {
    throw new Error('API 호출 실패 (HTTP ' + statusCode + '). 서버 응답 본문: ' + text.substring(0, 1000));
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('JSON 파싱 실패. 응답 원문(최대 1000자): ' + text.substring(0, 1000));
  }

  const header = json.response && json.response.header;
  if (!header || header.resultCode !== '00') {
    throw new Error(
      'API 결과 오류 [' + (header ? header.resultCode : 'N/A') + '] ' +
      (header ? header.resultMsg : '알 수 없음')
    );
  }

  const body = json.response.body;

  let items = [];
  if (Array.isArray(body.items)) {
    items = body.items;
  } else if (body.items && body.items.item) {
    items = Array.isArray(body.items.item) ? body.items.item : [body.items.item];
  }

  return items;
}

/**
 * B20의 입찰공고번호로 개찰결과 원본을 E~I열에 채우고,
 * 이어서 updateBiddingResults()를 호출한다.
 */
function fillOpengResultToSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const layout = OPENG_SHEET_LAYOUT;

  const bidNtceNo = sheet.getRange(layout.BID_NTCE_NO_CELL).getValue();

  if (!bidNtceNo || typeof bidNtceNo !== 'string' || bidNtceNo.trim() === '') {
    SpreadsheetApp.getUi().alert(
      layout.BID_NTCE_NO_CELL + ' 셀에 입찰공고번호를 먼저 입력해주세요. (예: R26BK01590022 또는 20240112345-01)'
    );
    return;
  }

  // B20 셀에서 가져온 공고번호 정제 (예: "20240112345-01" -> 공고번호: "20240112345", 차수: "001")
  let rawBidNo = String(bidNtceNo).trim();
  let cleanBidNo = rawBidNo;
  let bidOrd = '000'; // 기본값

  if (rawBidNo.includes('-')) {
    const parts = rawBidNo.split('-');
    cleanBidNo = parts[0]; // - 앞부분 (공고번호)
    
    // - 뒷부분이 숫자인 경우 3자리 포맷(001, 002 등)으로 맞춤
    if (parts[1]) {
      bidOrd = String(parts[1]).padStart(3, '0');
    }
  }

  let items = [];

  if (rawBidNo.includes('-')) {
    // 1. 차수가 명시되어 있으면(-01 등) 해당 차수로 1번만 바로 조회 (가장 빠름)
    try {
      items = getOpengCompleteList_(cleanBidNo, bidOrd, '0', '000');
    } catch (e) {
      SpreadsheetApp.getUi().alert('개찰결과 조회 실패: ' + e.message);
      return;
    }
  } else {
    // 2. 차수가 없는 경우만 000부터 순차 탐색 (000 -> 001 -> 002 -> 003)
    const candidateOrders = ['000', '001', '002', '003'];

    for (let ord of candidateOrders) {
      try {
        const res = getOpengCompleteList_(cleanBidNo, ord, '0', '000');
        if (res && res.length > 0) {
          items = res;
          break; // 찾으면 바로 종료하여 시간 단축
        }
      } catch (e) {
        // 해당 차수에 없으면 다음 차수 시도
      }
    }
  }

  if (items.length === 0) {
    SpreadsheetApp.getUi().alert(
      '조회 결과가 없습니다. 아직 개찰이 완료되지 않았거나, 입찰공고번호(' + bidNtceNo + ')를 확인해주세요.'
    );
    return;
  }

  // opengRank 기준 오름차순 정렬 (원본 순서를 그대로 유지하기 위함,
  // 실제 재정렬/재분류는 updateBiddingResults()가 다시 수행함)
  items.sort(function (a, b) {
    return Number(a.opengRank) - Number(b.opengRank);
  });

  // 기존 E~I 데이터 영역 지우기 (이전 공고 잔여 데이터 방지)
  const lastRow = sheet.getLastRow();
  const clearRowCount = Math.max(lastRow - layout.DATA_START_ROW + 1, items.length);
  if (clearRowCount > 0) {
    sheet
      .getRange(layout.DATA_START_ROW, layout.COL_RANK, clearRowCount, 5) // E~I 5개열
      .clearContent();
  }

  // 원본 그대로 채우기 (가공/재계산 없음)
  const rows = items.map(function (item) {
    return [
      Number(item.opengRank),        // E: 순위 (원본)
      item.prcbdrBizno,              // F: 사업자등록번호 (원본)
      item.prcbdrNm,                 // G: 업체명 (원본)
      Number(item.bidprcrt) / 100,   // H: 투찰률 (원본, % 서식에 맞춰 소수로 변환만)
      Number(item.bidprcAmt),        // I: 투찰금액 (원본)
    ];
  });

  sheet.getRange(layout.DATA_START_ROW, layout.COL_RANK, rows.length, 5).setValues(rows);

  // 원본 입력이 끝났으니, 정렬/분류/재계산/서식은 기존 함수에 위임
  updateBiddingResults();
}