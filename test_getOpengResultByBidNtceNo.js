/**
 * [승인받은 오퍼레이션으로 재작성] 개찰결과 용역 목록 조회
 * ------------------------------------------------------------
 * 사용 오퍼레이션: getOpengResultListInfoServcPPSSrch
 *   (나라장터 검색조건에 의한 개찰결과 용역 목록 조회 - 승인된 5번 항목)
 *
 * ⚠️ 이전 시도(getOpengResultListInfoServc)는 승인받지 않은 오퍼레이션이라
 *    403 Forbidden이 발생했음. 이 버전은 마이페이지에서 승인 확인된
 *    getOpengResultListInfoServcPPSSrch로 교체함.
 *
 * ⚠️ inqryDiv 코드값이 다른 오퍼레이션과 다름에 주의:
 *    이 오퍼레이션: 1=공고일시, 2=개찰일시, 3=입찰공고번호
 *    (다른 오퍼레이션은 1=등록일시, 2=공고일시, 3=개찰일시, 4=입찰공고번호였음)
 *
 * ⚠️ 검증 필요: opengCorpInfo 필드가 참가업체 전원(20~60개사)의
 *    순위별 정보를 다 담고 있는지, 아니면 낙찰예정자 요약만 주는지는
 *    실제 호출 결과로 확인해야 함. 문서 설명상 "다수 낙찰자의 경우
 *    낙찰예정자 다수 + 1위 정보만" 표시된다고 되어 있어, 이 필드가
 *    전체 순위표를 안 줄 가능성도 있음 -> 아래 테스트로 확인.
 *
 * 사전 준비:
 *   GAS 편집기 > 프로젝트 설정 > 스크립트 속성에 아래 키 등록 완료 상태여야 함:
 *     이름: G2B_SERVICE_KEY
 *     값  : 공공데이터포털에서 발급받은 인증키
 */

const G2B_BASE_URL = 'https://apis.data.go.kr/1230000/as/ScsbidInfoService';

/**
 * 서비스키를 Script Properties에서 가져옴
 */
function getServiceKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('G2B_SERVICE_KEY');
  if (!key) {
    throw new Error('G2B_SERVICE_KEY가 설정되지 않았습니다. 스크립트 속성에 먼저 등록하세요.');
  }
  return key;
}

/**
 * 특정 입찰공고번호의 개찰결과를 조회한다.
 * @param {string} bidNtceNo - 입찰공고번호 (예: R26BK01590022)
 * @return {Object} API 응답의 response.body 전체 (items, totalCount 등)
 */
function getOpengResultByBidNtceNo(bidNtceNo) {
  const serviceKey = getServiceKey_();

  const queryParts = [
    'ServiceKey=' + serviceKey, // 포털에서 받은 Encoding된 키를 그대로 사용 (추가 인코딩 안 함)
    'type=json',
    'inqryDiv=3', // 이 오퍼레이션은 3 = 입찰공고번호로 조회
    'bidNtceNo=' + encodeURIComponent(bidNtceNo),
    'pageNo=1',
    'numOfRows=100', // 참가업체 20~60개사까지 커버하도록 여유있게 100
  ];

  const url = G2B_BASE_URL + '/getOpengResultListInfoServcPPSSrch?' + queryParts.join('&');

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const text = response.getContentText();

  if (statusCode !== 200) {
    throw new Error(
      'API 호출 실패 (HTTP ' + statusCode + '). 서버 응답 본문: ' + text.substring(0, 1000) +
      ' / 요청 URL(키 일부 마스킹): ' + url.replace(serviceKey, serviceKey.substring(0, 4) + '...')
    );
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
      (header ? header.resultMsg : '알 수 없음') + ' / 요청 URL: ' + url
    );
  }

  return json.response.body;
}

/**
 * 테스트: 실제 개찰완료 공고번호로 호출해서 구조를 확인한다.
 * 실행 방법: GAS 편집기에서 이 함수 실행 > 실행 로그(Ctrl+Enter)에서 확인
 */
function test_getOpengResultByBidNtceNo() {
  // 스크린샷에서 확인한 실제 개찰완료 공고번호
  const bidNtceNo = 'R26BK01590022'; // 26년 표준설계도 제정 설계용역, 개찰일 2026-06-29

  Logger.log('조회할 입찰공고번호: ' + bidNtceNo);

  try {
    const body = getOpengResultByBidNtceNo(bidNtceNo);

    Logger.log('totalCount(응답에 담긴 건수): ' + body.totalCount);

    // items.item이 배열일 수도, 단일 객체일 수도 있음 (건수가 1건이면 배열이 아닐 수 있음)
    const items = body.items && body.items.item
      ? (Array.isArray(body.items.item) ? body.items.item : [body.items.item])
      : [];

    Logger.log('실제로 받은 item 개수: ' + items.length);

    items.forEach(function (item, idx) {
      Logger.log('--- item[' + idx + '] ---');
      Logger.log('입찰공고명: ' + item.bidNtceNm);
      Logger.log('개찰일시: ' + item.opengDt);
      Logger.log('진행구분: ' + item.progrsDivCdNm);
      Logger.log('참가업체수: ' + item.prtcptCnum);
      Logger.log('개찰업체정보(원본, 길이 ' + (item.opengCorpInfo || '').length + '자): ');
      Logger.log(item.opengCorpInfo);
    });
  } catch (e) {
    Logger.log('오류 발생: ' + e.message);
  }
}