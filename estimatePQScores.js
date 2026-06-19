// =====================================================================
// 사전 PQ 예측 자동화 스크립트 v3
//
// [참조 시트] (같은 파일 내)
//   경쟁업체실적5000이상  : 정확한 실적 DB (발주면적 5,000㎡ 이상)
//   경쟁업체실적추정치    : 임계점 DB (투찰 패턴 기반 ㎡ 기준 추정치)
//
// [대상 시트] 용역별 개찰결과분석 시트
//   B6  : 유사 면적 (㎡)
//   B19 : 공고일
//   G열 : 업체명 (단축명)
//   L열 : 추정PQ점수 (구 O열) ← 여기에 결과 기록
// =====================================================================

// ── 시트명 설정 ─────────────────────────────────────────────────────────
const SHEET_실적    = "경쟁업체실적5000이상";
const SHEET_임계점  = "경쟁업체실적추정치";

// 경쟁업체실적5000이상 열 (0-indexed)
// A:용역명 B:공고일 C:개찰일 D:용역기간 E:예상준공일 F:발주면적 G:낙찰업체 H:비고
const IDX_준공일   = 4;  // E열
const IDX_발주면적 = 5;  // F열
const IDX_낙찰업체 = 6;  // G열

// 경쟁업체실적추정치 열 (0-indexed)
// A:업체명 B:첫이탈 C:첫이상 D:이탈고착 E:이상고착
const IDX_이탈고착 = 3;  // D열
const IDX_이상고착 = 4;  // E열

const VALIDITY_YEARS = 5;
// ──────────────────────────────────────────────────────────────────────


/**
 * 메인 함수 — 현재 활성 시트의 O열(추정PQ점수)을 자동 계산
 * 버튼 또는 onOpen 커스텀 메뉴에서 호출
 */
function estimatePQScores() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  // ── ① 기준값 읽기 ────────────────────────────────────────────────
  const 유사면적Raw = activeSheet.getRange("B6").getValue();
  const 공고일Raw   = activeSheet.getRange("B19").getValue();

  if (!유사면적Raw || !공고일Raw) {
    ui.alert("오류: B6(유사면적) 또는 B19(공고일)이 비어 있습니다.");
    return;
  }

  const 유사면적 = parseFloat(유사면적Raw);
  const 공고일   = new Date(공고일Raw);
  const 유효기간시작 = new Date(공고일);
  유효기간시작.setFullYear(유효기간시작.getFullYear() - VALIDITY_YEARS);

  // ── ② 참조 시트 가져오기 ─────────────────────────────────────────
  const sheet실적   = ss.getSheetByName(SHEET_실적);
  const sheet임계점 = ss.getSheetByName(SHEET_임계점);

  if (!sheet실적) {
    ui.alert(`'${SHEET_실적}' 시트를 찾을 수 없습니다.`); return;
  }
  if (!sheet임계점) {
    ui.alert(`'${SHEET_임계점}' 시트를 찾을 수 없습니다.`); return;
  }

  // ── ③ 데이터 한 번에 로드 ────────────────────────────────────────
  const data실적   = sheet실적.getDataRange().getValues();
  const data임계점 = sheet임계점.getDataRange().getValues();

  // ── ④ 업체별 O열 계산 ────────────────────────────────────────────
  const lastRow = activeSheet.getLastRow();
  let updated = 0;

  for (let row = 2; row <= lastRow; row++) {
    const 업체명 = String(activeSheet.getRange(row, 7).getValue()).trim(); // G열
    if (!업체명) continue;

    let 건수;

    if (유사면적 >= 5000) {
      // ── Case A: 5,000㎡ 이상 ─────────────────────────────────────
      // 실적 DB에 정확한 데이터가 있으므로 직접 조회
      건수 = countFromDB(data실적, 업체명, 유사면적, 유효기간시작, 공고일);

    } else {
      // ── Case B: 5,000㎡ 미만 ─────────────────────────────────────
      // B-1) 실적 DB에서 ≥5,000㎡ 실적 카운트 (유사면적 < 5000이므로 DB의 모든 레코드가 조건 충족)
      const dbCount = countFromDB(data실적, 업체명, 5000, 유효기간시작, 공고일);

      // B-2) 임계점 DB로 해당 유사면적 구간 패턴 추정
      const estimate = estimateFromThreshold(data임계점, 업체명, 유사면적);

      // B-3) DB 하한 보장 + 추정값 중 큰 쪽 채택 (최대 5건)
      건수 = Math.min(5, Math.max(dbCount, estimate));
    }

    const pqScore = calculatePQ(건수);
    activeSheet.getRange(row, 12).setValue(pqScore); // L열(구 O열)
    updated++;
  }

  ui.alert(
    `완료: ${updated}개 업체 추정PQ점수(L열, 구 O열) 업데이트\n` +
    `유사면적 기준: ${유사면적.toLocaleString()}㎡`
  );
}


// ── 실적 DB 조회: 유효 건수 카운트 ───────────────────────────────────────
/**
 * @param {Array[][]} data      경쟁업체실적5000이상 전체 데이터
 * @param {string}    업체명    조회할 단축명
 * @param {number}    최소면적  발주면적 최소 기준(㎡)
 * @param {Date}      시작일    유효기간 시작 (공고일 - 5년)
 * @param {Date}      공고일    유효기간 종료 (이 날짜 이전 준공만 유효)
 * @returns {number}  유효 건수 (0~5)
 */
function countFromDB(data, 업체명, 최소면적, 시작일, 공고일) {
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // 업체명 매칭
    if (String(row[IDX_낙찰업체]).trim() !== 업체명) continue;

    // 준공일 유효기간 체크
    const 준공일Raw = row[IDX_준공일];
    if (!준공일Raw) continue;
    const 준공일 = new Date(준공일Raw);
    if (준공일 < 시작일 || 준공일 > 공고일) continue;

    // 발주면적 기준 충족 여부
    const 발주면적 = parseFloat(String(row[IDX_발주면적]).replace(/,/g, ""));
    if (isNaN(발주면적) || 발주면적 < 최소면적) continue;

    if (++count >= 5) break; // 5건이면 만점, 더 볼 필요 없음
  }
  return count;
}


// ── 임계점 DB 조회: 투찰 패턴 기반 건수 추정 ─────────────────────────────
/**
 * 유사면적(㎡)을 업체의 이탈/이상 고착 면적과 직접 비교하여 추정 건수 반환
 *
 * 판단 로직:
 *   유사면적 ≥ 이상 고착 면적  →  이상투찰 구간  →  1건
 *   유사면적 ≥ 이탈 고착 면적  →  경계투찰 구간  →  3건
 *   그 외                      →  정상투찰 구간  →  5건
 *   DB에 없는 업체             →  보수적 기본값  →  3건
 *
 * @param {Array[][]} data    경쟁업체실적추정치 전체 데이터
 * @param {string}    업체명  조회할 단축명
 * @param {number}    유사면적 이번 입찰의 유사면적 기준(㎡)
 * @returns {number}  추정 건수 (1 / 3 / 5)
 */
function estimateFromThreshold(data, 업체명, 유사면적) {
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() !== 업체명) continue;

    const 이탈고착 = row[IDX_이탈고착];
    const 이상고착 = row[IDX_이상고착];

    if (이상고착 !== "" && 이상고착 != null && 유사면적 >= Number(이상고착)) return 1;
    if (이탈고착 !== "" && 이탈고착 != null && 유사면적 >= Number(이탈고착)) return 3;
    return 5;
  }
  return 3; // DB에 없는 업체: 보수적 처리
}


// ── PQ 점수 계산 (국방부 기준) ───────────────────────────────────────────
/**
 * 만점 5건, 1건 부족 시 -1.8점
 * 0건 = 1건 동일 취급 (최대 감점 -7.2점)
 *
 * 건수  →  점수
 *  5+   →  100.0
 *  4    →   98.2
 *  3    →   96.4
 *  2    →   94.6
 *  0~1  →   92.8
 */
function calculatePQ(건수) {
  const 부족 = Math.max(0, Math.min(4, 5 - Math.max(0, 건수)));
  return Math.round((100 - 부족 * 1.8) * 10) / 10;
}
