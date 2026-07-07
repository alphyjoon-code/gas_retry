/**
 * 인터랙티브 시뮬레이션 모달 실행
 */
function showInteractiveSimulation() {
  const html = HtmlService.createHtmlOutputFromFile("InteractiveSimulation")
    .setWidth(1100)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, "📊 인터랙티브 수주 시뮬레이션");
}

/**
 * HTML에서 google.script.run으로 호출
 * 모든 용역 시트를 스캔해서 {개찰일, 구간, 분류, 금액} 반환
 * JS 쪽에서 기간 필터 + 집계를 담당하므로 GAS는 최초 1회만 호출됨
 */
function getProjectsForInteractive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const EXCLUDE_KEYWORDS = [
    "시뮬레이션", "차트", "백데이터", "대시보드", "업체매핑",
    "🏠", "RawData", "통계", "현황", "Form", "GT", "exc", "실적"
  ];

  // A열 예가율 기준선 (전기간 시뮬레이션 시트에서 가져옴)
  const masterSheet = ss.getSheetByName("전기간 예가율 수주 시뮬레이션");
  if (!masterSheet) return { rateAxis: [], projects: [] };

  const lastMasterRow = masterSheet.getLastRow();
  const rateAxis = lastMasterRow >= 2
    ? masterSheet.getRange(2, 1, lastMasterRow - 1, 1).getValues().flat()
    : [];

  // 용역 시트 스캔
  const projects = [];

  ss.getSheets().forEach(s => {
    const sName = s.getName();
    if (EXCLUDE_KEYWORDS.some(kw => sName.includes(kw))) return;

    try {
            // ★ B8→B10, B9→B11, B11→B13, B12→B14 으로 수정
            const ranges     = s.getRangeList(["B10", "B11", "B13", "B14", "J2"]).getRanges();
      const code       = parseInt(ranges[0].getValue())  || 0;
      const price      = Number(ranges[1].getValue())    || 0;
      const actualRate = parseFloat(ranges[2].getValue());
      const dateVal    = ranges[3].getValue();
      const winnerRate = parseFloat(ranges[4].getValue());

      if (isNaN(actualRate) || isNaN(winnerRate)) return;
      if (!dateVal) return;

      const dateStr = dateVal instanceof Date
        ? Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd")
        : String(dateVal).trim();

      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return;

      projects.push({
        date:  dateStr,    // "yyyy-MM-dd"
        min:   actualRate, // 실예가율 (하한)
        max:   winnerRate, // 1순위 판단예가율 (상한, 미만)
        price: price,      // 기초가격
        code:  code,       // 분류코드
        name:  sName,
      });
    } catch (e) {
      console.log(sName + " 스킵: " + e.message);
    }
  });

  // 날짜 오름차순 정렬 (슬라이더 min/max 계산용)
  projects.sort((a, b) => a.date.localeCompare(b.date));

  return { rateAxis, projects };
}
