/**
 * 수주 시뮬레이션 차트를 모달로 표시
 * ★ 수정: 템플릿 주입(<?= ?>) 대신 google.script.run으로 데이터 전달
 *         → JSON 이스케이프 문제 원천 차단
 */
function showSimulationChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("전기간 예가율 수주 시뮬레이션");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("'전기간 예가율 수주 시뮬레이션' 시트가 없습니다.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("데이터가 없습니다. 먼저 시뮬레이션을 실행해주세요.");
    return;
  }

  // 모달 HTML만 띄우고, 데이터는 HTML 쪽에서 google.script.run으로 요청
  const html = HtmlService.createHtmlOutputFromFile("SimulationChart")
    .setWidth(1100)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(html, "📊 전기간 수주 시뮬레이션");
}

/**
 * HTML에서 google.script.run으로 호출하는 데이터 반환 함수
 * 반드시 전역 함수로 존재해야 함
 */
function getSimulationData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("전기간 예가율 수주 시뮬레이션");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rawData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  return rawData
    .filter(row => row[0] !== "" && row[0] !== null)
    .map(row => ({
      rate:  row[0],
      count: Number(row[1]),
      names: row[2] ? row[2].toString() : "",
      total: Number(row[3]),
      cat1:  Number(row[4]),
      cat2:  Number(row[5]),
      cat5:  Number(row[6]),
      cat10: Number(row[7]),
      cat30: Number(row[8]),
    }));
}
