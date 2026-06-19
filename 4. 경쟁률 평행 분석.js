/**
 * 2. 5가지 조건 병렬 분석 실행 (오류 수정 버전)
 */
function generateParallelDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("분석용_RawData");
  let dashSheet = ss.getSheetByName("통계_대시보드");
  
  if (!dashSheet) dashSheet = ss.insertSheet("통계_대시보드");
  dashSheet.clear();

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 2) {
    SpreadsheetApp.getUi().alert("분석할 데이터가 로데이터 시트에 없습니다.");
    return;
  }
  
// 1. RawData에서 헤더를 제외한 모든 데이터 가져오기
  const allRows = rawData.slice(1);
  
  // 2. 특이 용역(건축물등재 등)을 분석에서 즉시 제외 (키워드 필터링)
  const filteredForTrend = allRows.filter(row => {
    const serviceName = row[1].toString();
    // 제외하고 싶은 키워드들을 여기에 추가하세요.
    const excludeKeywords = ["exc"]; 
    
    // 키워드가 하나라도 포함되어 있으면 false를 반환하여 제거
    return !excludeKeywords.some(keyword => serviceName.includes(keyword));
  });

  const now = new Date();
  
  // 이후 로직에서 'allRows' 대신 'filteredForTrend'를 사용하여 분석 진행

  // 분석 조건 정의
  const conditions = [
    { title: "전체 기간 분석", filter: function(rows) { return rows; } },
    { title: "최근 1개월 분석", filter: function(rows) {
        const limit = new Date(); limit.setMonth(limit.getMonth() - 1);
        return rows.filter(function(r) { return new Date(r[0]) >= limit; });
      }
    },
    { title: "최근 2주 분석", filter: function(rows) {
        const limit = new Date(); limit.setDate(limit.getDate() - 14);
        return rows.filter(function(r) { return new Date(r[0]) >= limit; });
      }
    },
    { title: "최신 10건 분석", filter: function(rows) {
        const uniqueBids = [...new Set(rows.map(function(r) { return r[1]; }))].slice(0, 10);
        return rows.filter(function(r) { return uniqueBids.includes(r[1]); });
      }
    },
    { title: "최신 5건 분석", filter: function(rows) {
        const uniqueBids = [...new Set(rows.map(function(r) { return r[1]; }))].slice(0, 5);
        return rows.filter(function(r) { return uniqueBids.includes(r[1]); });
      }
    }
  ];

  // 병렬 배치 실행
  conditions.forEach((cond, index) => {
    const startCol = (index * 5) + 1; // A, F, K, P, U열
    
    // 안전한 필터링 실행
    let filteredRows = [];
    try {
      filteredRows = cond.filter(allRows);
    } catch (e) {
      filteredRows = [];
    }

    const resultTable = runAnalysis(filteredRows, cond.title);
    
    // 데이터 기록
    dashSheet.getRange(1, startCol, resultTable.length, 4).setValues(resultTable);
    
    // 서식 및 정렬
    dashSheet.setColumnWidth(startCol, 120); 
    const headerRange = dashSheet.getRange(3, startCol, 1, 4);
    headerRange.setBackground("#4A86E8").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    
    if (resultTable.length > 3) {
      const dataRange = dashSheet.getRange(4, startCol, resultTable.length - 3, 4);
      dataRange.sort({column: startCol + 3, ascending: false});
      dataRange.setHorizontalAlignment("center");
    }
    dashSheet.getRange(1, startCol).setFontSize(12).setFontWeight("bold");
  });

  dashSheet.setFrozenRows(3);
  SpreadsheetApp.getUi().alert("병렬 분석이 완료되었습니다.");
}

/**
 * 데이터가 없을 때를 대비한 안전한 분석 로직
 */
function runAnalysis(rows, title) {
  // 안전 장치: rows가 없거나 비어있는 경우 처리
  if (!rows || rows.length === 0) {
    return [
      [title, "", "", ""],
      ["대상 용역 수:", "0 건", "", ""],
      ["데이터가 없습니다.", "", "", ""]
    ];
  }

  const summary = {};
  const uniqueBids = new Set();
  
  rows.forEach(row => {
    const name = row[1];
    const interval = row[4];
    const intensity = parseFloat(row[5]);
    const isWin = row[6];

    uniqueBids.add(name);
    if (!summary[interval]) summary[interval] = { win: 0, totalInt: 0, count: 0 };
    if (isWin === "YES") summary[interval].win++;
    summary[interval].totalInt += (isNaN(intensity) ? 0 : intensity);
    summary[interval].count++;
  });

  const output = [
    [title, "", "", ""],
    ["대상 용역 수:", uniqueBids.size + " 건", "", ""],
    ["판단예가율 구간", "낙찰 빈도", "평균 경쟁", "가치 지수"]
  ];

  const sortedIntervals = Object.keys(summary).sort((a, b) => parseFloat(b) - parseFloat(a));
  sortedIntervals.forEach(key => {
    const item = summary[key];
    const avgInt = item.totalInt / item.count;
    const score = item.win > 0 ? (item.win / (avgInt + 0.1)) : 0;
    output.push([key, item.win, avgInt.toFixed(2), score.toFixed(2)]);
  });

  return output;
}

/**
 * 공통 분석 로직 (데이터셋 -> 2D 배열)
 */
function runAnalysis(rows, title) {
  const summary = {};
  const uniqueBids = new Set();
  
  rows.forEach(row => {
    const [date, name, code, rate, interval, intensity, isWin] = row;
    uniqueBids.add(name);
    if (!summary[interval]) summary[interval] = { win: 0, totalInt: 0, count: 0 };
    if (isWin === "YES") summary[interval].win++;
    summary[interval].totalInt += parseFloat(intensity);
    summary[interval].count++;
  });

  const output = [
    [title, "", "", ""],
    ["대상 용역 수:", uniqueBids.size + " 건", "", ""],
    ["판단예가율 구간", "낙찰 빈도", "평균 경쟁", "가치 지수"]
  ];

  const sortedIntervals = Object.keys(summary).sort((a, b) => parseFloat(b) - parseFloat(a));
  sortedIntervals.forEach(key => {
    const item = summary[key];
    const avgInt = item.totalInt / item.count;
    const score = item.win > 0 ? (item.win / (avgInt + 0.1)) : 0;
    output.push([key, item.win, avgInt.toFixed(2), score.toFixed(2)]);
  });

  return output;
}

// 기존 보조 함수 유지
function checkRateInInterval(rate, intervalText) {
  const numbers = intervalText.match(/[\d.]+/g);
  if (numbers && numbers.length === 2) {
    const upper = parseFloat(numbers[0]) / 100, lower = parseFloat(numbers[1]) / 100;
    const rateVal = typeof rate === 'string' ? parseFloat(rate.replace('%','')) / 100 : rate;
    return rateVal <= upper && rateVal > lower;
  }
  return false;
}