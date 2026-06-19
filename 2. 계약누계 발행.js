/**
 * [기능 4] PQ 입찰 계약누계 리포트 갱신
 */
function updateContractorDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gtSheet = ss.getSheetByName("GT");
  if (!gtSheet) {
    SpreadsheetApp.getUi().alert("'GT' 시트를 찾을 수 없습니다.");
    return;
  }

  const targetFileId = "1mdXYtzi2yWBViAqFvVyzmpaT5EqGF5HWCFmC4KD49Y8";
  const targetSS = SpreadsheetApp.openById(targetFileId);

  let dash = targetSS.getSheetByName("계약확정누계") || targetSS.insertSheet("계약확정누계");
  dash.clear();
  targetSS.setActiveSheet(dash);
  targetSS.moveActiveSheet(1);

  // ── 제외 시트 설정 ──────────────────────────────────────────
  const EXCLUDE_KEYWORDS   = ["시뮬레이션", "차트", "백데이터", "대시보드", "실예가", "(exc)", "분석용", "업체매핑"];
  const EXCLUDE_EXACT_NAMES = [
    "GT", "Form", "🏠 네비게이션", "분석용_RawData", "누적현황", "통합",
    "PQ백데이터", "PQ백데이터(분석용)", "통계_대시보드",
    "최근 30일 실예가", "전체 예가율 차트", "전기간 예가율 차트", "최근 30일 백데이터", "실적"
  ];
  const DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
  // ────────────────────────────────────────────────────────────

  // ── GT 시트에서 nameMap 로드 ──────────────────────────────────
  const nameMap = {};
  const mappingRows = gtSheet.getRange("F2:G" + gtSheet.getLastRow()).getValues();
  mappingRows.forEach(r => {
    if (r[0] && r[1]) nameMap[r[0].toString().trim()] = r[1].toString().trim();
  });
  // ────────────────────────────────────────────────────────────

  const confirmedData    = [];
  const pendingData      = [];
  const disqualifiedData = [];

  ss.getSheets().forEach(sheet => {
    const sName = sheet.getName();

    // 제외 시트 필터링
    if (
      EXCLUDE_EXACT_NAMES.includes(sName) ||
      EXCLUDE_KEYWORDS.some(kw => sName.includes(kw)) ||
      DATE_PATTERN.test(sName)
    ) return;

    // ── 메타 정보 일괄 읽기 (B6:B18) ─────────────────────────────
    // B6=유사면적, B10=분류코드, B11=기초가격, B12=예정가격,
    // B16=1순위, B17=낙찰업체, B18=연번
    const metaValues  = sheet.getRange("B6:B18").getValues();
    const similarArea = metaValues[0][0];                            // B6  : index 0
    const code        = (metaValues[4][0] || "미분류").toString().trim(); // B10 : index 4
    const basePrice   = Number(metaValues[5][0]) || 0;              // B11 : index 5
    const estPrice    = Number(metaValues[6][0]) || 0;              // B12 : index 6
    const rank1Raw    = (metaValues[10][0] || "").toString().trim(); // B16 : index 10
    const winnerRaw   = (metaValues[11][0] || "").toString().trim(); // B17 : index 11
    // ────────────────────────────────────────────────────────────

    // 개찰 여부: 예정가격이 입력되어 있어야 개찰된 것으로 판단
    const isOpened = estPrice > 0;
    if (!isOpened) return;

    // ── 참가업체 데이터 일괄 읽기 (E~K열: K=필요PQ점수 포함 7열) ──
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const participantData = sheet.getRange(2, 5, lastRow - 1, 7).getValues();

    const validRows = participantData.filter(r => {
      const rank = r[0];
      return typeof rank === "number" && Number.isInteger(rank) && rank > 0;
    });
    // ────────────────────────────────────────────────────────────

    // ── 낙찰 확정 판정 ────────────────────────────────────────────
    const normalizedWinner       = nameMap[winnerRaw] || winnerRaw;
    const normalizedParticipants = validRows.map(r => {
      const n = r[2].toString().trim();
      return nameMap[n] || n;
    });

    const winnerIdx   = normalizedParticipants.indexOf(normalizedWinner);
    const isConfirmed = winnerRaw !== "" && winnerIdx !== -1;
    // ────────────────────────────────────────────────────────────

    if (isConfirmed) {
      // ── 계약 확정 데이터 ────────────────────────────────────────
      confirmedData.push({
        price:  basePrice,
        code:   code,
        winner: normalizedWinner
      });

      // ── 선순위 탈락자 추출 ─────────────────────────────────────
      validRows.slice(0, winnerIdx).forEach(r => {
        const pName = r[2].toString().trim();
        const pqScore = r[6]; // K열: 필요PQ점수

        disqualifiedData.push({
          projectName: sName,
          company:     nameMap[pName] || pName,
          pqScore:     pqScore, // 필요PQ점수 (점추정)
          area:        similarArea
        });
      });

    } else {
      // ── 계약 미확정 데이터 ─────────────────────────────────────
      const rank1Normalized = nameMap[rank1Raw] || rank1Raw;
      pendingData.push({
        projectName: sName,
        rank1:       rank1Normalized,
        basePrice:   basePrice,
        status:      winnerRaw === "" ? "절차개시중" : winnerRaw,
        code:        code
      });
    }
  });

  // ════════════════════════════════════════════════════════════
  // 리포트 렌더링
  // ════════════════════════════════════════════════════════════

  dash.getRange("B1")
    .setValue("2026 PQ입찰 계약확정 누계 현황")
    .setFontSize(22)
    .setFontWeight("bold");

  let curRow = 4;

  // ── 1. 분류코드별 계약 확정 누계 ──────────────────────────────
  const TARGET_CODES = ["전체", "1", "2", "5", "10", "30"];

  TARGET_CODES.forEach(targetCode => {
    const filtered = targetCode === "전체"
      ? confirmedData
      : confirmedData.filter(d => d.code === targetCode);

    dash.getRange(curRow, 2)
      .setValue("■ [" + targetCode + "] 입찰 계약누계")
      .setFontWeight("bold")
      .setFontSize(14);
    curRow++;

    const sectionTotal = filtered.reduce((acc, d) => acc + d.price, 0);
    const jwSum = filtered
      .filter(d => d.winner === "정우" || d.winner === "정우종합")
      .reduce((acc, d) => acc + d.price, 0);
    const shareRate = sectionTotal > 0 ? jwSum / sectionTotal : 0;

    dash.getRange(curRow, 2, 1, 3)
      .setValues([["총 계약 규모", "정우·정우종합 기초가 합계", "점유율"]])
      .setBackground("#2c3e50")
      .setFontColor("white")
      .setHorizontalAlignment("center")
      .setFontSize(11);
    dash.getRange(curRow + 1, 2, 1, 3)
      .setValues([[sectionTotal, jwSum, shareRate]])
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment("center")
      .setFontSize(11);
    dash.getRange(curRow + 1, 2, 1, 2).setNumberFormat('#,##0"원"');
    dash.getRange(curRow + 1, 4)
      .setNumberFormat("0.0%")
      .setBackground("#fff2cc")
      .setFontWeight("bold");

    const winMap = {};
    filtered.forEach(d => {
      if (!winMap[d.winner]) winMap[d.winner] = { sum: 0, count: 0 };
      winMap[d.winner].sum   += d.price;
      winMap[d.winner].count += 1;
    });
    const sortedRank = Object.entries(winMap)
      .sort((a, b) => b[1].sum - a[1].sum)
      .slice(0, 5);

    if (sortedRank.length > 0) {
      dash.getRange(curRow, 6, 1, 5)
        .setValues([["순위", "업체명", "건수", "기초금액 합계", "점유율"]])
        .setBackground("#34495e")
        .setFontColor("white")
        .setHorizontalAlignment("center")
        .setFontSize(11);

      const rankRows = sortedRank.map(([name, v], i) => [
        i + 1, name, v.count, v.sum,
        sectionTotal > 0 ? v.sum / sectionTotal : 0
      ]);
      dash.getRange(curRow + 1, 6, rankRows.length, 5)
        .setValues(rankRows)
        .setBorder(true, true, true, true, true, true)
        .setFontSize(11);
      dash.getRange(curRow + 1, 9, rankRows.length, 1).setNumberFormat('#,##0"원"');
      dash.getRange(curRow + 1, 10, rankRows.length, 1).setNumberFormat("0.0%");
    }

    curRow = dash.getLastRow() + 3;
  });

  // ── 2. 계약상대방 미확정 현황 ──────────────────────────────────
  curRow = dash.getLastRow() + 2;
  dash.getRange(curRow, 2)
    .setValue("■ 계약상대방 미확정 용역 현황")
    .setFontWeight("bold")
    .setFontSize(16)
    .setFontColor("#e67e22");
  dash.getRange(curRow + 1, 2)
    .setValue("※ 적격심사가 진행 중이거나 낙찰자가 최종 결정되지 않은 용역입니다.")
    .setFontColor("#7f8c8d")
    .setFontSize(11);
  dash.getRange(curRow + 2, 2, 1, 4)
    .setValues([["용역명", "1순위 업체", "기초금액", "현재 진행상태"]])
    .setBackground("#d35400")
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setFontSize(11);

  if (pendingData.length > 0) {
    const pendingRows = pendingData.map(d => [d.projectName, d.rank1, d.basePrice, d.status]);
    dash.getRange(curRow + 3, 2, pendingRows.length, 4)
      .setValues(pendingRows)
      .setBorder(true, true, true, true, true, true)
      .setFontSize(11);
    dash.getRange(curRow + 3, 4, pendingRows.length, 1).setNumberFormat('#,##0"원"');
  } else {
    dash.getRange(curRow + 3, 2)
      .setValue("현재 미확정된 계약 건이 없습니다.")
      .setFontSize(11);
  }

  // ── 3. 선순위 업체 부적격 현황 ────────────────────────────────
  curRow = dash.getLastRow() + 3;
  dash.getRange(curRow, 2)
    .setValue("■ 선순위 업체 부적격(탈락) 정보")
    .setFontWeight("bold")
    .setFontSize(16)
    .setFontColor("#c0392b");
  dash.getRange(curRow + 1, 2)
    .setValue("※ 낙찰자보다 투찰순위가 높았으나 적격심사 등에서 탈락한 업체 리스트입니다.")
    .setFontColor("#7f8c8d")
    .setFontSize(11);
  dash.getRange(curRow + 2, 2, 1, 4)
    .setValues([["대상 용역명", "부적격 업체명", "필요PQ점수", "유사면적(B6)"]])  // 주석도 B5→B6 반영
    .setBackground("#c0392b")
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setFontSize(11);

  if (disqualifiedData.length > 0) {
    const disqRows = disqualifiedData.map(d => [d.projectName, d.company, d.pqScore, d.area]);
    dash.getRange(curRow + 3, 2, disqRows.length, 4)
      .setValues(disqRows)
      .setBorder(true, true, true, true, true, true)
      .setFontSize(11)
      .setVerticalAlignment("middle");
    dash.getRange(curRow + 3, 4, disqRows.length, 1)
      .setNumberFormat("0.00")
      .setHorizontalAlignment("center");
    dash.getRange(curRow + 3, 5, disqRows.length, 1)
      .setNumberFormat("#,##0.00")
      .setHorizontalAlignment("center");
  } else {
    dash.getRange(curRow + 3, 2)
      .setValue("부적격 업체 없었습니다.")
      .setFontSize(11);
  }

  // ── 열 너비 최종 조정 ─────────────────────────────────────────
  dash.setColumnWidth(2, 350);
  dash.setColumnWidth(3, 200);
  dash.setColumnWidth(4, 150);
  dash.setColumnWidth(5, 150);

  ss.toast("PQ입찰 계약 누계가 발행되었습니다.", "알림", 5);
}