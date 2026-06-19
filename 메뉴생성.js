// @ts-nocheck
/**
 * 상단 통합 메뉴 생성
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ 입찰도구')
      .addItem('1-1. 업체명 단축', 'convertCompanyNamesToShort')
      .addItem('1-2, 입찰결과 분석', 'runAllBidAnalysis')
      .addItem('1-2. 예가율 계산', 'updateBiddingResults')
      .addItem('1-3. 가산점 기입', 'fillBonusScores')
      .addItem('1-3. 연번부여', 'updateSheetSerialNumbers')
      .addSeparator()
      .addItem('2-1. 일일현황 발행', 'createLatestPQStatusSheetWithProfessionalSummary')
      .addItem('2-2. 입찰누계 발행', 'updateMasterDashboard')
      .addItem('2-3. 계약누계 발행', 'updateContractorDashboard')
      .addItem('2-4. 네비게이션 갱신', 'createNavigationDashboard')
      .addItem('2-5. 입찰결과 전송', 'updateResultFile')
      .addItem('2-6. 업체별 정보 취합', 'extractAndPrepareDotPlot')
      .addSeparator()
      .addItem('3-1. PQ백데이터 취합', 'collectJungwooData')
      .addItem('3-2. PQ백데이터(분석용) 취합', 'collectJungwooDataforanalysis')
      .addItem('3-3. 업체수 복수예가 파일로 전송', 'syncCompanyCountToMaster')
      .addSeparator() // 이 부분을 수정했습니다.
      .addItem('4-1. 경쟁강도 분석', 'analyzeIntensityToColumns')
      .addItem('4-2. 경쟁강도 취합', 'collectBiddingData')
      .addItem('4-3. 가치구간 분석', 'generateParallelDashboard')
      .addSeparator()
      .addItem('5-1. 전기간 예가 시뮬레이션', 'updateSimulationWithRevenue')
      .addItem('5-2. 특정기간 예가 시뮬레이션', 'updateSimulationByPeriod')
      .addItem("📊 시뮬레이션 차트 보기", "showSimulationChart")
      .addItem("🎛️ 인터랙티브 시뮬레이션", "showInteractiveSimulation")
      .addSeparator()
      .addItem('PQ 예측(사전)', 'runPreBidPQEstimate')
      // 2026-06-19 롤백: L열이 점추정(추정PQ점수)으로 복원되어 'PQ 예측(사후)'
      // (estimatePQScores, L열에 기록)와의 충돌이 해소되어 재활성화함.
      .addItem('PQ 예측(사후)', 'estimatePQScores')
      .addToUi();
}
