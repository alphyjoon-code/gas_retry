// @ts-nocheck
/**
 * 상단 통합 메뉴 생성
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ 입찰도구')
      .addItem('0-1. 면적조회', 'analyzeSizeData')
      .addSeparator()
      .addItem('1-1. 입찰분석', 'runAllBidAnalysis')
      .addItem('1-2. 연번부여', 'updateSheetSerialNumbers')
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
      .addSeparator() // 
      .addItem('4-1. 경쟁강도 분석', 'analyzeIntensityToColumns')
      .addItem('4-2. 경쟁강도 취합', 'collectBiddingData')
      .addItem('4-3. 가치구간 분석', 'generateParallelDashboard')
      .addSeparator()
      .addItem('5-1. 전기간 예가 시뮬레이션', 'updateSimulationWithRevenue')
      .addItem('5-2. 특정기간 예가 시뮬레이션', 'updateSimulationByPeriod')
      .addItem("📊 시뮬레이션 차트 보기", "showSimulationChart")
      .addItem("🎛️ 인터랙티브 시뮬레이션", "showInteractiveSimulation")
      .addSeparator()
      // 2026-06-23 기능 비활성화(사용자 결정): L열이 PQ하한(필요PQ점수) 전용으로
      // 재정의되어, L열에 기록하는 'PQ 예측(사후)'(estimatePQScores)를 실행하면
      // updatePQTierColumns의 결과를 덮어써 충돌함. 같은 결정으로 'PQ 예측(사전)'
      // (runPreBidPQEstimate, 실제로는 외부 파일에만 기록해 구조적 충돌은 없음)도
      // 함께 비활성화함. estimatePQScores.js/runPreBidPQEstimate.js 파일은 이후
      // 실제로 삭제됨 — 재활성화하려면 아래 두 줄의 주석 해제만으로는 안 되고
      // 해당 함수를 새로 작성해야 함.
      // .addItem('PQ 예측(사전)', 'runPreBidPQEstimate')
      // .addItem('PQ 예측(사후)', 'estimatePQScores')
      .addToUi();
}
