function runAllProcessesThrough() {
  ///console.log('--- 処理を開始します ---');
  
unmergeAndFillValues();
transferData();
freeSheetCleaningAllProcesses();
lookupCourseData2();
convertAndFillDates();
splitRowsByQuantityAdvanced2();
applyAssortPattern2();
updateConcatColumn();
runAllProcessesCrossTable2555();
generatePivotSummary();
runAllProcesses2();
transferToLabels3();
coloring4AllSheets();


  ///console.log('--- すべての処理が正常に完了しました ---');
}