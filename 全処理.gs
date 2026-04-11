function runAllProcessesThrough() {
  ///console.log('--- 処理を開始します ---');
  
  createFinalCustomTable();
  runAllProcesses();             
  transferToLabels(); 
  applyColoring();            
  copyWithFormat();



  ///console.log('--- すべての処理が正常に完了しました ---');
}