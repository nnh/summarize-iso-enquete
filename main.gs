/**
 * Consolidate the results of questionnaires received from customers.
 * @param none
 * @return none
 */
function summarizeEnquete(){
  const INPUT_SHEET_NAME = 'フォームの回答 1'; 
  const USAGE_PERIOD = {0:'半年以下',
                        1:'1年～3年程度',
                        2:'3年～5年程度',
                        3:'5年以上'}
  const OUTPUT_HEADER = ['使用期間', '初期メニューの構成、操作の分かり易さ', '画面遷移の分かり易さ', '操作中の反応'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);
  const inputValues = inputSheet.getDataRange().getValues();
  const targetYear = (new Date()).getFullYear();
  //const fromDate = new Date(targetYear, 4, 1, 0, 0, 0);
  const fromDate = new Date(2020, 4, 1, 0, 0, 0);
  // Extract only the records for the current year. 
  const targetValues = inputValues.filter(x => x[0] > fromDate);
  let outputValues = targetValues.map(x => {
    let arr = x.slice(2, 6);
    let sortOrder = -1;
    sortOrder = Object.keys(USAGE_PERIOD).filter(val => USAGE_PERIOD[val] == arr[0])[0];    
    if (!sortOrder){
      sortOrder = '999';
    }
    arr.push(parseInt(sortOrder));
    return arr;
  });
  // Sort by "usage period".
  const sortOrderIdx = outputValues[0].length - 1;
  outputValues.sort((x, y) => x[sortOrderIdx] - y[sortOrderIdx]);
  outputValues = outputValues.map(x => x.slice(0, x.length - 1));
  // Output charts.
  const outputChartSheet = addSheet(ss, 'グラフ（' + targetYear + '年度）')
  // Remove charts.
  outputChartSheet.getCharts().forEach(x => outputChartSheet.removeChart(x));
  const outputHeaderAndValues = [OUTPUT_HEADER].concat(outputValues);
  const outputSourceRange = outputChartSheet.getRange(1, 1, outputHeaderAndValues.length, outputHeaderAndValues[0].length);
  outputSourceRange.setValues(outputHeaderAndValues);
  let outputStartRow = outputHeaderAndValues.length + 2;
  for (let i = 0; i < 3; i++){
    outputStartRow = editTable(outputChartSheet, outputValues, OUTPUT_HEADER, USAGE_PERIOD, i + 1, outputStartRow);
  }
  // Output probrems.
  const outputProbremSheet = addSheet(ss, '画面回答（' + targetYear + '年度）');
  let probremEditValues = [['使用期間', '画面URL', '問題点']];
  for (let i = 6; i <= 14; i = i + 2){
    let temp = targetValues.map(x => [x[2]].concat(x.slice(i, i + 2))).filter(x => x[2] != '');
    probremEditValues = probremEditValues.concat(temp);
  }
  temp = targetValues.map(x => [x[2]].concat(x.slice(16, 18))).filter(x => x[1] != '');
  probremEditValues = probremEditValues.concat(temp);
  outputProbremSheet.getRange(1, 1, probremEditValues.length, probremEditValues[0].length).setValues(probremEditValues);
}
/**
 * Create a table to output a graph.
 * @param {sheet} Sheet to output the table.
 * @param {Array.<string>} Values to output.
 * @param {Array.<string>} Values of the column heading.
 * @param {Array.<string>} Values of the row heading.
 * @param {number} Index of target column.
 * @param {number} Output start row.
 * @return {number} Next output start row.
 */
function editTable(outputSheet, targetValues, OUTPUT_HEADER, USAGE_PERIOD, headerIdx, outputStartRow){
  const SCALE_1_5 = {1:'悪い',
                     2:'やや悪い',
                     3:'普通',
                     4:'やや良い',
                     5:'大変良い'}
  outputSheet.getRange(outputStartRow, 1).setValue(OUTPUT_HEADER[headerIdx]);
  const sourceValues = targetValues.map(x => [x[0], SCALE_1_5[x[headerIdx]]]);
  let tableValues = Object.entries(USAGE_PERIOD).map(x => {
    const periodValue = x[1];
    const periodRow = Object.entries(SCALE_1_5).map(x => {
      const scaleValue = x[1];
      return sourceValues.filter(x => x[0] == periodValue && x[1] == scaleValue).length;
    });
    return periodRow;
  });
  // sum
  const sumValues = tableValues.reduce((accumlator, currentValue) => accumlator.map((x, idx) => x + currentValue[idx]));
  tableValues = [...tableValues, sumValues];
  const tableStartRow = outputStartRow;
  let tempOutputRow = outputStartRow;
  let tableRowHeader = Object.entries(USAGE_PERIOD).map(x => x[1]);
  tableRowHeader = ['', ...tableRowHeader, '計'];
  tableRowHeader.forEach(x => {
    tempOutputRow++;
    outputSheet.getRange(tempOutputRow, 1).setValue(x);
  });
  // Output table.
  const tableColHeader = Object.entries(SCALE_1_5).map(x => x[1]);
  outputStartRow++;
  outputSheet.getRange(outputStartRow, 2, 1, tableColHeader.length).setValues([tableColHeader]);
  outputStartRow++;
  outputSheet.getRange(outputStartRow, 2, tableValues.length, tableValues[0].length).setValues(tableValues);
  const chartSourceRange = outputSheet.getRange(tableStartRow + 1, 1, tableValues.length, tableRowHeader.length);
  chartSourceRange.setBorder(true, true, true, true, true, true);
  const sumRow = outputSheet.getRange(chartSourceRange.getLastRow() + 1, 1, 1, tableRowHeader.length);
  sumRow.setBorder(true, true, true, true, true, true);
  sumRow.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  const nextStartRow = editChart(outputSheet, chartSourceRange);
  return nextStartRow;
}
/**
 * Create a chart.
 * @param {sheet} Sheet to output the table.
 * @param {Array.<string>} Values to output.
 * @return {number} Next output start row.
 */
function editChart(outputSheet, targetRange){
  const CHART_HEIGHT = 19;
  const strTitle = targetRange.offset(-1, 0, 1, 1).getValues()[0][0];
  const outputChartRow = targetRange.getLastRow() + 3;
  let newChart = outputSheet.newChart()
    .addRange(targetRange)
    .setTransposeRowsAndColumns(true)
    .setPosition(outputChartRow, 1, 0, 0)
    .asColumnChart()
    .setOption('title', strTitle)
    .setOption('annotations.total.enabled', true)
    .setNumHeaders(1)
    .setStacked();
  outputSheet.insertChart(newChart.build());
  return outputChartRow + CHART_HEIGHT;
}
/**
 * Add a sheet for output.
 * @param {spreadsheet} Spreadsheet to add sheets to output.
 * @param {string} A sheet name.
 * @return {sheet} a sheet for output.
 */
function addSheet(ss, outputSheetName){
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet){
    outputSheet = ss.insertSheet();
    outputSheet.setName(outputSheetName);
  }
  outputSheet.clear();
  return outputSheet;
}