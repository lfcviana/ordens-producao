 
function ENTREGAFINAL() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ENC DIA'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setName('ENT DIA');  
  spreadsheet.getRange('M:M').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('L1'));
  spreadsheet.getActiveSheet().insertColumnsAfter(spreadsheet.getActiveRange().getLastColumn(), 1);
  spreadsheet.getActiveRange().offset(0, spreadsheet.getActiveRange().getNumColumns(), spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getActiveSheet().insertColumnsAfter(spreadsheet.getActiveRange().getLastColumn(), 1);
  spreadsheet.getActiveRange().offset(0, spreadsheet.getActiveRange().getNumColumns(), spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getRange('4:157').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('G4'));
  spreadsheet.getRange('4:157').createFilter();
  spreadsheet.getRange('L4').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['Balcão', 'Cancelado'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(12, criteria);
  spreadsheet.getRange('N4').activate();
  spreadsheet.getCurrentCell().setValue('R$');
  spreadsheet.getRange('O4').activate();
  spreadsheet.getCurrentCell().setValue('OK');
  spreadsheet.getRange('L2:R2').activate();
  spreadsheet.getCurrentCell().setFormula('=E5');
  spreadsheet.getActiveRangeList().setNumberFormat('dd/MM/yyyy')
  .setHorizontalAlignment('right');
  spreadsheet.getRange('R:S').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('L1'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('O:O').activate();
};


function VALOR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('P:Q').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('M:M').activate();
  spreadsheet.getActiveSheet().insertColumnsAfter(spreadsheet.getActiveRange().getLastColumn(), 1);
  spreadsheet.getActiveRange().offset(0, spreadsheet.getActiveRange().getNumColumns(), spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getActiveSheet().insertColumnsAfter(spreadsheet.getActiveRange().getLastColumn(), 1);
  spreadsheet.getActiveRange().offset(0, spreadsheet.getActiveRange().getNumColumns(), spreadsheet.getActiveRange().getNumRows(), 1).activate();
  spreadsheet.getRange('N1').activate();
  spreadsheet.getCurrentCell().setValue('R$');
  spreadsheet.getRange('O1').activate();
  spreadsheet.getCurrentCell().setValue('OK');
};

function EXE_ENTD_8_Final() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L2:R2').activate();
  spreadsheet.getCurrentCell().setFormula('=E5');
  spreadsheet.getActiveRangeList().setNumberFormat('dd/MM/yyyy')
  .setHorizontalAlignment('right');
  spreadsheet.getRange('U:U').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('O:O').activate();
  spreadsheet.getRange('N4').activate();
  spreadsheet.getCurrentCell().setValue('r$');
  spreadsheet.getRange('O4').activate();
  spreadsheet.getCurrentCell().setValue('OK');
  spreadsheet.getRange('O8').activate();
};

function EXE_ENTD_9_Valor_ok() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('4:158').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('G4'));
  spreadsheet.getRange('4:158').createFilter();
  spreadsheet.getRange('L4').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', '16:00', 'Balcão', 'Cancelado'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(12, criteria);
};


function Selecao() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:4').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('G1'));
};


