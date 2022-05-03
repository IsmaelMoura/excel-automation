function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();
  // Get the 'Gestor' column
  let manager_column = selectedSheet.getRange('A2:A1000').getValues()
  // Get table lenght
  let tableLength = selectedSheet.getTable('Tabela1').getRowCount()
  // Get 'Job Code' column index
  let jobCodeIndex = selectedSheet.getTable('Tabela1').getColumnByName('Job Code').getIndex()
  // Get 'Diretoria' column index
  let diretoriaIndex = selectedSheet.getTable('Tabela1').getColumnByName('Diretoria').getIndex()

  let teste = workbook


  for (let i = 0; i < tableLength + 1; i++) {
    switch (manager_column[i].toString()) {
      case 'teste2' && 'teste1':
        selectedSheet.getCell(i + 1, jobCodeIndex).setValue('1234')
        selectedSheet.getCell(i + 1, diretoriaIndex).setValue('4567')
        break;
    }
  }
}
