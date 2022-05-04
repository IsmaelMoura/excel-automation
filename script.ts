function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const selectedSheet = workbook.getActiveWorksheet();
  // Get the 'Gestor' column
  let manager_column = selectedSheet.getRange('A2:A1000').getValues();
  // Get table lenght
  let tableLength = selectedSheet.getTable('Tabela1').getRowCount();
  // Get 'Job Code' column index
  let jobCodeIndex = selectedSheet.getTable('Tabela1').getColumnByName('Job Code').getIndex();
  // Get 'Diretoria' column index
  let diretoriaIndex = selectedSheet.getTable('Tabela1').getColumnByName('Diretoria').getIndex();
  // Get all managers informations table
  let all_managers_infos = selectedSheet.getTable('Tabela2').getRange().getValues();

  // [linha][coluna - 'Gestores' = 0; Diretorias = '1'; 'Job Codes' = 2]

  // 
  for (let i = 1; i < all_managers_infos.length; i++) {
    for (let count = 1; count < tableLength; count++) {
      if (manager_column[count].toString() == all_managers_infos[i][0]) {
        selectedSheet.getCell(count, jobCodeIndex).setValue(all_managers_infos[i][2]);
        selectedSheet.getCell(count, diretoriaIndex).setValue(all_managers_infos[i][1]);
      } else if (manager_column[count].toString() == '') {
        selectedSheet.getCell(count, jobCodeIndex).setValue('');
        selectedSheet.getCell(count, diretoriaIndex).setValue('');
      }
    }
  }
}
