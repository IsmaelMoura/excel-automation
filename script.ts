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
  // Get all managers informations column
  // [linha][coluna - 'Gestores' = 0; Diretorias = '1'; 'Job Codes' = 2]
  let all_managers_infos = selectedSheet.getTable('Tabela2').getRange().getValues();


  // Run everything inside as long as the counter is less than 'all_managers_infos.length
  for (let i = 1; i < all_managers_infos.length; i++) {
    for (let count = 1; count < tableLength; count++) {
      // if name of the manager is the same as the other table executes what is inside {}
      if (manager_column[count].toString() == all_managers_infos[i][0]) {
        // Add the 'Job Code' value in the cell
        selectedSheet.getCell(count, jobCodeIndex).setValue(all_managers_infos[i][2]);
        // Add the 'Diretoria' value in the cell
        selectedSheet.getCell(count, diretoriaIndex).setValue(all_managers_infos[i][1]);
      } // if 'manager_column' is null
       else if (manager_column[count].toString() == '') {
        // Add null value in the 'Job Code' cell
        selectedSheet.getCell(count, jobCodeIndex).setValue('');
        // Add null value in the 'Diretoria' cell
        selectedSheet.getCell(count, diretoriaIndex).setValue('');
      }
    }
  }
}
