function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  const selectedSheet = workbook.getWorksheet('Informações Pessoais');
  // Get the values of 'Gestão Direta' column
  let manager_column = selectedSheet.getTable('Infos_candidatos').getColumnByName('Gestão direta').getRange().getValues();
  // Get table lenght
  let tableLength = selectedSheet.getTable('Infos_candidatos').getRowCount();
  // Get 'Diretoria' column index
  let diretoria_index = selectedSheet.getTable('Infos_candidatos').getColumnByName('Diretoria Sinch').getIndex();
  // Get 'Business Unit' column index
  let business_unit_index = selectedSheet.getTable('Infos_candidatos').getColumnByName('Business Unit Sinch').getIndex();
  // Get 'Centro de Custo' column index
  let id_centro_de_custo_index = selectedSheet.getTable('Infos_candidatos').getColumnByName('ID Centro de Custo Sinch').getIndex();
  // Get 'Nome do Centro de Custo' column index
  let nome_centro_de_custo_index = selectedSheet.getTable('Infos_candidatos').getColumnByName('Nome do Centro de Custo').getIndex();
  // Get all managers informations column
  // [linha][coluna - 'Gestores' = 0; Diretorias = '1'; 'Business Unit Sinch' = 2; 'ID Centro de Custo Sinch' = 3; 'Nome do Centro de Custo' = 4]
  let all_managers_infos = workbook.getWorksheet('Referências').getTable('managers_infos').getRange().getValues();

  // Run everything inside as long as the counter is less than 'all_managers_infos.length'
  for (let i = 1; i < all_managers_infos.length; i++) {
    for (let count = 0; count <= tableLength; count++) {
      // if name of the manager is the same as the other table executes what is inside {}
      if (manager_column[count].toString() == all_managers_infos[i][0]) {
        // Add the 'Diretoria Sinch' value in the cell
        selectedSheet.getCell(count + 1, diretoria_index).setValue(all_managers_infos[i][1]);
        // Add the 'Business Unit Sinch' value in the cell
        selectedSheet.getCell(count + 1, business_unit_index).setValue(all_managers_infos[i][2]);
        // Add the 'ID Centro de Custo Sinch' value in the cell
        selectedSheet.getCell(count + 1, id_centro_de_custo_index).setValue(all_managers_infos[i][3]);
        // Add the 'Nome do Centro de Custo' value in the cell
        selectedSheet.getCell(count + 1, nome_centro_de_custo_index).setValue(all_managers_infos[i][4])
      }
    }
  }
}