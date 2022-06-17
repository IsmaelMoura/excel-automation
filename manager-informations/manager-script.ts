function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getWorksheet("Informações Pessoais");
  let manager_column = selectedSheet
    .getTable("Infos_candidatos")
    .getColumnByName("Gestão direta")
    .getRange()
    .getValues();
  let tableLength = selectedSheet.getTable("Infos_candidatos").getRowCount();
  let diretoria_index = selectedSheet
    .getTable("Infos_candidatos")
    .getColumnByName("Diretoria Sinch")
    .getIndex();
  let business_unit_index = selectedSheet
    .getTable("Infos_candidatos")
    .getColumnByName("Business Unit Sinch")
    .getIndex();
  let id_centro_de_custo_index = selectedSheet
    .getTable("Infos_candidatos")
    .getColumnByName("ID Centro de Custo Sinch")
    .getIndex();
  let nome_centro_de_custo_index = selectedSheet
    .getTable("Infos_candidatos")
    .getColumnByName("Nome do Centro de Custo")
    .getIndex();
  // [linha][coluna - 'Gestores' = 0; Diretorias = '1'; 'Business Unit Sinch' = 2; 'ID Centro de Custo Sinch' = 3; 'Nome do Centro de Custo' = 4]
  let all_managers_infos = workbook
    .getWorksheet("Referências")
    .getTable("managers_infos")
    .getRange()
    .getValues();

  for (let i = 1; i < all_managers_infos.length; i++) {
    for (let count = 0; count <= tableLength; count++) {
      if (manager_column[count].toString() == all_managers_infos[i][0]) {

        selectedSheet
          .getCell(count + 1, diretoria_index)
          .setValue(all_managers_infos[i][1]);
        
          selectedSheet
          .getCell(count + 1, business_unit_index)
          .setValue(all_managers_infos[i][2]);
        
          selectedSheet
          .getCell(count + 1, id_centro_de_custo_index)
          .setValue(all_managers_infos[i][3]);
        
          selectedSheet
          .getCell(count + 1, nome_centro_de_custo_index)
          .setValue(all_managers_infos[i][4]);
      }
    }
  }
}
