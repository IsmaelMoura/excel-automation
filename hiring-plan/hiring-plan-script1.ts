function main(workbook: ExcelScript.Workbook) {
  /**
   * Hiring Plan Report
   */
  const hiringPlanReportSheet = workbook.getWorksheet(
    "hiring_plan_report_2022-01-01_2"
  );

  function createJobTable() {
    if (hiringPlanReportSheet.getTable("hiring_plan_report")) {
      let clearSheet: void = hiringPlanReportSheet
        .getRanges()
        .clear(ExcelScript.ClearApplyTo.all);

      throw new Error(
        "A table already exists. Paste your informations, and run the script again"
      );
    } else if (hiringPlanReportSheet.getCell(0, 0).getValue() !== "Code") {
      let clearSheet: void = hiringPlanReportSheet
        .getRanges()
        .clear(ExcelScript.ClearApplyTo.all);

      throw new Error(
        "There is no data in the spreadsheet. Paste your informations and run the script again"
      );
    }

    hiringPlanReportSheet.getAutoFilter().remove();

    hiringPlanReportSheet
      .addTable(
        hiringPlanReportSheet
          .getRange("A1:Z1")
          .getExtendedRange(ExcelScript.KeyboardDirection.down),
        true
      )
      .setName("hiring_plan_report");

    createHiringManagerInfosTable();
  }

  function createHiringManagerInfosTable() {
    let hiringManagerTableHeaders = [
      "BU",
      "Director",
      "Opened Date",
      "Closing Date",
      "gap to fill",
      "Complexity",
      "TECH or not",
      "Region",
    ];

    let sheetRowLength = hiringPlanReportSheet
      .getUsedRange()
      .getLastColumn()
      .getColumnIndex();

    hiringManagerTableHeaders.map((text, index) => {
      hiringPlanReportSheet
        .getCell(0, sheetRowLength + index + 1)
        .setValue(text);
    });

    let hiringPlanTable: ExcelScript.Table =
      hiringPlanReportSheet.getTable("hiring_plan_report");

    setFormula(hiringPlanTable);
  }

  function setFormula(hiringPlanTable: ExcelScript.Table) {
    let buColumn = hiringPlanTable
      .getColumnByName("BU")
      .getRangeBetweenHeaderAndTotal();

    let directorColumn = hiringPlanTable
      .getColumnByName("Director")
      .getRangeBetweenHeaderAndTotal();

    let openedDateColumn = hiringPlanTable
      .getColumnByName("Closing Date")
      .getRangeBetweenHeaderAndTotal();

    buColumn.setFormulaLocal("=VLOOKUP(K2;'To. For'!A:B;2;0)");
    directorColumn.setFormulaLocal("=VLOOKUP(K2;'To. For'!A:C;3;0)");
    openedDateColumn.setFormulaLocal("=LEFT(S2;7)");

    createReasonTable(hiringPlanTable);
  }

  /**
   * PivotTables
   */
  const pivotTablesSheet = workbook.getWorksheet("pivot");

  function createReasonTable(hiringPlanTable: ExcelScript.Table) {
    let reasonPivotTableName = "Filled-Per-Reason";

    if (pivotTablesSheet.getPivotTable(reasonPivotTableName)) {
      pivotTablesSheet.getPivotTable(reasonPivotTableName).refresh();
      return;
    }

    const filledPerReasonPivotTable = pivotTablesSheet.addPivotTable(
      reasonPivotTableName,
      hiringPlanTable,
      pivotTablesSheet.getCell(0, 0)
    );

    filledPerReasonPivotTable.addFilterHierarchy(
      filledPerReasonPivotTable.getHierarchy("Closing Date")
    );
    filledPerReasonPivotTable.addRowHierarchy(
      filledPerReasonPivotTable.getHierarchy("Reason")
    );
    filledPerReasonPivotTable.addDataHierarchy(
      filledPerReasonPivotTable.getHierarchy("Code")
    );

    directorPivotTable(hiringPlanTable);
  }

  function directorPivotTable(hiringPlanTable: ExcelScript.Table) {
    let directorPivotTableName = "Filled-Per-Director";

    if (pivotTablesSheet.getPivotTable(directorPivotTableName)) {
      pivotTablesSheet.getPivotTable(directorPivotTableName).refresh();
      return;
    }

    let filledPerDirectorPivotTable = pivotTablesSheet.addPivotTable(
      directorPivotTableName,
      hiringPlanTable,
      pivotTablesSheet.getCell(0, 3)
    );

    filledPerDirectorPivotTable.addFilterHierarchy(
      filledPerDirectorPivotTable.getHierarchy("Closing Date")
    );

    filledPerDirectorPivotTable.addRowHierarchy(
      filledPerDirectorPivotTable.getHierarchy("Director")
    );

    filledPerDirectorPivotTable.addDataHierarchy(
      filledPerDirectorPivotTable.getHierarchy("Code")
    );
  }

  createJobTable();

  console.log(
    "If you need support, send an email to: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura."
  );
}