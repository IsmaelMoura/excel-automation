function main(workbook: ExcelScript.Workbook) {
  /**
   * Hiring Plan Report
   */
  const HIRING_PLAN_REPORT_TABLE_NAME = "hiring_plan_report";
  let hiringPlanReportSheet = workbook.getWorksheet("hiring-plan-report");

  function createJobTable() {
    if (hiringPlanReportSheet.getTable(HIRING_PLAN_REPORT_TABLE_NAME)) {
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

      throw new RangeError(
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
      .setName(HIRING_PLAN_REPORT_TABLE_NAME);

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

    let hiringPlanTable: ExcelScript.Table = hiringPlanReportSheet.getTable(
      HIRING_PLAN_REPORT_TABLE_NAME
    );

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
    const REASON_PIVOT_TABLE_NAME = "Filled-Per-Reason";

    if (pivotTablesSheet.getPivotTable(REASON_PIVOT_TABLE_NAME)) {
      pivotTablesSheet.getPivotTable(REASON_PIVOT_TABLE_NAME).refresh();
      return;
    }

    const filledPerReasonPivotTable = pivotTablesSheet.addPivotTable(
      REASON_PIVOT_TABLE_NAME,
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
    const DIRECTOR_PIVOT_TABLE_NAME = "Filled-Per-Director";

    if (pivotTablesSheet.getPivotTable(DIRECTOR_PIVOT_TABLE_NAME)) {
      pivotTablesSheet.getPivotTable(DIRECTOR_PIVOT_TABLE_NAME).refresh();
      return;
    }

    let filledPerDirectorPivotTable = pivotTablesSheet.addPivotTable(
      DIRECTOR_PIVOT_TABLE_NAME,
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
