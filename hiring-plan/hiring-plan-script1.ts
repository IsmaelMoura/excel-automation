function main(workbook: ExcelScript.Workbook) {
  /**
   * Hiring Plan Report
   */
  const HIRING_PLAN_REPORT_TABLE_NAME = "hiring_plan_report";
  const HIRING_PLAN_REPORT_SHEET = workbook.getWorksheet("hiring-plan-report");

  function createJobTable() {
    HIRING_PLAN_REPORT_SHEET.getRange().getFormat().getFill().clear();

    if (HIRING_PLAN_REPORT_SHEET.getTable(HIRING_PLAN_REPORT_TABLE_NAME)) {
      let clearSheet: void = HIRING_PLAN_REPORT_SHEET.getRanges().clear(
        ExcelScript.ClearApplyTo.all
      );

      throw new Error(
        "A table already exists. Paste your informations, and run the script again"
      );
    } else if (HIRING_PLAN_REPORT_SHEET.getCell(0, 0).getValue() !== "Code") {
      let clearSheet: void = HIRING_PLAN_REPORT_SHEET.getRanges().clear(
        ExcelScript.ClearApplyTo.all
      );

      throw new RangeError(
        "There is no data in the spreadsheet. Paste your informations and run the script again"
      );
    }

    HIRING_PLAN_REPORT_SHEET.getAutoFilter().remove();

    HIRING_PLAN_REPORT_SHEET.addTable(
      HIRING_PLAN_REPORT_SHEET.getRange("A1:Z1").getExtendedRange(
        ExcelScript.KeyboardDirection.down
      ),
      true
    ).setName(HIRING_PLAN_REPORT_TABLE_NAME);

    createHiringManagerInfosTable();
  }

  function createHiringManagerInfosTable() {
    const HIRING_MANAGER_TABLE_HEADERS = [
      "BU",
      "Director",
      "Opened Date",
      "Closing Date",
      "gap to fill",
      "Complexity",
      "TECH or not",
      "Region",
    ];

    let sheetRowLength = HIRING_PLAN_REPORT_SHEET.getUsedRange()
      .getLastColumn()
      .getColumnIndex();

    HIRING_MANAGER_TABLE_HEADERS.map((text, index) => {
      HIRING_PLAN_REPORT_SHEET.getCell(0, sheetRowLength + index + 1).setValue(
        text
      );
    });

    let hiringPlanTable: ExcelScript.Table = HIRING_PLAN_REPORT_SHEET.getTable(
      HIRING_PLAN_REPORT_TABLE_NAME
    );

    setFormulas(hiringPlanTable);
  }

  function setFormulas(hiringPlanTable: ExcelScript.Table) {
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
    openedDateColumn.setFormulaLocal('=IF(S2="";"";TEXT(S2;"yyyy-mm"))');

    createReasonPivotTable(hiringPlanTable);
  }

  /**
   * Pivot Tables
   */
  const PIVOT_TABLES_SHEET = workbook.getWorksheet("pivot");

  function createReasonPivotTable(hiringPlanTable: ExcelScript.Table) {
    const REASON_PIVOT_TABLE_NAME = "Filled-Per-Reason";

    if (PIVOT_TABLES_SHEET.getPivotTable(REASON_PIVOT_TABLE_NAME)) {
      PIVOT_TABLES_SHEET.getPivotTable(REASON_PIVOT_TABLE_NAME).refresh();
      return;
    }

    const filledPerReasonPivotTable = PIVOT_TABLES_SHEET.addPivotTable(
      REASON_PIVOT_TABLE_NAME,
      hiringPlanTable,
      PIVOT_TABLES_SHEET.getCell(0, 0)
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

    createDirectorPivotTable(hiringPlanTable);
  }

  function createDirectorPivotTable(hiringPlanTable: ExcelScript.Table) {
    const DIRECTOR_PIVOT_TABLE_NAME = "Filled-Per-Director";

    if (PIVOT_TABLES_SHEET.getPivotTable(DIRECTOR_PIVOT_TABLE_NAME)) {
      PIVOT_TABLES_SHEET.getPivotTable(DIRECTOR_PIVOT_TABLE_NAME).refresh();
      return;
    }

    let filledPerDirectorPivotTable = PIVOT_TABLES_SHEET.addPivotTable(
      DIRECTOR_PIVOT_TABLE_NAME,
      hiringPlanTable,
      PIVOT_TABLES_SHEET.getCell(0, 3)
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
