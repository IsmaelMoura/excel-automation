function main(workbook: ExcelScript.Workbook) {

	/**
	 * Hiring Plan Report
	 */
  const HIRING_PLAN_REPORT_SHEET = workbook.getWorksheet('hiring_plan_report_2022-01-01_2');

  function createJobTable() {
    if (workbook.getTable('Hiring_plan_report')) {
      let clearSheet: void = HIRING_PLAN_REPORT_SHEET.getRanges().clear(ExcelScript.ClearApplyTo.all)
      throw new Error('A table already exists. Paste your informations, and run the script again')
    } else if (HIRING_PLAN_REPORT_SHEET.getCell(0, 0).getValue() !== 'Code') {
      let clearSheet: void = HIRING_PLAN_REPORT_SHEET.getRanges().clear(ExcelScript.ClearApplyTo.all)
      throw new Error('There is no data in the spreadsheet. Paste your informations and run the script again')
    }

    HIRING_PLAN_REPORT_SHEET.getAutoFilter().remove()

    workbook.addTable(HIRING_PLAN_REPORT_SHEET.getRange('A1:Z1').getExtendedRange(ExcelScript.KeyboardDirection.down), true).setName('Hiring_plan_report')

    createHiringManagerInfosTable()
  }

  function createHiringManagerInfosTable() {
    let hiringManagerTableHeaders = [
      'BU',
      'Director',
      'Opened Date',
      'Closing Date',
      'gap to fill',
      'Complexity',
      'TECH or not',
      'Region'
    ]

    let sheetRowLength = HIRING_PLAN_REPORT_SHEET.getUsedRange().getLastColumn().getColumnIndex();

    hiringManagerTableHeaders.map((text, index) => {
      HIRING_PLAN_REPORT_SHEET
        .getCell(0, sheetRowLength + index + 1)
        .setValue(text)
    });

    let jobInformationsTable: ExcelScript.Table = workbook.getTable('Hiring_plan_report');

    setFormula(jobInformationsTable);
  }

  function setFormula(table: ExcelScript.Table) {
    let buColumn = table.getColumnByName('BU')
    let directorColumn = table.getColumnByName('Director')
    let openedDateColumn = table.getColumnByName('Closing Date')

    buColumn.getRangeBetweenHeaderAndTotal().setFormulaLocal("=VLOOKUP(K2;'To. For'!A:B;2;0)")
    directorColumn.getRangeBetweenHeaderAndTotal().setFormulaLocal("=VLOOKUP(K2;'To. For'!A:C;3;0)")
    openedDateColumn.getRangeBetweenHeaderAndTotal().setFormulaLocal("=LEFT(S2;7)")

    reasonPivotTable()
  }


	/**
	 * PivotTable
	 */
  function reasonPivotTable() {
    const PIVOT_SHEET = workbook.getWorksheet("pivot")
    const HIRING_PLAN_REPORT_TABLE = workbook.getTable("Hiring_plan_report")

    if (PIVOT_SHEET.getPivotTable('Filled-Per-Reason')) {
      PIVOT_SHEET.getPivotTable('Filled-Per-Reason').refresh()
      return
    }

    const PIVOT_TABLE = workbook.addPivotTable(
      "Filled-Per-Reason", HIRING_PLAN_REPORT_TABLE, PIVOT_SHEET.getCell(0, 0)
    )

    PIVOT_TABLE.addFilterHierarchy(PIVOT_TABLE.getHierarchy("Closing Date"));
    PIVOT_TABLE.addRowHierarchy(PIVOT_TABLE.getHierarchy("Reason"));
    PIVOT_TABLE.addDataHierarchy(PIVOT_TABLE.getHierarchy("Code"));

    directorPivotTable()
  }

  function directorPivotTable() {
    const PIVOT_SHEET = workbook.getWorksheet("pivot")
    const HIRING_PLAN_REPORT_TABLE = workbook.getTable("Hiring_plan_report")

    if (PIVOT_SHEET.getPivotTable('Filled-Per-Director')) {
      PIVOT_SHEET.getPivotTable('Filled-Per-Director').refresh()
      return
    }

    const PIVOT_TABLE = workbook.addPivotTable(
      "Filled-Per-Director", HIRING_PLAN_REPORT_TABLE, PIVOT_SHEET.getCell(0, 3)
    )

    PIVOT_TABLE.addFilterHierarchy(PIVOT_TABLE.getHierarchy("Closing Date"));
    PIVOT_TABLE.addRowHierarchy(PIVOT_TABLE.getHierarchy("Director"));
    PIVOT_TABLE.addDataHierarchy(PIVOT_TABLE.getHierarchy("Code"));
  }


  try {
    createJobTable()
  } catch (err) {
    throw err
  } finally {
    console.log('If you need support, email me at: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura')
  }
}
