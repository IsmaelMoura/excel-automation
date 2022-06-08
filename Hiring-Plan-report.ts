function main(workbook: ExcelScript.Workbook) {
	/**
	 * Hiring Plan Report
	 */
	const HIRING_PLAN_REPORT_SHEET = workbook.getWorksheet('hiring_plan_report_2022-01-01_2');

	const createJobTable = () => {
		if (workbook.getTable('Hiring_plan_report')) {
			throw new Error('A table already exists. Delete it, paste your informations, and run the script again')
		} else if (!HIRING_PLAN_REPORT_SHEET.getUsedRange()) {
			throw new Error('There is no data in the spreadsheet. Clean it completely, paste your information and run the script again')
		}

		HIRING_PLAN_REPORT_SHEET.getAutoFilter().remove()
		// Add a new table at extended range obtained by extending right, then down from range A1:Z1 on HIRING_PLAN_REPORT_SHEET
		workbook.addTable(HIRING_PLAN_REPORT_SHEET.getRange('A1:Z1').getExtendedRange(ExcelScript.KeyboardDirection.down), true).setName('Hiring_plan_report')

		createHiringManagerInfosTable()
	}

	const createHiringManagerInfosTable = () => {
		const HIRING_MANAGER_TABLE_HEADERS = [
			'BU',
			'Director',
			'Opened Date',
			'Closing Date',
			'gap to fill',
			'Complexity',
			'TECH or not',
			'Region'
		]

		let sheetRowLength = HIRING_PLAN_REPORT_SHEET.getUsedRange().getLastColumn().getColumnIndex()
		let hiringManagerRange = HIRING_PLAN_REPORT_SHEET.getRangeByIndexes(0, sheetRowLength + 1, 1, HIRING_MANAGER_TABLE_HEADERS.length)

		for (let i = 0; i < HIRING_MANAGER_TABLE_HEADERS.length; i++) {
			HIRING_PLAN_REPORT_SHEET.getCell(0, sheetRowLength + i + 1).setValue(HIRING_MANAGER_TABLE_HEADERS[i])
		}

		let jobInformationsTable: ExcelScript.Table = workbook.getTable('Hiring_plan_report')

		setFormula(jobInformationsTable)
	}

	const setFormula = (table: ExcelScript.Table) => {
		let buColumn = table.getColumnByName('BU')
		let directorColumn = table.getColumnByName('Director')
		let openedDateColumn = table.getColumnByName('Closing Date')

		// add formula into BU, Director and Closing Date column cells
		buColumn.getRangeBetweenHeaderAndTotal().setFormulaLocal("=PROCV(K2;'To. For'!A:B;2;0)")
		directorColumn.getRangeBetweenHeaderAndTotal().setFormulaLocal("=PROCV(K2;'To. For'!A:C;3;0)")
		openedDateColumn.getRangeBetweenHeaderAndTotal().setFormulaLocal("=ESQUERDA(S2;7)")

		newPivotTable()
	}


	/**
	 * Pivotable
	 */

	const newPivotTable = () => {
		const PIVOT_TABLE_SHEET = workbook.getWorksheet("pivot")
		const HIRING_PLAN_REPORT_TABLE = workbook.getTable("Hiring_plan_report")

		workbook.addPivotTable("PivotTable HiringPlan", HIRING_PLAN_REPORT_TABLE, PIVOT_TABLE_SHEET.getCell(3, 0))
	}

	/**
	 * Inicialize principal function
	 */
	createJobTable()


	/**
	 * Suport messaging
	 */
	console.log('If you need support, email me at: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura')

}
