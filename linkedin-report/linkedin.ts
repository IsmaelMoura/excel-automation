function main(workbook: ExcelScript.Workbook) {
	const SELECTED_WORKSHEET = workbook.getActiveWorksheet();
	const today = new Date();
	const monthName = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

	function splitInformations() {
		if (!SELECTED_WORKSHEET.getUsedRange()) {
			let clearSheet: void = SELECTED_WORKSHEET.getRanges().clear(ExcelScript.ClearApplyTo.all);
			throw new Error('There is no data in the spreadsheet. Paste your informations and run the script again');
		}

		let header = SELECTED_WORKSHEET.getRange('A1').getValues().toString().split(/[,]/);
		let infos = SELECTED_WORKSHEET.getRange('A2').getValues().toString().split(/[,]/);

		header.map((text, index) => {
			SELECTED_WORKSHEET.getCell(0, index).setValue(text);
		});

		infos.map((text, index) => {
			SELECTED_WORKSHEET.getCell(1, index).setValue(text);
		});

	}

	function copyInformations() {
		let profileViews: ExcelScript.Range = SELECTED_WORKSHEET.getRange('K1:K2');
		let profileSaved: ExcelScript.Range = SELECTED_WORKSHEET.getRange('N1:N2');
		let messagesSent: ExcelScript.Range = SELECTED_WORKSHEET.getRange('R1:R2');
		let messagesAccepted: ExcelScript.Range = SELECTED_WORKSHEET.getRange('S1:S2');

		// insert infornations in the cells 
		SELECTED_WORKSHEET.getRange("A5").copyFrom(profileViews, ExcelScript.RangeCopyType.all, false, false);
		SELECTED_WORKSHEET.getRange("B5").copyFrom(profileSaved, ExcelScript.RangeCopyType.all, false, false);
		SELECTED_WORKSHEET.getRange("C5").copyFrom(messagesSent, ExcelScript.RangeCopyType.all, false, false);
		SELECTED_WORKSHEET.getRange("D5").copyFrom(messagesAccepted, ExcelScript.RangeCopyType.all, false, false);
	}

	function createChart() {
		let chartTitle: string = `LinkedIn source - ${monthName[today.getMonth() - 1]}`;

		if (SELECTED_WORKSHEET.getChart(chartTitle)) {
			SELECTED_WORKSHEET.getChart(chartTitle).delete();
		}

		let chart = SELECTED_WORKSHEET.addChart(ExcelScript.ChartType.pie, SELECTED_WORKSHEET.getRange("A5:D6"));

		chart.setName(chartTitle);
		chart.getTitle().setText(chartTitle);
		chart.getSeries()[0].setHasDataLabels(true);

		// set chart position
		chart.setLeft(50);
		chart.setTop(100);
	};

	try {
		splitInformations();
		copyInformations();
		createChart();
	} catch (err) {
		throw err;
	} finally {
		console.log('If you need support, email me at: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura');
	};
};