function main(workbook: ExcelScript.Workbook) {
  const SELECTED_WORKSHEET = workbook.getActiveWorksheet();
  const TODAY = new Date();
  const MONTH_NAME = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  function splitInformations() {
    if (!SELECTED_WORKSHEET.getUsedRange()) {
      let clearSheet: void = SELECTED_WORKSHEET.getRanges().clear(
        ExcelScript.ClearApplyTo.all
      );
      throw new Error(
        "There is no data in the spreadsheet. Paste your informations and run the script again"
      );
    }

    let header = SELECTED_WORKSHEET.getRange("A1")
      .getValues()
      .toString()
      .split(/[,]/);
    let infos = SELECTED_WORKSHEET.getRange("A2")
      .getValues()
      .toString()
      .split(/[,]/);

    header.map((text, index) => {
      SELECTED_WORKSHEET.getCell(0, index).setValue(text);
    });

    infos.map((text, index) => {
      SELECTED_WORKSHEET.getCell(1, index).setValue(text);
    });
  }

  function copyInformations() {
    let profileViews: ExcelScript.Range = SELECTED_WORKSHEET.getRange("K1:K2");
    let profileSaved: ExcelScript.Range = SELECTED_WORKSHEET.getRange("N1:N2");
    let messagesSent: ExcelScript.Range = SELECTED_WORKSHEET.getRange("R1:R2");
    let messagesAccepted: ExcelScript.Range = SELECTED_WORKSHEET.getRange("S1:S2");

    // insert infornations in the cells
    SELECTED_WORKSHEET.getRange("A5").copyFrom(
      profileViews,
      ExcelScript.RangeCopyType.all,
      false,
      false
    );
    SELECTED_WORKSHEET.getRange("B5").copyFrom(
      profileSaved,
      ExcelScript.RangeCopyType.all,
      false,
      false
    );
    SELECTED_WORKSHEET.getRange("C5").copyFrom(
      messagesSent,
      ExcelScript.RangeCopyType.all,
      false,
      false
    );
    SELECTED_WORKSHEET.getRange("D5").copyFrom(
      messagesAccepted,
      ExcelScript.RangeCopyType.all,
      false,
      false
    );
  }

  function createChart() {
    const CHART_TITLE: string = `LinkedIn source - ${
      MONTH_NAME[TODAY.getMonth() - 1]
    }`;

    if (SELECTED_WORKSHEET.getChart(CHART_TITLE)) {
      SELECTED_WORKSHEET.getChart(CHART_TITLE).delete();
    }

    let chart = SELECTED_WORKSHEET.addChart(
      ExcelScript.ChartType.pie,
      SELECTED_WORKSHEET.getRange("A5:D6")
    );

    chart.setName(CHART_TITLE);
    chart.getTitle().setText(CHART_TITLE);
    chart.getSeries()[0].setHasDataLabels(true);

    // set chart position
    chart.setLeft(50);
    chart.setTop(100);
  }

  splitInformations();
  copyInformations();
  createChart();

  console.log(
    "If you need support, email me at: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura"
  );
}
