function main(workbook: ExcelScript.Workbook) {
  const PIVOT_SHEET = workbook.getWorksheet("pivot");
  const REASON_TABLE = PIVOT_SHEET.getTable("Reason_Table");
  const DIRECTOR_TABLE = PIVOT_SHEET.getTable("Director_Table");
  let today = new Date();
  let monthName = [
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

  function createReasonChart() {
    let chartTitle = `Filled per Reason - ${monthName[today.getMonth() - 1]}`;
    if (PIVOT_SHEET.getChart(chartTitle)) {
      PIVOT_SHEET.getChart(chartTitle).delete();
    } else if (
      REASON_TABLE.getRangeBetweenHeaderAndTotal().getUsedRange() === undefined
    ) {
      throw new Error(
        "There is no data filled per reason table. Paste your informations and run the script again"
      );
    }

    let reasonChart = PIVOT_SHEET.addChart(
      ExcelScript.ChartType.columnClustered,
      REASON_TABLE.getRangeBetweenHeaderAndTotal()
    );

    reasonChart.setName(chartTitle);
    reasonChart.getTitle().setText(chartTitle);
    reasonChart.getSeries()[0].setHasDataLabels(true);
    reasonChart.getAxes().getValueAxis().getMajorGridlines().setVisible(false);
    reasonChart.getAxes().getValueAxis().getMinorGridlines().setVisible(false);
    reasonChart.getAxes().getValueAxis().setVisible(false);
    reasonChart.getLegend().setVisible(false);

    // set chart position
    reasonChart.setLeft(450);
    reasonChart.setTop(70);
  }

  function createDirectorChart() {
    let chartTitle: string = `Filled per Director - ${
      monthName[today.getMonth() - 1]
    }`;

    if (PIVOT_SHEET.getChart(chartTitle)) {
      PIVOT_SHEET.getChart(chartTitle).delete();
    } else if (
      DIRECTOR_TABLE.getRangeBetweenHeaderAndTotal().getUsedRange() ===
      undefined
    ) {
      throw new Error(
        "There is no data filled per director table. Paste your informations and run the script again"
      );
    }

    let reasonChart = PIVOT_SHEET.addChart(
      ExcelScript.ChartType.columnClustered,
      DIRECTOR_TABLE.getRangeBetweenHeaderAndTotal()
    );

    reasonChart.setName(chartTitle);
    reasonChart.getTitle().setText(chartTitle);
    reasonChart.getSeries()[0].setHasDataLabels(true);
    reasonChart.getAxes().getValueAxis().getMajorGridlines().setVisible(false);
    reasonChart.getAxes().getValueAxis().getMinorGridlines().setVisible(false);
    reasonChart.getAxes().getValueAxis().setVisible(false);
    reasonChart.getLegend().setVisible(false);

    // set char position
    reasonChart.setLeft(850);
    reasonChart.setTop(70);
  }

  try {
    createReasonChart();
    createDirectorChart();
  } catch (e) {
    throw e;
  } finally {
    console.log(
      "If you need support, email me at: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura"
    );
  }
}
