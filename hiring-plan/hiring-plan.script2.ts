function main(workbook: ExcelScript.Workbook) {
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

  let pivotSheet = workbook.getWorksheet("pivot");
  let reasonTable = pivotSheet.getTable("Reason_Table");
  let directorTable = pivotSheet.getTable("directorTable");

  function createReasonChart() {
    let chartTitle = `Filled per Reason - ${MONTH_NAME[TODAY.getMonth() - 1]}`;

    if (pivotSheet.getChart(chartTitle)) {
      pivotSheet.getChart(chartTitle).delete();
    } else if (
      reasonTable.getRangeBetweenHeaderAndTotal().getUsedRange() === undefined
    ) {
      throw new Error(
        "There is no data in the table filled per reason. Paste your information and run the script again"
      );
    }

    let reasonChart = pivotSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      reasonTable.getRangeBetweenHeaderAndTotal()
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
      MONTH_NAME[TODAY.getMonth() - 1]
    }`;

    if (pivotSheet.getChart(chartTitle)) {
      pivotSheet.getChart(chartTitle).delete();
    } else if (
      directorTable.getRangeBetweenHeaderAndTotal().getUsedRange() === undefined
    ) {
      throw new Error(
        "There is no data in the table filled per director. Paste your information and run the script again"
      );
    }

    let reasonChart = pivotSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      directorTable.getRangeBetweenHeaderAndTotal()
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

  createReasonChart();
  createDirectorChart();

  console.log(
    "If you need support, send an email to: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura."
  );
}
