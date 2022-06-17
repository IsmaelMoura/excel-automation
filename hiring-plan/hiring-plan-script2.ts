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

  const REASON_TABLE_NAME = "reason_table";
  const DIRECTOR_TABLE_NAME = "director_table";
  let pivotSheet = workbook.getWorksheet("pivot");
  let reasonTable = pivotSheet.getTable(REASON_TABLE_NAME);
  let directorTable = pivotSheet.getTable(DIRECTOR_TABLE_NAME);

  function createReasonChart() {
    let chartTitle = `Filled per Reason - ${MONTH_NAME[TODAY.getMonth() - 1]}`;

    if (!!reasonTable === false) {
      let reasonTableHeaders = ["Reason", "Count Of Code"];

      reasonTableHeaders.map((text, index) => {
        pivotSheet.getCell(0, 6 + index).setValue(text);
      });

      reasonTable = pivotSheet.addTable("G1:H1", true);
      reasonTable.setName(REASON_TABLE_NAME);
    }

    if (pivotSheet.getChart(chartTitle)) {
      pivotSheet.getChart(chartTitle).delete();
    }

    if (
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

    if (!!directorTable === false) {
      let directorTableHeaders = ["Director", "Count Of Code"];

      directorTableHeaders.map((text, index) => {
        pivotSheet.getCell(0, 11 + index).setValue(text);
      });

      directorTable = pivotSheet.addTable("L1:M1", true);
      directorTable.setName(DIRECTOR_TABLE_NAME);
    }

    if (pivotSheet.getChart(chartTitle)) {
      pivotSheet.getChart(chartTitle).delete();
    }

    if (
      directorTable.getRangeBetweenHeaderAndTotal().getUsedRange() === undefined
    ) {
      throw new Error(
        "There is no data in the table filled per director. Paste your information and run the script again"
      );
    }

    let directorChart = pivotSheet.addChart(
      ExcelScript.ChartType.columnClustered,
      directorTable.getRangeBetweenHeaderAndTotal()
    );

    directorChart.setName(chartTitle);
    directorChart.getTitle().setText(chartTitle);
    directorChart.getSeries()[0].setHasDataLabels(true);
    directorChart
      .getAxes()
      .getValueAxis()
      .getMajorGridlines()
      .setVisible(false);

    directorChart
      .getAxes()
      .getValueAxis()
      .getMinorGridlines()
      .setVisible(false);

    directorChart.getAxes().getValueAxis().setVisible(false);

    directorChart.getLegend().setVisible(false);

    // set char position
    directorChart.setLeft(850);
    directorChart.setTop(70);
  }

  createReasonChart();
  createDirectorChart();

  console.log(
    "If you need support, send an email to: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura."
  );
}
