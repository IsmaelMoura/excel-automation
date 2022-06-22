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
  const PIVOT_TABLES_SHEET = workbook.getWorksheet("pivot");
  let reasonTable = PIVOT_TABLES_SHEET.getTable(REASON_TABLE_NAME);
  let directorTable = PIVOT_TABLES_SHEET.getTable(DIRECTOR_TABLE_NAME);

  function createReasonChart() {
    const CHART_TITLE = `Filled per Reason - ${
      MONTH_NAME[TODAY.getMonth() - 1]
    }`;

    const REASON_TABLE_VALUES = reasonTable
      .getRangeBetweenHeaderAndTotal()
      .getUsedRange();

    if (PIVOT_TABLES_SHEET.getChart(CHART_TITLE)) {
      PIVOT_TABLES_SHEET.getChart(CHART_TITLE).delete();
    }

    if (REASON_TABLE_VALUES === undefined) {
      throw new RangeError(
        "There is no data in the table filled per reason. Paste your information and run the script again"
      );
    }

    let reasonChart = PIVOT_TABLES_SHEET.addChart(
      ExcelScript.ChartType.columnClustered,
      reasonTable.getRangeBetweenHeaderAndTotal()
    );

    reasonChart.setName(CHART_TITLE);
    reasonChart.getTitle().setText(CHART_TITLE);
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
    const CHART_TITLE: string = `Filled per Director - ${
      MONTH_NAME[TODAY.getMonth() - 1]
    }`;

    const DIRECTOR_TABLE_VALUES = directorTable
      .getRangeBetweenHeaderAndTotal()
      .getUsedRange();

    if (PIVOT_TABLES_SHEET.getChart(CHART_TITLE)) {
      PIVOT_TABLES_SHEET.getChart(CHART_TITLE).delete();
    }

    if (DIRECTOR_TABLE_VALUES === undefined) {
      throw new RangeError(
        "There is no data in the table filled per director. Paste your information and run the script again"
      );
    }

    let directorChart = PIVOT_TABLES_SHEET.addChart(
      ExcelScript.ChartType.columnClustered,
      directorTable.getRangeBetweenHeaderAndTotal()
    );

    directorChart.setName(CHART_TITLE);
    directorChart.getTitle().setText(CHART_TITLE);
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

  function createReasonTable() {
    const REASON_TABLE_HEADERS = ["Reason", "Count Of Code"];

    REASON_TABLE_HEADERS.map((text, index) => {
      PIVOT_TABLES_SHEET.getCell(0, 6 + index).setValue(text);
    });

    reasonTable = PIVOT_TABLES_SHEET.addTable("G1:H1", true);
    reasonTable.setName(REASON_TABLE_NAME);
  }

  function createDirectorTable() {
    const DIRECTOR_TABLE_HEADERS = ["Director", "Count Of Code"];

    DIRECTOR_TABLE_HEADERS.map((text, index) => {
      PIVOT_TABLES_SHEET.getCell(0, 11 + index).setValue(text);
    });

    directorTable = PIVOT_TABLES_SHEET.addTable("L1:M1", true);
    directorTable.setName(DIRECTOR_TABLE_NAME);
  }

  if (!!reasonTable === false) {
    createReasonTable();
  }

  if (!!directorTable === false) {
    createDirectorTable();
  }

  createReasonChart();
  createDirectorChart();

  console.log(
    "If you need support, send an email to: ismael.moura@sinch.com or send a message in Microsoft Teams to: Ismael de Sousa Paulino Moura."
  );
}
