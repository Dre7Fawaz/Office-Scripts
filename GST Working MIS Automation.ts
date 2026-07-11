function main(workbook: ExcelScript.Workbook) {

  const mainSheetName: string = "GST_Data";
  const localStateCode: string = "33";
  const clientName: string = "Client Name";
  const financialYear: string = "2025-26";
  const accFormat: string = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-';

  let sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(mainSheetName);

  if (!sheet) {
    sheet = workbook.getActiveWorksheet();
    sheet.setName(mainSheetName);
  }

  let usedRange: ExcelScript.Range | undefined = sheet.getUsedRange();
  if (!usedRange) return;

  let lastRow: number = usedRange.getLastRow().getRowIndex() + 1;
  if (lastRow < 2) return;

  function clearSheet(name: string): ExcelScript.Worksheet {
    let ws: ExcelScript.Worksheet | undefined = workbook.getWorksheet(name);
    if (!ws) ws = workbook.addWorksheet(name);

    let rng: ExcelScript.Range | undefined = ws.getUsedRange();
    if (rng) rng.clear();

    let charts: ExcelScript.Chart[] = ws.getCharts();
    for (let i = 0; i < charts.length; i++) {
      charts[i].delete();
    }

    return ws;
  }

  function styleHeader(ws: ExcelScript.Worksheet, headerRange: string, color: string): void {
    let range: ExcelScript.Range = ws.getRange(headerRange);
    range.getFormat().getFont().setBold(true);
    range.getFormat().getFont().setColor("white");
    range.getFormat().getFill().setColor(color);
    range.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }

  function styleBody(ws: ExcelScript.Worksheet, bodyRange: string, color: string): void {
    let range: ExcelScript.Range = ws.getRange(bodyRange);
    range.getFormat().getFill().setColor(color);
  }

  function applyBorders(rng: ExcelScript.Range): void {
    rng.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
    rng.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
    rng.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
    rng.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
    rng.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
    rng.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
  }

  function finalFit(ws: ExcelScript.Worksheet): void {
    ws.setShowGridlines(false);

    let rng: ExcelScript.Range | undefined = ws.getUsedRange();

    if (rng) {
      rng.getFormat().getFont().setName("Aptos");
      rng.getFormat().autofitColumns();
      rng.getFormat().autofitRows();
      applyBorders(rng);
    }
  }

  // ========================
  // GST DATA
  // ========================

  sheet.getRange("G1:N1").setValues([[
    "CGST", "SGST", "IGST", "Invoice Value", "GST % Check", "Supply Type", "Month", "Month Name"
  ]]);

  styleHeader(sheet, "A1:N1", "#1F4E78");

  sheet.getRange("G2").setFormula(`=IF(TRIM(C2)="",F2*E2/200,IF(LEFT(SUBSTITUTE(C2," ",""),2)="${localStateCode}",F2*E2/200,0))`);
  sheet.getRange("H2").setFormula(`=IF(TRIM(C2)="",F2*E2/200,IF(LEFT(SUBSTITUTE(C2," ",""),2)="${localStateCode}",F2*E2/200,0))`);
  sheet.getRange("I2").setFormula(`=IF(TRIM(C2)="",0,IF(LEFT(SUBSTITUTE(C2," ",""),2)<>"${localStateCode}",F2*E2/100,0))`);
  sheet.getRange("J2").setFormula(`=F2+G2+H2+I2`);
  sheet.getRange("K2").setFormula(`=IF(F2=0,0,IF(I2>0,I2/F2*100,(G2+H2)/F2*100))`);
  sheet.getRange("L2").setFormula(`=IF(TRIM(C2)="","B2C","B2B")`);
  sheet.getRange("M2").setFormula(`=MONTH(D2)`);
  sheet.getRange("N2").setFormula(`=TEXT(D2,"mmm-yy")`);

  let helperCols: string[] = ["G", "H", "I", "J", "K", "L", "M", "N"];

  for (let i = 0; i < helperCols.length; i++) {
    let col: string = helperCols[i];
    sheet.getRange(`${col}2`).autoFill(`${col}2:${col}${lastRow}`, ExcelScript.AutoFillType.fillDefault);
  }

  sheet.getRange(`A2:N${lastRow}`).getFormat().getFill().setColor("#F8FBFD");
  sheet.getRange(`F2:J${lastRow}`).setNumberFormat(accFormat);

  let totalRow: number = lastRow + 1;

  sheet.getRange(`D${totalRow}`).setValue("TOTAL");
  sheet.getRange(`F${totalRow}`).setFormula(`=SUM(F2:F${lastRow})`);
  sheet.getRange(`G${totalRow}`).setFormula(`=SUM(G2:G${lastRow})`);
  sheet.getRange(`H${totalRow}`).setFormula(`=SUM(H2:H${lastRow})`);
  sheet.getRange(`I${totalRow}`).setFormula(`=SUM(I2:I${lastRow})`);
  sheet.getRange(`J${totalRow}`).setFormula(`=SUM(J2:J${lastRow})`);

  sheet.getRange(`D${totalRow}:J${totalRow}`).getFormat().getFont().setBold(true);
  sheet.getRange(`D${totalRow}:J${totalRow}`).getFormat().getFill().setColor("#FFF2CC");
  sheet.getRange(`F${totalRow}:J${totalRow}`).setNumberFormat(accFormat);

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  let dataValues: (string | number | boolean)[][] = sheet.getRange(`A2:N${lastRow}`).getValues();

  // ========================
  // B2B
  // ========================

  let b2b: ExcelScript.Worksheet = clearSheet("GSTR1_B2B");

  b2b.getRange("A1:J1").setValues([[
    "Invoice No", "Customer", "GSTIN", "Invoice Date", "Rate", "Taxable", "CGST", "SGST", "IGST", "Invoice Value"
  ]]);

  styleHeader(b2b, "A1:J1", "#1F4E78");

  b2b.getRange("A2").setFormula(
    `=FILTER(GST_Data!A2:J${lastRow},GST_Data!L2:L${lastRow}="B2B","No B2B Data")`
  );

  b2b.getRange("D:D").setNumberFormat("dd-mm-yyyy");
  b2b.getRange("F:J").setNumberFormat(accFormat);
  styleBody(b2b, "A2:J500", "#EAF3F8");

  // ========================
  // B2C SUMMARY
  // ========================

  let b2c: ExcelScript.Worksheet = clearSheet("GSTR1_B2C");

  b2c.getRange("A1:G1").setValues([[
    "Month", "Rate", "Taxable", "CGST", "SGST", "IGST", "Invoice Value"
  ]]);

  styleHeader(b2c, "A1:G1", "#548235");

  let b2cMap: { [key: string]: number[] } = {};

  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];

    let rateVal: number = Number(row[4]);
    let taxableVal: number = Number(row[5]);
    let cgstVal: number = Number(row[6]);
    let sgstVal: number = Number(row[7]);
    let igstVal: number = Number(row[8]);
    let invoiceVal: number = Number(row[9]);
    let supplyType: string = String(row[11] ?? "").trim();
    let monthName: string = String(row[13] ?? "").trim();

    if (supplyType === "B2C" && taxableVal !== 0) {
      let key: string = monthName + "|" + rateVal;

      if (!b2cMap[key]) b2cMap[key] = [0, 0, 0, 0, 0];

      b2cMap[key][0] += taxableVal;
      b2cMap[key][1] += cgstVal;
      b2cMap[key][2] += sgstVal;
      b2cMap[key][3] += igstVal;
      b2cMap[key][4] += invoiceVal;
    }
  }

  let b2cSummary: (string | number)[][] = [];

  for (let key in b2cMap) {
    let parts: string[] = key.split("|");
    let vals: number[] = b2cMap[key];

    b2cSummary.push([
      parts[0],
      Number(parts[1]),
      vals[0],
      vals[1],
      vals[2],
      vals[3],
      vals[4]
    ]);
  }

  if (b2cSummary.length > 0) {
    b2c.getRangeByIndexes(1, 0, b2cSummary.length, 7).setValues(b2cSummary);
    b2c.getRange(`C2:G${b2cSummary.length + 1}`).setNumberFormat(accFormat);
    b2c.getRange(`A2:G${b2cSummary.length + 1}`).getFormat().getFill().setColor("#E2F0D9");
  } else {
    b2c.getRange("A2").setValue("No B2C Data");
    b2c.getRange("A2").getFormat().getFill().setColor("#E2F0D9");
  }

  // ========================
  // RATE WISE SUMMARY
  // ========================

  let rate: ExcelScript.Worksheet = clearSheet("Rate_Wise_Summary");

  rate.getRange("A1:F1").setValues([[
    "Rate", "Taxable", "CGST", "SGST", "IGST", "Invoice Value"
  ]]);

  styleHeader(rate, "A1:F1", "#7030A0");

  let rateMap: { [key: string]: number[] } = {};

  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];

    let rateVal: number = Number(row[4]);
    let taxableVal: number = Number(row[5]);
    let cgstVal: number = Number(row[6]);
    let sgstVal: number = Number(row[7]);
    let igstVal: number = Number(row[8]);
    let invoiceVal: number = Number(row[9]);

    if (taxableVal !== 0) {
      let key: string = String(rateVal);

      if (!rateMap[key]) rateMap[key] = [0, 0, 0, 0, 0];

      rateMap[key][0] += taxableVal;
      rateMap[key][1] += cgstVal;
      rateMap[key][2] += sgstVal;
      rateMap[key][3] += igstVal;
      rateMap[key][4] += invoiceVal;
    }
  }

  let rateSummary: (string | number)[][] = [];

  for (let key in rateMap) {
    let vals: number[] = rateMap[key];

    rateSummary.push([
      Number(key),
      vals[0],
      vals[1],
      vals[2],
      vals[3],
      vals[4]
    ]);
  }

  if (rateSummary.length > 0) {
    rate.getRangeByIndexes(1, 0, rateSummary.length, 6).setValues(rateSummary);
    rate.getRange(`B2:F${rateSummary.length + 1}`).setNumberFormat(accFormat);
    rate.getRange(`A2:F${rateSummary.length + 1}`).getFormat().getFill().setColor("#EFE4F8");
  } else {
    rate.getRange("A2").setValue("No Data");
  }

  // ========================
  // TOP CUSTOMERS
  // ========================

  let top: ExcelScript.Worksheet = clearSheet("Top_Customers");

  top.getRange("A1:F1").setValues([[
    "Customer", "Taxable", "GST", "Invoice Value", "Invoice Count", "Share %"
  ]]);

  styleHeader(top, "A1:F1", "#C65911");

  let customerMap: { [key: string]: number[] } = {};

  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];

    let customer: string = String(row[1] ?? "").trim();
    let taxableVal: number = Number(row[5]);
    let cgstVal: number = Number(row[6]);
    let sgstVal: number = Number(row[7]);
    let igstVal: number = Number(row[8]);
    let invoiceVal: number = Number(row[9]);

    if (customer !== "" && taxableVal !== 0) {
      if (!customerMap[customer]) customerMap[customer] = [0, 0, 0, 0];

      customerMap[customer][0] += taxableVal;
      customerMap[customer][1] += cgstVal + sgstVal + igstVal;
      customerMap[customer][2] += invoiceVal;
      customerMap[customer][3] += 1;
    }
  }

  let customerSummary: (string | number)[][] = [];
  let totalTaxableForShare: number = 0;

  for (let cust in customerMap) {
    totalTaxableForShare += customerMap[cust][0];
  }

  for (let cust in customerMap) {
    let vals: number[] = customerMap[cust];
    let share: number = totalTaxableForShare === 0 ? 0 : vals[0] / totalTaxableForShare;

    customerSummary.push([
      cust,
      vals[0],
      vals[1],
      vals[2],
      vals[3],
      share
    ]);
  }

  customerSummary.sort((a, b) => Number(b[1]) - Number(a[1]));

  if (customerSummary.length > 0) {
    top.getRangeByIndexes(1, 0, customerSummary.length, 6).setValues(customerSummary);
    top.getRange(`B2:D${customerSummary.length + 1}`).setNumberFormat(accFormat);
    top.getRange(`F2:F${customerSummary.length + 1}`).setNumberFormat("0.00%");
    top.getRange(`A2:F${customerSummary.length + 1}`).getFormat().getFill().setColor("#FCE4D6");

    let chartRows: number = Math.min(customerSummary.length + 1, 11);
    let topChart: ExcelScript.Chart = top.addChart(ExcelScript.ChartType.barClustered, top.getRange(`A1:B${chartRows}`));
    topChart.setPosition("H2", "T35");
    topChart.getTitle().setText("Top Customers by Taxable Value");
  } else {
    top.getRange("A2").setValue("No Customer Data");
  }

  // ========================
  // ERROR CHECK
  // ========================

  let err: ExcelScript.Worksheet = clearSheet("Error_Check");

  err.getRange("A1:H1").setValues([[
    "Invoice No", "Customer", "GSTIN", "Invoice Date", "Rate", "Taxable", "Issue", "Supply Type"
  ]]);

  styleHeader(err, "A1:H1", "#C00000");

  let errorRows: (string | number | boolean)[][] = [];

  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];

    let invoiceNo = row[0];
    let customer = row[1];
    let gstin: string = String(row[2] ?? "").trim();
    let cleanGstin: string = gstin.replace(/\s/g, "");
    let invoiceDate = row[3];
    let rateVal = row[4];
    let taxableVal = row[5];
    let gstCheck: number = Number(row[10]);
    let supplyType: string = String(row[11] ?? "").trim();

    let issue: string = "";

    if (invoiceDate === "" || invoiceDate === null) {
      issue = "Missing Date";
    } else if (rateVal === "" || rateVal === null) {
      issue = "Missing GST Rate";
    } else if (taxableVal === "" || taxableVal === null) {
      issue = "Missing Taxable Value";
    } else if (cleanGstin !== "" && cleanGstin.length !== 15) {
      issue = "Invalid GSTIN Length";
    } else if (Math.round(gstCheck * 100) / 100 !== Number(rateVal)) {
      issue = "GST Rate Mismatch";
    }

    if (issue !== "") {
      errorRows.push([
        invoiceNo,
        customer,
        gstin,
        invoiceDate as string | number | boolean,
        rateVal as string | number | boolean,
        taxableVal as string | number | boolean,
        issue,
        supplyType
      ]);
    }
  }

  if (errorRows.length > 0) {
    err.getRangeByIndexes(1, 0, errorRows.length, 8).setValues(errorRows);
    err.getRange(`D2:D${errorRows.length + 1}`).setNumberFormat("dd-mm-yyyy");
    err.getRange(`F2:F${errorRows.length + 1}`).setNumberFormat(accFormat);
    err.getRange(`A2:H${errorRows.length + 1}`).getFormat().getFill().setColor("#F4CCCC");
  } else {
    err.getRange("A2").setValue("No Errors Found");
    err.getRange("A2").getFormat().getFill().setColor("#E2F0D9");
    err.getRange("A2").getFormat().getFont().setBold(true);
  }

  // ========================
  // DASHBOARD
  // ========================

  let dash: ExcelScript.Worksheet = clearSheet("Dashboard");

  dash.getRange("A1:H1").merge();
  dash.getRange("A1").setValue("GST SALES DASHBOARD");
  dash.getRange("A1").getFormat().getFont().setBold(true);
  dash.getRange("A1").getFormat().getFont().setSize(18);
  dash.getRange("A1").getFormat().getFill().setColor("#1F4E78");
  dash.getRange("A1").getFormat().getFont().setColor("white");
  dash.getRange("A1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  dash.getRange("A2:H2").merge();
  dash.getRange("A2").setValue(`Client: ${clientName} | FY: ${financialYear}`);
  dash.getRange("A2").getFormat().getFill().setColor("#D9EAF7");
  dash.getRange("A2").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  dash.getRange("A4:B5").setValues([["Taxable Value", ""], ["", ""]]);
  dash.getRange("D4:E5").setValues([["Total GST", ""], ["", ""]]);
  dash.getRange("G4:H5").setValues([["Invoice Value", ""], ["", ""]]);

  dash.getRange("A5").setFormula(`=SUM(GST_Data!F2:F${lastRow})`);
  dash.getRange("D5").setFormula(`=SUM(GST_Data!G2:G${lastRow})+SUM(GST_Data!H2:H${lastRow})+SUM(GST_Data!I2:I${lastRow})`);
  dash.getRange("G5").setFormula(`=SUM(GST_Data!J2:J${lastRow})`);

  dash.getRange("A4:B5").getFormat().getFill().setColor("#D9EAF7");
  dash.getRange("D4:E5").getFormat().getFill().setColor("#E2F0D9");
  dash.getRange("G4:H5").getFormat().getFill().setColor("#FFF2CC");

  dash.getRange("A4:H5").getFormat().getFont().setBold(true);
  dash.getRange("A5:H5").getFormat().getFont().setSize(14);
  dash.getRange("A5").setNumberFormat(accFormat);
  dash.getRange("D5").setNumberFormat(accFormat);
  dash.getRange("G5").setNumberFormat(accFormat);

  dash.getRange("A8:B12").setValues([
    ["Tax Type", "Amount"],
    ["CGST", ""],
    ["SGST", ""],
    ["IGST", ""],
    ["Total GST", ""]
  ]);

  dash.getRange("B9").setFormula(`=SUM(GST_Data!G2:G${lastRow})`);
  dash.getRange("B10").setFormula(`=SUM(GST_Data!H2:H${lastRow})`);
  dash.getRange("B11").setFormula(`=SUM(GST_Data!I2:I${lastRow})`);
  dash.getRange("B12").setFormula(`=SUM(B9:B11)`);
  dash.getRange("B9:B12").setNumberFormat(accFormat);

  dash.getRange("D8:E11").setValues([
    ["Supply Type", "Taxable"],
    ["B2B", ""],
    ["B2C", ""],
    ["Total", ""]
  ]);

  dash.getRange("E9").setFormula(`=SUMIFS(GST_Data!F:F,GST_Data!L:L,"B2B")`);
  dash.getRange("E10").setFormula(`=SUMIFS(GST_Data!F:F,GST_Data!L:L,"B2C")`);
  dash.getRange("E11").setFormula(`=SUM(E9:E10)`);
  dash.getRange("E9:E11").setNumberFormat(accFormat);

  dash.getRange("G8:H13").setValues([
    ["Check", "Status"],
    ["Error Count", ""],
    ["Result", ""],
    ["Total Invoices", ""],
    ["B2B Invoices", ""],
    ["B2C Invoices", ""]
  ]);

  dash.getRange("H9").setFormula(`=COUNTA(Error_Check!A2:A500)-COUNTIF(Error_Check!A2:A500,"No Errors Found")`);
  dash.getRange("H10").setFormula(`=IF(H9=0,"Clean","Need Check")`);
  dash.getRange("H11").setFormula(`=COUNTA(GST_Data!A2:A${lastRow})`);
  dash.getRange("H12").setFormula(`=COUNTIF(GST_Data!L2:L${lastRow},"B2B")`);
  dash.getRange("H13").setFormula(`=COUNTIF(GST_Data!L2:L${lastRow},"B2C")`);

  styleHeader(dash, "A8:B8", "#1F4E78");
  styleHeader(dash, "D8:E8", "#548235");
  styleHeader(dash, "G8:H8", "#C00000");

  dash.getRange("A9:B12").getFormat().getFill().setColor("#EAF3F8");
  dash.getRange("D9:E11").getFormat().getFill().setColor("#E2F0D9");
  dash.getRange("G9:H13").getFormat().getFill().setColor("#F4CCCC");

  let taxChart: ExcelScript.Chart = dash.addChart(ExcelScript.ChartType.doughnut, dash.getRange("A8:B11"));
  taxChart.setPosition("A15", "D30");
  taxChart.getTitle().setText("GST Breakup");

  let supplyChart: ExcelScript.Chart = dash.addChart(ExcelScript.ChartType.columnClustered, dash.getRange("D8:E10"));
  supplyChart.setPosition("E15", "H30");
  supplyChart.getTitle().setText("B2B vs B2C Taxable");

  dash.getRange("A32:E32").setValues([[
    "Highest Invoice Customer", "Invoice No", "Date", "Taxable", "Invoice Value"
  ]]);

  styleHeader(dash, "A32:E32", "#7030A0");

  dash.getRange("A33").setFormula(`=XLOOKUP(MAX(GST_Data!J2:J${lastRow}),GST_Data!J2:J${lastRow},GST_Data!B2:B${lastRow})`);
  dash.getRange("B33").setFormula(`=XLOOKUP(MAX(GST_Data!J2:J${lastRow}),GST_Data!J2:J${lastRow},GST_Data!A2:A${lastRow})`);
  dash.getRange("C33").setFormula(`=XLOOKUP(MAX(GST_Data!J2:J${lastRow}),GST_Data!J2:J${lastRow},GST_Data!D2:D${lastRow})`);
  dash.getRange("D33").setFormula(`=XLOOKUP(MAX(GST_Data!J2:J${lastRow}),GST_Data!J2:J${lastRow},GST_Data!F2:F${lastRow})`);
  dash.getRange("E33").setFormula(`=MAX(GST_Data!J2:J${lastRow})`);

  dash.getRange("A33:E33").getFormat().getFill().setColor("#EFE4F8");
  dash.getRange("C33").setNumberFormat("dd-mm-yyyy");
  dash.getRange("D33:E33").setNumberFormat(accFormat);

  // ========================
  // MONTHLY REPORT
  // ========================

  let monthly: ExcelScript.Worksheet = clearSheet("Monthly_Report");

  monthly.getRange("A1:F1").setValues([[
    "Month", "Taxable", "GST", "Invoice Value", "B2B Taxable", "B2C Taxable"
  ]]);

  styleHeader(monthly, "A1:F1", "#1F4E78");

  for (let m = 1; m <= 12; m++) {
    let r: number = m + 1;

    monthly.getRange(`A${r}`).setValue(m);
    monthly.getRange(`B${r}`).setFormula(`=SUMIF(GST_Data!M2:M${lastRow},${m},GST_Data!F2:F${lastRow})`);
    monthly.getRange(`C${r}`).setFormula(`=SUMIF(GST_Data!M2:M${lastRow},${m},GST_Data!G2:G${lastRow})+SUMIF(GST_Data!M2:M${lastRow},${m},GST_Data!H2:H${lastRow})+SUMIF(GST_Data!M2:M${lastRow},${m},GST_Data!I2:I${lastRow})`);
    monthly.getRange(`D${r}`).setFormula(`=SUMIF(GST_Data!M2:M${lastRow},${m},GST_Data!J2:J${lastRow})`);
    monthly.getRange(`E${r}`).setFormula(`=SUMIFS(GST_Data!F:F,GST_Data!M:M,${m},GST_Data!L:L,"B2B")`);
    monthly.getRange(`F${r}`).setFormula(`=SUMIFS(GST_Data!F:F,GST_Data!M:M,${m},GST_Data!L:L,"B2C")`);
  }

  monthly.getRange("A2:F13").getFormat().getFill().setColor("#EAF3F8");
  monthly.getRange("B2:F13").setNumberFormat(accFormat);

  let mchart: ExcelScript.Chart = monthly.addChart(ExcelScript.ChartType.columnClustered, monthly.getRange("A1:D13"));
  mchart.setPosition("H2", "P20");
  mchart.getTitle().setText("Monthly GST Sales Report");

  // ========================
  // FINAL CLEAN STYLE + AUTOFIT
  // ========================

  let allSheets: ExcelScript.Worksheet[] = workbook.getWorksheets();

  for (let i = 0; i < allSheets.length; i++) {
    finalFit(allSheets[i]);
  }
}