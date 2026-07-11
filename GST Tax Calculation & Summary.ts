function main(workbook: ExcelScript.Workbook) {
  const accFormat: string = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-';
  const months: string[] = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"];

  const inputColor = "#FFF2CC";
  const calcColor = "#EAF3F8";
  const goodColor = "#E2F0D9";
  const warnColor = "#F4CCCC";

  function getSheet(name: string): ExcelScript.Worksheet {
    let ws: ExcelScript.Worksheet | undefined = workbook.getWorksheet(name);
    if (!ws) ws = workbook.addWorksheet(name);
    return ws;
  }

  function clearSheet(ws: ExcelScript.Worksheet): void {
    let used: ExcelScript.Range | undefined = ws.getUsedRange();
    if (used) used.clear();

    let charts: ExcelScript.Chart[] = ws.getCharts();
    for (let i = 0; i < charts.length; i++) {
      charts[i].delete();
    }

    ws.setShowGridlines(false);
  }

  function title(ws: ExcelScript.Worksheet, range: string, text: string, color: string): void {
    let r: ExcelScript.Range = ws.getRange(range);
    r.merge();
    r.setValue(text);
    r.getFormat().getFill().setColor(color);
    r.getFormat().getFont().setColor("white");
    r.getFormat().getFont().setBold(true);
    r.getFormat().getFont().setSize(18);
    r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }

  function header(ws: ExcelScript.Worksheet, range: string, color: string): void {
    let r: ExcelScript.Range = ws.getRange(range);
    r.getFormat().getFill().setColor(color);
    r.getFormat().getFont().setColor("white");
    r.getFormat().getFont().setBold(true);
    r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }

  function back(ws: ExcelScript.Worksheet, cell: string): void {
    let r: ExcelScript.Range = ws.getRange(cell);
    r.setFormula(`=HYPERLINK("#'Summary'!A1","⬅ Back to Summary")`);
    r.getFormat().getFill().setColor("#FFF2CC");
    r.getFormat().getFont().setColor("#0563C1");
    r.getFormat().getFont().setBold(true);
  }

  function border(rng: ExcelScript.Range): void {
    const borders: ExcelScript.BorderIndex[] = [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft,
      ExcelScript.BorderIndex.edgeRight,
      ExcelScript.BorderIndex.insideHorizontal,
      ExcelScript.BorderIndex.insideVertical
    ];

    for (let i = 0; i < borders.length; i++) {
      rng.getFormat().getRangeBorder(borders[i]).setStyle(ExcelScript.BorderLineStyle.continuous);
    }
  }

  function fit(ws: ExcelScript.Worksheet): void {
    ws.setShowGridlines(false);

    let used: ExcelScript.Range | undefined = ws.getUsedRange();

    if (used) {
      used.getFormat().getFont().setName("Aptos");
      used.getFormat().autofitColumns();
      used.getFormat().autofitRows();
      border(used);
    }

    ws.getRange("A:Z").getFormat().autofitColumns();
  }

  // ========================
  // CREATE ALL SHEETS FIRST
  // ========================

  let settings: ExcelScript.Worksheet = getSheet("Settings");
  let nav: ExcelScript.Worksheet = getSheet("Summary");
  let out: ExcelScript.Worksheet = getSheet("Output_Working");
  let amend: ExcelScript.Worksheet = getSheet("Output_Amendment");
  let g3bm: ExcelScript.Worksheet = getSheet("GSTR3B_Monthwise");
  let inp: ExcelScript.Worksheet = getSheet("Input_Working");
  let diff: ExcelScript.Worksheet = getSheet("Difference_Check");
  let sum: ExcelScript.Worksheet = getSheet("GST_3B_Summary");
  let dash: ExcelScript.Worksheet = getSheet("Dashboard");
  let annual: ExcelScript.Worksheet = getSheet("Annual_Summary");
  let print: ExcelScript.Worksheet = getSheet("Printable_Report");

  const allBuildSheets: ExcelScript.Worksheet[] = [
    settings, nav, out, amend, g3bm, inp, diff, sum, dash, annual, print
  ];

  for (let i = 0; i < allBuildSheets.length; i++) {
    clearSheet(allBuildSheets[i]);
  }

  // ========================
  // SETTINGS
  // ========================

  settings.getRange("A1:B1").setValues([["Setting", "Value"]]);
  settings.getRange("A2:B6").setValues([
    ["Client Name", "Client Name"],
    ["Financial Year", "2025-26"],
    ["State Code", "33"],
    ["Prepared By", "Your Name"],
    ["Remarks", "GST Working Template"]
  ]);

  header(settings, "A1:B1", "#1F4E78");
  settings.getRange("A2:A6").getFormat().getFill().setColor("#D9EAF7");
  settings.getRange("B2:B6").getFormat().getFill().setColor(inputColor);
  back(settings, "D1");

  // ========================
  // SUMMARY
  // ========================

  title(nav, "A1:D1", "GST WORKING NAVIGATION", "#1F4E78");

  nav.getRange("A3:B7").setValues([
    ["Client Name", "=Settings!B2"],
    ["Financial Year", "=Settings!B3"],
    ["State Code", "=Settings!B4"],
    ["Prepared By", "=Settings!B5"],
    ["Purpose", "Enter values only in yellow cells"]
  ]);

  nav.getRange("A3:A7").getFormat().getFill().setColor("#D9EAF7");
  nav.getRange("A3:A7").getFormat().getFont().setBold(true);
  nav.getRange("B3:B7").getFormat().getFill().setColor("#F8FBFD");

  nav.getRange("A9:C9").setValues([["No", "Sheet", "Purpose"]]);
  header(nav, "A9:C9", "#548235");

  const links: string[][] = [
    ["Output_Working", "Enter GSTR-1 outward values"],
    ["Output_Amendment", "Enter GSTR-1 amendment values"],
    ["GSTR3B_Monthwise", "Enter 3B values similar to office format"],
    ["Input_Working", "Enter 2B / ITC / RCM / reversal values"],
    ["Difference_Check", "Auto difference review"],
    ["GST_3B_Summary", "Auto GST 3B summary"],
    ["Dashboard", "Auto GST dashboard"],
    ["Annual_Summary", "Auto yearly summary"],
    ["Printable_Report", "Client/auditor printable report"],
    ["Settings", "Client details"]
  ];

  for (let i = 0; i < links.length; i++) {
    let r: number = i + 10;
    nav.getRange(`A${r}`).setValue(i + 1);
    nav.getRange(`B${r}`).setFormula(`=HYPERLINK("#'${links[i][0]}'!A1","${links[i][0]}")`);
    nav.getRange(`C${r}`).setValue(links[i][1]);
  }

  nav.getRange("A10:C25").getFormat().getFill().setColor("#F8FBFD");

  // ========================
  // GSTR3B MONTHWISE FIRST
  // ========================

  title(g3bm, "A1:M1", "AS PER 3B - MONTHWISE WORKING", "#1F4E78");
  back(g3bm, "O1");

  g3bm.getRange("A3:M3").setValues([[
    "Month",
    "B2B Taxable", "B2B IGST", "B2B CGST", "B2B SGST",
    "Total Output GST",
    "Export Taxable", "Export IGST",
    "RCM Taxable", "RCM IGST", "RCM CGST", "RCM SGST",
    "Status"
  ]]);

  header(g3bm, "A3:M3", "#C65911");

  for (let i = 0; i < months.length; i++) {
    let r: number = i + 4;
    g3bm.getRange(`A${r}`).setValue(months[i]);
    g3bm.getRange(`F${r}`).setFormula(`=SUM(C${r}:E${r})+H${r}+SUM(J${r}:L${r})`);
    g3bm.getRange(`M${r}`).setFormula(`=IF(COUNTA(B${r}:E${r},G${r}:L${r})=0,"No Data","Entered")`);
  }

  g3bm.getRange("A16").setValue("TOTAL");

  const g3bmCols: string[] = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];

  for (let i = 0; i < g3bmCols.length; i++) {
    let c: string = g3bmCols[i];
    g3bm.getRange(`${c}16`).setFormula(`=SUM(${c}4:${c}15)`);
  }

  g3bm.getRange("M16").setFormula(`=IF(COUNTA(B4:E15,G4:L15)=0,"No Data","Entered")`);

  g3bm.getRange("B4:E15").getFormat().getFill().setColor(inputColor);
  g3bm.getRange("G4:L15").getFormat().getFill().setColor(inputColor);
  g3bm.getRange("F4:F15").getFormat().getFill().setColor(calcColor);
  g3bm.getRange("M4:M15").getFormat().getFill().setColor(calcColor);
  g3bm.getRange("A16:M16").getFormat().getFill().setColor("#FFF2CC");
  g3bm.getRange("A16:M16").getFormat().getFont().setBold(true);
  g3bm.getRange("B4:L16").setNumberFormat(accFormat);

  // ========================
  // OUTPUT WORKING
  // ========================

  title(out, "A1:O1", "AS PER GSTR-1 - OUTPUT WORKING", "#1F4E78");
  back(out, "Q1");

  out.getRange("A3:O3").setValues([[
    "Month",
    "B2B Taxable", "B2B IGST", "B2B CGST", "B2B SGST",
    "Export Taxable", "Export IGST",
    "CDNR Taxable", "CDNR IGST", "CDNR CGST", "CDNR SGST",
    "GSTR-1 GST", "As per 3B GST", "Difference", "Status"
  ]]);

  header(out, "A3:O3", "#1F4E78");

  for (let i = 0; i < months.length; i++) {
    let r: number = i + 4;

    out.getRange(`A${r}`).setValue(months[i]);
    out.getRange(`L${r}`).setFormula(`=SUM(C${r}:E${r})+G${r}+SUM(I${r}:K${r})`);
    out.getRange(`M${r}`).setFormula(`='GSTR3B_Monthwise'!F${r}`);
    out.getRange(`N${r}`).setFormula(`=L${r}-M${r}`);
    out.getRange(`O${r}`).setFormula(`=IF(COUNTA(B${r}:K${r})=0,"No Data",IF(ABS(N${r})<=1,"Clean","Need Check"))`);
  }

  out.getRange("A16").setValue("TOTAL");

  const outCols: string[] = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"];

  for (let i = 0; i < outCols.length; i++) {
    let c: string = outCols[i];
    out.getRange(`${c}16`).setFormula(`=SUM(${c}4:${c}15)`);
  }

  out.getRange("O16").setFormula(`=IF(COUNTA(B4:K15)=0,"No Data",IF(ABS(N16)<=1,"Clean","Need Check"))`);

  out.getRange("B4:K15").getFormat().getFill().setColor(inputColor);
  out.getRange("L4:O15").getFormat().getFill().setColor(calcColor);
  out.getRange("A16:O16").getFormat().getFill().setColor("#FFF2CC");
  out.getRange("A16:O16").getFormat().getFont().setBold(true);
  out.getRange("B4:N16").setNumberFormat(accFormat);

  // ========================
  // OUTPUT AMENDMENT
  // ========================

  title(amend, "A1:K1", "AS PER GSTR-1 AMENDMENT", "#7030A0");
  back(amend, "M1");

  amend.getRange("A3:K3").setValues([[
    "Month",
    "B2B Amend Taxable", "B2B Amend IGST", "B2B Amend CGST", "B2B Amend SGST",
    "CDNR Amend Taxable", "CDNR Amend IGST", "CDNR Amend CGST", "CDNR Amend SGST",
    "Net Amendment GST", "Remarks"
  ]]);

  header(amend, "A3:K3", "#7030A0");

  for (let i = 0; i < months.length; i++) {
    let r: number = i + 4;
    amend.getRange(`A${r}`).setValue(months[i]);
    amend.getRange(`J${r}`).setFormula(`=SUM(C${r}:E${r})+SUM(G${r}:I${r})`);
  }

  amend.getRange("A16").setValue("TOTAL");

  const amendCols: string[] = ["B", "C", "D", "E", "F", "G", "H", "I", "J"];

  for (let i = 0; i < amendCols.length; i++) {
    let c: string = amendCols[i];
    amend.getRange(`${c}16`).setFormula(`=SUM(${c}4:${c}15)`);
  }

  amend.getRange("B4:I15").getFormat().getFill().setColor(inputColor);
  amend.getRange("J4:J15").getFormat().getFill().setColor(calcColor);
  amend.getRange("K4:K15").getFormat().getFill().setColor(inputColor);
  amend.getRange("A16:K16").getFormat().getFill().setColor("#FFF2CC");
  amend.getRange("A16:K16").getFormat().getFont().setBold(true);
  amend.getRange("B4:J16").setNumberFormat(accFormat);

  // ========================
  // INPUT WORKING
  // ========================

  title(inp, "A1:Q1", "AS PER 2B / ITC WORKING", "#548235");
  back(inp, "S1");

  inp.getRange("A3:Q3").setValues([[
    "Month",
    "2B IGST", "2B CGST", "2B SGST",
    "IMPG IGST",
    "RCM IGST", "RCM CGST", "RCM SGST",
    "All Other ITC IGST", "All Other ITC CGST", "All Other ITC SGST",
    "ITC Reversed IGST", "ITC Reversed CGST", "ITC Reversed SGST",
    "As per 3B ITC", "2B vs 3B Difference", "Status"
  ]]);

  header(inp, "A3:Q3", "#548235");

  for (let i = 0; i < months.length; i++) {
    let r: number = i + 4;
    inp.getRange(`A${r}`).setValue(months[i]);
    inp.getRange(`P${r}`).setFormula(`=(SUM(B${r}:K${r})-SUM(L${r}:N${r}))-O${r}`);
    inp.getRange(`Q${r}`).setFormula(`=IF(COUNTA(B${r}:O${r})=0,"No Data",IF(ABS(P${r})<=1,"Clean","Need Check"))`);
  }

  inp.getRange("A16").setValue("TOTAL");

  const inpCols: string[] = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"];

  for (let i = 0; i < inpCols.length; i++) {
    let c: string = inpCols[i];
    inp.getRange(`${c}16`).setFormula(`=SUM(${c}4:${c}15)`);
  }

  inp.getRange("Q16").setFormula(`=IF(COUNTA(B4:O15)=0,"No Data",IF(ABS(P16)<=1,"Clean","Need Check"))`);

  inp.getRange("B4:O15").getFormat().getFill().setColor(inputColor);
  inp.getRange("P4:Q15").getFormat().getFill().setColor(calcColor);
  inp.getRange("A16:Q16").getFormat().getFill().setColor("#FFF2CC");
  inp.getRange("A16:Q16").getFormat().getFont().setBold(true);
  inp.getRange("B4:P16").setNumberFormat(accFormat);

  // ========================
  // DIFFERENCE CHECK
  // ========================

  title(diff, "A1:H1", "GST DIFFERENCE CHECK", "#C00000");
  back(diff, "J1");

  diff.getRange("A3:H3").setValues([[
    "Month", "GSTR-1 GST", "3B Output GST", "GSTR-1 vs 3B Diff",
    "2B ITC", "3B ITC", "2B vs 3B Diff", "Status"
  ]]);

  header(diff, "A3:H3", "#C00000");

  for (let i = 0; i < months.length; i++) {
    let r: number = i + 4;

    diff.getRange(`A${r}`).setValue(months[i]);
    diff.getRange(`B${r}`).setFormula(`=Output_Working!L${r}+Output_Amendment!J${r}`);
    diff.getRange(`C${r}`).setFormula(`='GSTR3B_Monthwise'!F${r}`);
    diff.getRange(`D${r}`).setFormula(`=B${r}-C${r}`);
    diff.getRange(`E${r}`).setFormula(`=SUM(Input_Working!B${r}:K${r})-SUM(Input_Working!L${r}:N${r})`);
    diff.getRange(`F${r}`).setFormula(`=Input_Working!O${r}`);
    diff.getRange(`G${r}`).setFormula(`=E${r}-F${r}`);
    diff.getRange(`H${r}`).setFormula(`=IF(COUNTA(Output_Working!B${r}:K${r},'GSTR3B_Monthwise'!B${r}:L${r},Input_Working!B${r}:O${r})=0,"No Data",IF(ABS(D${r})+ABS(G${r})<=1,"Clean","Need Check"))`);
  }

  diff.getRange("A16").setValue("TOTAL");

  const diffCols: string[] = ["B", "C", "D", "E", "F", "G"];

  for (let i = 0; i < diffCols.length; i++) {
    let c: string = diffCols[i];
    diff.getRange(`${c}16`).setFormula(`=SUM(${c}4:${c}15)`);
  }

  diff.getRange("H16").setFormula(`=IF(COUNTIF(H4:H15,"Need Check")>0,"Need Check",IF(COUNTIF(H4:H15,"Clean")>0,"Clean","No Data"))`);

  diff.getRange("A4:H15").getFormat().getFill().setColor(warnColor);
  diff.getRange("A16:H16").getFormat().getFill().setColor("#FFF2CC");
  diff.getRange("A16:H16").getFormat().getFont().setBold(true);
  diff.getRange("B4:G16").setNumberFormat(accFormat);

  // ========================
  // GST 3B SUMMARY
  // ========================

  title(sum, "A1:F1", "GST 3B SUMMARY", "#1F4E78");
  back(sum, "H1");

  sum.getRange("A3:F3").setValues([["Particulars", "Taxable", "IGST", "CGST", "SGST", "Total GST"]]);
  header(sum, "A3:F3", "#C65911");

  sum.getRange("A4:A10").setValues([
    ["As per GSTR-1"],
    ["GSTR-1 Amendment"],
    ["As per 3B Output"],
    ["GSTR-1 vs 3B Difference"],
    ["As per 2B ITC"],
    ["As per 3B ITC"],
    ["2B vs 3B Difference"]
  ]);

  sum.getRange("B4").setFormula("=Output_Working!B16+Output_Working!F16+Output_Working!H16");
  sum.getRange("C4").setFormula("=Output_Working!C16+Output_Working!G16+Output_Working!I16");
  sum.getRange("D4").setFormula("=Output_Working!D16+Output_Working!J16");
  sum.getRange("E4").setFormula("=Output_Working!E16+Output_Working!K16");
  sum.getRange("F4").setFormula("=Output_Working!L16");

  sum.getRange("B5").setFormula("=Output_Amendment!B16+Output_Amendment!F16");
  sum.getRange("C5").setFormula("=Output_Amendment!C16+Output_Amendment!G16");
  sum.getRange("D5").setFormula("=Output_Amendment!D16+Output_Amendment!H16");
  sum.getRange("E5").setFormula("=Output_Amendment!E16+Output_Amendment!I16");
  sum.getRange("F5").setFormula("=Output_Amendment!J16");

  sum.getRange("B6").setFormula("='GSTR3B_Monthwise'!B16+'GSTR3B_Monthwise'!G16+'GSTR3B_Monthwise'!I16");
  sum.getRange("C6").setFormula("='GSTR3B_Monthwise'!C16+'GSTR3B_Monthwise'!H16+'GSTR3B_Monthwise'!J16");
  sum.getRange("D6").setFormula("='GSTR3B_Monthwise'!D16+'GSTR3B_Monthwise'!K16");
  sum.getRange("E6").setFormula("='GSTR3B_Monthwise'!E16+'GSTR3B_Monthwise'!L16");
  sum.getRange("F6").setFormula("='GSTR3B_Monthwise'!F16");

  sum.getRange("F7").setFormula("=F4+F5-F6");

  sum.getRange("C8").setFormula("=Input_Working!B16+Input_Working!E16+Input_Working!F16+Input_Working!I16-Input_Working!L16");
  sum.getRange("D8").setFormula("=Input_Working!C16+Input_Working!G16+Input_Working!J16-Input_Working!M16");
  sum.getRange("E8").setFormula("=Input_Working!D16+Input_Working!H16+Input_Working!K16-Input_Working!N16");
  sum.getRange("F8").setFormula("=SUM(C8:E8)");

  sum.getRange("F9").setFormula("=Input_Working!O16");
  sum.getRange("F10").setFormula("=F8-F9");

  sum.getRange("A4:F10").getFormat().getFill().setColor("#F8FBFD");
  sum.getRange("B4:F10").setNumberFormat(accFormat);

  // ========================
  // DASHBOARD
  // ========================

  title(dash, "A1:H1", "GST WORKING DASHBOARD", "#1F4E78");
  back(dash, "J1");

  dash.getRange("A3:B4").setValues([["As per 3B Output", ""], ["", ""]]);
  dash.getRange("D3:E4").setValues([["As per 3B ITC", ""], ["", ""]]);
  dash.getRange("G3:H4").setValues([["Net Payable", ""], ["", ""]]);

  dash.getRange("A4").setFormula("=GST_3B_Summary!F6");
  dash.getRange("D4").setFormula("=GST_3B_Summary!F9");
  dash.getRange("G4").setFormula("=A4-D4");

  dash.getRange("A3:B4").getFormat().getFill().setColor("#D9EAF7");
  dash.getRange("D3:E4").getFormat().getFill().setColor(goodColor);
  dash.getRange("G3:H4").getFormat().getFill().setColor("#FFF2CC");
  dash.getRange("A3:H4").getFormat().getFont().setBold(true);
  dash.getRange("A4").setNumberFormat(accFormat);
  dash.getRange("D4").setNumberFormat(accFormat);
  dash.getRange("G4").setNumberFormat(accFormat);

  dash.getRange("A7:B11").setValues([
    ["Review", "Status"],
    ["GSTR-1 vs 3B Difference", ""],
    ["2B vs 3B Difference", ""],
    ["Months Need Check", ""],
    ["Overall Status", ""]
  ]);

  header(dash, "A7:B7", "#C00000");

  dash.getRange("B8").setFormula("=GST_3B_Summary!F7");
  dash.getRange("B9").setFormula("=GST_3B_Summary!F10");
  dash.getRange("B10").setFormula(`=COUNTIF(Difference_Check!H4:H15,"Need Check")`);
  dash.getRange("B11").setFormula(`=IF(B10>0,"Need Check",IF(COUNTIF(Difference_Check!H4:H15,"Clean")>0,"Clean","No Data"))`);

  dash.getRange("B8:B9").setNumberFormat(accFormat);
  dash.getRange("A8:B11").getFormat().getFill().setColor(warnColor);

  // ========================
  // ANNUAL SUMMARY
  // ========================

  title(annual, "A1:F1", "ANNUAL GST SUMMARY", "#1F4E78");
  back(annual, "H1");

  annual.getRange("A3:F3").setValues([["Particulars", "Taxable", "IGST", "CGST", "SGST", "Total GST"]]);
  header(annual, "A3:F3", "#548235");

  annual.getRange("A4:A9").setValues([
    ["As per GSTR-1"],
    ["GSTR-1 Amendment"],
    ["As per 3B Output"],
    ["As per 2B ITC"],
    ["As per 3B ITC"],
    ["Net Payable"]
  ]);

  for (let r = 4; r <= 8; r++) {
    const cols: string[] = ["B", "C", "D", "E", "F"];
    for (let i = 0; i < cols.length; i++) {
      let c: string = cols[i];
      annual.getRange(`${c}${r}`).setFormula(`=GST_3B_Summary!${c}${r}`);
    }
  }

  annual.getRange("F9").setFormula("=F6-F8");

  annual.getRange("A4:F9").getFormat().getFill().setColor(goodColor);
  annual.getRange("B4:F9").setNumberFormat(accFormat);

  // ========================
  // PRINTABLE REPORT
  // ========================

  title(print, "A1:F1", "GST WORKING PRINTABLE REPORT", "#1F4E78");
  back(print, "H1");

  print.getRange("A3:B7").setValues([
    ["Client Name", "=Settings!B2"],
    ["Financial Year", "=Settings!B3"],
    ["State Code", "=Settings!B4"],
    ["Prepared By", "=Settings!B5"],
    ["Report Status", "=Dashboard!B11"]
  ]);

  header(print, "A3:A7", "#7030A0");
  print.getRange("B3:B7").getFormat().getFill().setColor("#EFE4F8");

  print.getRange("D3:E8").setValues([
    ["Particulars", "Value"],
    ["Output as per 3B", ""],
    ["ITC as per 3B", ""],
    ["Net Payable", ""],
    ["GSTR-1 vs 3B Diff", ""],
    ["2B vs 3B Diff", ""]
  ]);

  header(print, "D3:E3", "#548235");

  print.getRange("E4").setFormula("=GST_3B_Summary!F6");
  print.getRange("E5").setFormula("=GST_3B_Summary!F9");
  print.getRange("E6").setFormula("=E4-E5");
  print.getRange("E7").setFormula("=GST_3B_Summary!F7");
  print.getRange("E8").setFormula("=GST_3B_Summary!F10");

  print.getRange("D4:E8").getFormat().getFill().setColor(goodColor);
  print.getRange("E4:E8").setNumberFormat(accFormat);

  // ========================
  // FINAL
  // ========================

  nav.setPosition(0);

  const allSheets: ExcelScript.Worksheet[] = workbook.getWorksheets();

  for (let i = 0; i < allSheets.length; i++) {
    fit(allSheets[i]);
  }

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}