const XlsxPopulate = require('xlsx-populate');

XlsxPopulate.fromBlankAsync()
  .then(workbook => {
    workbook.sheet("Sheet1").cell("A1").value("This is neat!").style({
      bold: true,
      italic: true,
      wrapText: true
    });

    return workbook.toFileAsync("./tmp/out.xlsx");
  });