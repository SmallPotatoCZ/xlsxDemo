import {
  saveAs
} from 'file-saver';
const XLSX = require('xlsx');
const wb = {
  SheetNames: [],
  Sheets: {}
};
const ws_name = 'Report';
/* create worksheet: */
const ws: any = {};
/* the range object is used to keep track of the range of the sheet */
const range = {
  s: {
    c: 0,
    r: 0
  },
  e: {
    c: 0,
    r: 0
  }
};

/* Iterate through each element in the structure */
for (let R = 0; R < data.length; R++) {
  if (range.e.r < R) {
    range.e.r = R;
  }
  for (let C = 0; C < data[R].length; C++) {
    if (range.e.c < C) {
      range.e.c = C;
    }
    // tslint:disable-next-line:max-line-length
    const cell = {
      v: data[R][C],
      s: {
        alignment: {
          textRotation: 90
        },
        font: {
          sz: 16,
          bold: true,
          color: '#FF00FF'
        },
        fill: {
          bgColor: '#FFFFFF'
        }
      },
      t: 's'
    };
    if (cell.v == null) {
      continue;
    }
    /* create the correct cell reference */
    const cell_ref = XLSX.utils.encode_cell({
      c: C,
      r: R
    });
    /* determine the cell type */
    if (typeof cell.v === 'number') {
      cell.t = 'n';
    } else if (typeof cell.v === 'boolean') {
      cell.t = 'b';
    } else {
      cell.t = 's';
    }
    /* add to structure */
    ws[cell_ref] = cell;
  }
}
ws['!ref'] = XLSX.utils.encode_range(range);
const wscols = [{
    wch: 4
  }, // "characters"
  {
    wch: 3
  },
  {
    wch: 20
  },
  {
    wch: 45
  },
  {
    wch: 13
  },
  {
    wch: 30
  },
  {
    wch: 15
  },
  {
    wch: 5
  },
  {
    wch: 20
  },
  {
    wch: 3
  },
  {
    wch: 15
  },
  {
    wch: 15
  },
  {
    wch: 12
  },
  {
    wch: 75
  }
];

ws['!cols'] = wscols;
wb.SheetNames.push(ws_name);
/**
 * Set worksheet sheet to "narrow".
 */
ws['!margins'] = {
  left: 0.25,
  right: 0.25,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3
};
wb.Sheets[ws_name] = ws;
const wbout = XLSX.write(wb, {
  type: 'binary',
  bookSST: true,
  bookType: 'xlsx',
  cellStyles: true
});
saveAs(new Blob([this.s2ab(wbout)], {
  type: 'application/octet-stream'
}), 'report.xlsx');