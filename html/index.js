// 测试数据
const json = {
  "sheet1": [{
      "列1": "第1行单元格内容dddddddddddddddddddd",
      "列2": "第1行单元格内容",
      "列3": "第1行单元格内容",
      "列4": "第1行单元格内容"
    },
    {
      "列1": "第2行单元格内容",
      "列2": "第2行单元格内容",
      "列3": "第2行单元格内容",
      "列4": "第2行单元格内容"
    },
    {
      "列1": "第3行单元格内容",
      "列2": "第3行单元格内容",
      "列3": "第3行单元格内容",
      "列4": "第3行单元格内容"
    }
  ],
  "sheet2": [{
      "列1": "第1行单元格内容",
      "列2": "第1行单元格内容",
      "列3": "第1行单元格内容",
      "列4": "第1行单元格内容"
    },
    {
      "列1": "第2行单元格内容",
      "列2": "第2行单元格内容",
      "列3": "第2行单元格内容",
      "列4": "第2行单元格内容"
    },
    {
      "列1": "第3行单元格内容",
      "列2": "第3行单元格内容",
      "列3": "第3行单元格内容",
      "列4": "第3行单元格内容"
    }
  ]
};

function getCharCol(num) {
  let _res = '',
    tmp = 0,
    _sum = 0,
    n = 1,
    tar = 26;
  while (num >= _sum) {
    _sum = (Math.pow(tar, n + 1) - tar) / (tar - 1);
    n += 1;
    if (num >= _sum) {
      tmp = _sum;
    }
  }

  num -= tmp;
  while (num > 0) {
    tmp = num % tar;
    _res = String.fromCharCode(tmp + 65) + _res;
    num = (num - tmp) / tar;
  }

  while (n - 1 > _res.length) {
    _res = 'A' + _res;
  }


  return _res;
}

function generateWS(sheets) {
  let tmpdata = {},
    scol = [],
    scolNames = {},
    scolW = [];
  sheets.map((v) => {
    Object.keys(v).map(k => {
      let _res = scol.find(v => {
        return v === k
      });
      if (!_res) {
        scol.push(k);
      }
    });
  });
  scol.forEach((value, index) => {
    let _name = getCharCol(index)
    scolNames[value] = {
      name: _name,
      mw: 5
    };
    tmpdata[_name + 1] = {
      v: value
    };
  });
  sheets.forEach((_v, _i) => {
    for (var key in _v) {
      tmpdata[scolNames[key].name + (_i + 2)] = {
        v: _v[key]
      };
      if (_v[key].length * 2 > scolNames[key].mw) {
        scolNames[key].mw = _v[key].length * 2;
      }
    }
  })
  for (let _v in scolNames) {
    scolW.push({
      width: scolNames[_v].mw
    })
  }
  tmpdata['!cols'] = scolW;
  tmpdata['!ref'] = `A1:${getCharCol(scol.length)}${sheets.length + 1}`;
  return tmpdata;
}

function generateWB(json) {
  let sheetnames = [],
    sheets = {};
  Object.keys(json).map((v) => {
    sheetnames.push(v);
    sheets[v] = generateWS(json[v]);
  })

  return {
    "SheetNames": sheetnames,
    "Sheets": sheets
  };
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

function download() {
  let tmpDown = new Blob([s2ab(XLSX.write(generateWB(json), {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  }))], {
    type: ""
  });

  var href = URL.createObjectURL(tmpDown);
  var aEle = document.createElement('a');
  aEle.href = href;
  aEle.download = 'test.xlsx';
  aEle.click();
  setTimeout(function () {
    URL.revokeObjectURL(tmpDown);
  }, 100);
}




console.log('-------', generateWB(json));