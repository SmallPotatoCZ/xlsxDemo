const XlsxPopulate = require('xlsx-populate');
const xlsx = require('xlsx');

// 测试数据
const json = {
  "测试数据": [{
    "111": "ssssssssssssssssssss\nyyyyyyyyyyyy",
    "测试字段": "衡阳车务段\n，增加的内容"
  }, {
    "111": "ssssssssssssssssssss\nyyyyyyyyyyyy",
    "测试字段": "衡阳车务段,,,,,,\n增加的内容"
  }]
};

function charLength(content) {
  var length = 0;
  var _strArr = content.split('\n');
  for (let i = 0; i < _strArr.length; i++) {
    _length = _strArr[i].length + _strArr[i].replace(/[\x00-\xff]/g, '').length;
    if (_length > length) {
      length = _length
    }
  }
  return length;
}

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
      mw: charLength(value)
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
      if (charLength(_v[key]) > scolNames[key].mw) {
        scolNames[key].mw = charLength(_v[key]);
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

let _result = generateWB(json);

xlsx.writeFile(_result, './tmp/out.xlsx', {
  type: 'binary',
  bookSST: true,
  bookType: 'xlsx'
});

XlsxPopulate.fromFileAsync('./tmp/out.xlsx').then(wb => {
  _result.SheetNames.forEach(v => {
    wb.sheet(v).range(_result.Sheets[v]['!ref']).style("wrapText", true);
  })

  return wb.toFileAsync("./tmp/out.xlsx");
})