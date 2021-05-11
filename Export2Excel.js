/* eslint-disable */
require('script-loader!file-saver');
// require('script-loader!../../assets/js/Blob'); 测试不记载,也没什么问题
// require('script-loader!xlsx/dist/xlsx.core.min');
/* eslint-disable */
import XLSX from 'xlsx-style'

// 配置参数文档  https://github.com/pikaz-18/pikaz-excel-js

function generateArray(table) {
  var out = [];
  var rows = table.querySelectorAll('tr');
  var ranges = [];
  for (var R = 0; R < rows.length; ++R) {
    var outRow = [];
    var row = rows[R];
    var columns = row.querySelectorAll('td');
    for (var C = 0; C < columns.length; ++C) {
      var cell = columns[C];
      var colspan = cell.getAttribute('colspan');
      var rowspan = cell.getAttribute('rowspan');
      var cellValue = cell.innerText;
      if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

      //Skip ranges
      ranges.forEach(function (range) {
        if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
          for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
        }
      });

      //Handle Row Span
      if (rowspan || colspan) {
        rowspan = rowspan || 1;
        colspan = colspan || 1;
        ranges.push({
          s: {
            r: R,
            c: outRow.length
          },
          e: {
            r: R + rowspan - 1,
            c: outRow.length + colspan - 1
          }
        });
      }
      ;

      //Handle Value
      outRow.push(cellValue !== "" ? cellValue : null);

      //Handle Colspan
      if (colspan)
        for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
    }
    out.push(outRow);
  }
  return [out, ranges];
};

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };
  for (var R = 0; R != data.length; ++R) {
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      var cell = {
        v: data[R][C]
      };
      if (cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

function converter(value) {
  // A-Z计数器转换器
  const type = typeof value;
  let finaly = null;
  switch (type) {
    case "number":
      finaly = "";
      let divisor = Math.floor(value / 26),
        remainder = value % 26 ? [value % 26] : value <= 26 ? [value] : [];
      while (divisor > 26) {
        divisor = Math.floor(divisor / 26);
        remainder.unshift(divisor % 26);
      }
      value > 26 && remainder.unshift(divisor);
      for (let val of remainder) {
        finaly += String.fromCharCode(val + 64);
      }
      break;
    case "string":
      finaly = 0;
      const length = value.length;
      for (let len = 0; len < length; len++) {
        finaly +=
          (value.charAt(len).charCodeAt() - 64) *
          Math.pow(26, length - len - 1);
      }
      break;
  }
  return finaly;
}


export function export_table_title_custom(data,name) {  //id选择器第一行灰色
  var ws_name = "SheetJS";
  var wb = new Workbook();
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = data;
  var dataInfo = wb.Sheets[wb.SheetNames[0]];

  dataInfo['!merges'] && addRangeBorder(dataInfo['!merges'],dataInfo)//定义全局边框

  function addRangeBorder(range,ws){
    /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
    let arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"];

    range.forEach(item=>{
      let startRowNumber = Number(item.s.r),
        endRowNumber = Number(item.e.r),
        startColNumber = Number(item.s.c),
        endColNumber = Number(item.e.c);
      for(let i = startRowNumber;i<= endRowNumber+1;i++){
        for(let j = startColNumber;j<= endColNumber;j++){
          if(!ws[arr[j]+i]) ws[arr[j]+i] = {}
          ws[arr[j]+i].s = {border:{top:{style:'thin'}, left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}};
        }
      }
    })
    
    return ws;
  }

  
  const borderAll = {  //单元格外侧框线
    top: {
      style: 'thin'
    },
    bottom: {
      style: 'thin'
    },
    left: {
      style: 'thin'
    },
    right: {
      style: 'thin'
    }
  };
  //给所以单元格加上边框
  for (var i in dataInfo) {
    if (i == '!ref' || i == '!merges' || i == '!cols' || i == 'A1') {

    } else {
      dataInfo[i + ''].s = {
        border: borderAll,
        alignment: {
          horizontal: "center",
          vertical: "center"
        }
      }
    }
  }
  
  // console.log('AA'.charCodeAt(),String.fromCharCode(65))
  /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
  let map_arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"];  
  // let numberRang = data['!ref'].match(/[A-Z]+/g)[1].charCodeAt() - data['!ref'].match(/[A-Z]+/g)[0].charCodeAt() +1
  let indexRang = map_arr.indexOf(data['!ref'].match(/[A-Z]+/g)[1])
  
  for(let i=0;i<=indexRang;i++){
    //设置副主标题样式
    dataInfo[map_arr[i]+'1'].s = {
      font: {
        // name: '宋体',
        sz: 14,
        // color: {rgb: "ff0000"},
        bold: true,
        italic: false,
        underline: false
      },
      fill: {
        fgColor: {rgb: "cccccc"},
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
    };
  }
 

  dataInfo["!cols"] = [  //单元格列宽
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    },
    {
        wpx: 170
    }
  ];
  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  });

  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), name)
}



export function export_table_to_excel_custom(data,name) {   //id选择器第一行黄色,第二行灰色
  var ws_name = "SheetJS";
  var wb = new Workbook();
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = data;
  var dataInfo = wb.Sheets[wb.SheetNames[0]];


  addRangeBorder(dataInfo['!merges'],dataInfo)//定义全局边框

  function addRangeBorder(range,ws){
    /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
    let arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"];

    range.forEach(item=>{
      let startRowNumber = Number(item.s.r),
        endRowNumber = Number(item.e.r),
        startColNumber = Number(item.s.c),
        endColNumber = Number(item.e.c);
      for(let i = startRowNumber;i<= endRowNumber+1;i++){
        for(let j = startColNumber;j<= endColNumber;j++){
          if(!ws[arr[j]+i]) ws[arr[j]+i] = {}
          ws[arr[j]+i].s = {border:{top:{style:'thin'}, left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}};
        }
      }
    })
    
    return ws;
  }

  
  const borderAll = {  //单元格外侧框线
    top: {
      style: 'thin'
    },
    bottom: {
      style: 'thin'
    },
    left: {
      style: 'thin'
    },
    right: {
      style: 'thin'
    }
  };
  //给所以单元格加上边框
  for (var i in dataInfo) {
    if (i == '!ref' || i == '!merges' || i == '!cols' || i == 'A1') {

    } else {
      dataInfo[i + ''].s = {
        border: borderAll,
        alignment: {
          horizontal: "center",
          vertical: "center"
        }
      }
    }
  }


  //设置主标题样式
  dataInfo["A1"].s = {
    font: {
      name: '宋体',
      sz: 18,
      // color: {rgb: "ff0000"},
      bold: true,
      italic: false,
      underline: false
    },
    alignment: {
      horizontal: "center",
      vertical: "center"
    },
    fill: {
      fgColor: {rgb: "FCFF40"},
    }
  };

  
  // console.log('AA'.charCodeAt(),String.fromCharCode(65))
  /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
  let map_arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"];  
  // let numberRang = data['!ref'].match(/[A-Z]+/g)[1].charCodeAt() - data['!ref'].match(/[A-Z]+/g)[0].charCodeAt() +1
  let indexRang = map_arr.indexOf(data['!ref'].match(/[A-Z]+/g)[1])
  
  for(let i=0;i<=indexRang;i++){
    //设置副主标题样式
    dataInfo[map_arr[i]+'2'].s = {
      font: {
        // name: '宋体',
        sz: 14,
        // color: {rgb: "ff0000"},
        bold: true,
        italic: false,
        underline: false
      },
      fill: {
        fgColor: {rgb: "cccccc"},
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
    };
  }
 

  dataInfo["!cols"] = [  //单元格列宽
    {
        wpx: 170
    }, {
        wpx: 130
    }, {
        wpx: 170
    }
  ];
  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  });

  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), name)
}

export function export_table_to_excel(id) {     //id选择器无样式表格
  var theTable = document.getElementById(id);
  var oo = generateArray(theTable);
  var ranges = oo[1];

  /* original data */
  var data = oo[0];
  var ws_name = "SheetJS";

  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data);

  /* add ranges to worksheet */
  // ws['!cols'] = ['apple', 'banan'];
  ws['!merges'] = ranges;

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;

  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  });

  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), "test.xlsx")
}

export function export_json_to_excel(th, jsonData, defaultTitle) {  //json数据无样式表格

  /* original data */

  var data = jsonData;
  data.unshift(th);
  var ws_name = "SheetJS";

  var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);


  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;

  var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: false, type: 'binary'});
  var title = defaultTitle || '列表'
  saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), title + ".xlsx")
}

export function export_json_to_excel_custom({   //json数据设置主题样式
   title,
   multiHeader = [],
   header,
   data,
   filename,
   merges = [],
   autoWidth = true,
   bookType = 'xlsx'
  } = {}) {
  /* original data */
  filename = filename || 'excel-list'
  data = [...data]
  data.unshift(header);
  data.unshift(title);
  for (let i = multiHeader.length - 1; i > -1; i--) {
    data.unshift(multiHeader[i])
  }

  var ws_name = "SheetJS";
  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data);

  if (merges.length > 0) {
    if (!ws['!merges']) ws['!merges'] = [];
    merges.forEach(item => {
      ws['!merges'].push(XLSX.utils.decode_range(item))
    })
  }

  if (autoWidth) {
    /*设置worksheet每列的最大宽度*/
    const colWidth = data.map(row => row.map(val => {
      /*先判断是否为null/undefined*/
      if (val == null) {
        return {
          'wch': 10
        };
      }
      /*再判断是否为中文*/
      else if (val.toString().charCodeAt(0) > 255) {
        return {
          'wch': val.toString().length * 2
        };
      } else {
        return {
          'wch': val.toString().length
        };
      }
    }))
    /*以第一行为初始值*/
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['wch'] < colWidth[i][j]['wch']) {
          result[j]['wch'] = colWidth[i][j]['wch'];
        }
      }
    }
    ws['!cols'] = result;
  }

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;
  var dataInfo = wb.Sheets[wb.SheetNames[0]];

  const borderAll = {  //单元格外侧框线
    top: {
      style: 'thin'
    },
    bottom: {
      style: 'thin'
    },
    left: {
      style: 'thin'
    },
    right: {
      style: 'thin'
    }
  };
  //给所以单元格加上边框
  for (var i in dataInfo) {
    if (i == '!ref' || i == '!merges' || i == '!cols' || i == 'A1') {

    } else {
      dataInfo[i + ''].s = {
        border: borderAll
      }
    }
  }

  // 去掉标题边框
  /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
  let arr = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1", "U1", "V1", "W1", "X1", "Y1", "Z1"];
  arr.some(function (v) {
    let a = merges[0].split(':')
    if (v == a[1]) {
      dataInfo[v].s = {}
      return true;
    } else {
      dataInfo[v].s = {}
    }
  })

  //设置主标题样式
  dataInfo["A1"].s = {
    font: {
      name: '宋体',
      sz: 18,
      color: {rgb: "ff0000"},
      bold: true,
      italic: false,
      underline: false
    },
    alignment: {
      horizontal: "center",
      vertical: "center"
    },
    // fill: {
    //   fgColor: {rgb: "008000"},
    // },
  };

  // console.log(merges)
  // console.log(dataInfo)


  var wbout = XLSX.write(wb, {
    bookType: bookType,
    bookSST: false,
    type: 'binary'
  });
  saveAs(new Blob([s2ab(wbout)], {
    type: "application/octet-stream"
  }), `${filename}.${bookType}`);
}




export function export_json_title_custom({    //json数据设置第一行黄色第二行灰色
  title,
  multiHeader = [],
  header,
  data,
  filename,
  merges = [],
  autoWidth = true,
  bookType = 'xlsx'
 } = {}) {
 /* original data */
 filename = filename || 'excel-list'
 data = [...data]
 data.unshift(header);
 data.unshift(title);
 for (let i = multiHeader.length - 1; i > -1; i--) {
   data.unshift(multiHeader[i])
 }

 var ws_name = "SheetJS";
 var wb = new Workbook(),
   ws = sheet_from_array_of_arrays(data);

 if (merges.length > 0) {
   if (!ws['!merges']) ws['!merges'] = [];
   merges.forEach(item => {
     ws['!merges'].push(XLSX.utils.decode_range(item))
   })
 }


 if (autoWidth) {
   /*设置worksheet每列的最大宽度*/
   data.shift(title)
   const colWidth = data.map(row => row.map(val => {
     /*先判断是否为null/undefined*/
     if (val == null) {
       return {
         'wch': 10
       };
     }
     /*再判断是否为中文*/
     else if (val.toString().charCodeAt(0) > 255) {
       return {
         'wch': val.toString().length * 2
       };
     } else {
       return {
         'wch': val.toString().length
       };
     }
   }))
   /*以第一行为初始值*/
   let result = colWidth[0];
   for (let i = 1; i < colWidth.length; i++) {
     for (let j = 0; j < colWidth[i].length; j++) {
       if (result[j]['wch'] < colWidth[i][j]['wch']) {
         result[j]['wch'] = colWidth[i][j]['wch'];
       }
     }
   }
   ws['!cols'] = result;
 }

 /* add worksheet to workbook */
 wb.SheetNames.push(ws_name);
 wb.Sheets[ws_name] = ws;
 var dataInfo = wb.Sheets[wb.SheetNames[0]];

 const borderAll = {  //单元格外侧框线
   top: {
     style: 'thin'
   },
   bottom: {
     style: 'thin'
   },
   left: {
     style: 'thin'
   },
   right: {
     style: 'thin'
   }
 };
 //给所以单元格加上边框
 for (var i in dataInfo) {
   if (i == '!ref' || i == '!merges' || i == '!cols' || i == 'A1') {

   } else {
     dataInfo[i + ''].s = {
       border: borderAll,
       alignment: {
          horizontal: "center",
          vertical: "center"
        }
     }
   }
 }

 // 去掉标题边框
 /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
 let arr = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1", "U1", "V1", "W1", "X1", "Y1", "Z1"];
 arr.some(function (v) {
   let a = merges[0].split(':')
   if (v == a[1]) {
     dataInfo[v].s = {}
     return true;
   } else {
     dataInfo[v].s = {}
   }
 })

 //设置主标题样式
 dataInfo["A1"].s = {
   font: {
     name: '宋体',
     sz: 18,
    //  color: {rgb: "ff0000"},
     bold: true,
     italic: false,
     underline: false
   },
   alignment: {
     horizontal: "center",
     vertical: "center"
   },
   fill: {
      fgColor: {rgb: "FCFF40"},
   },
 };

   // console.log('AA'.charCodeAt(),String.fromCharCode(65))
  /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
  let map_arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"];  
  // let numberRang = data['!ref'].match(/[A-Z]+/g)[1].charCodeAt() - data['!ref'].match(/[A-Z]+/g)[0].charCodeAt() +1
  let indexRang = map_arr.indexOf(dataInfo['!ref'].match(/[A-Z]+/g)[1])
  
  for(let i=0;i<=indexRang;i++){
    //设置副主标题样式
    dataInfo[map_arr[i]+'2'].s = {
      font: {
        // name: '宋体',
        sz: 14,
        // color: {rgb: "ff0000"},
        bold: true,
        italic: false,
        underline: false
      },
      fill: {
        fgColor: {rgb: "cccccc"},
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
    };
  }
  
  // dataInfo["!cols"] = [  //单元格列宽
  //   {
  //       wpx: 70
  //   }
  // ];


 var wbout = XLSX.write(wb, {
   bookType: bookType,
   bookSST: false,
   type: 'binary'
 });
 saveAs(new Blob([s2ab(wbout)], {
   type: "application/octet-stream"
 }), `${filename}.${bookType}`);
}

export function export_json_top_custom({    //json数据设置第一行为灰色
  title,
  multiHeader = [],
  header,
  data,
  filename,
  merges = [],
  autoWidth = true,
  bookType = 'xlsx'
 } = {}) {
 /* original data */
 filename = filename || 'excel-list'
 data = [...data]
 data.unshift(header);
 data.unshift(title);
 for (let i = multiHeader.length - 1; i > -1; i--) {
   data.unshift(multiHeader[i])
 }

 var ws_name = "SheetJS";
 var wb = new Workbook(),
   ws = sheet_from_array_of_arrays(data);

 if (merges.length > 0) {
   if (!ws['!merges']) ws['!merges'] = [];
   merges.forEach(item => {
     ws['!merges'].push(XLSX.utils.decode_range(item))
   })
 }


 if (autoWidth) {
   /*设置worksheet每列的最大宽度*/
   data.shift(title)
   const colWidth = data.map(row => row.map(val => {
     /*先判断是否为null/undefined*/
     if (val == null) {
       return {
         'wch': 10
       };
     }
     /*再判断是否为中文*/
     else if (val.toString().charCodeAt(0) > 255) {
       return {
         'wch': val.toString().length * 2
       };
     } else {
       return {
         'wch': val.toString().length
       };
     }
   }))
   /*以第一行为初始值*/
   let result = colWidth[0];
   for (let i = 1; i < colWidth.length; i++) {
     for (let j = 0; j < colWidth[i].length; j++) {
       if (result[j]['wch'] < colWidth[i][j]['wch']) {
         result[j]['wch'] = colWidth[i][j]['wch'];
       }
     }
   }
   ws['!cols'] = result;
 }

 /* add worksheet to workbook */
 wb.SheetNames.push(ws_name);
 wb.Sheets[ws_name] = ws;
 var dataInfo = wb.Sheets[wb.SheetNames[0]];
 console.log(dataInfo)

 const borderAll = {  //单元格外侧框线
   top: {
     style: 'thin'
   },
   bottom: {
     style: 'thin'
   },
   left: {
     style: 'thin'
   },
   right: {
     style: 'thin'
   }
 };
 //给所以单元格加上边框
 for (var i in dataInfo) {
   if (i == '!ref' || i == '!merges' || i == '!cols' || i == 'A1') {

   } else {
     dataInfo[i + ''].s = {
       border: borderAll,
       alignment: {
          horizontal: "center",
          vertical: "center"
        }
     }
   }
 }

 // 去掉标题边框
 /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
 let arr = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1", "U1", "V1", "W1", "X1", "Y1", "Z1"];
 arr.some(function (v) {
   let a = merges[0].split(':')
   if (v == a[1]) {
     dataInfo[v].s = {}
     return true;
   } else {
     dataInfo[v].s = {}
   }
 })


   // console.log('AA'.charCodeAt(),String.fromCharCode(65))
  /* 写死的数组不能灵活适应26进制表格数据,对于本项目问题不大 */
  let map_arr = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"];  
  // let numberRang = data['!ref'].match(/[A-Z]+/g)[1].charCodeAt() - data['!ref'].match(/[A-Z]+/g)[0].charCodeAt() +1
  let indexRang = map_arr.indexOf(dataInfo['!ref'].match(/[A-Z]+/g)[1])
  
  for(let i=0;i<=indexRang;i++){
    //设置副主标题样式
    dataInfo[map_arr[i]+'1'].s = {
      font: {
        // name: '宋体',
        sz: 14,
        // color: {rgb: "ff0000"},
        bold: true,
        italic: false,
        underline: false
      },
      fill: {
        fgColor: {rgb: "cccccc"},
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
    };
  }
  
  // dataInfo["!cols"] = [  //单元格列宽
  //   {
  //       wpx: 70
  //   }
  // ];


 var wbout = XLSX.write(wb, {
   bookType: bookType,
   bookSST: false,
   type: 'binary'
 });
 saveAs(new Blob([s2ab(wbout)], {
   type: "application/octet-stream"
 }), `${filename}.${bookType}`);
}


// 通用配置函数
export function export_json_common_custom({
  title,
  multiHeader = [],
  header,
  data,
  filename,
  cellStyle,
  merges = [],
  autoWidth = true,
  bookType = "xlsx"
} = {}) {
  /* original data */
  filename = filename || "excel-list";
  data = [...data];
  data.unshift(header);
  data.unshift(title);
  for (let i = multiHeader.length - 1; i > -1; i--) {
    data.unshift(multiHeader[i]);
  }

  var ws_name = "SheetJS";
  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data);

  if (merges.length > 0) {
    if (!ws["!merges"]) ws["!merges"] = [];
    merges.forEach(item => {
      ws["!merges"].push(XLSX.utils.decode_range(item));
    });
  }

  if (autoWidth) {
    /*设置worksheet每列的最大宽度*/
    data.shift(title);
    const colWidth = data.map(row =>
      row.map(val => {
        /*先判断是否为null/undefined*/
        if (val == null) {
          return {
            wch: 10
          };
        } else if (val.toString().charCodeAt(0) > 255) {
          /*再判断是否为中文*/
          return {
            wch: val.toString().length * 2 + 5
          };
        } else {
          return {
            wch: val.toString().length + 5
          };
        }
      })
    );
    /*以第一行为初始值*/
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]["wch"] < colWidth[i][j]["wch"]) {
          result[j]["wch"] = colWidth[i][j]["wch"];
        }
      }
    }
    ws["!cols"] = result;
  }

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;
  var dataInfo = wb.Sheets[wb.SheetNames[0]];

  const borderAll = {
    //单元格外侧框线
    top: {
      style: "thin"
    },
    bottom: {
      style: "thin"
    },
    left: {
      style: "thin"
    },
    right: {
      style: "thin"
    }
  };
  //给所以单元格加上边框
  for (var i in dataInfo) {
    if (i == "!ref" || i == "!merges" || i == "!cols" || i == "A1") {
    } else {
      dataInfo[i + ""].s = {
        border: borderAll,
        alignment: {
          horizontal: "center",
          vertical: "center"
        }
      };
    }
  }

  merges.map(mer => {
    const a = mer.split(":"),
      col = [];
    let row = null;
    for (let i of a) {
      col.push(i.match(/^[A-Z]+/gi)[0]);
      row = i.match(/\d+$/gi)[0];
    }
    const before = col[0].charCodeAt() - 64;
    const length = col[1].charCodeAt() - col[0].charCodeAt();
    for (let len = 0; len <= length; len++) {
      dataInfo[converter(before + len) + row].s = {};
    }
  });

  cellStyle.map(cell => {
    cell.range.map(mer => {
      const a = mer.split(":"),
        col = [];
      let row = null;
      for (let i of a) {
        col.push(i.match(/^[A-Z]+/gi)[0]);
        row = i.match(/\d+$/gi)[0];
      }
      console.log(col)
      switch(col.length) {
        case 1:
          dataInfo[mer].s = cell.style;
          break;
        case 2:
          const before = col[0].charCodeAt() - 64;
          const length = col[1].charCodeAt() - col[0].charCodeAt();
          for (let len = 0; len <= length; len++) {
            dataInfo[converter(before + len) + row].s = cell.style;
          }
          break;
      }
      
    });
  })

  var wbout = XLSX.write(wb, {
    bookType: bookType,
    bookSST: false,
    type: "binary"
  });
  saveAs(
    new Blob([s2ab(wbout)], {
      type: "application/octet-stream"
    }),
    `${filename}.${bookType}`
  );
}

