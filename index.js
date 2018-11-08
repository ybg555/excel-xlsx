var XLSX = require('xlsx');

var excelHandle = {
  _convert(num) {
    return num <= 26 ? String.fromCharCode(num + 64) : excelHandle._convert(parseInt((num - 1) / 26)) + excelHandle._convert(num % 26 || 26);
  },
  handleDownload(headList, list, excelName) {
    var headersLetters = {};
    var sheetData = {};

    headList.forEach((head, index) => {
      var colLetter = excelHandle._convert(index + 1);
      headersLetters[ head.prop ] = colLetter;
      sheetData[ colLetter + '1' ] = {
        v: head.label
      };
      if (index === 0) {
        sheetData[ '!ref' ] = colLetter + '1';
      }
    });

    list.forEach((item, index) => {
      var newIndex = index + 2;
      var headersLettersKeys = Object.keys(headersLetters);
      headersLettersKeys.forEach((key, keyIndex) => {
        sheetData[ headersLetters[ key ] + '' + newIndex ] = {
          v: item[ key ] || '-'
        };
        if ((list.length - 1) === index && (headersLettersKeys.length - 1) === keyIndex) {
          sheetData[ '!ref' ] += ':' + headersLetters[ key ] + '' + newIndex;
        }
      });
    });

    var wb = {
      SheetNames: [ 'sheet1' ],
      Sheets: {
        sheet1: sheetData
      }
    };
    XLSX.writeFile(wb, excelName + '.xlsx');
  }
};

module.exports = excelHandle.handleDownload;