const XLSX = require('xlsx');

const excelHandle = {
  _convert(num) {
    return num <= 26 ? String.fromCharCode(num + 64) : excelHandle._convert(parseInt((num - 1) / 26)) + excelHandle._convert(num % 26 || 26);
  },
  handleDownload(headList, list, excelName) {
    const headersLetters = {};
    const sheetData = {};

    headList.forEach((head, index) => {
      const colLetter = excelHandle._convert(index + 1);
      headersLetters[ head.prop ] = colLetter;
      sheetData[ colLetter + '1' ] = {
        v: head.label
      };
      if (index === 0) {
        sheetData[ '!ref' ] = colLetter + '1';
      }
    });

    list.forEach((item, index) => {
      const newIndex = index + 2;
      const headersLettersKeys = Object.keys(headersLetters);
      headersLettersKeys.forEach((key, keyIndex) => {
        sheetData[ headersLetters[ key ] + '' + newIndex ] = {
          v: item[ key ] || '-'
        };
        if ((list.length - 1) === index && (headersLettersKeys.length - 1) === keyIndex) {
          sheetData[ '!ref' ] += ':' + headersLetters[ key ] + '' + newIndex;
        }
      });
    });

    const wb = {
      SheetNames: [ 'sheet1' ],
      Sheets: {
        sheet1: sheetData
      }
    };
    XLSX.writeFile(wb, excelName + '.xlsx');
  }
};

module.exports = excelHandle.handleDownload;