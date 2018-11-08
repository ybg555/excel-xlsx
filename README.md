excel-xlsx
==========================

[![Travis](https://img.shields.io/travis/ybg555/excel-xlsx.svg)](https://travis-ci.org/ybg555/excel-xlsx) [![npm](https://img.shields.io/npm/dm/excel-xlsx.svg)](https://www.npmjs.com/package/excel-xlsx) [![npm](https://img.shields.io/npm/v/excel-xlsx.svg)](https://www.npmjs.com/package/excel-xlsx)

导出 xlsx 格式的 excel 文件


### Features

---

* 准备数据开箱即用
* 适合纯前端场景、数据量较小


### Installation

---

```shell
npm install excel-xlsx --save
```


### Usages

---

```js
const EXPORT_XLSX = require('excel-xlsx');

const headerData = [ { label: '日期', prop: 'time' }, { label: '金额', prop: 'money' } ];
const listData = [ { time: '2018', money: '234.21' }, { time: '2019', money: '1234.21' } ];
EXPORT_XLSX(headerData, listData, 'tableName');
```

### Issues

---

Submit the [issues](https://github.com/ybg555/excel-xlsx/issues) if you find any bug or have any suggestion.


### Contribution

---

Fork the [repository](https://github.com/ybg555/excel-xlsx) and submit pull requests.


### Release Notes

---

see [CHANGELOG](https://github.com/ybg555/excel-xlsx/blob/master/CHANGELOG.md)


### License

---

[![npm](https://img.shields.io/npm/l/excel-xlsx.svg)]()

