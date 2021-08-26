const fs = require('fs');
const xlsx = require('xlsx');

function find(type) {
    const path = './data';
    return new Promise ((resolve, reject) => {
        fs.readdir (path, function(e, files) {
          if (e) throw e;
          files.forEach((v) => {
              if (v.startsWith(type)) {
                  resolve(path + '/' + v);
              }
          });
        });
    });
}

function readXLSX(file) {
    return new Promise (function (resolve, reject) {
      const workbook = xlsx.readFile(file);
      if (!workbook) {
        reject(new Error('Unable to read xlsx file!'));
      } else {
        resolve (workbook);
      }
    });
}

function getDataFromWb(workbook) {
    const sheetList = workbook.SheetNames;
    return workbook.Sheets[sheetList[0]];
}

module.exports = { find, readXLSX, getDataFromWb }