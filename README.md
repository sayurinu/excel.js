Excel.js
========

Native node.js Excel file parser. Only supports xlsx for now.

Install
=======
    git clone https://github.com/shibucafe/excel.js.git excel

Use
====
(edited by @shibucafe)

    var parseXlsx = require('excel'); 

    parseXlsx('Spreadsheet.xlsx', function(err, data) {
      if(err) throw err;
      console.log(data);
      // [ { sheetNum: 1, , sheetName: 'hoge', contents: (array of arrays) }, ... ]
    });

If you have multiple sheets in your spreadsheet,

    parseXlsx('Spreadsheet.xlsx', [2, 3], function(err, data) {
      if (err) throw err; // if sheet2 or sheet3 does not exist, an error occurs
      console.log(data);
      // [ { sheetNum: 2, sheetName: 'hoge', contents: (array of arrays) }, { sheetNum: 3, sheetName: 'hoge', contents: (array of arrays) } ]
    });
    

MIT License.

*Author: Trevor Dixon <trevordixon@gmail.com>*

Contributors: 
- Jake Scott <scott.iroh@gmail.com>
- Fabian Tollenaar <fabian@startingpoint.nl> (Just a small contribution, really)
- amakhrov
