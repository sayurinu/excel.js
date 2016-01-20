var _ = require('lodash'),
    async = require('async'),
    fs = require('fs'),
    JSZip = require('node-zip');

function extractFiles(path, sheets, callback) {
    var files = {
        strings: {},
        sheets: [],
        book: {}
    };

    fs.readFile(path, 'binary', function(err, data) {
        if (err || !data) {
            return callback(err || new Error('data not exists'));
        }

        try {
            var zip = new JSZip(data, { base64: false });
        } catch (e) {
            return callback(e);
        }

        // get contents by xml path
        function getContents_(path) {
            var raw = zip && zip.files && zip.files[path];
            return raw && (typeof raw.asText === 'function') && raw.asText();
        }

        var contents;
        contents = getContents_('xl/sharedStrings.xml');
        if (!contents) {
            return callback(new Error('xl/sharedStrings.xml not exists (maybe not xlsx file)'));
        }
        files.strings.contents = contents;

        contents = getContents_('xl/workbook.xml');
        if (!contents) {
            return callback(new Error('xl/workbook.xml not exists (maybe not xlsx file)'));
        }
        files.book.contents = contents;

        var sheetNum;
        if (sheets) {
            for (var i = 0; i < sheets.length; i++) {
                sheetNum = sheets[i];
                contents = getContents_('xl/worksheets/sheet' + sheetNum + '.xml');
                if (!contents) {
                    return callback(new Error('sheet ' + sheetNum + ' not exists'));
                }
                files.sheets.push({
                    sheetNum: sheetNum,
                    contents: contents
                });
            }
        } else {
            // push contents to sheets array
            function sheetsPush_(contents) {
                files.sheets.push({
                    sheetNum: files.sheets.length + 1,
                    contents: contents
                });
            }

            // for google spreadsheet
            var firstSheetContents = getContents_('xl/worksheets/sheet.xml');
            if (firstSheetContents) {
                sheetsPush_(firstSheetContents);
            }

            sheetNum = 1;
            while (true) {
                contents = getContents_('xl/worksheets/sheet' + sheetNum + '.xml');
                if (!contents) break;
                sheetsPush_(contents);
                sheetNum++;
            }
        }
        callback(null, files);
    });
}

function calculateDimensions (cells) {
    var comparator = function (a, b) { return a-b; };
    var allRows = _.map(cells, function (cell) { return cell.row; }).sort(comparator),
        allCols = _.map(cells, function (cell) { return cell.column; }).sort(comparator),
        minRow = allRows[0],
        maxRow = _.last(allRows),
        minCol = allCols[0],
        maxCol = _.last(allCols);

    return [
        {row: minRow, column: minCol},
        {row: maxRow, column: maxCol}
    ];
}

function extractData(files, callback) {
    var libxmljs = require('libxmljs');
    var sheets;
    try {
        var strings = libxmljs.parseXml(files.strings.contents),
            book = libxmljs.parseXml(files.book.contents),
            //styles = libxmljs.parseXml(files.styles.contents),
            ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};

        var b = book.find('//a:sheets//a:sheet', ns);
        var sheetNames = _.map(b, function(tag) {
            return tag.attr('name').value();
        });

        //sheets and sheetNames were retained the arrangement.
        sheets = _.map(files.sheets, function(sheetObj) {
            return {
                sheetNum: sheetObj.sheetNum,
                sheetName: sheetNames[sheetObj.sheetNum - 1],
                xml: libxmljs.parseXml(sheetObj.contents)
            };
        });
    } catch (e) {
        return callback(e);
    }

    var colToInt = function(col) {
        var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
        col = col.trim().split('');

        var n = 0;

        for (var i = 0; i < col.length; i++) {
            n *= 26;
            n += letters.indexOf(col[i]);
        }

        return n;
    };

    var CellCoords = function(cell) {
        cell = cell.split(/([0-9]+)/);
        this.row = parseInt(cell[1]);
        this.column = colToInt(cell[0]);
    };

    var na = { value: function() { return ''; },
        text:  function() { return ''; } };

    var Cell = function(cellNode) {
        var r = cellNode.attr('r').value(),
            type = (cellNode.attr('t') || na).value(),
            value = (cellNode.get('a:v', ns) || na ).text(),
            coords = new CellCoords(r);

        this.column = coords.column;
        this.row = coords.row;
        this.value = value;
        this.type = type;
    };

    async.mapSeries(sheets, function(sheetObj, next) {
        var sheet = sheetObj.xml;
        var cellNodes, cells, d;
        var onedata = [];

        async.series([
            function(_next) {
                cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);
                async.setImmediate(_next);
            },
            function(_next) {
                var count = 0;
                async.mapSeries(cellNodes, function(node, __next) {
                    // use setImmediate every 100 times
                    count = (count + 1) % 100;
                    if (count === 0) {
                        async.setImmediate(function() {
                            __next(null, new Cell(node));
                        });
                    } else {
                        __next(null, new Cell(node));
                    }
                }, function(err, results) {
                    cells = results;
                    _next();
                });
            },
            function(_next) {
                d = sheet.get('//a:dimension/@ref', ns);
                if (d) {
                    d = _.map(d.value().split(':'), function(v) { return new CellCoords(v); });
                } else {
                    d = calculateDimensions(cells);
                }
                async.setImmediate(_next);
            },
            function(_next) {
                var cols = d[1].column - d[0].column + 1,
                    rows = d[1].row - d[0].row + 1;
                _(rows).times(function() {
                    var _row = [];
                    _(cols).times(function() { _row.push(''); });
                    onedata.push(_row);
                });
                async.setImmediate(_next);
            },
            function(_next) {
                _(cells).each(function(cell) {
                    var value = cell.value;

                    if (cell.type == 's') {
                        var tmp = '';
                        _(strings.find('//a:si[' + (parseInt(value) + 1) + ']//a:t', ns)).each(function(t) {
                            if (t.get('..').name() != 'rPh') {
                                tmp += t.text();
                            }
                        });
                        value = tmp;
                    }

                    onedata[cell.row - d[0].row][cell.column - d[0].column] = value;

                });
                async.setImmediate(_next);
            }
        ], function() {
            next(null, {
                sheetNum: sheetObj.sheetNum,
                sheetName: sheetObj.sheetName,
                contents: onedata
            });
        });
    }, callback);
}

module.exports = function parseXlsx() {
    var path, sheets, callback;
    if (arguments.length == 2) {
        path = arguments[0];
        sheets = null;
        callback = arguments[1];
    }
    else if (arguments.length == 3) {
        path = arguments[0];
        sheets = arguments[1];
        if (typeof sheets === 'number') sheets = [sheets];
        callback = arguments[2];
    }
    extractFiles(path, sheets, function(err, files) {
        if (err) return callback(err);
        extractData(files, callback);
    });
};
