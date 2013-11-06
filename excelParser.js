var _ = require('underscore'),
    fs = require('fs'),
    JSZip = require('node-zip');

function extractFiles(path, sheets, callback) {
    var files = {
        strings: {},
        sheets: [],
        // styles: {},
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
        var raw, contents;
        raw = zip && zip.files && zip.files['xl/sharedStrings.xml'];
        contents = raw && (typeof raw.asText === 'function') && raw.asText();
        if (!contents) {
            return callback(new Error('xl/sharedStrings.xml not exists (maybe not xlsx file)'));
        }
        files.strings.contents = contents;

        // raw = zip && zip.files && zip.files['xl/styles.xml'];
        // contents = raw && (typeof raw.asText === 'function') && raw.asText();
        // if (!contents) {
        //     return callback(new Error('xl/styles.xml not exists (maybe not xlsx file)'));
        // }
        // files.styles.contents = contents;

        var sheetNum;
        if (sheets) {
            for (i = 0; i < sheets.length; i++) {
                sheetNum = sheets[i];
                raw = zip.files['xl/worksheets/sheet' + sheetNum + '.xml'];
                contents = raw && (typeof raw.asText === 'function') && raw.asText();
                if (!contents) {
                    return callback(new Error('sheet ' + sheetNum + ' not exists'));
                }
                files.sheets.push({
                    sheetNum: sheetNum,
                    contents: contents
                });
            }
        } else {
            sheetNum = 1;
            while (true) {
                raw = zip.files['xl/worksheets/sheet' + sheetNum + '.xml'];
                contents = raw && (typeof raw.asText === 'function') && raw.asText();
                if (!contents) break;
                files.sheets.push({
                    sheetNum: sheetNum,
                    contents: contents
                });
                sheetNum++;
            }
        }
        callback(null, files);
    });
}

function calculateDimensions (cells) {
    var comparator = function (a, b) { return a-b; };
    var allRows = _(cells).map(function (cell) { return cell.row; }).sort(comparator),
        allCols = _(cells).map(function (cell) { return cell.column; }).sort(comparator),
        minRow = allRows[0],
        maxRow = _.last(allRows),
        minCol = allCols[0],
        maxCol = _.last(allCols);

    return [
        {row: minRow, column: minCol},
        {row: maxRow, column: maxCol}
    ];
}

function extractData(files) {
    var libxmljs = require('libxmljs');
    try {
        var strings = libxmljs.parseXml(files.strings.contents),
            // styles = libxmljs.parseXml(files.styles.contents),
            ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'},
            data = [];
        var sheets = _(files.sheets).map(function(sheetObj) {
            return {
                sheetNum: sheetObj.sheetNum,
                xml: libxmljs.parseXml(sheetObj.contents)
            };
        });

    } catch (e) {
        return [];
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

    _(sheets).each(function(sheetObj) {
        var sheet = sheetObj.xml;
        var cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);
        var cells = _(cellNodes).map(function (node) {
            return new Cell(node);
        });
        var onedata = [];

        var d = sheet.get('//a:dimension/@ref', ns);
        if (d) {
            d = _.map(d.value().split(':'), function(v) { return new CellCoords(v); });
        } else {
            d = calculateDimensions(cells)
        }

        var cols = d[1].column - d[0].column + 1,
            rows = d[1].row - d[0].row + 1;

        _(rows).times(function() {
            var _row = [];
            _(cols).times(function() { _row.push(''); });
            onedata.push(_row);
        });

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
        data.push({
            sheetNum: sheetObj.sheetNum,
            // TODO: sheetName: get from styles
            contents: onedata
        });
    });
    return data;
}

module.exports = function parseXlsx() {
    var path, sheets, cb;
    if (arguments.length == 2) {
        path = arguments[0];
        sheets = null;
        cb = arguments[1];
    }
    else if (arguments.length == 3) {
        path = arguments[0];
        sheets = arguments[1];
        if (typeof sheets === 'number') sheets = [sheets];
        cb = arguments[2];
    }
    extractFiles(path, sheets, function(err, files) {
        if (err) return cb(err);
        cb(null, extractData(files));
    });
};
