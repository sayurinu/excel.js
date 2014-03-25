var should = require('should');
var parseXlsx = require('../excelParser');
var excel = './xlsx/test.xlsx';

describe('execl.js', function() {

    describe('parseXlsx', function() {
        it('excelの全てのシートが読み込める', function(done) {
            parseXlsx(excel, function(err, result) {
                should.not.exist(err);
                should.exist(result);
                
                result.length.should.equal(3);
                var sheet1 = result[0];
                var sheet2 = result[1];
                var sheet3 = result[2];

                sheet1.sheetNum.should.equal(1);
                sheet1.sheetName.should.equal('Sheet1');
                sheet1.contents.length.should.equal(5);
                
                sheet2.sheetNum.should.equal(2);
                sheet2.sheetName.should.equal('Sheet2');
                sheet2.contents.length.should.equal(5);


                sheet3.sheetNum.should.equal(3);
                sheet3.sheetName.should.equal('Sheet3');
                sheet3.contents.length.should.equal(2);
                
                var contents1 = sheet1.contents; 
                var contents2 = sheet2.contents;
                var contents3 = sheet3.contents;

                contents1[0][0].should.equal('hoge');
                contents1[0][1].should.equal('hoge');
                contents1[0][2].should.equal('hoge');

                contents2[0][0].should.equal('foo');
                contents2[0][1].should.equal('bar');
                
                contents3[0][0].should.equal('hoge');

                contents1[1][0].should.equal('1');
                contents1[1][1].should.equal('2');
                contents1[1][2].should.equal('3');

                contents2[1][0].should.equal('1');
                contents2[1][1].should.equal('1');

                contents3[1][0].should.equal('1');
                done();
            });
        });

        it('excelの特定のシート番号を読み込める', function(done) {
            parseXlsx(excel, [1], function(err, result) {
                should.not.exist(err);
                should.exist(result);

                var sheet1 = result[0];
                sheet1.sheetNum.should.equal(1);
                sheet1.sheetName.should.equal('Sheet1');
                var contents1 = sheet1.contents;

                contents1[0][0].should.equal('hoge');
                contents1[0][1].should.equal('hoge');
                contents1[0][2].should.equal('hoge');

                contents1[1][0].should.equal('1');
                contents1[1][1].should.equal('2');
                contents1[1][2].should.equal('3');
                done();
            });
        });

        it('excelの特定のシート番号を読み込める', function(done) {
            parseXlsx(excel, [2], function(err, result) {
                should.not.exist(err);
                should.exist(result);

                var sheet2 = result[0];
                sheet2.sheetNum.should.equal(2);
                sheet2.sheetName.should.equal('Sheet2');
                var contents2 = sheet2.contents;
                
                contents2[0][0].should.equal('foo');
                contents2[0][1].should.equal('bar');
                
                contents2[1][0].should.equal('1');
                contents2[1][1].should.equal('1');
                done();
            });
        });

        it('excelの特定のシート番号の範囲で読み込める', function(done) {
            parseXlsx(excel, [1, 2], function(err, result) {
                should.not.exist(err);
                should.exist(result);
                should.not.exist(result[2]);
                done();
            });
        });

        it('excelの特定のシート番号が飛び飛びでも読み込める', function(done) {
            parseXlsx(excel, [1, 3], function(err, result) {
                should.not.exist(err);
                should.exist(result);
                should.not.exist(result[2]);

                var sheet1 = result[0];
                var sheet3 = result[1];
                sheet1.sheetNum.should.equal(1);
                sheet1.sheetName.should.equal('Sheet1');
                var contents1 = sheet1.contents;

                sheet3.sheetNum.should.equal(3);
                sheet3.sheetName.should.equal('Sheet3');
                sheet3.contents.length.should.equal(2);
                var contents3 = sheet3.contents;


                contents1[0][0].should.equal('hoge');
                contents1[0][1].should.equal('hoge');
                contents1[0][2].should.equal('hoge');

                contents3[0][0].should.equal('hoge');
                
                contents1[1][0].should.equal('1');
                contents1[1][1].should.equal('2');
                contents1[1][2].should.equal('3');

                contents3[1][0].should.equal('1');
                done();
            });
        });
    });
});
