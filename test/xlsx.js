var fs = require('fs');
var path = require('path');

var PullStream = require('pullstream');
var XLSX = require('../lib/xlsx');
var Push = require('../lib/push');

describe('XLSX', function() {
  var xlsx = new XLSX();
  describe('#_init', function() {
    it('should init the xlsx instance', function(done) {
      this.timeout(10000);
      xlsx.setFilePath(path.join(__dirname, 'worksheets.xlsx'));
      xlsx._init(done);
    });
  });
  var row;
  var json = {
    grade: '高一'
  };
  describe('#_substitute', function() {
    it('should substitue all placeholders', function() {
      row = xlsx._substitute(xlsx._sheetsMeta[2].template, json);
      row.should.eql({
        customHeight: 1,
        ht: 15,
        r: 6,
        spans: '1:201',
        'x14ac:dyDescent': 0.15,
        c: [{
          r: 'A6',
          t: 'str',
          v: {
            '$t': '高一'
          }
        }, {
          r: 'B6'
        }, {
          r: 'C6'
        }, {
          r: 'D6'
        }, {
          r: 'E6'
        }, {
          r: 'F6'
        }, {
          r: 'G6'
        }, {
          r: 'H6'
        }, {
          r: 'I6'
        }, {
          r: 'J6'
        }]
      });
    });
  });
  describe('#_setRow', function() {
    it('should set substitued row, row number plus one', function() {
      var prevLen = xlsx._sheets[2].worksheet.sheetData.row.length;
      xlsx._setRow(2, row);
      var nextLen = xlsx._sheets[2].worksheet.sheetData.row.length;
      (nextLen - prevLen).should.equal(1);
    });
    it('should set a duplicated row, row number keep', function() {
      var prevLen = xlsx._sheets[2].worksheet.sheetData.row.length;
      xlsx._setRow(2, row);
      var nextLen = xlsx._sheets[2].worksheet.sheetData.row.length;
      (nextLen - prevLen).should.equal(0);
    });
  });
  describe('#_match', function() {
    it('should match the row to json', function() {
      xlsx._match(xlsx._sheetsMeta[2].template, row).should.eql(json);
    });
  });
});