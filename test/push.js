var fs = require('fs');
var path = require('path');

var XLSX = require('../lib/xlsx');
var Push = require('../lib/push');

describe('Push', function() {
  var xlsx = new XLSX();
  var push = new Push();
  var excelFilePath = path.join(__dirname, 'worksheets.xlsx');
  var records = require('./records.json');
  describe('#setFilePath', function() {
    it('should set file path to ' + excelFilePath, function() {
      push.setFilePath(excelFilePath);
      push._filePath.should.equal(excelFilePath);
    });
  });
  describe('#record', function() {
    it('should add a record to push', function() {
      push.record(records[0], 2);
      push._records.should.eql({
        2: records.slice(0, 1)
      });
    });
  });
  describe('#records', function () {
    it('should add some records to push', function() {
      push.records(records.slice(1), 2);
      push._records.should.eql({
        2: records
      });
    });
  });
  // describe('#finalize', function () {
  //   it('should finalize push instance', function (done) {
  //     push.finalize(done);
  //   });
  // });
  describe('#pipe', function () {
    it('should pipe to fstream', function(done) {
      this.timeout(100000);
      var ws = fs.createWriteStream(path.join(__dirname, 'output-worksheets.xlsx'));
      push.pipe(ws);
      ws.on('close', done);
    });
  });
});
