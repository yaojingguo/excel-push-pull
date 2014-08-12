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
      excelFilePath.should.equal(push._filePath);
    });
  });
  describe('#record', function() {
    it('should add a record to push', function() {
      push.record(records[0], 2);
      ({
        2: records.slice(0, 1)
      }).should.eql(push._records);
    });
  });
  describe('#records', function() {
    it('should add some records to push', function() {
      push.records(records.slice(1), 2);
      ({
        2: records
      }).should.eql(push._records);
    });
  });
  describe('#pipe', function() {
    it('should pipe to fstream', function(done) {
      var ws = fs.createWriteStream(path.join(__dirname, '.ignore.worksheets.xlsx'));
      push.pipe(ws);
      ws.on('close', done);
    });
  });
});