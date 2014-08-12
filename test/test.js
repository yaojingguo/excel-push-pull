var fs = require('fs');
var path = require('path');

var Push = require('../lib/push');
var Pull = require('../lib/pull');
var concat = require('concat-stream');

var tests = [];
fs.readdirSync(__dirname).forEach(function(fil) {
  var match = /^test-(\w+)\.xlsx$/.exec(fil);
  if (match) {
    tests.push(match[1]);
  }
});
tests.forEach(function(test) {
  var xlsxFile = path.join(__dirname, 'test-' + test + '.xlsx');
  var jsonFile = path.join(__dirname, 'data-' + test + '.json');
  var inputJson = require(jsonFile);
  describe('Test case ' + test, function() {
    var zipBuffer;
    var push = new Push();
    var pull = new Pull();
    it('should push records', function (done) {
      this.timeout(1500);
      push.setXLSXStream(fs.createReadStream(xlsxFile));
      push.records(inputJson);
      push.pipe(concat(function(buf) {
        zipBuffer = buf;
        done();
      }));
    });
    it('should pull the pushed value from zipBuffer', function (done) {
      this.timeout(1500);
      pull.setXLSXBuffer(zipBuffer);
      pull.records(function(err, records) {
        if (err) return done(err);
        inputJson.should.eql(records);
        done();
      });
    });
  });
});

