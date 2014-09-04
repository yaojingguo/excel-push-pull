var fs = require('fs');
var path = require('path');

var Push = require('../lib/push');
var Pull = require('../lib/pull');

var tests = [];
fs.readdirSync(__dirname).forEach(function(fil) {
  var match = /^test-(\w+)\.xlsx$/.exec(fil);
  if (match) {
    tests.push(match[1]);
  }
});
tests.forEach(function(test) {
  var xlsxFile = path.join(__dirname, 'test-' + test + '.xlsx');
  var tempFile = path.join(__dirname, '.ignore-' + test + '.xlsx');
  var jsonFile = path.join(__dirname, 'data-' + test + '.json');
  var inputJson = require(jsonFile);
  describe('Test case ' + test + ' with writing file', function() {
    var push = new Push();
    var pull = new Pull();
    it('should push records', function (done) {
      this.timeout(3000);
      push.setXLSXStream(fs.createReadStream(xlsxFile));
      push.records(inputJson || []);
      var ws = fs.createWriteStream(tempFile, {
        encoding: null
      });
      push.pipe(ws);
      ws.on('error', function(err) {
        done(err);
      });
      ws.on('close', function() {
        done();
      });
    });
    it('should pull the pushed value', function (done) {
      this.timeout(3000);
      pull.setFilePath(tempFile);
      pull.records(function(err, records) {
        if (err) return done(err);
        inputJson === records || inputJson.should.eql(records);
        done();
      });
    });
  });
});

