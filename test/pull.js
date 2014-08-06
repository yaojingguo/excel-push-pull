var path = require('path');
var fs = require('fs');
var Pull = require('../lib/pull.js');
describe('Pull', function () {
  var pull = new Pull();
  describe('#setXLSXStream', function() {
    it('should set xml file stream', function() {
      pull.setXLSXStream(fs.createReadStream(path.join(__dirname, 'std-worksheets.xlsx')));
    })  
  });
  describe('#records', function () {
    it('should pull records from std-worksheets.xlsx', function (done) {
      pull.records(2, function(err, records) {
        if (err) return done(err);
        records.should.eql(require('./records.json'));
        done();
      });
    });
  });
});
