#excel-push-pull

##Push lines to pre-defined excel files

```javascript
var Push = require('excel-push-pull').Push;
var push = new Push();
var fs = require('fs');
var rs = fs.createReadStream('PATH_TO_INPUT_XLSX_FILE');
push.setXLSXStream(rs);

// Use this method to push a single record
push.record({
  // json data
});

// Use this method to push multiple records
push.records([{
  // json data
}]);

var ws = fs.createWriteStream('PATH_TO_OUTPUT_XLSX_FILE');
push.pipe(ws);
ws.on('close', function() {
  // all data piped to ws
});
```


##Pull records from pre-defined and pre-filled excel files
```javascript
var Pull = require('excel-push-pull').Pull;
var pull = new Pull();
pull.setFilePath('PATH_TO_INPUT_XLSX_FILE');

pull.records(function(err, records) {
  // records is an array contains json data
});
```

##API
- `setXLSXStream(readStream)` set a read stream
- `setXLSXBuffer(buffer)` set a zip buffer
- `setFilePath(filePath)` set a file path

###Push
- `record(json, sheetId=1)` push a json to sheet in _sheetId_
- `records(array, sheetId=1)` push an array of json to sheet in _sheetId_
- `pipe(writeStream)` pipe out the records added xlsx binary stream

###Pull
- `records(sheedId=1, callback)` pull an array of json from sheet in _sheetId_

  callback is in nodejs style, `callback(err, records)` where records is an array 
of data, empty line is trimed off.

