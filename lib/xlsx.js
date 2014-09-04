'use strict';
var unzip = require("unzip");
var path = require('path');
var inherits = require('util').inherits;
var xmlParser;
try {
  xmlParser = require('xml2json-hotfix');
} catch (e) {
  xmlParser = require('xml2json');
}
var fs = require('fs');
var concat = require('concat-stream');
var extend = require('extend');
var debug = require('debug')('excel-push-pull');
var EventEmitter = require('events').EventEmitter;
var PullStream = require('pullstream');

module.exports = XLSX;

function XLSX(opts) {
  this._sheets = {};
  this._rows = {};
  this._sheetsMeta = {};
  this._strings;
  this._opts = opts || {};
  this._opts.templatePrefix = this._opts.templatePrefix || '##';
  this._emitter = new EventEmitter();
  this._picking = 0;
  this._pickingList = {};
  this._pickedList = {};
}

XLSX.prototype.setFilePath = function(filePath) {
  this._filePath = filePath;
}
XLSX.prototype.setXLSXStream = function(rs) {
  this._readStream = rs;
}
XLSX.prototype.setXLSXBuffer = function(buffer) {
  this._buffer = buffer;
}
XLSX.prototype._init = function(cb) {
  var self = this;
  var callbacked = false;
  this._initBuffer(function(err, buf) {
    if (err) return callback(err);
    var rs = new PullStream();
    rs.write(buf);
    rs.end();
    rs = rs.pipe(unzip.Parse());
    rs.on('error', callback);
    rs.on('entry', self._bypassEntry.bind(self));
    rs.on('close', function() {
      debug('Picking left when closed: %d', self._picking);
      if (self._picking === 0) {
        self._initAll(callback);
      } else {
        self._emitter.once('pick', self._initAll.bind(self, callback));
      }
    });
  });

  function callback(err) {
    if (callbacked) {
      return;
    }
    callbacked = true;
    if (self._err || err) return cb(self._err || err);
    cb();
  }
}
XLSX.prototype._initBuffer = function(cb) {
  if (this._buffer) return cb(null, this._buffer);
  if (this._filePath) {
    fs.readFile(this._filePath, cb);
  } else if (this._readStream) {
    this._readStream.pipe(concat(function(buf) {
      return cb(null, buf);
    }));
  } else {
    return cb(new Error('Cannot init buffer'));
  }
}
XLSX.prototype._initAll = function(cb) {
  if (this._err) return cb(this._err);
  try {
    this._initSheetsMeta();
    this._initRows();
  } catch (err) {
    return cb(err);
  }
  cb();
}
XLSX.prototype._initRows = function() {
  for (var k in this._sheets) {
    var rowIndex = this._rows[k] = {};
    var rows = this._sheets[k].worksheet.sheetData.row;
    if (rows) {
      for (var r = 0; r < rows.length; ++r) {
        var row = rows[r];
        rowIndex[row.r] = r;
      }
    }
  }
  debug('Init row index');
}
XLSX.prototype._initSheetsMeta = function() {
  for (var k in this._sheets) {
    this._sheetsMeta[k] = this._parseSheet(this._sheets[k]);
  }
  debug('Init sheets meta');
}
XLSX.prototype._bypassEntry = function(entry) {
  var self = this;
  var entryType = this._entryType(entry);
  if (entryType === 'strings') {
    debug('Found strings entry: %s', entry.path);
    this._pick(entry, function(err, json) {
      if (err) {
        debug('Error: %s', err.toString(), err.stack);
        self._err = err;
      } else {
        debug('Pick: %s', entry.path);
        self._strings = json;  
      }
      if (self._picking === 0) {
        self._emitter.emit('pick');
      }
    });
  } else if (entryType === 'worksheet') {
    debug('Found worksheet entry: %s', entry.path);
    this._pick(entry, function(err, json) {
      if (err) {
        debug('Error: %s', err.toString(), err.stack);
        self._err = err;
      } else {
        debug('Pick: %s', entry.path);
        var match = /sheet(\d+)\.xml$/.exec(entry.path);
        self._sheets[match[1]] = json;
      }
      if (self._picking === 0) {
        self._emitter.emit('pick');
      }
    });
  } else {
    debug('Bypass entry: %s', entry.path);
    this._drain(entry);
  }
}
XLSX.prototype._drain = function(entry) {
  entry.autodrain();
}
XLSX.prototype._pick = function(entry, cb) {
  var self = this;
  if (!this._pickingList[entry.path]) {
    this._pickingList[entry.path] = true;
    this._picking++;
    debug('Picking: %s', entry.path);
  } else {
    debug('Picking duplicated: %s', entry.path);
  }
  entry.pipe(concat(function(data) {
    if (self._pickedList[entry.path]) {
      debug('Picked duplicated: %s', entry.path);
      return;
    }
    self._picking --;
    self._pickedList[entry.path] = true;
    try {
      var buf = data;
      data = xmlParser.toJson(buf.toString('utf-8'), {
        object: true,
        reversible: true,
        coerce: false,
        sanitize: false,
        trim: false
      });
    } catch (err) {
      // console.log(buf.toString());
      // process.exit(0);
      cb(err);
      return;
    }
    cb(null, data);
  }));
}
XLSX.prototype._entryType = function(entry) {
  if (entry.path === 'xl/sharedStrings.xml') {
    return 'strings';
  }
  if (entry.path.indexOf('xl/worksheets/sheet') === 0 && entry.type === 'File') {
    return 'worksheet';
  }
}

XLSX.prototype._parseSheet = function(sheetJson) {
  var self = this;
  var template;
  var templatePrefix = this._opts.templatePrefix;
  var rows = sheetJson.worksheet.sheetData.row;
  if (rows) {
    for (var i = 0; i < rows.length; ++i) {
      var row = rows[i];
      for (var j = 0; j < row.c.length; ++j) {
        var cell = row.c[j];
        var val = this._parseCellValue(cell);
        if (!val) continue;
        if (val.indexOf(templatePrefix) === 0) {
          template = row;
          break;
        }
      }
      if (template) break;
    }
  }
  if (!template) 
    return {
      template: null
    };
  return {
    template: extend(true, {}, template)
  };
}
XLSX.prototype._parseCellValue = function(cell) {
  if (!cell.v) return;
  if (cell.t === 's') return this._derefValue(cell.v.$t);
  if (cell.t === 'n') return parseFloat(cell.v.$t);
  if (cell.t === 'b') return Boolean(cell.v.$t);
  return cell.v.$t;
}
XLSX.prototype._derefValue = function(id) {
  if (!this._strings) return '';
  var string = this._strings.sst.si[id];
  if (string && string.t) {
    return string.t.$t;
  }
  return '';
}
XLSX.prototype._getEntities = function() {
  var entities = [];
  if (this._strings) {
    var stringsXML = xmlParser.toXml(this._strings);
    entities.push({
      xml: stringsXML,
      path: 'xl/sharedStrings.xml'
    });
    debug('Convert to xml: xl/sharedStrings.xml');
  }
  for (var k in this._sheets) {
    var sheet = this._sheets[k];
    entities.push({
      xml: xmlParser.toXml(sheet),
      path: 'xl/worksheets/sheet' + k + '.xml'
    });
    debug('Convert to xml: xl/worksheets/sheet%d.xml', k);
  }
  return entities;
}

XLSX.prototype._setRow = function(sheetId, row) {
  var rows = this._sheets[sheetId].worksheet.sheetData.row;
  var rowIndex = this._rows[sheetId];
  if (row.r in rowIndex) {
    var r = rowIndex[row.r];
    rows[r] = row;
  } else {
    rowIndex[row.r] = rows.length;
    rows.push(row);
  }
}

XLSX.prototype._substitute = function(rowTemplate, record) {
  rowTemplate.r++;
  var row = extend(true, {}, rowTemplate);
  var cells = row.c;
  for (var i = 0; i < cells.length; ++i) {
    var cell = cells[i];
    cell.r = this._setCellRow(cell.r, row.r)
    var val = this._parseCellValue(cell);
    if (val && val.indexOf(this._opts.templatePrefix) === 0) {
      var key = val.slice(this._opts.templatePrefix.length);
      var val = this._getValue(record, key);
      delete cell.v;
      delete cell.t;
      if (val === undefined || val === null) continue;
      cell.t = this._getType(val);
      cell.v = {
        $t: val
      };
    }
  }
  return row;
}
XLSX.prototype._match = function(rowTemplate, row) {
  var dict = {};
  var isEmpty = true;
  for (var c = 0; c < row.c.length; ++c) {
    var cell = row.c[c];
    if (!cell.v) continue;
    var match = /([A-Z]+)(\d+)/.exec(cell.r);
    dict[match[1]] = this._parseCellValue(cell);
  }
  var json = {};
  var templatePrefix = this._opts.templatePrefix;
  for (var c = 0; c < rowTemplate.c.length; ++c) {
    var cell = rowTemplate.c[c];
    var match = /([A-Z]+)(\d+)/.exec(cell.r);
    var val = dict[match[1]];
    if (!cell.v) continue;
    if (val === undefined) val = '';
    var placeholder = this._parseCellValue(cell);
    if (placeholder.indexOf(templatePrefix) === 0) {
      var field = placeholder.substr(templatePrefix.length);
      if (val) isEmpty = false;
      this._setValue(json, field, val);
    }
  }
  return isEmpty ? null : json;
}
XLSX.prototype._getValue = function(json, key) {
  key.split('.').some(function(k) {
    if (json === undefined || json === null) return true;
    json = json[k];
  });
  return json;
}
XLSX.prototype._setValue = function(json, key, val) {
  var keys = key.split('.');
  var lastKey = keys[keys.length - 1];
  keys.slice(0, -1).forEach(function(k) {
    if (!json[k]) json[k] = {};
    json = json[k];
  });
  json[lastKey] = val;
}
XLSX.prototype._getType = function(v) {
  if (typeof v === 'boolean') return 'b';
  if (typeof v === 'number') return 'n';
  return 'str';
}
XLSX.prototype._setCellRow = function(ref, r) {
  var match = /([A-Z]+)(\d+)/.exec(ref);
  return match[1] + r;
}