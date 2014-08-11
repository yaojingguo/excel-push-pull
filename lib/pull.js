'use strict'
var XLSX = require('./xlsx');
var xlsx = require('xlsx');
var inherits = require('util').inherits;
var archiver = require('archiver');
var concat = require('concat-stream');
var debug = require('debug')('excel-push-pull:pull');
inherits(Pull, XLSX);

function Pull(opts) {
  XLSX.call(this, opts);
  this._records = {};
}

Pull.prototype.records = function(sheetId, cb) {
  var self = this;
  if (sheetId instanceof Function) {
    cb = sheetId;
    sheetId = 1;
  }
  this._init(function(err) {
    debug('Init');
    if (err) {
      debug('Error: %s', err.toString());
      return cb(err);
    }
    cb(null, self._pull(sheetId));
  });
};

Pull.prototype._pull = function(sheetId) {
  if (this._records[sheetId]) {
    return this._records[sheetId];
  }
  var sheet = this._sheets[sheetId];
  if (!sheet) return [];
  var meta = this._sheetsMeta[sheetId];
  if (!meta.template) return [];

  var records = [];
  var rowTemplate = meta.template;
  var rows = sheet.worksheet.sheetData.row;
  for (var r = 0; r < rows.length; ++r) {
    if (r < rowTemplate.r) continue;
    var row = rows[r];
    var json = this._match(rowTemplate, row);
    records.push({
      r: r,
      data: json
    });
  }
  return records.sort(function(lhv, rhv) {
    return lhv.r - rhv.r;
  }).map(function(it) {
    return it.data;
  });
}
module.exports = Pull;

