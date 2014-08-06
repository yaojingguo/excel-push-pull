'use strict'
var XLSX = require('./xlsx');
var xlsx = require('xlsx');
var inherits = require('util').inherits;
var archiver = require('archiver');
var concat = require('concat-stream');
inherits(Push, XLSX);

function Push(opts) {
  XLSX.call(this, opts);
  this._archive = archiver('zip');
  this._records = {};
}
Push.prototype.record = function(record, sheetId) {
  sheetId = sheetId || 1;
  if (!this._records[sheetId]) this._records[sheetId] = [];
  this._records[sheetId].push(record);
}
Push.prototype.records = function(records, sheetId) {
  for (var i = 0; i < records.length; ++i)
    this.record(records[i], sheetId);
}
Push.prototype.pipe = function(ws) {
  var self = this;
  this._init(function(err) {
    if (err) return;
    self._archive.pipe(ws);
    self._push();
    self._archive.finalize();
  });
  return this;
}
Push.prototype._push = function(err) {
  var self = this;
  Object.keys(this._records).forEach(function(sheetId) {
    var records = self._records[sheetId];
    self._pushRecords(sheetId, records);
  });
  var entities = this._getEntities();
  for (var i = 0; i < entities.length; ++i) {
    var entity = entities[i];
    this._archive.append(entity.xml, {
      name: entity.path
    });
  }
  this._archive.finalize();
}
Push.prototype._pushRecords = function(sheetId, records) {
  var meta = this._sheetsMeta[sheetId];
  if (!meta.template) return;
  for (var i = 0; i < records.length; ++i) {
    var record = records[i];
    var row = this._substitute(meta.template, record);
    this._setRow(sheetId, row);
  }
}
Push.prototype._drain = function(entry) {
  var self = this;
  if (entry.type === 'File') {
    entry.pipe(concat(function(data) {
      self._archive.append(data.toString(), {
        name: entry.path
      });
    }));
  } else if (entry.type === 'Directory') {
    this._archive.append(null, { name: entry.path });
    entry.autodrain();
  } else {
    entry.autodrain();
  }
}
module.exports = Push;