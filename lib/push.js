'use strict'
var XLSX = require('./xlsx');
var xlsx = require('xlsx');
var inherits = require('util').inherits;
var archiver = require('archiver');
var concat = require('concat-stream');
var debug = require('debug')('excel-push-pull:push');
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
    debug('Init');
    if (err) {
      debug('Error: %s', err.toString());
      ws.end();
      return;
    }
    self._archive.pipe(ws);
    debug('Archive pipe to write stream');
    self._push();
    debug('Push records to excel and append entries to archive');
    self._archive.finalize();
    debug('Finalize archive');
  });
  return this;
}
Push.prototype._push = function(err) {
  var self = this;
  Object.keys(this._records).forEach(function(sheetId) {
    var records = self._records[sheetId];
    self._pushRecords(sheetId, records);
    debug('Push %d records to sheet %d', records.length, sheetId);
  });
  var entities = this._getEntities();
  for (var i = 0; i < entities.length; ++i) {
    var entity = entities[i];
    debug('Append: %s', entity.path);
    this._archive.append(entity.xml, {
      name: entity.path
    });
  }
}
Push.prototype._pushRecords = function(sheetId, records) {
  var meta = this._sheetsMeta[sheetId];
  if (!meta.template) return;
  for (var i = 0; i < records.length; ++i) {
    var record = records[i];
    var row = this._substitute(meta.template, record);
    delete row.hidden;
    this._setRow(sheetId, row);
  }
}
Push.prototype._drain = function(entry) {
  var self = this;
  if (entry.type === 'File') {
    debug('Pipe file entry: %s', entry.path);
    entry.pipe(concat(function(data) {
      debug('Append file entry: %s', entry.path);
      self._archive.append(data.toString(), {
        name: entry.path
      });
    }));
  } else if (entry.type === 'Directory') {
    debug('Append dir entry: %s', entry.path);
    this._archive.append(null, { name: entry.path });
    entry.autodrain();
  } else {
    debug('Ignore entry: %s', entry.path);
    entry.autodrain();
  }
}
module.exports = Push;