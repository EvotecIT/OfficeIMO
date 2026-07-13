using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing.Internal;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        /// <summary>
        /// Returns the workbook-level cache of table names, initializing it from the current
        /// document if needed. Case-insensitive comparison.
        /// </summary>
        internal HashSet<string> GetOrInitTableNameCache() {
            // Fast path without locking
            if (_tableNameCache != null) return _tableNameCache;

            // Initialize without taking a new lock if we're already in a write scope
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                var set = new HashSet<string>(_tableNameComparer);
                var wb = WorkbookPartRoot;
                foreach (var ws in wb.WorksheetParts) {
                    foreach (var tdp in ws.TableDefinitionParts) {
                        var n = tdp.Table?.Name?.Value;
                        if (!string.IsNullOrEmpty(n)) set.Add(n!);
                    }
                }
                _tableNameCache = set;
                return _tableNameCache!;
            }

            // Otherwise, use write lock for thread safety
            return Locking.ExecuteWrite(EnsureLock(), () => {
                if (_tableNameCache != null) return _tableNameCache;
                var set = new HashSet<string>(_tableNameComparer);
                var wb = WorkbookPartRoot;
                foreach (var ws in wb.WorksheetParts) {
                    foreach (var tdp in ws.TableDefinitionParts) {
                        var n = tdp.Table?.Name?.Value;
                        if (!string.IsNullOrEmpty(n)) set.Add(n!);
                    }
                }
                _tableNameCache = set;
                return _tableNameCache;
            });
        }

        /// <summary>
        /// Adds the given table name to the cache. Should be called once the name is finalized.
        /// </summary>
        internal void ReserveTableName(string name) {
            if (string.IsNullOrWhiteSpace(name)) return;
            var cache = GetOrInitTableNameCache();
            cache.Add(name);
        }

        /// <summary>
        /// Removes the given table name from the cache. Intended for future table deletion APIs.
        /// Safe to call even if the cache hasn't been initialized.
        /// </summary>
        internal void RemoveReservedTableName(string name) {
            if (string.IsNullOrWhiteSpace(name)) return;
            if (_tableNameCache == null) return;
            _tableNameCache.Remove(name);
        }

        internal uint AllocateTableId() {
            lock (_tableMetadataLock) {
                if (_nextTableId == null) {
                    uint maxExistingId = 0;
                    var workbookPart = WorkbookPartRoot;
                    foreach (var worksheetPart in workbookPart.WorksheetParts) {
                        foreach (var part in worksheetPart.TableDefinitionParts) {
                            var idValue = part.Table?.Id?.Value;
                            if (idValue != null && idValue.Value > maxExistingId) {
                                maxExistingId = idValue.Value;
                            }
                        }
                    }

                    _nextTableId = maxExistingId + 1;
                }

                uint tableId = _nextTableId.Value;
                _nextTableId = tableId + 1;
                return tableId;
            }
        }
    }
}
