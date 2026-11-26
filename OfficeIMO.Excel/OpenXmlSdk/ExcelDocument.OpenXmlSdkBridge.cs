using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    /// <summary>
    /// The bridge between OfficeIMO's own Open XML layer and the Open XML SDK model. The owned engine hands the bridge
    /// the package bytes it holds; the bridge opens the SDK package over them with the load's settings, binds the
    /// worksheets the document has handed out, writes the values recorded so far through the model, and from then on
    /// the SDK package owns the document's saving. Everything the engine does without the SDK lives outside this folder.
    /// </summary>
    public partial class ExcelDocument {
        // The settings a deferred load opens the SDK package with; set while the SDK has not been opened.
        private OpenSettings? _deferredOpenSettings;
        private readonly object _deferredOpenGate = new();
        // Set while one thread opens the package, so the SDK getters on other threads wait for it to finish.
        private volatile bool _deferredSdkOpening;

        /// <summary>
        /// Opens the SDK package of a deferred load over the bytes the load read, with its settings, and binds the
        /// worksheets the document has handed out. Does nothing when the package is already open.
        /// </summary>
        internal void OpenDeferredSdk(string? reason) {
            if (_deferredWorkbook == null && !_deferredSdkOpening) {
                if (_disposed && _spreadSheetDocumentValue == null) throw new ObjectDisposedException(nameof(ExcelDocument));
                return;
            }
            WithDeferredOpenGate(() => {
                OpenDeferredSdkGated(reason);
                return true;
            });
        }

        /// <summary>
        /// Runs <paramref name="action"/> under the open gate. A thread holding no document lock takes the write lock
        /// first, the order every writer takes them in, so replaying recorded values never waits on a writer that is
        /// itself waiting on the gate.
        /// </summary>
        private T WithDeferredOpenGate<T>(Func<T> action) {
            ReaderWriterLockSlim lck = EnsureLock();
            if (!Locking.IsNoLock && !lck.IsReadLockHeld && !lck.IsWriteLockHeld && !lck.IsUpgradeableReadLockHeld
                && !Monitor.IsEntered(_deferredOpenGate)) {
                return Locking.ExecuteWrite(lck, () => {
                    lock (_deferredOpenGate) return action();
                });
            }
            lock (_deferredOpenGate) return action();
        }

        private void OpenDeferredSdkGated(string? reason) {
            // One thread opens the package; others wait on the gate until its recorded changes are in the model. The
            // opening thread finds the workbook already taken when binding and replay come back through the SDK getters.
            ExcelDeferredWorkbook? deferred = _deferredWorkbook;
            if (_deferredSdkOpening) return;
            if (deferred == null) {
                if (_disposed && _spreadSheetDocumentValue == null) throw new ObjectDisposedException(nameof(ExcelDocument));
                return;
            }
            _deferredSdkOpening = true;
            try {
                OpenDeferredSdkLocked(deferred, reason);
            } finally {
                _deferredSdkOpening = false;
            }
        }

        private void OpenDeferredSdkLocked(ExcelDeferredWorkbook deferred, string? reason) {
            byte[] bytes = _deferredPackageBytes!;
            Stream stream = OfficePackageLoadSource.CreatePackageStream(bytes, DocumentAccessMode.ReadWrite);
            SpreadsheetDocument document;
            try {
                document = _deferredReadOnly
                    ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenSpreadsheet(stream, _deferredOpenSettings!)
                    : SpreadsheetDocument.Open(stream, true, _deferredOpenSettings!);
            } catch {
                stream.Dispose();
                throw;
            }

            _deferredPackageBytes = null;
            _spreadSheetDocumentValue = document;
            _workBookPartValue = GetWorkbookPartOrThrow(document);
            _packageCopyOnWriteStream = stream as OfficeCopyOnWriteStream;
            _packageOnCopyOnWriteStream = _packageCopyOnWriteStream != null ? document : null;
            ExcelChartAxisIdGenerator.ReadChartRelationships(document);
            // Binding hides each sheet's workbook entry, so the changes recorded there move to the sheets that replay them.
            lock (_deferredEditsGate) {
                // Recording checks this under the same gate, so no change lands in an entry after it moves.
                _deferredWorkbook = null;
                foreach (ExcelSheet edited in _deferredEditedSheets) {
                    edited.DeferredCellsForReplay = edited.DeferredEntry?.Cells;
                    if (edited.DeferredEntry is { } entry) entry.Cells = null;
                }
            }
            // A read can open the package (the first SDK read of a workbook with recorded changes). Binding and replay
            // write to the model, which a read lock cannot be upgraded to; no writer can hold the lock meanwhile, and
            // other threads wait on the open gate, so they run without taking it.
            ReaderWriterLockSlim lck = EnsureLock();
            using (lck.IsReadLockHeld && !lck.IsWriteLockHeld && !lck.IsUpgradeableReadLockHeld ? Locking.EnterNoLockScope() : null) {
                BindDeferredSheets(deferred);
                _openXmlSdkFallbacks.Add(reason ?? DescribeSdkCaller());
                // Values written before the package opened go through the model now, in the order they were written.
                ReplayDeferredCellEdits();
            }
        }

        /// <summary>Binds each handed-out sheet to its SDK element and part, and rebuilds the sheet cache around them.</summary>
        private void BindDeferredSheets(ExcelDeferredWorkbook deferred) {
            var elements = ReadSheetElements();
            var byRelationship = new Dictionary<string, Sheet>(StringComparer.Ordinal);
            foreach (Sheet element in elements) {
                if (element.Id?.Value is { } id && !byRelationship.ContainsKey(id)) byRelationship[id] = element;
            }

            lock (_sheetWrapperCacheLock) {
                for (int index = _sheetWrappers.Count - 1; index >= 0; index--) {
                    if (!_sheetWrappers[index].TryGetTarget(out ExcelSheet? wrapper)) {
                        _sheetWrappers.RemoveAt(index);
                        continue;
                    }
                    ExcelDeferredSheet? entry = wrapper.DeferredEntry;
                    if (entry == null || !byRelationship.TryGetValue(entry.RelationshipId, out Sheet? element)) continue;
                    if (_workBookPartValue.GetPartById(entry.RelationshipId) is WorksheetPart part) {
                        wrapper.BindToSdk(_spreadSheetDocumentValue, element, part);
                    }
                }
            }

            lock (_sheetCacheRebuildGate) {
                UpdateSheetIdCache(elements);
                _cachedSheetElements = elements.ConvertAll(element => new CachedSheetElement(element));
                if (_cachedSheets != null) {
                    // The cached wrappers are the ones callers hold; they stay the document's sheets.
                    _cachedSheets = _cachedSheets.Where(sheet => sheet.DeferredEntry == null).ToList();
                    if (_cachedSheets.Count != elements.Count(IsWorksheetElement)) _sheetCacheDirty = true;
                }
            }
            _ = deferred;
        }

        private bool IsWorksheetElement(Sheet sheet) =>
            sheet.Id?.Value is { } id && _workBookPartValue.GetPartById(id) is WorksheetPart;

        // Settings beyond the defaults change how the SDK reads parts, so those loads open it right away.
        private static bool IsPlainOpenSettings(OpenSettings settings) =>
            settings.MaxCharactersInPart == 0
            && settings.MarkupCompatibilityProcessSettings.ProcessMode == MarkupCompatibilityProcessMode.NoProcess;
    }
}
