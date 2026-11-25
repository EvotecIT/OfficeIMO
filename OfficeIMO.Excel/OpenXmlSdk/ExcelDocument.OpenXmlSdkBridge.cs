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

        /// <summary>
        /// Opens the SDK package of a deferred load over the bytes the load read, with its settings, and binds the
        /// worksheets the document has handed out. Does nothing when the package is already open.
        /// </summary>
        internal void OpenDeferredSdk(string? reason) {
            ExcelDeferredWorkbook? deferred = _deferredWorkbook;
            if (deferred == null) {
                if (_disposed && _spreadSheetDocumentValue == null) throw new ObjectDisposedException(nameof(ExcelDocument));
                return;
            }
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

            _deferredWorkbook = null;
            _deferredPackageBytes = null;
            _spreadSheetDocumentValue = document;
            _workBookPartValue = GetWorkbookPartOrThrow(document);
            _packageCopyOnWriteStream = stream as OfficeCopyOnWriteStream;
            _packageOnCopyOnWriteStream = _packageCopyOnWriteStream != null ? document : null;
            ExcelChartAxisIdGenerator.ReadChartRelationships(document);
            BindDeferredSheets(deferred);
            _openXmlSdkFallbacks.Add(reason ?? DescribeSdkCaller());
            // Values written before the package opened go through the model now, in the order they were written.
            ReplayDeferredCellEdits();
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
