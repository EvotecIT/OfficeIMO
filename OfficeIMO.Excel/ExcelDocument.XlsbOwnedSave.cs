using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Writes scalar changes from the deferred model when they leave the native package structure intact.</summary>
        private bool TryPrepareOwnedXlsbEdit(ExcelSaveOptions? options, out byte[] bytes, CancellationToken cancellationToken = default) {
            long maximumBytes = Math.Min(GetPackageMaterializationLimit(options) ?? int.MaxValue, int.MaxValue);
            byte[] prepared = Array.Empty<byte>();
            bool success = WithDeferredOpenGate(() => {
                if (!IsSdkDeferred || _xlsbDeferredBaseline == null || !ReferenceEquals(_deferredPackageBytes, _xlsbDeferredBaseline)
                    || _xlsbOriginalPackageBytes == null || _xlsbAdvancedWorkbook == null
                    || _packagePropertiesDirty || _customDocumentPropertiesDirty || HasXlsbTransformSaveWork(options)
                    || options?.DisableFastPackageWriter == true || HasCalculationSaveWork(options)) return false;

                cancellationToken.ThrowIfCancellationRequested();
                var taken = TakeDeferredEdits();
                if (taken.Length == 0) return false;
                bool adopted = false;
                try {
                    var sheetEdits = new Dictionary<string, IReadOnlyDictionary<(int Row, int Column), ExcelDeferredCellEdit>>(StringComparer.Ordinal);
                    foreach ((ExcelSheet sheet, ExcelDeferredCellStore store) in taken) {
                        var edits = new Dictionary<(int Row, int Column), ExcelDeferredCellEdit>();
                        for (int index = 0; index < store.Count; index++) {
                            ExcelDeferredCellEdit edit = store[index];
                            if (edit.Kind is not (ExcelDeferredCellKind.Number or ExcelDeferredCellKind.Integer
                                or ExcelDeferredCellKind.Decimal or ExcelDeferredCellKind.Boolean or ExcelDeferredCellKind.Text)) return false;
                            edits[(edit.Row, edit.Column)] = edit;
                        }
                        sheetEdits.Add(sheet.Name, edits);
                    }
                    if (!XlsbNativePackageWriter.TryRewriteOwnedCells(_xlsbOriginalPackageBytes, _xlsbAdvancedWorkbook,
                            sheetEdits, maximumBytes, out byte[] nativeBytes)) return false;
                    XlsbWorkbook native = ReferenceEquals(nativeBytes, _xlsbOriginalPackageBytes)
                        ? _xlsbAdvancedWorkbook
                        : XlsbWorkbookReader.Load(nativeBytes, new XlsbImportOptions {
                            MaxPackageBytes = _xlsbAdvancedWorkbook.MaxPackageBytes,
                            MaxPartBytes = _xlsbAdvancedWorkbook.MaxPartBytes,
                            CancellationToken = cancellationToken
                        });
                    if (!TryCreateDeferredPackageBytesCore(taken, maximumBytes, out byte[] projection)) return false;
                    cancellationToken.ThrowIfCancellationRequested();

                    // Like an owned XLSX save, publish the edited document state before writing to a destination.
                    // A failing or re-entrant destination cannot drop the taken edits, which now live in the packages.
                    CompleteDeferredSave(projection);
                    _xlsbDeferredBaseline = projection;
                    _xlsbOriginalPackageBytes = nativeBytes;
                    if (!ReferenceEquals(native, _xlsbAdvancedWorkbook)) {
                        foreach (var mapping in _xlsbAdvancedWorkbook.ProjectionPartUris) {
                            native.ProjectionPartUris[mapping.Key] = mapping.Value;
                        }
                    }
                    _xlsbAdvancedWorkbook = native;
                    _xlsbImportDiagnostics = native.Diagnostics.ToArray();
                    _xlsbPreservedRecords = native.PreservedRecords.ToArray();
                    _openXmlOriginalPackageBytes = null;
                    _unchangedPackageBytes = null;
                    prepared = nativeBytes;
                    adopted = true;
                    return true;
                } finally {
                    if (!adopted) RestoreDeferredEdits(taken);
                }
            });
            bytes = prepared;
            return success;
        }
    }
}
