using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private byte[]? _xlsbDeferredBaseline;

        private static ExcelDocument? TryProjectXlsbOnOwnedPackage(XlsbWorkbook source, bool readOnly,
            CancellationToken cancellationToken, int maxCells) {
            if (!XlsbOwnedWorkbookProjector.CanWrite(source)) return null;
            cancellationToken.ThrowIfCancellationRequested();
            XlsbWorkbookProjector.EnsureArrayProjectionBudget(source, maxCells, cancellationToken);
            ExcelDocumentType type = source.VbaProjectPartName == null
                ? ExcelDocumentType.Workbook : ExcelDocumentType.MacroEnabledWorkbook;
            byte[] bytes;
            long inputVersion = 0, authoredVersion = 0;
            Dictionary<(Uri WorksheetUri, string CellReference), FormulaAuthoredVersions>? formulas = null;
            var worksheetUris = new Dictionary<string, Uri>(StringComparer.Ordinal);

            using (OfficeOpcEditablePackage package = CreateWorkbookPackage(CreatedMainContentType(type)!, theme: null)) {
                bytes = XlsbOwnedWorkbookProjector.Write(package, source, SanitizeWorksheetName, RecordMutation, cancellationToken);
            }

            ExcelDeferredWorkbook? workbook = TryReadDeferredWorkbook(bytes);
            if (workbook == null) return null;
            ExcelDocument result = CreateDeferredDocument(workbook, bytes, readOnly);
            result._deferredPackageWrittenHere = true;
            result._formulaInputMutationVersion = inputVersion;
            result._formulaAuthoredMutationVersion = authoredVersion;
            if (formulas != null) result._formulaAuthoredVersions.Restore(formulas);
            return result;

            void RecordMutation(string partName, string? formula, bool hasValue) {
                if (hasValue) inputVersion++;
                if (formula == null) return;
                if (!worksheetUris.TryGetValue(partName, out Uri? uri)) {
                    uri = new Uri(partName, UriKind.Relative);
                    worksheetUris.Add(partName, uri);
                }
                (formulas ??= new(FormulaCellKeyComparer.Instance))[(uri, formula)] =
                    new FormulaAuthoredVersions(inputVersion, ++authoredVersion);
                if (hasValue) inputVersion++;
            }
        }

        private void MarkLoadedFromOwnedXlsb(string? sourcePath, XlsbWorkbook workbook) {
            ArgumentNullException.ThrowIfNull(workbook);

            SourceFormat = ExcelFileFormat.Xlsb;
            _legacyXlsPreservationSnapshot = null;
            _xlsbSourcePath = sourcePath;
            _xlsbOriginalPackageBytes = workbook.OriginalPackageBytes;
            _legacyXlsWasEncryptedSource = false;
            _openXmlOriginalPackageBytes = null;
            _xlsbAdvancedWorkbook = workbook;
            _xlsbImportDiagnostics = workbook.Diagnostics.ToArray();
            CaptureBinaryLoadDiagnostics(workbook);
            _xlsbPreservedRecords = workbook.PreservedRecords.ToArray();
            _packageDirty = false;
            _packagePropertiesDirty = false;
            _xlsbDeferredBaseline = _deferredPackageBytes;
            _requiresSavePreflight = false;
            _unchangedPackageBytes = null;
            _packageContentTypesKnownNormalized = false;
            if (!string.IsNullOrEmpty(sourcePath)) {
                FilePath = sourcePath!;
            }
        }

        // An explicit SDK fallback needs the same original state as a workbook that was projected
        // through the SDK. Capture it before the bridge replays the edits on a deferred workbook.
        private void RecordOpenedDeferredPackage(string reason, byte[] bytes) {
            _openXmlSdkFallbacks.Add(reason);
            if (_xlsbDeferredBaseline is not { } original) return;
            if (ReferenceEquals(original, bytes)) {
                _xlsbPreservationGuard = XlsbProjectionPreservationGuard.Capture(this);
                _xlsbDocumentPropertiesSnapshot = XlsbDocumentPropertyParts.Snapshot.Take(this);
            } else {
                using ExcelDocument baseline = LoadFromByteArray(original,
                    new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly }, filePath: null);
                _xlsbPreservationGuard = XlsbProjectionPreservationGuard.Capture(baseline);
                _xlsbDocumentPropertiesSnapshot = XlsbDocumentPropertyParts.Snapshot.Take(baseline);
            }
            _xlsbDeferredBaseline = null;
        }
    }
}
