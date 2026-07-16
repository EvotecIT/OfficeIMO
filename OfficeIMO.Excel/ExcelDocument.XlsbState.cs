using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private void EnsureXlsbFileTargetSupported(
            string path,
            ExcelSaveOptions? options,
            bool allowUnchangedCopy = true) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path)) {
                return;
            }

            if (allowUnchangedCopy && CanWriteNativeXlsb(options)) {
                return;
            }

            throw new NotSupportedException(GetXlsbWriteUnsupportedMessage());
        }

        private void EnsureXlsbStreamTargetSupported(ExcelFileFormat format, ExcelSaveOptions? options) {
            if (format != ExcelFileFormat.Xlsb) {
                return;
            }

            if (CanWriteNativeXlsb(options)) {
                return;
            }

            throw new NotSupportedException(GetXlsbWriteUnsupportedMessage());
        }

        private bool CanCopyUnchangedXlsb(ExcelSaveOptions? options) {
            return SourceFormat == ExcelFileFormat.Xlsb
                && !_packageDirty
                && _xlsbOriginalPackageBytes != null
                && !HasXlsbTransformSaveWork(options);
        }

        private bool CanWriteNativeXlsb(ExcelSaveOptions? options) {
            return SourceFormat == ExcelFileFormat.Xlsb
                && _xlsbOriginalPackageBytes != null
                && _xlsbAdvancedWorkbook != null
                && !HasXlsbTransformSaveWork(options);
        }

        private static bool HasXlsbTransformSaveWork(ExcelSaveOptions? options) {
            return options?.SafeRepairDefinedNames == true
                || options?.ValidateOpenXml == true
                || options?.SafePreflight == true
                || options?.EvaluateFormulasBeforeSave == true
                || options?.ClearCachedFormulaResultsBeforeSave == true
                || options?.MarkFormulasDirtyBeforeSave == true
                || options?.ForceFullCalculationOnOpen == true;
        }

        private string GetXlsbWriteUnsupportedMessage() {
            return "Native XLSB generation currently requires an existing XLSB source and supports preservation-aware cell-value rewrites. New XLSB generation and requested save-time transforms are rejected before writing so XLSX bytes are never mislabeled as .xlsb.";
        }

        private bool TrySaveUnchangedXlsbToFile(string path, ExcelSaveOptions? options) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path) || !CanWriteNativeXlsb(options)) {
                return false;
            }

            bool unchanged = CanCopyUnchangedXlsb(options);
            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb();
            CommitPreparedPackageToFile(path, bytes);
            FilePath = path;
            _xlsbSourcePath = path;
            RefreshXlsbStateAfterNativeWrite(bytes);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(unchanged
                ? "Unmodified XLSB source copied byte-for-byte with all package parts preserved."
                : "XLSB worksheet cell records rewritten while all other package parts were preserved.");
            return true;
        }

        private async Task<bool> TrySaveUnchangedXlsbToFileAsync(
            string path,
            ExcelSaveOptions? options,
            CancellationToken cancellationToken) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path) || !CanWriteNativeXlsb(options)) {
                return false;
            }

            bool unchanged = CanCopyUnchangedXlsb(options);
            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb();
            await CommitPreparedPackageToFileAsync(path, bytes, cancellationToken).ConfigureAwait(false);
            FilePath = path;
            _xlsbSourcePath = path;
            RefreshXlsbStateAfterNativeWrite(bytes);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(unchanged
                ? "Unmodified XLSB source copied byte-for-byte with all package parts preserved."
                : "XLSB worksheet cell records rewritten while all other package parts were preserved.");
            return true;
        }

        private bool TrySaveUnchangedXlsbToStream(
            Stream destination,
            ExcelFileFormat format,
            ExcelSaveOptions? options) {
            if (format != ExcelFileFormat.Xlsb || !CanWriteNativeXlsb(options)) {
                return false;
            }

            bool unchanged = CanCopyUnchangedXlsb(options);
            PrepareDestinationStreamForWrite(destination);
            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb();
            destination.Write(bytes, 0, bytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            RefreshXlsbStateAfterNativeWrite(bytes);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(unchanged
                ? "Unmodified XLSB source copied byte-for-byte with all package parts preserved."
                : "XLSB worksheet cell records rewritten while all other package parts were preserved.");
            return true;
        }

        private async Task<bool> TrySaveUnchangedXlsbToStreamAsync(
            Stream destination,
            ExcelFileFormat format,
            ExcelSaveOptions? options,
            CancellationToken cancellationToken) {
            if (format != ExcelFileFormat.Xlsb || !CanWriteNativeXlsb(options)) {
                return false;
            }

            bool unchanged = CanCopyUnchangedXlsb(options);
            PrepareDestinationStreamForWrite(destination);
            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb();
            await destination.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            RefreshXlsbStateAfterNativeWrite(bytes);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(unchanged
                ? "Unmodified XLSB source copied byte-for-byte with all package parts preserved."
                : "XLSB worksheet cell records rewritten while all other package parts were preserved.");
            return true;
        }

        private byte[] RewriteNativeXlsb() {
            return XlsbNativePackageWriter.Rewrite(
                this,
                _xlsbOriginalPackageBytes ?? throw new InvalidOperationException("The XLSB source package is unavailable."),
                _xlsbAdvancedWorkbook ?? throw new InvalidOperationException("The XLSB source model is unavailable."));
        }

        private void RefreshXlsbStateAfterNativeWrite(byte[] bytes) {
            XlsbWorkbook workbook = XlsbWorkbookReader.Load(bytes);
            _xlsbOriginalPackageBytes = bytes;
            _xlsbAdvancedWorkbook = workbook;
            _xlsbImportDiagnostics = workbook.Diagnostics.ToArray();
            _xlsbPreservedRecords = workbook.PreservedRecords.ToArray();
            _packageDirty = false;
            _packagePropertiesDirty = false;
            _requiresSavePreflight = false;
            _unchangedPackageBytes = null;
            _packageContentTypesKnownNormalized = false;
        }
    }
}
