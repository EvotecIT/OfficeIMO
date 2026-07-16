using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private void EnsureXlsbFileTargetSupported(
            string path,
            ExcelSaveOptions? options,
            bool allowUnchangedCopy = true) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path)) {
                return;
            }

            if (allowUnchangedCopy && CanCopyUnchangedXlsb(options)) {
                return;
            }

            throw new NotSupportedException(GetXlsbWriteUnsupportedMessage());
        }

        private void EnsureXlsbStreamTargetSupported(ExcelFileFormat format, ExcelSaveOptions? options) {
            if (format != ExcelFileFormat.Xlsb) {
                return;
            }

            if (CanCopyUnchangedXlsb(options)) {
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
            if (SourceFormat == ExcelFileFormat.Xlsb && _packageDirty) {
                return "Native XLSB rewriting is not available for modified workbooks in this build. Save to .xlsx, or reload the source and copy it unchanged to preserve every XLSB package part.";
            }

            return "Native XLSB generation is not available in this build. Unmodified XLSB sources can be copied byte-for-byte, and all other XLSB targets are rejected before writing so XLSX bytes are never mislabeled as .xlsb.";
        }

        private bool TrySaveUnchangedXlsbToFile(string path, ExcelSaveOptions? options) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path) || !CanCopyUnchangedXlsb(options)) {
                return false;
            }

            byte[] bytes = _xlsbOriginalPackageBytes!;
            CommitPreparedPackageToFile(path, bytes);
            FilePath = path;
            _xlsbSourcePath = path;
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Unmodified XLSB source copied byte-for-byte with all package parts preserved.");
            return true;
        }

        private async Task<bool> TrySaveUnchangedXlsbToFileAsync(
            string path,
            ExcelSaveOptions? options,
            CancellationToken cancellationToken) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path) || !CanCopyUnchangedXlsb(options)) {
                return false;
            }

            byte[] bytes = _xlsbOriginalPackageBytes!;
            await CommitPreparedPackageToFileAsync(path, bytes, cancellationToken).ConfigureAwait(false);
            FilePath = path;
            _xlsbSourcePath = path;
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Unmodified XLSB source copied byte-for-byte with all package parts preserved.");
            return true;
        }

        private bool TrySaveUnchangedXlsbToStream(
            Stream destination,
            ExcelFileFormat format,
            ExcelSaveOptions? options) {
            if (format != ExcelFileFormat.Xlsb || !CanCopyUnchangedXlsb(options)) {
                return false;
            }

            PrepareDestinationStreamForWrite(destination);
            byte[] bytes = _xlsbOriginalPackageBytes!;
            destination.Write(bytes, 0, bytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Unmodified XLSB source copied byte-for-byte with all package parts preserved.");
            return true;
        }

        private async Task<bool> TrySaveUnchangedXlsbToStreamAsync(
            Stream destination,
            ExcelFileFormat format,
            ExcelSaveOptions? options,
            CancellationToken cancellationToken) {
            if (format != ExcelFileFormat.Xlsb || !CanCopyUnchangedXlsb(options)) {
                return false;
            }

            PrepareDestinationStreamForWrite(destination);
            byte[] bytes = _xlsbOriginalPackageBytes!;
            await destination.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Unmodified XLSB source copied byte-for-byte with all package parts preserved.");
            return true;
        }
    }
}
