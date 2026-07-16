using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
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
            if (HasXlsbTransformSaveWork(options)) return false;
            return SourceFormat != ExcelFileFormat.Xlsb
                || (_xlsbOriginalPackageBytes != null && _xlsbAdvancedWorkbook != null);
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
            return "Native XLSB writing does not support the requested save-time transforms. XLSB output is rejected before writing so XLSX bytes are never mislabeled as .xlsb.";
        }

        private bool TrySaveUnchangedXlsbToFile(string path, ExcelSaveOptions? options) {
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path) || !CanWriteNativeXlsb(options)) {
                return false;
            }

            bool existingSource = SourceFormat == ExcelFileFormat.Xlsb;
            bool unchanged = existingSource && CanCopyUnchangedXlsb(options);
            byte[] bytes = existingSource
                ? unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb()
                : GenerateNativeXlsb(options);
            CommitPreparedPackageToFile(path, bytes);
            FilePath = path;
            _xlsbSourcePath = path;
            AdoptNativeXlsbAfterWrite(bytes, path);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(!existingSource
                ? "New workbook encoded with the first-party BIFF12 XLSB writer."
                : unchanged
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

            bool existingSource = SourceFormat == ExcelFileFormat.Xlsb;
            bool unchanged = existingSource && CanCopyUnchangedXlsb(options);
            byte[] bytes = existingSource
                ? unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb()
                : GenerateNativeXlsb(options);
            await CommitPreparedPackageToFileAsync(path, bytes, cancellationToken).ConfigureAwait(false);
            FilePath = path;
            _xlsbSourcePath = path;
            AdoptNativeXlsbAfterWrite(bytes, path);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(!existingSource
                ? "New workbook encoded with the first-party BIFF12 XLSB writer."
                : unchanged
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

            bool existingSource = SourceFormat == ExcelFileFormat.Xlsb;
            bool unchanged = existingSource && CanCopyUnchangedXlsb(options);
            if (!existingSource) {
                PrepareWorkbookForSave(options);
                if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
                XlsbNewPackageWriter.Write(this, destination);
                if (destination.CanSeek) destination.SetLength(destination.Position);
                try { destination.Flush(); } catch (NotSupportedException) { }
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("New workbook streamed with the first-party BIFF12 XLSB writer.");
                return true;
            }

            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb();
            PrepareDestinationStreamForWrite(destination);
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

            bool existingSource = SourceFormat == ExcelFileFormat.Xlsb;
            bool unchanged = existingSource && CanCopyUnchangedXlsb(options);
            if (!existingSource) {
                cancellationToken.ThrowIfCancellationRequested();
                PrepareWorkbookForSave(options);
                cancellationToken.ThrowIfCancellationRequested();
                if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
                XlsbNewPackageWriter.Write(this, destination);
                if (destination.CanSeek) destination.SetLength(destination.Position);
                try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("New workbook streamed with the first-party BIFF12 XLSB writer.");
                return true;
            }

            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb();
            PrepareDestinationStreamForWrite(destination);
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

        private byte[] GenerateNativeXlsb(ExcelSaveOptions? options) {
            PrepareWorkbookForSave(options);
            using var output = new MemoryStream();
            XlsbNewPackageWriter.Write(this, output);
            byte[] bytes = output.ToArray();
            XlsbWorkbookReader.Load(bytes, new XlsbImportOptions { ReportPreservedRecords = false });
            return bytes;
        }

        private void AdoptNativeXlsbAfterWrite(byte[] bytes, string? sourcePath) {
            XlsbWorkbook workbook = XlsbWorkbookReader.Load(bytes);
            if (workbook.Stylesheet != null) XlsbStylesheetProjector.Install(this, workbook.Stylesheet);
            MarkLoadedFromXlsb(sourcePath, workbook);
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
