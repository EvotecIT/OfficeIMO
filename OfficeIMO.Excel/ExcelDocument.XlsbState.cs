using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private XlsbProjectionPreservationGuard? _xlsbPreservationGuard;

        internal XlsbProjectionPreservationGuard NativeXlsbPreservation => _xlsbPreservationGuard
            ?? throw new InvalidOperationException("The XLSB source bindings are unavailable.");
        internal void ValidateNativeXlsbProjection() => _xlsbPreservationGuard?.Validate(this);
        private void EnsureXlsbFileTargetSupported(
            string path,
            ExcelSaveOptions? options,
            bool allowUnchangedCopy = true) {
            EnsureXlsbStringStorageSupported(ExcelDocumentLoadRouting.HasXlsbExtension(path), options);
            if (!ExcelDocumentLoadRouting.HasXlsbExtension(path)) {
                return;
            }

            if (allowUnchangedCopy && CanWriteNativeXlsb(options)) {
                return;
            }

            throw new NotSupportedException(GetXlsbWriteUnsupportedMessage());
        }

        private void EnsureXlsbStreamTargetSupported(ExcelFileFormat format, ExcelSaveOptions? options) {
            EnsureXlsbStringStorageSupported(format == ExcelFileFormat.Xlsb, options);
            if (format != ExcelFileFormat.Xlsb) {
                return;
            }

            if (CanWriteNativeXlsb(options)) {
                return;
            }

            throw new NotSupportedException(GetXlsbWriteUnsupportedMessage());
        }

        private void EnsureXlsbStringStorageSupported(bool isXlsbTarget, ExcelSaveOptions? options) {
            if (options?.XlsbUseSharedStrings == null) return;
            if (!isXlsbTarget) {
                throw new NotSupportedException("XlsbUseSharedStrings applies only to native XLSB output.");
            }
            if (SourceFormat == ExcelFileFormat.Xlsb) {
                throw new NotSupportedException("Changing string storage in an imported XLSB workbook is not supported because its preserved records may reference the original shared-string table. Leave XlsbUseSharedStrings unset to preserve that storage.");
            }
        }

        private bool CanCopyUnchangedXlsb(ExcelSaveOptions? options) {
            ValidateNativeXlsbProjection();
            return SourceFormat == ExcelFileFormat.Xlsb
                && !_packageDirty
                && _xlsbPreservationGuard?.HasWorksheetChanges(this) != true
                && _xlsbOriginalPackageBytes != null
                && !HasXlsbTransformSaveWork(options)
                && !HasCalculationSaveWork(options);
        }

        private bool CanWriteNativeXlsb(ExcelSaveOptions? options) {
            if (HasXlsbTransformSaveWork(options)) return false;
            return SourceFormat != ExcelFileFormat.Xlsb
                || (_xlsbOriginalPackageBytes != null && _xlsbAdvancedWorkbook != null);
        }

        private static bool HasXlsbTransformSaveWork(ExcelSaveOptions? options) {
            return options?.SafeRepairDefinedNames == true
                || options?.ValidateOpenXml == true
                || options?.SafePreflight == true;
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
            bool directWrite = false;
            XlsbWorkbook? generatedWorkbook = null;
            byte[] bytes = existingSource
                ? unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb(options)
                : GenerateNativeXlsb(options, CancellationToken.None, out directWrite, out generatedWorkbook);
            CommitPreparedPackageToFile(path, bytes);
            FilePath = path;
            _xlsbSourcePath = path;
            AdoptNativeXlsbAfterWrite(bytes, path, generatedWorkbook);
            LastSaveDiagnostics = directWrite
                ? ExcelSaveDiagnostics.NativeBinaryDirectPackage()
                : ExcelSaveDiagnostics.Standard(!existingSource
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
            bool directWrite = false;
            XlsbWorkbook? generatedWorkbook = null;
            byte[] bytes = existingSource
                ? unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb(options)
                : GenerateNativeXlsb(options, cancellationToken, out directWrite, out generatedWorkbook);
            await CommitPreparedPackageToFileAsync(path, bytes, cancellationToken).ConfigureAwait(false);
            FilePath = path;
            _xlsbSourcePath = path;
            AdoptNativeXlsbAfterWrite(bytes, path, generatedWorkbook);
            LastSaveDiagnostics = directWrite
                ? ExcelSaveDiagnostics.NativeBinaryDirectPackage()
                : ExcelSaveDiagnostics.Standard(!existingSource
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
                if (TryGetDirectTabularSaveSource(options, CancellationToken.None, "Xlsb", out ExcelDirectTabularSource source)) {
                    if (XlsbNewPackageWriter.TryWriteDirectTabular(this, source, destination, CancellationToken.None, options?.XlsbUseSharedStrings == true)) {
                        try { destination.Flush(); } catch (NotSupportedException) { }
                        LastSaveDiagnostics = ExcelSaveDiagnostics.NativeBinaryDirectPackage();
                        return true;
                    }

                    Execution.ReportInfo("Save.Xlsb.Direct skipped: the tabular values require the general BIFF12 value encoder.");
                }

                PrepareNativeCalculationForSave(options);
                PrepareWorkbookForSave(options, applyCalculationPolicy: false);
                if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
                XlsbNewPackageWriter.Write(this, destination, options?.XlsbUseSharedStrings == true);
                if (destination.CanSeek) destination.SetLength(destination.Position);
                try { destination.Flush(); } catch (NotSupportedException) { }
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("New workbook streamed with the first-party BIFF12 XLSB writer.");
                return true;
            }

            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb(options);
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
                if (TryGetDirectTabularSaveSource(options, cancellationToken, "Xlsb", out ExcelDirectTabularSource source)) {
                    if (XlsbNewPackageWriter.TryWriteDirectTabular(this, source, destination, cancellationToken, options?.XlsbUseSharedStrings == true)) {
                        try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
                        LastSaveDiagnostics = ExcelSaveDiagnostics.NativeBinaryDirectPackage();
                        return true;
                    }

                    Execution.ReportInfo("Save.Xlsb.Direct skipped: the tabular values require the general BIFF12 value encoder.");
                }

                PrepareNativeCalculationForSave(options);
                PrepareWorkbookForSave(options, applyCalculationPolicy: false);
                cancellationToken.ThrowIfCancellationRequested();
                if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
                XlsbNewPackageWriter.Write(this, destination, options?.XlsbUseSharedStrings == true, cancellationToken);
                if (destination.CanSeek) destination.SetLength(destination.Position);
                try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("New workbook streamed with the first-party BIFF12 XLSB writer.");
                return true;
            }

            byte[] bytes = unchanged ? _xlsbOriginalPackageBytes! : RewriteNativeXlsb(options);
            PrepareDestinationStreamForWrite(destination);
            await destination.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            RefreshXlsbStateAfterNativeWrite(bytes);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(unchanged
                ? "Unmodified XLSB source copied byte-for-byte with all package parts preserved."
                : "XLSB worksheet cell records rewritten while all other package parts were preserved.");
            return true;
        }

        private byte[] RewriteNativeXlsb(ExcelSaveOptions? options, long maximumBytes = int.MaxValue) {
            var calculation = PrepareNativeCalculationForSave(options);
            maximumBytes = Math.Min(maximumBytes, Math.Min(GetPackageMaterializationLimit(options) ?? int.MaxValue, int.MaxValue));
            return XlsbNativePackageWriter.Rewrite(
                this,
                _xlsbOriginalPackageBytes ?? throw new InvalidOperationException("The XLSB source package is unavailable."),
                _xlsbAdvancedWorkbook ?? throw new InvalidOperationException("The XLSB source model is unavailable."), maximumBytes, calculation);
        }

        private byte[] GenerateNativeXlsb(
            ExcelSaveOptions? options,
            CancellationToken cancellationToken,
            out bool directWrite,
            out XlsbWorkbook generatedWorkbook) {
            directWrite = false;
            if (TryGetDirectTabularSaveSource(options, cancellationToken, "Xlsb", out ExcelDirectTabularSource source)) {
                using var directOutput = new MemoryStream();
                if (XlsbNewPackageWriter.TryWriteDirectTabular(this, source, directOutput, cancellationToken, options?.XlsbUseSharedStrings == true)) {
                    byte[] directBytes = directOutput.ToArray();
                    generatedWorkbook = XlsbWorkbookReader.Load(
                        directBytes,
                        new XlsbImportOptions { ReportPreservedRecords = false });
                    directWrite = true;
                    return directBytes;
                }

                Execution.ReportInfo("Save.Xlsb.Direct skipped: the tabular values require the general BIFF12 value encoder.");
            }

            PrepareNativeCalculationForSave(options);
            PrepareWorkbookForSave(options, applyCalculationPolicy: false);
            using var output = new MemoryStream();
            XlsbNewPackageWriter.Write(this, output, options?.XlsbUseSharedStrings == true, cancellationToken);
            byte[] bytes = output.ToArray();
            generatedWorkbook = XlsbWorkbookReader.Load(
                bytes,
                new XlsbImportOptions { ReportPreservedRecords = false });
            return bytes;
        }

        private void AdoptNativeXlsbAfterWrite(
            byte[] bytes,
            string? sourcePath,
            XlsbWorkbook? generatedWorkbook = null) {
            XlsbWorkbook workbook = generatedWorkbook ?? XlsbWorkbookReader.Load(bytes);
            RestoreNativeXlsbProjectionIdentity(workbook);
            if (workbook.Stylesheet != null) XlsbStylesheetProjector.Install(this, workbook.Stylesheet);
            MarkLoadedFromXlsb(sourcePath, workbook);
        }

        private void RestoreNativeXlsbProjectionIdentity(XlsbWorkbook workbook) {
            XlsbDefinedNameProjector.SynchronizeExisting(this, workbook.DefinedNames);
            if (workbook.VbaProjectPartName != null && WorkbookPartRoot.VbaProjectPart != null) {
                using var input = new MemoryStream(workbook.RelatedParts[workbook.VbaProjectPartName].Bytes, writable: false);
                WorkbookPartRoot.VbaProjectPart.FeedData(input);
            }
            if (_xlsbAdvancedWorkbook != null)
                foreach (var mapping in _xlsbAdvancedWorkbook.ProjectionPartUris) workbook.ProjectionPartUris[mapping.Key] = mapping.Value;
            foreach (var sheet in Sheets) {
                var native = workbook.Worksheets.FirstOrDefault(item => item.Name == sheet.Name);
                if (native == null) continue;
                var root = sheet.WorksheetPart.Worksheet!;
                if (native.UsedRange != null)
                    root.SheetDimension = new DocumentFormat.OpenXml.Spreadsheet.SheetDimension { Reference = native.UsedRange.ToA1Reference() };
                var nativeFormulas = native.Cells.Where(cell => cell.FormulaText != null).ToDictionary(cell => ExcelA1.CellReference(cell.Row, cell.Column));
                foreach (var cell in root.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(cell => cell.CellFormula != null))
                    if (nativeFormulas.TryGetValue(cell.CellReference?.Value ?? string.Empty, out var formula)) cell.CellFormula!.Text = formula.FormulaText!;
                var clearedOpaque = native.Cells.Where(cell => cell.FormulaBytes != null && cell.FormulaText == null && cell.Kind == XlsbCellValueKind.Text && cell.Value is string text && text.Length == 0)
                    .ToDictionary(cell => ExcelA1.CellReference(cell.Row, cell.Column));
                if (clearedOpaque.Count != 0) foreach (var cell in root.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                    if (clearedOpaque.ContainsKey(cell.CellReference?.Value ?? string.Empty) && cell.CellFormula == null) {
                        cell.InlineString = null; cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(string.Empty);
                    }
                string? codeName = native.Properties?.CodeName;
                if (string.IsNullOrEmpty(codeName)) continue;
                var properties = root.SheetProperties;
                if (properties == null) { properties = new DocumentFormat.OpenXml.Spreadsheet.SheetProperties(); root.AddChild(properties, true); }
                properties.CodeName = codeName;
            }
        }

        private void RefreshXlsbStateAfterNativeWrite(byte[] bytes, XlsbWorkbook? prepared = null) {
            XlsbWorkbook workbook = prepared ?? XlsbWorkbookReader.Load(bytes);
            RestoreNativeXlsbProjectionIdentity(workbook);
            if (workbook.Stylesheet != null) XlsbStylesheetProjector.Install(this, workbook.Stylesheet);
            _xlsbOriginalPackageBytes = bytes;
            _xlsbAdvancedWorkbook = workbook;
            _xlsbPreservationGuard = XlsbProjectionPreservationGuard.Capture(this);
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
