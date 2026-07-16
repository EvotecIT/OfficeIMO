using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly string[] SupportedExcelConversionExtensions = { ".xls", ".xlsb", ".xlsx" };

        /// <summary>
        /// Converts an Excel file between XLS and XLSX and returns format and fidelity diagnostics.
        /// </summary>
        /// <param name="sourcePath">Path to the source XLS or XLSX file.</param>
        /// <param name="destinationPath">Path to the destination XLS or XLSX file.</param>
        /// <param name="options">Optional conversion policy settings.</param>
        /// <returns>The completed conversion result.</returns>
        public static ExcelDocumentConversionResult Convert(
            string sourcePath,
            string destinationPath,
            ExcelDocumentConversionOptions? options = null) {
            return ConvertAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();
        }

        /// <summary>Asynchronously converts an Excel file between XLS and XLSX.</summary>
        /// <param name="sourcePath">Path to the source XLS or XLSX file.</param>
        /// <param name="destinationPath">Path to the destination XLS or XLSX file.</param>
        /// <param name="options">Optional conversion policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>The completed conversion result.</returns>
        public static async Task<ExcelDocumentConversionResult> ConvertAsync(
            string sourcePath,
            string destinationPath,
            ExcelDocumentConversionOptions? options = null,
            CancellationToken cancellationToken = default) {
            options ??= new ExcelDocumentConversionOptions();
            OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
                sourcePath,
                destinationPath,
                SupportedExcelConversionExtensions,
                "Excel workbook");

            using ExcelDocument document = await LoadExcelConversionSourceAsync(paths.Source, options, cancellationToken).ConfigureAwait(false);
            ExcelFileFormat sourceFormat = document.SourceFormat;
            ExcelFileFormat destinationFormat = GetExcelFormat(paths.Destination);
            List<ExcelConversionDiagnostic> diagnostics = CreateExcelConversionDiagnostics(document, paths.Source, sourceFormat);
            ExcelDocumentConversionResult assessment = CreateExcelConversionResult(
                paths,
                sourceFormat,
                destinationFormat,
                diagnostics,
                outputCreated: false,
                replacedExistingFile: false);

            if (sourceFormat == destinationFormat) {
                throw new ExcelDocumentConversionException(
                    ExcelDocumentConversionFailureReason.SameFormat,
                    assessment,
                    $"The source is already {sourceFormat}. Convert requires different physical source and destination formats.");
            }

            if (assessment.HasLoss && options.LossPolicy == ExcelConversionLossPolicy.Block) {
                throw new ExcelDocumentConversionException(
                    ExcelDocumentConversionFailureReason.DataLossBlocked,
                    assessment,
                    $"Excel conversion is blocked because {diagnostics.Count(diagnostic => diagnostic.RepresentsDataLoss)} source feature(s) would not survive conversion. Inspect Result.Report.Diagnostics or set LossPolicy to Allow when that loss is intentional.");
            }

            if (File.Exists(paths.Destination)
                && options.FileConflictPolicy == ExcelConversionFileConflictPolicy.FailIfExists) {
                throw new ExcelDocumentConversionException(
                    ExcelDocumentConversionFailureReason.DestinationExists,
                    assessment,
                    $"The destination file '{paths.Destination}' already exists. Set FileConflictPolicy to Replace to replace it atomically.");
            }

            EnsureDestinationFileWritable(paths.Destination);

            OfficeFileConversion.EnsureDestinationDirectory(paths.Destination);
            string stagingPath = OfficeFileCommit.CreateStagingPath(paths.Destination);
            try {
                try {
                    ExcelSaveOptions conversionSaveOptions = (options.SaveOptions ?? new ExcelSaveOptions()).WithLossPolicy(options.LossPolicy);
                    await document.SaveAsync(stagingPath, conversionSaveOptions, cancellationToken).ConfigureAwait(false);
                } catch (NotSupportedException exception) {
                    diagnostics.Add(new ExcelConversionDiagnostic(
                        "Excel.DestinationFeatureUnsupported",
                        ExcelConversionDiagnosticCategory.DestinationFormat,
                        ExcelConversionDiagnosticSeverity.Error,
                        exception.Message,
                        representsDataLoss: false));
                    throw new ExcelDocumentConversionException(
                        ExcelDocumentConversionFailureReason.DestinationFeatureUnsupported,
                        CreateExcelConversionResult(paths, sourceFormat, destinationFormat, diagnostics, false, false),
                        $"The workbook contains content that cannot be written as {destinationFormat}. See Result.Report.Diagnostics for the specific unsupported feature.",
                        exception);
                }

                bool replacesExistingFile = File.Exists(paths.Destination);
                try {
                    cancellationToken.ThrowIfCancellationRequested();
                    OfficeFileCommit.CommitTemporaryFile(
                        stagingPath,
                        paths.Destination,
                        options.FileConflictPolicy == ExcelConversionFileConflictPolicy.Replace
                            ? OfficeFileCommit.ConflictPolicy.Replace
                            : OfficeFileCommit.ConflictPolicy.FailIfExists);
                    stagingPath = string.Empty;
                } catch (IOException exception) when (
                    options.FileConflictPolicy == ExcelConversionFileConflictPolicy.FailIfExists
                    && File.Exists(paths.Destination)) {
                    throw new ExcelDocumentConversionException(
                        ExcelDocumentConversionFailureReason.DestinationExists,
                        assessment,
                        $"The destination file '{paths.Destination}' was created while conversion was running and was not replaced.",
                        exception);
                }

                return CreateExcelConversionResult(paths, sourceFormat, destinationFormat, diagnostics, true, replacesExistingFile);
            } finally {
                OfficeFileCommit.DeleteIfExists(stagingPath);
            }
        }

        private static async Task<ExcelDocument> LoadExcelConversionSourceAsync(
            string sourcePath,
            ExcelDocumentConversionOptions options,
            CancellationToken cancellationToken) {
            if (options.LegacyXlsImportOptions != null) {
                byte[] sourceBytes = await OfficeFileConversion.ReadAllBytesAsync(sourcePath, cancellationToken).ConfigureAwait(false);
                if (ExcelDocumentLoadRouting.IsLegacyXls(sourceBytes, sourcePath)) {
                    LegacyXlsImportOptions importOptions = CreateConversionImportOptions(options.LegacyXlsImportOptions);
                    return LoadLegacyXlsFromNormalFlow(
                        sourceBytes,
                        readOnly: false,
                        saveOnDispose: false,
                        filePath: sourcePath,
                        importOptions: importOptions);
                }
            }

            return await LoadAsync(
                sourcePath,
                new ExcelLoadOptions {
                    OpenSettings = CreateConversionOpenSettings(options.OpenSettings)
                },
                cancellationToken).ConfigureAwait(false);
        }

        private static OpenSettings? CreateConversionOpenSettings(OpenSettings? openSettings) {
            if (openSettings == null) {
                return null;
            }

            return new OpenSettings {
                AutoSave = false,
                CompatibilityLevel = openSettings.CompatibilityLevel,
                MarkupCompatibilityProcessSettings = openSettings.MarkupCompatibilityProcessSettings,
                MaxCharactersInPart = openSettings.MaxCharactersInPart,
            };
        }

        private static LegacyXlsImportOptions CreateConversionImportOptions(LegacyXlsImportOptions options) {
            return new LegacyXlsImportOptions {
                MaxInputBytes = options.MaxInputBytes,
                Password = options.Password,
                // Conversion loss policy depends on complete unsupported-content discovery.
                ReportUnsupportedContent = true
            };
        }

        private static List<ExcelConversionDiagnostic> CreateExcelConversionDiagnostics(
            ExcelDocument document,
            string sourcePath,
            ExcelFileFormat detectedFormat) {
            var diagnostics = new List<ExcelConversionDiagnostic>();
            ExcelFileFormat declaredFormat = GetExcelFormat(sourcePath);
            if (declaredFormat != detectedFormat) {
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.SourceExtensionMismatch",
                    ExcelConversionDiagnosticCategory.SourceFormat,
                    ExcelConversionDiagnosticSeverity.Warning,
                    $"The source extension declares {declaredFormat}, but its content is {detectedFormat}. Content detection was used.",
                    representsDataLoss: false));
            }

            foreach (LegacyXlsUnsupportedFeature feature in document.LegacyXlsUnsupportedFeatures) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(feature.Code, feature.Description));
            }
            foreach (LegacyXlsPreservedFeatureRecord feature in document.LegacyXlsPreservedFeatures) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(feature.Code, feature.Description));
            }
            foreach (LegacyXlsUnsupportedSheet sheet in document.LegacyXlsUnsupportedSheets) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(
                    $"Excel.LegacyXls.UnsupportedSheet.{sheet.Kind}",
                    $"Legacy sheet '{sheet.Name}' ({sheet.Kind}) was not projected as a normal worksheet."));
            }
            foreach (LegacyXlsCompoundFeatureRecord feature in document.LegacyXlsCompoundFeatures.Where(IsLossyExcelCompoundFeature)) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(
                    $"Excel.LegacyXls.Compound.{feature.Kind}",
                    $"Legacy compound feature '{feature.Kind}' with {feature.Entries.Count} entr{(feature.Entries.Count == 1 ? "y" : "ies")} is not projected into XLSX."));
            }

            return diagnostics
                .GroupBy(diagnostic => diagnostic.Code + "\0" + diagnostic.Message, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
        }

        private static bool IsLossyExcelCompoundFeature(LegacyXlsCompoundFeatureRecord feature) {
            return feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject
                || feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject;
        }

        private static ExcelConversionDiagnostic CreateExcelDataLossDiagnostic(string code, string message) {
            return new ExcelConversionDiagnostic(
                code,
                ExcelConversionDiagnosticCategory.DataLoss,
                ExcelConversionDiagnosticSeverity.Warning,
                message,
                representsDataLoss: true);
        }

        private static ExcelDocumentConversionResult CreateExcelConversionResult(
            OfficeFileConversion.Paths paths,
            ExcelFileFormat sourceFormat,
            ExcelFileFormat destinationFormat,
            IReadOnlyList<ExcelConversionDiagnostic> diagnostics,
            bool outputCreated,
            bool replacedExistingFile) {
            return new ExcelDocumentConversionResult(
                paths.Source,
                paths.Destination,
                sourceFormat,
                destinationFormat,
                diagnostics.ToArray(),
                outputCreated,
                replacedExistingFile);
        }

        private static ExcelFileFormat GetExcelFormat(string path) {
            string extension = Path.GetExtension(path);
            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)) {
                return ExcelFileFormat.Xls;
            }

            return string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase)
                ? ExcelFileFormat.Xlsb
                : ExcelFileFormat.Xlsx;
        }
    }
}
