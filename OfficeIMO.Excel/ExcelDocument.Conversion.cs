using OfficeIMO.Drawing.Internal;
using OfficeIMO.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Xlsb;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly string[] SupportedExcelConversionExtensions = ExcelFormatCatalog.All
            .Select(format => format.Extension)
            .ToArray();

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
            OfficeFormatDescriptor sourceDescriptor = document.SourceFormatDescriptor;
            OfficeFormatDescriptor destinationDescriptor = ExcelFormatCatalog.GetByExtension(paths.Destination);
            OfficeCompatibilityMode compatibilityMode = GetCompatibilityMode(options);
            bool allowsLoss = AllowsLoss(options, compatibilityMode);
            List<ExcelConversionDiagnostic> diagnostics = CreateExcelConversionDiagnostics(
                document,
                paths.Source,
                sourceDescriptor,
                destinationDescriptor,
                options,
                compatibilityMode,
                allowsLoss,
                out ExcelVisualFallbackPlan? visualFallback,
                out bool embedSourceCarrier);
            ExcelDocumentConversionResult assessment = CreateExcelConversionResult(
                paths,
                sourceFormat,
                destinationFormat,
                sourceDescriptor,
                destinationDescriptor,
                diagnostics,
                compatibilityMode,
                outputCreated: false,
                replacedExistingFile: false);

            if (sourceDescriptor.Equals(destinationDescriptor)) {
                throw new ExcelDocumentConversionException(
                    ExcelDocumentConversionFailureReason.SameFormat,
                    assessment,
                    $"The source is already {sourceDescriptor.Id}. Convert requires a different concrete source and destination format.");
            }

            if (diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss
                    && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked)) {
                throw new ExcelDocumentConversionException(
                    ExcelDocumentConversionFailureReason.DataLossBlocked,
                    assessment,
                    $"Excel conversion is blocked because {diagnostics.Count(diagnostic => diagnostic.RepresentsDataLoss && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked)} source feature(s) have no representation accepted by the selected compatibility policy. Inspect Result.Report.Compatibility or select an explicit fallback policy.");
            }

            if (assessment.Report.Compatibility.HasBlockedFeatures) {
                throw new ExcelDocumentConversionException(
                    ExcelDocumentConversionFailureReason.DestinationFeatureUnsupported,
                    assessment,
                    $"The requested destination {destinationDescriptor.Id} is classified but is not a supported native write target. Inspect Result.Report.Compatibility for details.");
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
                    bool requiresSourceBytes = visualFallback?.EmbedSource == true || embedSourceCarrier;
                    byte[]? sourceBytes = requiresSourceBytes
                        ? await OfficeFileConversion.ReadAllBytesAsync(paths.Source, cancellationToken).ConfigureAwait(false)
                        : null;
                    if (visualFallback != null) {
                        byte[] fallbackBytes = CreateExcelVisualFallbackBytes(
                            visualFallback,
                            sourceDescriptor,
                            destinationDescriptor,
                            compatibilityMode,
                            options,
                            paths.Source,
                            sourceBytes);
                        await OfficeFileCommit.WriteAllBytesAsync(
                            stagingPath,
                            fallbackBytes,
                            cancellationToken: cancellationToken).ConfigureAwait(false);
                    } else {
                        if (document.HasMacros && !destinationDescriptor.IsMacroEnabled && allowsLoss) {
                            document.RemoveMacros();
                        }
                        ExcelSaveOptions conversionSaveOptions = (options.SaveOptions ?? new ExcelSaveOptions()).WithLossPolicy(
                            allowsLoss ? ExcelConversionLossPolicy.Allow : ExcelConversionLossPolicy.Block);
                        await document.SaveAsync(stagingPath, conversionSaveOptions, cancellationToken).ConfigureAwait(false);
                        if (embedSourceCarrier) {
                            byte[] destinationBytes = await OfficeFileConversion.ReadAllBytesAsync(stagingPath, cancellationToken).ConfigureAwait(false);
                            byte[] carriedBytes = AttachExcelSourceCarrier(
                                destinationBytes,
                                destinationDescriptor,
                                sourceDescriptor.Id,
                                Path.GetFileName(paths.Source),
                                compatibilityMode,
                                sourceBytes ?? throw new InvalidOperationException("Embedded-source conversion requires source bytes."));
                            await OfficeFileCommit.WriteAllBytesAsync(
                                stagingPath,
                                carriedBytes,
                                cancellationToken: cancellationToken).ConfigureAwait(false);
                        }
                    }
                } catch (NotSupportedException exception) {
                    diagnostics.Add(new ExcelConversionDiagnostic(
                        "Excel.DestinationFeatureUnsupported",
                        ExcelConversionDiagnosticCategory.DestinationFormat,
                        ExcelConversionDiagnosticSeverity.Error,
                        exception.Message,
                        representsDataLoss: false));
                    throw new ExcelDocumentConversionException(
                        ExcelDocumentConversionFailureReason.DestinationFeatureUnsupported,
                        CreateExcelConversionResult(paths, sourceFormat, destinationFormat, sourceDescriptor, destinationDescriptor, diagnostics, compatibilityMode, false, false),
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

                return CreateExcelConversionResult(paths, sourceFormat, destinationFormat, sourceDescriptor, destinationDescriptor, diagnostics, compatibilityMode, true, replacesExistingFile);
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
            OfficeFormatDescriptor detectedFormat,
            OfficeFormatDescriptor destinationFormat,
            ExcelDocumentConversionOptions options,
            OfficeCompatibilityMode compatibilityMode,
            bool allowsLoss,
            out ExcelVisualFallbackPlan? visualFallback,
            out bool embedSourceCarrier) {
            var diagnostics = new List<ExcelConversionDiagnostic>();
            bool preserveLossySource = allowsLoss
                && (options.EmbedSourceWhenLossy || compatibilityMode == OfficeCompatibilityMode.PreservationOnly);
            OfficeFormatDescriptor declaredFormat = ExcelFormatCatalog.GetByExtension(sourcePath);
            if (!declaredFormat.Equals(detectedFormat)) {
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.SourceExtensionMismatch",
                    ExcelConversionDiagnosticCategory.SourceFormat,
                    ExcelConversionDiagnosticSeverity.Warning,
                    $"The source extension declares {declaredFormat.Id}, but its package declares {detectedFormat.Id}. Package content was used.",
                    representsDataLoss: false));
            }

            foreach (LegacyXlsUnsupportedFeature feature in document.LegacyXlsUnsupportedFeatures) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(
                    feature.Code,
                    feature.Description,
                    compatibilityMode,
                    preserveLossySource,
                    GetLegacyXlsUnsupportedImpact(feature.Kind),
                    feature.SheetName));
            }
            foreach (LegacyXlsPreservedFeatureRecord feature in document.LegacyXlsPreservedFeatures) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(
                    feature.Code,
                    feature.Description,
                    compatibilityMode,
                    preserveLossySource,
                    GetLegacyXlsUnsupportedImpact(feature.Kind),
                    feature.SheetName));
            }
            foreach (LegacyXlsUnsupportedSheet sheet in document.LegacyXlsUnsupportedSheets) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(
                    $"Excel.LegacyXls.UnsupportedSheet.{sheet.Kind}",
                    $"Legacy sheet '{sheet.Name}' ({sheet.Kind}) was not projected as a normal worksheet.",
                    compatibilityMode,
                    preserveLossySource,
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier
                        | OfficeCompatibilityImpact.Editability,
                    sheet.Name));
            }
            foreach (LegacyXlsCompoundFeatureRecord feature in document.LegacyXlsCompoundFeatures.Where(IsLossyExcelCompoundFeature)) {
                diagnostics.Add(CreateExcelDataLossDiagnostic(
                    $"Excel.LegacyXls.Compound.{feature.Kind}",
                    $"Legacy compound feature '{feature.Kind}' with {feature.Entries.Count} entr{(feature.Entries.Count == 1 ? "y" : "ies")} is not projected into XLSX.",
                    compatibilityMode,
                    preserveLossySource,
                    GetLegacyXlsCompoundImpact(feature.Kind),
                    feature.Entries.FirstOrDefault()));
            }

            foreach (XlsbImportDiagnostic diagnostic in document.XlsbImportDiagnostics) {
                bool representsLoss = diagnostic.Severity != XlsbImportDiagnosticSeverity.Information;
                diagnostics.Add(new ExcelConversionDiagnostic(
                    diagnostic.Code,
                    representsLoss ? ExcelConversionDiagnosticCategory.DataLoss : ExcelConversionDiagnosticCategory.SourceFormat,
                    diagnostic.Severity == XlsbImportDiagnosticSeverity.Error
                        ? ExcelConversionDiagnosticSeverity.Error
                        : diagnostic.Severity == XlsbImportDiagnosticSeverity.Warning
                            ? ExcelConversionDiagnosticSeverity.Warning
                            : ExcelConversionDiagnosticSeverity.Information,
                    diagnostic.Message,
                    representsLoss,
                    representsLoss
                        ? GetExcelGenericLossState(compatibilityMode, preserveLossySource)
                        : OfficeCompatibilityState.PreservedOpaque,
                    representsLoss
                        ? OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability
                        : OfficeCompatibilityImpact.None,
                    CreateXlsbSourceLocation(diagnostic.PartName, diagnostic.RecordOffset)));
            }
            foreach (XlsbPreservedRecordInfo record in document.XlsbPreservedRecords) {
                diagnostics.Add(new ExcelConversionDiagnostic(
                    $"Excel.Xlsb.UnprojectedRecord.0x{record.RecordType:X}",
                    ExcelConversionDiagnosticCategory.DataLoss,
                    ExcelConversionDiagnosticSeverity.Warning,
                    $"BIFF12 record 0x{record.RecordType:X} in '{record.PartName}' at offset {record.Offset} is retained in the XLSB source but is not projected into the editable workbook model or the converted target.",
                    representsDataLoss: true,
                    GetExcelGenericLossState(compatibilityMode, preserveLossySource),
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Editability,
                    CreateXlsbSourceLocation(record.PartName, record.Offset)));
            }

            if (document.HasMacros && !destinationFormat.IsMacroEnabled) {
                OfficeCompatibilityState macroState = GetExcelGenericLossState(
                    compatibilityMode,
                    preserveLossySource);
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.VbaProject.Removed",
                    ExcelConversionDiagnosticCategory.DataLoss,
                    macroState == OfficeCompatibilityState.Blocked
                        ? ExcelConversionDiagnosticSeverity.Error
                        : ExcelConversionDiagnosticSeverity.Warning,
                    $"The source contains VBA, but {destinationFormat.Extension} cannot carry a VBA project.",
                    representsDataLoss: true,
                    macroState,
                    OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Security,
                    sourceLocation: "xl/vbaProject.bin"));
            }

            if (document._legacyXlsWasEncryptedSource
                && destinationFormat.Generation == OfficeFormatGeneration.Modern) {
                OfficeCompatibilityState encryptionState = GetExcelGenericLossState(
                    compatibilityMode,
                    preserveLossySource);
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.PasswordEncryption.Removed",
                    ExcelConversionDiagnosticCategory.DataLoss,
                    encryptionState == OfficeCompatibilityState.Blocked
                        ? ExcelConversionDiagnosticSeverity.Error
                        : ExcelConversionDiagnosticSeverity.Warning,
                    "The legacy workbook was decrypted for import, but the requested modern destination is not password-encrypted. Save the converted artifact with an explicit encryption API if confidentiality must continue.",
                    representsDataLoss: true,
                    encryptionState,
                    OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier,
                    sourceLocation: "Workbook/FilePass"));
            }

            if (destinationFormat.Generation == OfficeFormatGeneration.Legacy
                && !string.Equals(destinationFormat.Extension, ".xls", StringComparison.Ordinal)) {
                diagnostics.Add(new ExcelConversionDiagnostic(
                    "Excel.LegacyDestination.NotWritable",
                    ExcelConversionDiagnosticCategory.DestinationFormat,
                    ExcelConversionDiagnosticSeverity.Error,
                    $"{destinationFormat.Extension} is classified for import and reporting, but native output is currently limited to .xls among legacy Excel formats.",
                    representsDataLoss: false,
                    OfficeCompatibilityState.Blocked,
                    OfficeCompatibilityImpact.None));
            }

            visualFallback = PlanExcelVisualFallback(
                document,
                destinationFormat,
                compatibilityMode,
                options,
                diagnostics);
            embedSourceCarrier = visualFallback == null
                && preserveLossySource
                && (compatibilityMode == OfficeCompatibilityMode.PreservationOnly
                    || diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss));
            if (embedSourceCarrier) {
                AddExcelSourceCarrierDiagnostic(
                    diagnostics,
                    embedded: true,
                    document.HasMacros,
                    visualFallback: false);
            }

            return diagnostics
                .GroupBy(diagnostic => diagnostic.Code + "\0" + diagnostic.Message, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
        }

        private static bool IsLossyExcelCompoundFeature(LegacyXlsCompoundFeatureRecord feature) {
            return feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject
                || feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject
                || feature.Kind == LegacyXlsCompoundFeatureRecordKind.DigitalSignature;
        }

        private static string? CreateXlsbSourceLocation(string? partName, long? offset) {
            if (string.IsNullOrWhiteSpace(partName)) return offset.HasValue ? $"offset:{offset.Value}" : null;
            return offset.HasValue ? $"{partName}@{offset.Value}" : partName;
        }

        private static ExcelConversionDiagnostic CreateExcelDataLossDiagnostic(
            string code,
            string message,
            OfficeCompatibilityMode mode,
            bool embedSource,
            OfficeCompatibilityImpact impact,
            string? sourceLocation = null) {
            OfficeCompatibilityState state = GetExcelGenericLossState(mode, embedSource);
            return new ExcelConversionDiagnostic(
                code,
                ExcelConversionDiagnosticCategory.DataLoss,
                state == OfficeCompatibilityState.Blocked
                    ? ExcelConversionDiagnosticSeverity.Error
                    : ExcelConversionDiagnosticSeverity.Warning,
                message,
                representsDataLoss: true,
                state,
                impact,
                sourceLocation);
        }

        private static OfficeCompatibilityState GetExcelGenericLossState(
            OfficeCompatibilityMode mode,
            bool embedSource) => mode switch {
                OfficeCompatibilityMode.BestEffort => embedSource
                    ? OfficeCompatibilityState.EmbeddedSource
                    : OfficeCompatibilityState.Dropped,
                OfficeCompatibilityMode.PreservationOnly => OfficeCompatibilityState.EmbeddedSource,
                _ => OfficeCompatibilityState.Blocked
            };

        private static OfficeCompatibilityImpact GetLegacyXlsUnsupportedImpact(
            LegacyXlsUnsupportedFeatureKind kind) {
            OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.Semantic
                | OfficeCompatibilityImpact.Carrier
                | OfficeCompatibilityImpact.Editability;
            if (kind is LegacyXlsUnsupportedFeatureKind.VbaProject
                or LegacyXlsUnsupportedFeatureKind.OleObject
                or LegacyXlsUnsupportedFeatureKind.DrawingObject) {
                impact |= OfficeCompatibilityImpact.Behavioral;
            }
            if (kind is LegacyXlsUnsupportedFeatureKind.VbaProject
                or LegacyXlsUnsupportedFeatureKind.OleObject
                or LegacyXlsUnsupportedFeatureKind.DigitalSignature
                or LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook) {
                impact |= OfficeCompatibilityImpact.Security;
            }
            return impact;
        }

        private static OfficeCompatibilityImpact GetLegacyXlsCompoundImpact(
            LegacyXlsCompoundFeatureRecordKind kind) {
            OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.Semantic
                | OfficeCompatibilityImpact.Carrier
                | OfficeCompatibilityImpact.Editability;
            if (kind is LegacyXlsCompoundFeatureRecordKind.VbaProject
                or LegacyXlsCompoundFeatureRecordKind.OleObject) {
                impact |= OfficeCompatibilityImpact.Behavioral;
            }
            if (kind is LegacyXlsCompoundFeatureRecordKind.VbaProject
                or LegacyXlsCompoundFeatureRecordKind.OleObject
                or LegacyXlsCompoundFeatureRecordKind.DigitalSignature) {
                impact |= OfficeCompatibilityImpact.Security;
            }
            return impact;
        }

        private static ExcelDocumentConversionResult CreateExcelConversionResult(
            OfficeFileConversion.Paths paths,
            ExcelFileFormat sourceFormat,
            ExcelFileFormat destinationFormat,
            OfficeFormatDescriptor sourceDescriptor,
            OfficeFormatDescriptor destinationDescriptor,
            IReadOnlyList<ExcelConversionDiagnostic> diagnostics,
            OfficeCompatibilityMode compatibilityMode,
            bool outputCreated,
            bool replacedExistingFile) {
            return new ExcelDocumentConversionResult(
                paths.Source,
                paths.Destination,
                sourceFormat,
                destinationFormat,
                sourceDescriptor,
                destinationDescriptor,
                diagnostics.ToArray(),
                compatibilityMode,
                outputCreated,
                replacedExistingFile);
        }

        private static OfficeCompatibilityMode GetCompatibilityMode(ExcelDocumentConversionOptions options) {
            if (options.CompatibilityMode != OfficeCompatibilityMode.StrictNative) return options.CompatibilityMode;
            return options.LossPolicy == ExcelConversionLossPolicy.Allow
                ? OfficeCompatibilityMode.BestEffort
                : OfficeCompatibilityMode.StrictNative;
        }

        private static bool AllowsLoss(
            ExcelDocumentConversionOptions options,
            OfficeCompatibilityMode mode) => options.LossPolicy == ExcelConversionLossPolicy.Allow
            || mode == OfficeCompatibilityMode.PreferEditable
            || mode == OfficeCompatibilityMode.PreferVisual
            || mode == OfficeCompatibilityMode.BestEffort
            || mode == OfficeCompatibilityMode.PreservationOnly;

        private static ExcelFileFormat GetExcelFormat(string path) {
            OfficeFormatDescriptor descriptor = ExcelFormatCatalog.GetByExtension(path);
            if (descriptor.Encoding == OfficeFormatEncoding.CompoundBinary) return ExcelFileFormat.Xls;
            return descriptor.Encoding == OfficeFormatEncoding.BinaryOpenXml
                ? ExcelFileFormat.Xlsb
                : ExcelFileFormat.Xlsx;
        }
    }
}
