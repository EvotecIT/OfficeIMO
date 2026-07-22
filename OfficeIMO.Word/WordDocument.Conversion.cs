using OfficeIMO.Drawing.Internal;
using OfficeIMO.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static readonly string[] SupportedWordConversionExtensions = WordFormatCatalog.All
            .Select(format => format.Extension)
            .ToArray();

        /// <summary>
        /// Converts a Word file between DOC and DOCX and returns format and fidelity diagnostics.
        /// </summary>
        /// <param name="sourcePath">Path to the source DOC or DOCX file.</param>
        /// <param name="destinationPath">Path to the destination DOC or DOCX file.</param>
        /// <param name="options">Optional conversion policy settings.</param>
        /// <returns>The completed conversion result.</returns>
        public static WordDocumentConversionResult Convert(
            string sourcePath,
            string destinationPath,
            WordDocumentConversionOptions? options = null) {
            return ConvertAsync(sourcePath, destinationPath, options).GetAwaiter().GetResult();
        }

        /// <summary>Asynchronously converts a Word file between DOC and DOCX.</summary>
        /// <param name="sourcePath">Path to the source DOC or DOCX file.</param>
        /// <param name="destinationPath">Path to the destination DOC or DOCX file.</param>
        /// <param name="options">Optional conversion policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>The completed conversion result.</returns>
        public static async Task<WordDocumentConversionResult> ConvertAsync(
            string sourcePath,
            string destinationPath,
            WordDocumentConversionOptions? options = null,
            CancellationToken cancellationToken = default) {
            options ??= new WordDocumentConversionOptions();
            OfficeFileConversion.Paths paths = OfficeFileConversion.ValidatePaths(
                sourcePath,
                destinationPath,
                SupportedWordConversionExtensions,
                "Word document");

            using WordDocument document = await LoadWordConversionSourceAsync(paths.Source, options, cancellationToken).ConfigureAwait(false);
            WordFileFormat sourceFormat = document.SourceFormat;
            WordFileFormat destinationFormat = GetWordFormat(paths.Destination);
            OfficeFormatDescriptor sourceDescriptor = document.SourceFormatDescriptor;
            OfficeFormatDescriptor destinationDescriptor = WordFormatCatalog.GetByExtension(paths.Destination);
            OfficeCompatibilityMode compatibilityMode = GetCompatibilityMode(options);
            bool allowsLoss = AllowsLoss(options, compatibilityMode);
            List<WordConversionDiagnostic> diagnostics = CreateWordConversionDiagnostics(
                document,
                paths.Source,
                sourceDescriptor,
                destinationDescriptor,
                options,
                compatibilityMode,
                allowsLoss,
                out WordLegacyVisualFallbackPlan? visualFallback,
                out bool embedSourceCarrier);
            WordDocumentConversionResult assessment = CreateWordConversionResult(
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
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.SameFormat,
                    assessment,
                    $"The source is already {sourceDescriptor.Id}. Convert requires a different concrete source and destination format.");
            }

            if (diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss
                    && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked)) {
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.DataLossBlocked,
                    assessment,
                    $"Word conversion is blocked because {diagnostics.Count(diagnostic => diagnostic.RepresentsDataLoss && diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked)} source feature(s) have no representation accepted by the selected compatibility policy. Inspect Result.Report.Compatibility or select an explicit fallback policy.");
            }

            if (assessment.Report.Compatibility.HasBlockedFeatures) {
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.DestinationFeatureUnsupported,
                    assessment,
                    $"The document contains content that cannot be represented as {destinationDescriptor.Id} under {compatibilityMode}. Inspect Result.Report.Compatibility.Findings for the blocked feature.");
            }

            if (File.Exists(paths.Destination)
                && options.FileConflictPolicy == WordConversionFileConflictPolicy.FailIfExists) {
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.DestinationExists,
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
                        byte[] fallbackBytes = CreateLegacyVisualFallbackBytes(
                            visualFallback,
                            sourceDescriptor,
                            destinationDescriptor,
                            compatibilityMode,
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
                        WordSaveOptions conversionSaveOptions = (options.SaveOptions ?? new WordSaveOptions()).WithLossPolicy(
                            allowsLoss ? WordConversionLossPolicy.Allow : WordConversionLossPolicy.Block);
                        await document.SaveAsync(stagingPath, conversionSaveOptions, cancellationToken).ConfigureAwait(false);
                        if (embedSourceCarrier) {
                            byte[] destinationBytes = await OfficeFileConversion.ReadAllBytesAsync(
                                stagingPath,
                                cancellationToken).ConfigureAwait(false);
                            byte[] carriedBytes = AttachWordSourceCarrier(
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
                    diagnostics.Add(new WordConversionDiagnostic(
                        "Word.DestinationFeatureUnsupported",
                        WordConversionDiagnosticCategory.DestinationFormat,
                        WordConversionDiagnosticSeverity.Error,
                        exception.Message,
                        representsDataLoss: false));
                    throw new WordDocumentConversionException(
                        WordDocumentConversionFailureReason.DestinationFeatureUnsupported,
                        CreateWordConversionResult(paths, sourceFormat, destinationFormat, sourceDescriptor, destinationDescriptor, diagnostics, compatibilityMode, false, false),
                        $"The document contains content that cannot be written as {destinationFormat}. See Result.Report.Diagnostics for the specific unsupported feature.",
                        exception);
                }

                bool replacesExistingFile = File.Exists(paths.Destination);
                try {
                    cancellationToken.ThrowIfCancellationRequested();
                    OfficeFileCommit.CommitTemporaryFile(
                        stagingPath,
                        paths.Destination,
                        options.FileConflictPolicy == WordConversionFileConflictPolicy.Replace
                            ? OfficeFileCommit.ConflictPolicy.Replace
                            : OfficeFileCommit.ConflictPolicy.FailIfExists);
                    stagingPath = string.Empty;
                } catch (IOException exception) when (
                    options.FileConflictPolicy == WordConversionFileConflictPolicy.FailIfExists
                    && File.Exists(paths.Destination)) {
                    throw new WordDocumentConversionException(
                        WordDocumentConversionFailureReason.DestinationExists,
                        assessment,
                        $"The destination file '{paths.Destination}' was created while conversion was running and was not replaced.",
                        exception);
                }

                return CreateWordConversionResult(paths, sourceFormat, destinationFormat, sourceDescriptor, destinationDescriptor, diagnostics, compatibilityMode, true, replacesExistingFile);
            } finally {
                OfficeFileCommit.DeleteIfExists(stagingPath);
            }
        }

        private static async Task<WordDocument> LoadWordConversionSourceAsync(
            string sourcePath,
            WordDocumentConversionOptions options,
            CancellationToken cancellationToken) {
            if (options.LegacyDocImportOptions != null) {
                byte[] sourceBytes = await OfficeFileConversion.ReadAllBytesAsync(sourcePath, cancellationToken).ConfigureAwait(false);
                if (WordDocumentLoadRouting.IsLegacyDoc(sourceBytes, sourcePath)) {
                    LegacyDocImportOptions importOptions = CreateConversionImportOptions(options.LegacyDocImportOptions);
                    return LoadLegacyDocFromNormalFlow(
                        sourceBytes,
                        sourcePath,
                        saveOnDispose: false,
                        readOnly: false,
                        importOptions: importOptions);
                }
            }

            return await LoadAsync(
                sourcePath,
                new WordLoadOptions {
                    OverrideStyles = options.OverrideStyles,
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

        private static LegacyDocImportOptions CreateConversionImportOptions(LegacyDocImportOptions options) {
            return new LegacyDocImportOptions {
                MaxInputBytes = options.MaxInputBytes,
                MaxDecodedImageBytes = options.MaxDecodedImageBytes,
                // Conversion loss policy depends on complete unsupported-content discovery.
                ReportUnsupportedContent = true
            };
        }

        private static List<WordConversionDiagnostic> CreateWordConversionDiagnostics(
            WordDocument document,
            string sourcePath,
            OfficeFormatDescriptor detectedFormat,
            OfficeFormatDescriptor destinationFormat,
            WordDocumentConversionOptions options,
            OfficeCompatibilityMode compatibilityMode,
            bool allowsLoss,
            out WordLegacyVisualFallbackPlan? visualFallback,
            out bool embedSourceCarrier) {
            var diagnostics = new List<WordConversionDiagnostic>();
            bool preserveLossySource = allowsLoss
                && (options.EmbedSourceWhenLossy
                    || compatibilityMode == OfficeCompatibilityMode.PreservationOnly);
            OfficeFormatDescriptor declaredFormat = WordFormatCatalog.GetByExtension(sourcePath);
            if (!declaredFormat.Equals(detectedFormat)) {
                diagnostics.Add(new WordConversionDiagnostic(
                    "Word.SourceExtensionMismatch",
                    WordConversionDiagnosticCategory.SourceFormat,
                    WordConversionDiagnosticSeverity.Warning,
                    $"The source extension declares {declaredFormat.Id}, but its package declares {detectedFormat.Id}. Package content was used.",
                    representsDataLoss: false));
            }

            foreach (LegacyDocUnsupportedFeature feature in document.LegacyDocUnsupportedFeatures) {
                diagnostics.Add(CreateWordDataLossDiagnostic(
                    feature.Code,
                    feature.Description,
                    compatibilityMode,
                    preserveLossySource,
                    GetLegacyDocUnsupportedImpact(feature.Kind),
                    feature.EntryPath));
            }
            foreach (LegacyDocPreservedFeature feature in document.LegacyDocPreservedFeatures) {
                diagnostics.Add(CreateWordDataLossDiagnostic(
                    feature.Code,
                    feature.Description,
                    compatibilityMode,
                    preserveLossySource,
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier
                        | OfficeCompatibilityImpact.Editability));
            }
            foreach (LegacyDocCompoundFeature feature in document.LegacyDocCompoundFeatures) {
                diagnostics.Add(CreateWordDataLossDiagnostic(
                    feature.Code,
                    feature.Description,
                    compatibilityMode,
                    preserveLossySource,
                    GetLegacyDocCompoundImpact(feature.Kind),
                    feature.EntryPath));
            }

            if (document.HasMacros && !destinationFormat.IsMacroEnabled) {
                OfficeCompatibilityState macroState = GetWordGenericLossState(
                    compatibilityMode,
                    preserveLossySource);
                diagnostics.Add(new WordConversionDiagnostic(
                    "Word.VbaProject.Removed",
                    WordConversionDiagnosticCategory.DataLoss,
                    macroState == OfficeCompatibilityState.Blocked
                        ? WordConversionDiagnosticSeverity.Error
                        : WordConversionDiagnosticSeverity.Warning,
                    $"The source contains VBA, but {destinationFormat.Extension} cannot carry a VBA project.",
                    representsDataLoss: true,
                    macroState,
                    OfficeCompatibilityImpact.Behavioral | OfficeCompatibilityImpact.Carrier | OfficeCompatibilityImpact.Security,
                    sourceLocation: "word/vbaProject.bin"));
            }

            if (detectedFormat.Generation == OfficeFormatGeneration.Modern) {
                WordSignatureInfo signatures = document.InspectSignatures();
                if (signatures.HasSignatures) {
                    bool signatureInvalidationAllowed = options.SaveOptions?.SignedDocumentPolicy
                        == WordSignedDocumentSavePolicy.AllowSignatureInvalidation;
                    OfficeCompatibilityState signatureState = signatureInvalidationAllowed
                        ? GetWordGenericLossState(compatibilityMode, preserveLossySource)
                        : OfficeCompatibilityState.Blocked;
                    diagnostics.Add(new WordConversionDiagnostic(
                        "Word.DigitalSignature.Invalidated",
                        WordConversionDiagnosticCategory.DataLoss,
                        signatureState == OfficeCompatibilityState.Blocked
                            ? WordConversionDiagnosticSeverity.Error
                            : WordConversionDiagnosticSeverity.Warning,
                        signatureInvalidationAllowed
                            ? "Saving the converted package invalidates its existing digital signature. Signature markup may remain, but it must no longer be trusted."
                            : "The source carries digital-signature metadata and the configured save policy blocks conversion because rewriting the package can invalidate the signature.",
                        representsDataLoss: true,
                        signatureState,
                        OfficeCompatibilityImpact.Security | OfficeCompatibilityImpact.Carrier,
                        sourceLocation: "_xmlsignatures"));
                }
            }

            visualFallback = PlanLegacyVisualFallback(
                document,
                destinationFormat,
                compatibilityMode,
                options,
                diagnostics);

            bool hasBlockedFeature = diagnostics.Any(diagnostic =>
                diagnostic.CompatibilityState == OfficeCompatibilityState.Blocked);
            bool hasLossyFeature = diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss);
            embedSourceCarrier = visualFallback == null
                && preserveLossySource
                && !hasBlockedFeature
                && (compatibilityMode == OfficeCompatibilityMode.PreservationOnly || hasLossyFeature);
            if (visualFallback == null
                && !hasBlockedFeature
                && (hasLossyFeature || compatibilityMode == OfficeCompatibilityMode.PreservationOnly)) {
                AddWordSourceCarrierDiagnostic(
                    diagnostics,
                    embedSourceCarrier,
                    document.HasMacros);
            }

            return diagnostics
                .GroupBy(diagnostic => diagnostic.Code + "\0" + diagnostic.Message, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
        }

        private static WordConversionDiagnostic CreateWordDataLossDiagnostic(
            string code,
            string message,
            OfficeCompatibilityMode mode,
            bool embedSource,
            OfficeCompatibilityImpact impact,
            string? sourceLocation = null) {
            OfficeCompatibilityState state = GetWordGenericLossState(mode, embedSource);
            return new WordConversionDiagnostic(
                code,
                WordConversionDiagnosticCategory.DataLoss,
                state == OfficeCompatibilityState.Blocked
                    ? WordConversionDiagnosticSeverity.Error
                    : WordConversionDiagnosticSeverity.Warning,
                message,
                representsDataLoss: true,
                state,
                impact,
                sourceLocation);
        }

        private static OfficeCompatibilityState GetWordGenericLossState(
            OfficeCompatibilityMode mode,
            bool embedSource) => mode switch {
                OfficeCompatibilityMode.BestEffort => embedSource
                    ? OfficeCompatibilityState.EmbeddedSource
                    : OfficeCompatibilityState.Dropped,
                OfficeCompatibilityMode.PreservationOnly => OfficeCompatibilityState.EmbeddedSource,
                _ => OfficeCompatibilityState.Blocked
            };

        private static OfficeCompatibilityImpact GetLegacyDocUnsupportedImpact(
            LegacyDocUnsupportedFeatureKind kind) {
            OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.Semantic
                | OfficeCompatibilityImpact.Carrier
                | OfficeCompatibilityImpact.Editability;
            if (kind is LegacyDocUnsupportedFeatureKind.VbaProject
                or LegacyDocUnsupportedFeatureKind.ActiveXControl
                or LegacyDocUnsupportedFeatureKind.OleObject
                or LegacyDocUnsupportedFeatureKind.EmbeddedPackage) {
                impact |= OfficeCompatibilityImpact.Behavioral;
            }
            if (kind is LegacyDocUnsupportedFeatureKind.VbaProject
                or LegacyDocUnsupportedFeatureKind.ActiveXControl
                or LegacyDocUnsupportedFeatureKind.DigitalSignature) {
                impact |= OfficeCompatibilityImpact.Security;
            }
            return impact;
        }

        private static OfficeCompatibilityImpact GetLegacyDocCompoundImpact(
            LegacyDocCompoundFeatureKind kind) {
            OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.Semantic
                | OfficeCompatibilityImpact.Carrier
                | OfficeCompatibilityImpact.Editability;
            if (kind is LegacyDocCompoundFeatureKind.VbaProject
                or LegacyDocCompoundFeatureKind.ActiveXControl
                or LegacyDocCompoundFeatureKind.OleObject
                or LegacyDocCompoundFeatureKind.EmbeddedPackage) {
                impact |= OfficeCompatibilityImpact.Behavioral;
            }
            if (kind is LegacyDocCompoundFeatureKind.VbaProject
                or LegacyDocCompoundFeatureKind.ActiveXControl
                or LegacyDocCompoundFeatureKind.DigitalSignature) {
                impact |= OfficeCompatibilityImpact.Security;
            }
            return impact;
        }

        private static WordDocumentConversionResult CreateWordConversionResult(
            OfficeFileConversion.Paths paths,
            WordFileFormat sourceFormat,
            WordFileFormat destinationFormat,
            OfficeFormatDescriptor sourceDescriptor,
            OfficeFormatDescriptor destinationDescriptor,
            IReadOnlyList<WordConversionDiagnostic> diagnostics,
            OfficeCompatibilityMode compatibilityMode,
            bool outputCreated,
            bool replacedExistingFile) {
            return new WordDocumentConversionResult(
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

        private static OfficeCompatibilityMode GetCompatibilityMode(WordDocumentConversionOptions options) {
            if (options.CompatibilityMode != OfficeCompatibilityMode.StrictNative) return options.CompatibilityMode;
            return options.LossPolicy == WordConversionLossPolicy.Allow
                ? OfficeCompatibilityMode.BestEffort
                : OfficeCompatibilityMode.StrictNative;
        }

        private static bool AllowsLoss(
            WordDocumentConversionOptions options,
            OfficeCompatibilityMode mode) => options.LossPolicy == WordConversionLossPolicy.Allow
            || mode == OfficeCompatibilityMode.PreferEditable
            || mode == OfficeCompatibilityMode.PreferVisual
            || mode == OfficeCompatibilityMode.BestEffort
            || mode == OfficeCompatibilityMode.PreservationOnly;

        private static WordFileFormat GetWordFormat(string path) {
            return WordDocumentLoadRouting.HasLegacyDocExtension(path)
                ? WordFileFormat.Doc
                : WordFileFormat.Docx;
        }
    }
}
