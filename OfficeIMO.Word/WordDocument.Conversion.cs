using OfficeIMO.Shared;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static readonly string[] SupportedWordConversionExtensions = { ".doc", ".docx" };

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
            List<WordConversionDiagnostic> diagnostics = CreateWordConversionDiagnostics(document, paths.Source, sourceFormat);
            WordDocumentConversionResult assessment = CreateWordConversionResult(
                paths,
                sourceFormat,
                destinationFormat,
                diagnostics,
                outputCreated: false,
                replacedExistingFile: false);

            if (sourceFormat == destinationFormat) {
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.SameFormat,
                    assessment,
                    $"The source is already {sourceFormat}. Convert requires different physical source and destination formats.");
            }

            if (assessment.HasDataLoss && options.LossPolicy == WordConversionLossPolicy.Block) {
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.DataLossBlocked,
                    assessment,
                    $"Word conversion is blocked because {diagnostics.Count(diagnostic => diagnostic.RepresentsDataLoss)} source feature(s) would not survive conversion. Inspect Result.Diagnostics or set LossPolicy to Allow when that loss is intentional.");
            }

            if (File.Exists(paths.Destination)
                && options.FileConflictPolicy == WordConversionFileConflictPolicy.FailIfExists) {
                throw new WordDocumentConversionException(
                    WordDocumentConversionFailureReason.DestinationExists,
                    assessment,
                    $"The destination file '{paths.Destination}' already exists. Set FileConflictPolicy to Replace to replace it atomically.");
            }

            OfficeFileConversion.EnsureDestinationDirectory(paths.Destination);
            string stagingPath = OfficeFileCommit.CreateStagingPath(paths.Destination);
            try {
                try {
                    WordSaveOptions conversionSaveOptions = (options.SaveOptions ?? new WordSaveOptions()).WithLossPolicy(options.LossPolicy);
                    await document.SaveAsync(stagingPath, openWord: false, conversionSaveOptions, cancellationToken).ConfigureAwait(false);
                } catch (NotSupportedException exception) {
                    diagnostics.Add(new WordConversionDiagnostic(
                        "Word.DestinationFeatureUnsupported",
                        WordConversionDiagnosticCategory.DestinationFormat,
                        WordConversionDiagnosticSeverity.Error,
                        exception.Message,
                        representsDataLoss: false));
                    throw new WordDocumentConversionException(
                        WordDocumentConversionFailureReason.DestinationFeatureUnsupported,
                        CreateWordConversionResult(paths, sourceFormat, destinationFormat, diagnostics, false, false),
                        $"The document contains content that cannot be written as {destinationFormat}. See Result.Diagnostics for the specific unsupported feature.",
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

                if (options.OpenAfterSave) document.Open(paths.Destination, true);
                return CreateWordConversionResult(paths, sourceFormat, destinationFormat, diagnostics, true, replacesExistingFile);
            } finally {
                OfficeFileCommit.DeleteIfExists(stagingPath);
            }
        }

        private static async Task<WordDocument> LoadWordConversionSourceAsync(
            string sourcePath,
            WordDocumentConversionOptions options,
            CancellationToken cancellationToken) {
            if (options.LegacyDocImportOptions != null && WordDocumentLoadRouting.HasLegacyDocExtension(sourcePath)) {
                return LoadLegacyDoc(sourcePath, options.LegacyDocImportOptions);
            }

            return await LoadAsync(
                sourcePath,
                readOnly: false,
                autoSave: false,
                overrideStyles: options.OverrideStyles,
                openSettings: options.OpenSettings,
                cancellationToken).ConfigureAwait(false);
        }

        private static List<WordConversionDiagnostic> CreateWordConversionDiagnostics(
            WordDocument document,
            string sourcePath,
            WordFileFormat detectedFormat) {
            var diagnostics = new List<WordConversionDiagnostic>();
            WordFileFormat declaredFormat = GetWordFormat(sourcePath);
            if (declaredFormat != detectedFormat) {
                diagnostics.Add(new WordConversionDiagnostic(
                    "Word.SourceExtensionMismatch",
                    WordConversionDiagnosticCategory.SourceFormat,
                    WordConversionDiagnosticSeverity.Warning,
                    $"The source extension declares {declaredFormat}, but its content is {detectedFormat}. Content detection was used.",
                    representsDataLoss: false));
            }

            foreach (LegacyDocUnsupportedFeature feature in document.LegacyDocUnsupportedFeatures) {
                diagnostics.Add(CreateWordDataLossDiagnostic(feature.Code, feature.Description));
            }
            foreach (LegacyDocPreservedFeature feature in document.LegacyDocPreservedFeatures) {
                diagnostics.Add(CreateWordDataLossDiagnostic(feature.Code, feature.Description));
            }
            foreach (LegacyDocCompoundFeature feature in document.LegacyDocCompoundFeatures) {
                diagnostics.Add(CreateWordDataLossDiagnostic(feature.Code, feature.Description));
            }

            return diagnostics
                .GroupBy(diagnostic => diagnostic.Code + "\0" + diagnostic.Message, StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
        }

        private static WordConversionDiagnostic CreateWordDataLossDiagnostic(string code, string message) {
            return new WordConversionDiagnostic(
                code,
                WordConversionDiagnosticCategory.DataLoss,
                WordConversionDiagnosticSeverity.Warning,
                message,
                representsDataLoss: true);
        }

        private static WordDocumentConversionResult CreateWordConversionResult(
            OfficeFileConversion.Paths paths,
            WordFileFormat sourceFormat,
            WordFileFormat destinationFormat,
            IReadOnlyList<WordConversionDiagnostic> diagnostics,
            bool outputCreated,
            bool replacedExistingFile) {
            return new WordDocumentConversionResult(
                paths.Source,
                paths.Destination,
                sourceFormat,
                destinationFormat,
                diagnostics.ToArray(),
                outputCreated,
                replacedExistingFile);
        }

        private static WordFileFormat GetWordFormat(string path) {
            return string.Equals(Path.GetExtension(path), ".doc", StringComparison.OrdinalIgnoreCase)
                ? WordFileFormat.Doc
                : WordFileFormat.Docx;
        }
    }
}
