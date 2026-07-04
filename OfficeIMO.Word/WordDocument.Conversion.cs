using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word.LegacyDoc;
using System.Runtime.InteropServices;

namespace OfficeIMO.Word {
    /// <summary>
    /// Controls file-to-file Word document conversion.
    /// </summary>
    public sealed class WordDocumentConversionOptions {
        /// <summary>
        /// Gets or sets whether an existing destination file may be overwritten. Defaults to <c>true</c>.
        /// </summary>
        public bool Overwrite { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to open Word after saving the converted file.
        /// </summary>
        public bool OpenWord { get; set; }

        /// <summary>
        /// Gets or sets whether OfficeIMO should override styles while loading Open XML sources.
        /// </summary>
        public bool OverrideStyles { get; set; }

        /// <summary>
        /// Gets or sets optional Open XML load settings for `.docx` sources.
        /// </summary>
        public OpenSettings? OpenSettings { get; set; }

        /// <summary>
        /// Gets or sets optional legacy `.doc` import settings.
        /// </summary>
        public LegacyDocImportOptions? LegacyDocImportOptions { get; set; }

        /// <summary>
        /// Gets or sets optional save settings for the destination file.
        /// </summary>
        public WordSaveOptions? SaveOptions { get; set; }

        /// <summary>
        /// Gets or sets whether conversion may continue when the legacy `.doc` importer reports unsupported or preserve-only content.
        /// </summary>
        public bool AllowLossyLegacyConversion { get; set; }
    }

    public partial class WordDocument {
        private static readonly string[] SupportedWordConversionExtensions = { ".doc", ".docx" };

        /// <summary>
        /// Converts a Word document between `.doc` and `.docx` using the normal OfficeIMO load and save paths.
        /// </summary>
        /// <param name="sourcePath">Path to the source `.doc` or `.docx` file.</param>
        /// <param name="destinationPath">Path to the destination `.doc` or `.docx` file.</param>
        /// <param name="options">Optional conversion policy settings.</param>
        public static void Convert(string sourcePath, string destinationPath, WordDocumentConversionOptions? options = null) {
            options ??= new WordDocumentConversionOptions();
            ValidateWordConversionPaths(sourcePath, destinationPath, options.Overwrite);
            EnsureWordConversionDirectory(destinationPath);

            using WordDocument document = LoadWordConversionSource(sourcePath, options);
            EnsureWordLegacyConversionIsSafe(document, options);
            document.Save(destinationPath, options.OpenWord, options.SaveOptions);
        }

        private static WordDocument LoadWordConversionSource(string sourcePath, WordDocumentConversionOptions options) {
            if (options.LegacyDocImportOptions != null && WordDocumentLoadRouting.HasLegacyDocExtension(sourcePath)) {
                return LoadLegacyDoc(sourcePath, options.LegacyDocImportOptions);
            }

            return Load(sourcePath, readOnly: false, autoSave: false, overrideStyles: options.OverrideStyles, openSettings: options.OpenSettings);
        }

        private static void EnsureWordLegacyConversionIsSafe(WordDocument document, WordDocumentConversionOptions options) {
            int lossyFeatureCount = document.LegacyDocUnsupportedFeatures.Count
                + document.LegacyDocPreservedFeatures.Count
                + document.LegacyDocCompoundFeatures.Count;
            if (!document.WasLoadedFromLegacyDoc
                || options.AllowLossyLegacyConversion
                || lossyFeatureCount == 0) {
                return;
            }

            throw new NotSupportedException($"Legacy DOC conversion is blocked because the source contains {lossyFeatureCount} unsupported or preserve-only feature(s). Review LegacyDocUnsupportedFeatures, LegacyDocPreservedFeatures, and LegacyDocCompoundFeatures, or set WordDocumentConversionOptions.AllowLossyLegacyConversion when that loss is intentional.");
        }

        private static void ValidateWordConversionPaths(string sourcePath, string destinationPath, bool overwrite) {
            ValidateWordConversionPath(sourcePath, nameof(sourcePath), SupportedWordConversionExtensions);
            ValidateWordConversionPath(destinationPath, nameof(destinationPath), SupportedWordConversionExtensions);

            string sourceFullPath = Path.GetFullPath(sourcePath);
            string destinationFullPath = Path.GetFullPath(destinationPath);

            if (!File.Exists(sourceFullPath)) {
                throw new FileNotFoundException("The source Word document was not found.", sourceFullPath);
            }

            if (PathsReferToSameFile(sourceFullPath, destinationFullPath)) {
                throw new IOException("The source and destination paths must be different for conversion.");
            }

            if (!overwrite && File.Exists(destinationFullPath)) {
                throw new IOException($"The destination file '{destinationFullPath}' already exists.");
            }
        }

        private static void ValidateWordConversionPath(string path, string parameterName, IReadOnlyCollection<string> supportedExtensions) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("A file path is required.", parameterName);
            }

            string extension = Path.GetExtension(path);
            if (!supportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase)) {
                throw new NotSupportedException($"Word conversion supports .doc and .docx files. The path '{path}' uses '{extension}'.");
            }
        }

        private static void EnsureWordConversionDirectory(string destinationPath) {
            string? directory = Path.GetDirectoryName(Path.GetFullPath(destinationPath));
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }
        }

        private static bool PathsReferToSameFile(string left, string right) {
            StringComparison comparison = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return string.Equals(NormalizeConversionPath(left), NormalizeConversionPath(right), comparison);
        }

        private static string NormalizeConversionPath(string path) {
            return path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }
}
