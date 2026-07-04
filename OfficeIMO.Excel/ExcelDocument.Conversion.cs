using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls file-to-file Excel workbook conversion.
    /// </summary>
    public sealed class ExcelDocumentConversionOptions {
        /// <summary>
        /// Gets or sets whether an existing destination file may be overwritten. Defaults to <c>true</c>.
        /// </summary>
        public bool Overwrite { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to open Excel after saving the converted file.
        /// </summary>
        public bool OpenExcel { get; set; }

        /// <summary>
        /// Gets or sets optional Open XML load settings for `.xlsx` sources.
        /// </summary>
        public OpenSettings? OpenSettings { get; set; }

        /// <summary>
        /// Gets or sets optional legacy `.xls` import settings.
        /// </summary>
        public LegacyXlsImportOptions? LegacyXlsImportOptions { get; set; }

        /// <summary>
        /// Gets or sets optional save settings for the destination file.
        /// </summary>
        public ExcelSaveOptions? SaveOptions { get; set; }

        /// <summary>
        /// Gets or sets whether conversion may continue when the legacy `.xls` importer reports unsupported or preserve-only content.
        /// </summary>
        public bool AllowLossyLegacyConversion { get; set; }
    }

    public partial class ExcelDocument {
        private static readonly string[] SupportedExcelConversionExtensions = { ".xls", ".xlsx" };

        /// <summary>
        /// Converts an Excel workbook between `.xls` and `.xlsx` using the normal OfficeIMO load and save paths.
        /// </summary>
        /// <param name="sourcePath">Path to the source `.xls` or `.xlsx` file.</param>
        /// <param name="destinationPath">Path to the destination `.xls` or `.xlsx` file.</param>
        /// <param name="options">Optional conversion policy settings.</param>
        public static void Convert(string sourcePath, string destinationPath, ExcelDocumentConversionOptions? options = null) {
            options ??= new ExcelDocumentConversionOptions();
            ValidateExcelConversionPaths(sourcePath, destinationPath, options.Overwrite);
            EnsureExcelConversionDirectory(destinationPath);

            using ExcelDocument document = LoadExcelConversionSource(sourcePath, options);
            EnsureExcelLegacyConversionIsSafe(document, options);
            document.Save(destinationPath, options.OpenExcel, options.SaveOptions);
        }

        private static ExcelDocument LoadExcelConversionSource(string sourcePath, ExcelDocumentConversionOptions options) {
            if (options.LegacyXlsImportOptions != null && ExcelDocumentLoadRouting.HasLegacyXlsExtension(sourcePath)) {
                return LoadLegacyXls(sourcePath, options.LegacyXlsImportOptions);
            }

            return Load(sourcePath, readOnly: false, autoSave: false, openSettings: options.OpenSettings);
        }

        private static void EnsureExcelLegacyConversionIsSafe(ExcelDocument document, ExcelDocumentConversionOptions options) {
            if (!document.WasLoadedFromLegacyXls || options.AllowLossyLegacyConversion) {
                return;
            }

            int unsupportedCount = document.LegacyXlsUnsupportedFeatures.Count + document.LegacyXlsUnsupportedSheets.Count;
            if (unsupportedCount == 0) {
                return;
            }

            throw new NotSupportedException($"Legacy XLS conversion is blocked because the source contains {unsupportedCount} unsupported or preserve-only feature(s). Review LegacyXlsUnsupportedFeatures and LegacyXlsUnsupportedSheets or set ExcelDocumentConversionOptions.AllowLossyLegacyConversion when that loss is intentional.");
        }

        private static void ValidateExcelConversionPaths(string sourcePath, string destinationPath, bool overwrite) {
            ValidateExcelConversionPath(sourcePath, nameof(sourcePath), SupportedExcelConversionExtensions);
            ValidateExcelConversionPath(destinationPath, nameof(destinationPath), SupportedExcelConversionExtensions);

            string sourceFullPath = Path.GetFullPath(sourcePath);
            string destinationFullPath = Path.GetFullPath(destinationPath);

            if (!File.Exists(sourceFullPath)) {
                throw new FileNotFoundException("The source Excel workbook was not found.", sourceFullPath);
            }

            if (PathsReferToSameFile(sourceFullPath, destinationFullPath)) {
                throw new IOException("The source and destination paths must be different for conversion.");
            }

            if (!overwrite && File.Exists(destinationFullPath)) {
                throw new IOException($"The destination file '{destinationFullPath}' already exists.");
            }
        }

        private static void ValidateExcelConversionPath(string path, string parameterName, IReadOnlyCollection<string> supportedExtensions) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("A file path is required.", parameterName);
            }

            string extension = Path.GetExtension(path);
            if (!supportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase)) {
                throw new NotSupportedException($"Excel conversion supports .xls and .xlsx files. The path '{path}' uses '{extension}'.");
            }
        }

        private static void EnsureExcelConversionDirectory(string destinationPath) {
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
