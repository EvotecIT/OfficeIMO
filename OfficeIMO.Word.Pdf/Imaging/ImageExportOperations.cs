using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public static partial class WordImageExportExtensions {
        /// <summary>Saves selected pages and retains per-image metadata for in-process consumers.</summary>
        internal static IReadOnlyList<OfficeImageExportResult> SaveImageResults(this WordDocument document,
            string folderPath, WordImageExportOptions options) =>
            CreateImageExportBatchOperation(document, options, options.Format).Save(folderPath);

        private static OfficeImageExportOperation<WordImageExportOptions> CreateImageExportOperation(
            WordDocument document, WordImageExportOptions? options, OfficeImageExportFormat format) {
            WordImageExportOptions snapshot = options?.Clone() ?? new WordImageExportOptions();
            return new OfficeImageExportOperation<WordImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy,
                (encoding, effective) => document.ExportImage(encoding, effective),
                (encoding, effective, token) => document.ExportImage(encoding, effective, token));
        }

        private static OfficeImageExportBatchOperation<WordImageExportOptions> CreateImageExportBatchOperation(
            WordDocument document, WordImageExportOptions? options, OfficeImageExportFormat format) {
            WordImageExportOptions snapshot = options?.Clone() ?? new WordImageExportOptions();
            return new OfficeImageExportBatchOperation<WordImageExportOptions>(snapshot, format,
                snapshot.FileConflictPolicy, document.ExportImages,
                exportEach: (encoding, effective, consumer, token) =>
                    document.ExportImages(encoding, consumer, effective, token));
        }
    }
}
