using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>Renders Word pages through the PDF layout adapter.</summary>
    public static partial class WordImageExportExtensions {
        /// <summary>
        /// Creates a format-neutral visual snapshot for the requested document page.
        /// </summary>
        public static WordDocumentVisualInfo CreateVisualSnapshot(this WordDocument document, WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordPageImageExport.CreateSnapshot(document, resolved);
        }

        /// <summary>Returns the shared layout's document page count.</summary>
        public static int GetEstimatedPageCount(this WordDocument document) => WordPageImageExport.EstimatePageCount(document);

        /// <summary>Creates format-neutral visual snapshots for a page range or the complete document.</summary>
        public static IReadOnlyList<WordDocumentVisualInfo> CreateVisualSnapshots(this WordDocument document, WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordPageImageExport.CreateSnapshots(document, resolved);
        }

        /// <summary>
        /// Creates format-neutral visual snapshots for a page range or the complete document
        /// with cooperative cancellation during pagination.
        /// </summary>
        public static IReadOnlyList<WordDocumentVisualInfo> CreateVisualSnapshots(this WordDocument document,
            WordImageExportOptions? options,
            CancellationToken cancellationToken) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordPageImageExport.CreateSnapshots(document, resolved, cancellationToken);
        }

        /// <summary>
        /// Exports the requested document page as a supported raster format or SVG.
        /// </summary>
        public static OfficeImageExportResult ExportImage(this WordDocument document,
            OfficeImageExportFormat format,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordPageImageExport.Render(document, format, resolved, cancellationToken);
        }

        /// <summary>Exports a page range or the complete document as PNG or SVG images.</summary>
        public static IReadOnlyList<OfficeImageExportResult> ExportImages(this WordDocument document, OfficeImageExportFormat format,
            WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordPageImageExport.RenderPages(document, format, resolved);
        }

        /// <summary>Streams selected page images to a consumer without retaining earlier payloads.</summary>
        public static void ExportImages(this WordDocument document,
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            WordPageImageExport.RenderPages(document, format, resolved, consumer, cancellationToken);
        }

        /// <summary>Saves every selected document page as a PNG file in a folder.</summary>
        public static IReadOnlyList<OfficeImageExportResult> SaveAsImages(this WordDocument document, string folderPath,
            WordImageExportOptions? options = null) =>
            new WordDocumentPageImageExportBuilder(document, options).AsPng().Save(folderPath);

        /// <summary>Saves every selected document page as PNG or SVG in a folder.</summary>
        public static IReadOnlyList<OfficeImageExportResult> SaveAsImages(this WordDocument document, string folderPath,
            OfficeImageExportFormat format, WordImageExportOptions? options = null) =>
            new WordDocumentPageImageExportBuilder(document, options).As(format).Save(folderPath);

        /// <summary>Asynchronously saves every selected document page as a PNG file in a folder.</summary>
        public static Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(this WordDocument document, string folderPath,
            WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentPageImageExportBuilder(document, options).AsPng().SaveAsync(folderPath, cancellationToken);

        /// <summary>Asynchronously saves every selected document page as PNG or SVG in a folder.</summary>
        public static Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(this WordDocument document, string folderPath,
            OfficeImageExportFormat format, WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentPageImageExportBuilder(document, options).As(format).SaveAsync(folderPath, cancellationToken);

        /// <summary>
        /// Returns how many pages the shared Word layout produces for this document.
        /// </summary>
        /// <remarks>
        /// This follows OfficeIMO's renderer layout model and is not a substitute for Word's application-owned pagination.
        /// </remarks>
        public static int GetEstimatedImagePageCount(this WordDocument document) => WordPageImageExport.EstimatePageCount(document);

        /// <summary>
        /// Renders the requested document page to dependency-free PNG bytes.
        /// </summary>
        public static byte[] ToPng(this WordDocument document, WordImageExportOptions? options = null) =>
            document.ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders the requested document page to dependency-free SVG text.
        /// </summary>
        public static string ToSvg(this WordDocument document, WordImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(document.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves the requested document page as a PNG file.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this WordDocument document, string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsPng().Save(path);

        /// <summary>
        /// Saves the requested document page as an SVG file.
        /// </summary>
        public static OfficeImageExportResult SaveAsSvg(this WordDocument document, string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsSvg().Save(path);

        /// <summary>
        /// Writes the requested document page as PNG to a stream.
        /// </summary>
        public static OfficeImageExportResult SaveAsPng(this WordDocument document, Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsPng().Save(stream);

        /// <summary>
        /// Writes the requested document page as SVG to a stream.
        /// </summary>
        public static OfficeImageExportResult SaveAsSvg(this WordDocument document, Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsSvg().Save(stream);

        /// <summary>
        /// Asynchronously saves the requested document page as a PNG file.
        /// </summary>
        public static Task<OfficeImageExportResult> SaveAsPngAsync(this WordDocument document,
            string path,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsPng().SaveAsync(path, cancellationToken);

        /// <summary>
        /// Asynchronously saves the requested document page as an SVG file.
        /// </summary>
        public static Task<OfficeImageExportResult> SaveAsSvgAsync(this WordDocument document,
            string path,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsSvg().SaveAsync(path, cancellationToken);

        /// <summary>
        /// Asynchronously writes the requested document page as PNG to a stream.
        /// </summary>
        public static Task<OfficeImageExportResult> SaveAsPngAsync(this WordDocument document,
            Stream stream,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsPng().SaveAsync(stream, cancellationToken);

        /// <summary>
        /// Asynchronously writes the requested document page as SVG to a stream.
        /// </summary>
        public static Task<OfficeImageExportResult> SaveAsSvgAsync(this WordDocument document,
            Stream stream,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsSvg().SaveAsync(stream, cancellationToken);

        private static WordImageExportOptions NormalizeImageExportOptions(WordImageExportOptions? options) {
            WordImageExportOptions resolved = options?.Clone() ?? new WordImageExportOptions();
            resolved.Validate();
            return resolved;
        }
    }
}
