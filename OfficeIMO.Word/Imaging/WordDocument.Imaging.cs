using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Creates a format-neutral visual snapshot for the requested document page.
        /// </summary>
        public WordDocumentVisualSnapshot CreateVisualSnapshot(WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordDocumentImageRenderer.CreateSnapshot(this, resolved);
        }

        /// <summary>Returns the renderer's estimated document page count.</summary>
        public int GetEstimatedPageCount() => WordDocumentImageRenderer.EstimatePageCount(this);

        /// <summary>Creates format-neutral visual snapshots for a page range or the complete document.</summary>
        public IReadOnlyList<WordDocumentVisualSnapshot> CreateVisualSnapshots(WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordDocumentImageRenderer.CreateSnapshots(this, resolved);
        }

        /// <summary>
        /// Exports the requested document page as a supported raster format or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(
            OfficeImageExportFormat format,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordDocumentImageRenderer.Render(this, format, resolved, cancellationToken);
        }

        /// <summary>Exports a page range or the complete document as PNG or SVG images.</summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format,
            WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordDocumentImageRenderer.RenderPages(this, format, resolved);
        }

        /// <summary>Streams selected page images to a consumer without retaining earlier payloads.</summary>
        public void ExportImages(
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            WordDocumentImageRenderer.RenderPages(this, format, resolved, consumer, cancellationToken);
        }

        /// <summary>Saves every selected document page as a PNG file in a folder.</summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath,
            WordImageExportOptions? options = null) =>
            new WordDocumentPageImageExportBuilder(this, options).AsPng().Save(folderPath);

        /// <summary>Saves every selected document page as PNG or SVG in a folder.</summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath,
            OfficeImageExportFormat format, WordImageExportOptions? options = null) =>
            new WordDocumentPageImageExportBuilder(this, options).As(format).Save(folderPath);

        /// <summary>Asynchronously saves every selected document page as a PNG file in a folder.</summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(string folderPath,
            WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentPageImageExportBuilder(this, options).AsPng().SaveAsync(folderPath, cancellationToken);

        /// <summary>Asynchronously saves every selected document page as PNG or SVG in a folder.</summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(string folderPath,
            OfficeImageExportFormat format, WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentPageImageExportBuilder(this, options).As(format).SaveAsync(folderPath, cancellationToken);

        /// <summary>
        /// Estimates how many pages the dependency-free image renderer will produce for this document.
        /// </summary>
        /// <remarks>
        /// This follows OfficeIMO's renderer layout model and is not a substitute for Word's application-owned pagination.
        /// </remarks>
        public int GetEstimatedImagePageCount() => WordDocumentImageRenderer.EstimatePageCount(this);

        /// <summary>
        /// Renders the requested document page to dependency-free PNG bytes.
        /// </summary>
        public byte[] ToPng(WordImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders the requested document page to dependency-free SVG text.
        /// </summary>
        public string ToSvg(WordImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves the requested document page as a PNG file.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsPng().Save(path);

        /// <summary>
        /// Saves the requested document page as an SVG file.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsSvg().Save(path);

        /// <summary>
        /// Writes the requested document page as PNG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsPng(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsPng().Save(stream);

        /// <summary>
        /// Writes the requested document page as SVG to a stream.
        /// </summary>
        public OfficeImageExportResult SaveAsSvg(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsSvg().Save(stream);

        /// <summary>
        /// Asynchronously saves the requested document page as a PNG file.
        /// </summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            string path,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsPng().SaveAsync(path, cancellationToken);

        /// <summary>
        /// Asynchronously saves the requested document page as an SVG file.
        /// </summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            string path,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsSvg().SaveAsync(path, cancellationToken);

        /// <summary>
        /// Asynchronously writes the requested document page as PNG to a stream.
        /// </summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            Stream stream,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsPng().SaveAsync(stream, cancellationToken);

        /// <summary>
        /// Asynchronously writes the requested document page as SVG to a stream.
        /// </summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            Stream stream,
            WordImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsSvg().SaveAsync(stream, cancellationToken);

        private static WordImageExportOptions NormalizeImageExportOptions(WordImageExportOptions? options) {
            WordImageExportOptions resolved = options?.Clone() ?? new WordImageExportOptions();
            resolved.Validate();
            return resolved;
        }
    }
}
