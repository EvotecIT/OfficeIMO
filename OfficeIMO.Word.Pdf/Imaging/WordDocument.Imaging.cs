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

        /// <summary>
        /// Returns how many pages the shared Word layout produces for this document.
        /// </summary>
        /// <remarks>
        /// This follows OfficeIMO's renderer layout model and is not a substitute for Word's application-owned pagination.
        /// </remarks>
        public static int GetEstimatedImagePageCount(this WordDocument document) => WordPageImageExport.EstimatePageCount(document);

        private static WordImageExportOptions NormalizeImageExportOptions(WordImageExportOptions? options) {
            WordImageExportOptions resolved = options?.Clone() ?? new WordImageExportOptions();
            resolved.Validate();
            return resolved;
        }
    }
}
