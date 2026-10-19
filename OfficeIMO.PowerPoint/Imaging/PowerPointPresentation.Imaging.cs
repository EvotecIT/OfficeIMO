using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        /// Exports presentation slides as supported raster formats or SVG images.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, PowerPointPresentationImageExportOptions? options = null) {
            var results = new List<OfficeImageExportResult>();
            ExportImages(format, results.Add, options);
            return results.AsReadOnly();
        }

        /// <summary>Streams selected slide images to a consumer without retaining earlier payloads.</summary>
        public void ExportImages(
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            PowerPointPresentationImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (consumer == null) throw new ArgumentNullException(nameof(consumer));
            ThrowIfDisposed();
            PowerPointPresentationImageExportOptions resolved = NormalizePresentationImageExportOptions(options);
            PowerPointPageImageExport.RenderPages(this, format, resolved, consumer, cancellationToken);
        }

        /// <summary>
        /// Saves visible presentation slides as PNG files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, PowerPointPresentationImageExportOptions? options = null) =>
            new PowerPointPresentationImageExportBuilder(this, options).AsPng().Save(folderPath);

        /// <summary>
        /// Saves visible presentation slides as image files in a folder.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> SaveAsImages(string folderPath, OfficeImageExportFormat format, PowerPointPresentationImageExportOptions? options = null) {
            PowerPointPresentationImageExportBuilder builder = new PowerPointPresentationImageExportBuilder(this, options);
            builder.As(format);
            return builder.Save(folderPath);
        }

        /// <summary>
        /// Asynchronously saves visible presentation slides as PNG files in a folder.
        /// </summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
            string folderPath,
            PowerPointPresentationImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new PowerPointPresentationImageExportBuilder(this, options).AsPng().SaveAsync(folderPath, cancellationToken);

        /// <summary>
        /// Asynchronously saves visible presentation slides as image files in a folder.
        /// </summary>
        public Task<IReadOnlyList<OfficeImageExportResult>> SaveAsImagesAsync(
            string folderPath,
            OfficeImageExportFormat format,
            PowerPointPresentationImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            PowerPointPresentationImageExportBuilder builder = new PowerPointPresentationImageExportBuilder(this, options);
            builder.As(format);
            return builder.SaveAsync(folderPath, cancellationToken);
        }

        private static PowerPointPresentationImageExportOptions NormalizePresentationImageExportOptions(PowerPointPresentationImageExportOptions? options) {
            PowerPointPresentationImageExportOptions resolved = options?.ClonePresentation() ?? new PowerPointPresentationImageExportOptions();
            resolved.Validate();
            if (resolved.SlideNumbers != null && resolved.SlideNumbers.Any(slideNumber => slideNumber <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Slide numbers must be 1-based positive values.");
            }

            return resolved;
        }

    }
}
