using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        /// Exports presentation slides as supported raster formats or SVG images.
        /// </summary>
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, PowerPointDocumentImageExportOptions? options = null) {
            var results = new List<OfficeImageExportResult>();
            ExportImages(format, results.Add, options);
            return results.AsReadOnly();
        }

        /// <summary>Streams selected slide images to a consumer without retaining earlier payloads.</summary>
        public void ExportImages(
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            PowerPointDocumentImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (consumer == null) throw new ArgumentNullException(nameof(consumer));
            ThrowIfDisposed();
            PowerPointDocumentImageExportOptions resolved = NormalizeDocumentImageExportOptions(options);
            PowerPointPageImageExport.RenderPages(this, format, resolved, consumer, cancellationToken);
        }

        private static PowerPointDocumentImageExportOptions NormalizeDocumentImageExportOptions(PowerPointDocumentImageExportOptions? options) {
            PowerPointDocumentImageExportOptions resolved = options?.CloneDocument() ?? new PowerPointDocumentImageExportOptions();
            resolved.Validate();
            if (resolved.SlideNumbers != null && resolved.SlideNumbers.Any(slideNumber => slideNumber <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Slide numbers must be 1-based positive values.");
            }

            return resolved;
        }

    }
}
