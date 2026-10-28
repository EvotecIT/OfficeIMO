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
        public IReadOnlyList<OfficeImageExportResult> ExportImages(OfficeImageExportFormat format, PowerPointImageExportOptions? options = null) {
            var results = new List<OfficeImageExportResult>();
            ExportImages(format, results.Add, options);
            return results.AsReadOnly();
        }

        /// <summary>Streams selected slide images to a consumer without retaining earlier payloads.</summary>
        public void ExportImages(
            OfficeImageExportFormat format,
            OfficeImageExportConsumer consumer,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (consumer == null) throw new ArgumentNullException(nameof(consumer));
            ThrowIfDisposed();
            PowerPointImageExportOptions resolved = NormalizeDocumentImageExportOptions(options);
            PowerPointPageImageExport.RenderPages(this, format, resolved, consumer, cancellationToken);
        }

        private static PowerPointImageExportOptions NormalizeDocumentImageExportOptions(PowerPointImageExportOptions? options) {
            PowerPointImageExportOptions resolved = options?.Clone() ?? new PowerPointImageExportOptions();
            resolved.Validate();
            if (resolved.SlideNumbers != null && resolved.SlideNumbers.Any(slideNumber => slideNumber <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Slide numbers must be 1-based positive values.");
            }

            return resolved;
        }

    }
}
