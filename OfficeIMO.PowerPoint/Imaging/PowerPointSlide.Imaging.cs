using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Creates a format-neutral visual snapshot for this slide.
        /// </summary>
        public PowerPointSlideVisualInfo CreateVisualSnapshot(PowerPointImageExportOptions? options = null) {
            PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
            return PowerPointPageImageExport.CreateSnapshot(this, resolved);
        }

        /// <summary>
        /// Exports this slide as a supported raster format or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(
            OfficeImageExportFormat format,
            PowerPointImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
            return PowerPointPageImageExport.Render(this, format, resolved, cancellationToken);
        }

        private static PowerPointImageExportOptions NormalizeImageExportOptions(PowerPointImageExportOptions? options) {
            PowerPointImageExportOptions resolved = options?.Clone() ?? new PowerPointImageExportOptions();
            resolved.Validate();
            return resolved;
        }

    }
}
