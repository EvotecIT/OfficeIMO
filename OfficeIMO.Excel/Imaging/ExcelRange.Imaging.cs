using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelRange {
        /// <summary>
        /// Creates a format-neutral visual snapshot for this range.
        /// </summary>
        public ExcelRangeVisualInfo CreateVisualSnapshot(ExcelImageExportOptions? options = null) =>
            ExcelRangeVisualSnapshotBuilder.Build(Sheet, Address, NormalizeOptions(options));

        /// <summary>
        /// Exports this range as a supported raster format or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(
            OfficeImageExportFormat format,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            ExcelImageExportOptions resolved = NormalizeOptions(options);
            return OfficeImageExportExecutionScope.Run(
                resolved,
                cancellationToken,
                token => {
                    ExcelRangeVisualInfo snapshot = ExcelRangeVisualSnapshotBuilder.Build(
                        Sheet,
                        Address,
                        resolved,
                        cancellationToken: token);
                    return ExcelRangeImageRenderer.Render(snapshot, format, resolved, token);
                });
        }

        private static ExcelImageExportOptions NormalizeOptions(ExcelImageExportOptions? options) {
            ExcelImageExportOptions resolved = options?.Clone() ?? new ExcelImageExportOptions();
            resolved.ConditionalFormattingDate ??= System.DateTime.Today;
            resolved.Validate();

            return resolved;
        }

    }
}
