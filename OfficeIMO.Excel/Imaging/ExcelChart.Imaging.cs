#pragma warning disable RS0026 // The shared chart options overloads take a required options argument, so existing calls keep binding to the Excel options forms.
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        /// <summary>
        /// Exports this chart through the shared dependency-free chart renderer.
        /// </summary>
        /// <param name="format">Target image format.</param>
        /// <param name="options">Shared Excel image-export settings.</param>
        /// <param name="cancellationToken">Token used to cancel rendering.</param>
        /// <returns>The encoded chart image and any fidelity diagnostics.</returns>
        /// <exception cref="NotSupportedException">The chart cannot be drawn by the shared chart renderer.</exception>
        public OfficeImageExportResult ToImage(
            OfficeImageExportFormat format,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            ExcelImageExportOptions resolved = NormalizeImageConversionOptions(options);
            return OfficeImageExportExecutionScope.Run(
                resolved,
                cancellationToken,
                token => ExportImageCore(format, resolved, token));
        }

        /// <summary>Renders this chart on its own, without the worksheet around it.</summary>
        /// <remarks>Scheme colors and series without explicit colors resolve against the workbook theme.</remarks>
        /// <param name="options">Image format, size and rendering settings; a null <see cref="OfficeChartImageExportOptions.Width"/> or <see cref="OfficeChartImageExportOptions.Height"/> uses the chart's anchor size.</param>
        /// <param name="cancellationToken">Token used to cancel rendering.</param>
        /// <returns>The encoded image, its size and the rendering diagnostics.</returns>
        /// <exception cref="NotSupportedException">The shared chart renderer cannot draw this chart; the message gives the reason.</exception>
        public OfficeImageExportResult ToImage(OfficeChartImageExportOptions options, CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(options);
            return OfficeChartImageExporter.ToImage(CreateImageSource(), options, cancellationToken);
        }

        /// <summary>Renders this chart and writes the image to a file.</summary>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Image format, size and rendering settings.</param>
        /// <param name="cancellationToken">Token used to cancel rendering.</param>
        /// <returns>The conversion report with the rendering diagnostics.</returns>
        /// <exception cref="NotSupportedException">The shared chart renderer cannot draw this chart; the message gives the reason.</exception>
        public OfficeConversionReport SaveAsImage(string path, OfficeChartImageExportOptions options,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(options);
            return OfficeChartImageExporter.SaveAsImage(CreateImageSource(), path, options, cancellationToken);
        }

        /// <summary>Renders this chart and writes the image to a caller-owned stream.</summary>
        /// <param name="stream">Destination stream.</param>
        /// <param name="options">Image format, size and rendering settings.</param>
        /// <param name="cancellationToken">Token used to cancel rendering.</param>
        /// <returns>The conversion report with the rendering diagnostics.</returns>
        /// <exception cref="NotSupportedException">The shared chart renderer cannot draw this chart; the message gives the reason.</exception>
        public OfficeConversionReport SaveAsImage(Stream stream, OfficeChartImageExportOptions options,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(options);
            return OfficeChartImageExporter.SaveAsImage(CreateImageSource(), stream, options, cancellationToken);
        }

        /// <summary>Tries to expose this chart through the shared dependency-free chart contract at its anchor size.</summary>
        /// <remarks>Reads the saved chart part and the workbook theme, so charts in loaded workbooks project the same way as charts built in this session.</remarks>
        /// <param name="snapshot">The cached chart data and supported presentation metadata.</param>
        /// <returns>True when the shared chart renderer can draw the chart.</returns>
        public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) {
            snapshot = null!;
            if (!TryGetSnapshot(out ExcelChartInfo info)) return false;
            if (!ExcelRangeImageRenderer.TryCreateOfficeChartSnapshot(info,
                ExcelImageExportLimits.ExtentPixels(info.Width), ExcelImageExportLimits.ExtentPixels(info.Height),
                null, _sheetName, out OfficeChartSnapshot? created) || created == null) return false;
            snapshot = created;
            return true;
        }

        private OfficeChartImageSource CreateImageSource() {
            (int widthPixels, int heightPixels) = GetAnchorSizePixels();
            string name = Name is { Length: > 0 } chartName ? chartName : "Chart";
            string source = _sheetName + "!" + name;
            return new OfficeChartImageSource("Excel", name, source, OfficeLength.Pixels(widthPixels), OfficeLength.Pixels(heightPixels),
                (width, height, diagnostics) => CreateRenderSnapshot(width, height, diagnostics, out _));
        }

        private OfficeChartSnapshot CreateRenderSnapshot(double? width, double? height,
            List<OfficeImageExportDiagnostic> diagnostics, out ExcelChartInfo info) {
            if (!TryGetSnapshot(out info)) {
                throw new InvalidOperationException(
                    "The chart data could not be resolved into a renderable snapshot.");
            }

            diagnostics.AddRange(info.Diagnostics);
            int firstNewDiagnostic = diagnostics.Count;
            if (!ExcelRangeImageRenderer.TryCreateOfficeChartSnapshot(
                info,
                width ?? ExcelImageExportLimits.ExtentPixels(info.Width),
                height ?? ExcelImageExportLimits.ExtentPixels(info.Height),
                diagnostics,
                _sheetName,
                out OfficeChartSnapshot? officeSnapshot) || officeSnapshot == null) {
                throw new NotSupportedException(diagnostics.Count > firstNewDiagnostic
                    ? OfficeChartImageExporter.UnsupportedChartMessage + diagnostics[diagnostics.Count - 1].Message
                    : "The chart type cannot be rendered by the shared Office chart renderer.");
            }
            return officeSnapshot;
        }

        private OfficeImageExportResult ExportImageCore(
            OfficeImageExportFormat format,
            ExcelImageExportOptions options,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeChartSnapshot officeSnapshot = CreateRenderSnapshot(null, null, diagnostics, out ExcelChartInfo info);
            return OfficeChartImageExporter.Export(officeSnapshot, format, options, diagnostics, info.Name,
                _sheetName + "!" + info.Name, cancellationToken);
        }
    }
}
