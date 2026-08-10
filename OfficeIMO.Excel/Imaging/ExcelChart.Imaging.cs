using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
        public OfficeImageExportResult ExportImage(
            OfficeImageExportFormat format,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            ExcelImageExportOptions resolved = options?.Clone() ?? new ExcelImageExportOptions();
            resolved.ConditionalFormattingDate ??= DateTime.Today;
            resolved.Validate();

            return OfficeImageExportExecutionScope.Run(
                resolved,
                cancellationToken,
                token => ExportImageCore(format, resolved, token));
        }

        /// <summary>Renders this chart to dependency-free PNG bytes.</summary>
        public byte[] ToPng(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>Renders this chart to dependency-free SVG text.</summary>
        public string ToSvg(ExcelImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>Saves this chart as a PNG file.</summary>
        public OfficeImageExportResult SaveAsPng(string path, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsPng().Save(path);

        /// <summary>Saves this chart as an SVG file.</summary>
        public OfficeImageExportResult SaveAsSvg(string path, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsSvg().Save(path);

        /// <summary>Writes this chart as PNG to a caller-owned stream.</summary>
        public OfficeImageExportResult SaveAsPng(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsPng().Save(stream);

        /// <summary>Writes this chart as SVG to a caller-owned stream.</summary>
        public OfficeImageExportResult SaveAsSvg(Stream stream, ExcelImageExportOptions? options = null) =>
            new ExcelChartImageExportBuilder(this, options).AsSvg().Save(stream);

        /// <summary>Asynchronously saves this chart as a PNG file.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            string path,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsPng().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves this chart as an SVG file.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            string path,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsSvg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes this chart as PNG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsPngAsync(
            Stream stream,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsPng().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes this chart as SVG to a stream.</summary>
        public Task<OfficeImageExportResult> SaveAsSvgAsync(
            Stream stream,
            ExcelImageExportOptions? options = null,
            CancellationToken cancellationToken = default) =>
            new ExcelChartImageExportBuilder(this, options).AsSvg().SaveAsync(stream, cancellationToken);

        private OfficeImageExportResult ExportImageCore(
            OfficeImageExportFormat format,
            ExcelImageExportOptions options,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                throw new InvalidOperationException(
                    "The chart data could not be resolved into a renderable snapshot.");
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            if (!ExcelRangeImageRenderer.TryCreateOfficeChartSnapshot(
                snapshot,
                snapshot.WidthPixels,
                snapshot.HeightPixels,
                diagnostics,
                _sheetName,
                out OfficeChartSnapshot? officeSnapshot) || officeSnapshot == null) {
                throw new NotSupportedException(
                    "The chart type cannot be rendered by the shared Office chart renderer.");
            }

            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
                officeSnapshot,
                useMinimumCanvas: false);
            drawing.Fonts.AddRange(options.Fonts);
            string source = _sheetName + "!" + snapshot.Name;
            drawing.AppendFontDiagnostics(diagnostics, source);

            if (format == OfficeImageExportFormat.Svg) {
                OfficeDrawing svgDrawing = new OfficeDrawing(drawing.Width, drawing.Height);
                OfficeShape background = OfficeShape.Rectangle(drawing.Width, drawing.Height);
                background.FillColor = options.BackgroundColor;
                background.StrokeWidth = 0D;
                svgDrawing.AddShape(background, 0D, 0D);
                svgDrawing.AddDrawing(drawing, 0D, 0D);
                double scale = options.GetEffectiveScale(svgDrawing.Width, svgDrawing.Height);
                byte[] svgBytes = OfficeDrawingSvgExporter.ToSvgBytes(
                    svgDrawing,
                    scale,
                    OfficeSvgSizeUnit.Pixel,
                    imageCodec: null,
                    resourceIdPrefix: null,
                    cancellationToken);
                if (!OfficeImageReader.TryIdentifyByContent(svgBytes, ".svg", out OfficeImageInfo svgInfo)) {
                    throw new InvalidDataException("The rendered chart SVG dimensions could not be identified.");
                }
                cancellationToken.ThrowIfCancellationRequested();
                return options.EnsureAccepted(new OfficeImageExportResult(
                    format,
                    svgInfo.Width,
                    svgInfo.Height,
                    svgBytes,
                    snapshot.Name,
                    source,
                    diagnostics.AsReadOnly()));
            }

            if (!format.IsRaster()) {
                throw new ArgumentOutOfRangeException(nameof(format), "The selected image format is not supported.");
            }

            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                drawing.Width,
                drawing.Height,
                format,
                options,
                source);
            if (plan.Diagnostic != null) {
                diagnostics.Add(plan.Diagnostic);
            }

            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(
                drawing,
                new OfficeDrawingRasterRenderOptions {
                    Scale = plan.Limit.Scale,
                    Background = options.BackgroundColor,
                    TextShapingProvider = options.TextShapingProvider,
                    TextShapingLanguage = options.TextShapingLanguage,
                    DiagnosticSink = diagnostics,
                    DiagnosticSource = source,
                    CancellationToken = cancellationToken
                });
            byte[] bytes = OfficeRasterImageEncoder.Encode(
                image,
                format,
                plan.CreateEncodingOptions());
            cancellationToken.ThrowIfCancellationRequested();
            return options.EnsureAccepted(new OfficeImageExportResult(
                format,
                image.Width,
                image.Height,
                bytes,
                snapshot.Name,
                source,
                diagnostics.AsReadOnly()));
        }
    }
}
