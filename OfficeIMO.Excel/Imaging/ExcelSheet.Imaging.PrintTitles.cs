using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private OfficeImageExportResult RenderWorksheetImageResult(
            OfficeImageExportFormat format,
            WorksheetImageRangeResolution range,
            ExcelWorksheetImageExportOptions options,
            int pageNumber,
            int pageCount,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportFormat workingFormat = format == OfficeImageExportFormat.Svg
                ? OfficeImageExportFormat.Svg
                : OfficeImageExportFormat.Png;
            OfficeImageExportResult result;
            if (options.SplitByManualPageBreaks &&
                TryCreatePrintTitleLayout(range.Range, out PrintTitleLayout layout)) {
                result = RenderPrintTitleLayout(workingFormat, range, options, layout);
            } else {
                ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(this, range.Range, options, range.Diagnostics);
                result = ExcelRangeImageRenderer.Render(snapshot, workingFormat, options, cancellationToken);
            }

            if (options.SplitByManualPageBreaks) {
                result = ApplyPageSetupCanvas(workingFormat, result, options);
                result = ApplyHeaderFooterTextChrome(workingFormat, result, options, pageNumber, pageCount);
            }

            if (format == workingFormat) return result;
            cancellationToken.ThrowIfCancellationRequested();
            if (!OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? image) || image == null) {
                throw new InvalidOperationException("The worksheet raster composition could not be decoded for final image encoding.");
            }

            return new OfficeImageExportResult(
                format,
                result.Width,
                result.Height,
                OfficeRasterImageEncoder.Encode(image, format, options.RasterEncoding),
                result.Name,
                result.Source,
                result.Diagnostics);
        }

        private OfficeImageExportResult RenderPrintTitleLayout(
            OfficeImageExportFormat format,
            WorksheetImageRangeResolution range,
            ExcelWorksheetImageExportOptions options,
            PrintTitleLayout layout) {
            var diagnostics = new List<OfficeImageExportDiagnostic>(range.Diagnostics);
            List<PrintTitleComponent> components = CreatePrintTitleComponents(layout, format, options, diagnostics);
            double width = components.Count == 0 ? 0D : components.Max(component => component.X + component.Width);
            double height = components.Count == 0 ? 0D : components.Max(component => component.Y + component.Height);
            int outputWidth = Math.Max(1, (int)Math.Ceiling(width));
            int outputHeight = Math.Max(1, (int)Math.Ceiling(height));

            if (format == OfficeImageExportFormat.Svg) {
                return new OfficeImageExportResult(
                    format,
                    outputWidth,
                    outputHeight,
                    OfficeImageComposer.ComposeSvgBytes(
                        outputWidth,
                        outputHeight,
                        options.BackgroundColor,
                        components.Select(component => component.ToLayer())),
                    Name,
                    Name + "!" + range.Range,
                    diagnostics.AsReadOnly());
            }

            return new OfficeImageExportResult(
                format,
                outputWidth,
                outputHeight,
                OfficeImageComposer.ComposePng(
                    outputWidth,
                    outputHeight,
                    options.BackgroundColor,
                    components.Select(component => component.ToLayer())),
                Name,
                Name + "!" + range.Range,
                diagnostics.AsReadOnly());
        }

        private List<PrintTitleComponent> CreatePrintTitleComponents(
            PrintTitleLayout layout,
            OfficeImageExportFormat format,
            ExcelWorksheetImageExportOptions options,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var components = new List<PrintTitleComponent>();
            PrintTitleComponent? corner = layout.CornerRange == null
                ? null
                : RenderPrintTitleComponent(layout.CornerRange, 0D, 0D, format, options, diagnostics);
            PrintTitleComponent? rowTitles = layout.RowTitleRange == null
                ? null
                : RenderPrintTitleComponent(layout.RowTitleRange, corner?.Width ?? 0D, 0D, format, options, diagnostics);
            PrintTitleComponent? columnTitles = layout.ColumnTitleRange == null
                ? null
                : RenderPrintTitleComponent(layout.ColumnTitleRange, 0D, rowTitles?.Height ?? 0D, format, options, diagnostics);
            PrintTitleComponent body = RenderPrintTitleComponent(
                layout.BodyRange,
                columnTitles?.Width ?? 0D,
                rowTitles?.Height ?? 0D,
                format,
                options,
                diagnostics);

            if (corner != null) {
                components.Add(corner);
            }

            if (rowTitles != null) {
                components.Add(rowTitles);
            }

            if (columnTitles != null) {
                components.Add(columnTitles);
            }

            components.Add(body);
            return components;
        }

        private PrintTitleComponent RenderPrintTitleComponent(
            string range,
            double x,
            double y,
            OfficeImageExportFormat format,
            ExcelWorksheetImageExportOptions options,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Scale <= 0D || double.IsNaN(options.Scale) || double.IsInfinity(options.Scale)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Scale must be a finite positive number.");
            }

            ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(this, range, options);
            if (snapshot.Diagnostics.Count > 0) {
                diagnostics.AddRange(snapshot.Diagnostics);
            }

            var componentDiagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeRasterImage? raster = null;
            string svgInner = string.Empty;
            int width = Math.Max(1, (int)Math.Ceiling(snapshot.Width * options.Scale));
            int height = Math.Max(1, (int)Math.Ceiling(snapshot.Height * options.Scale));
            if (format == OfficeImageExportFormat.Svg) {
                string svg = ExcelRangeImageRenderer.RenderSvg(snapshot, options, componentDiagnostics);
                svgInner = OfficeSvgFormatting.ExtractSvgInner(svg);
            } else {
                raster = ExcelRangeImageRenderer.RenderRaster(snapshot, options, componentDiagnostics);
                width = raster.Width;
                height = raster.Height;
            }

            if (componentDiagnostics.Count > 0) {
                diagnostics.AddRange(componentDiagnostics);
            }

            return new PrintTitleComponent(
                range,
                x,
                y,
                width,
                height,
                raster,
                svgInner);
        }

        private bool TryCreatePrintTitleLayout(string bodyRange, out PrintTitleLayout layout) {
            layout = default;
            ExcelPrintTitles titles = GetPrintTitles();
            if ((!titles.HasRows && !titles.HasColumns) ||
                !A1.TryParseRange(bodyRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return false;
            }

            bool repeatRows = titles.HasRows &&
                titles.FirstRow.GetValueOrDefault() > 0 &&
                titles.LastRow.GetValueOrDefault() >= titles.FirstRow.GetValueOrDefault() &&
                firstRow > titles.LastRow.GetValueOrDefault();
            bool repeatColumns = titles.HasColumns &&
                titles.FirstColumn.GetValueOrDefault() > 0 &&
                titles.LastColumn.GetValueOrDefault() >= titles.FirstColumn.GetValueOrDefault() &&
                firstColumn > titles.LastColumn.GetValueOrDefault();
            if (!repeatRows && !repeatColumns) {
                return false;
            }

            string? rowTitleRange = repeatRows
                ? ToRange(titles.FirstRow.GetValueOrDefault(), firstColumn, titles.LastRow.GetValueOrDefault(), lastColumn)
                : null;
            string? columnTitleRange = repeatColumns
                ? ToRange(firstRow, titles.FirstColumn.GetValueOrDefault(), lastRow, titles.LastColumn.GetValueOrDefault())
                : null;
            string? cornerRange = repeatRows && repeatColumns
                ? ToRange(titles.FirstRow.GetValueOrDefault(), titles.FirstColumn.GetValueOrDefault(), titles.LastRow.GetValueOrDefault(), titles.LastColumn.GetValueOrDefault())
                : null;
            layout = new PrintTitleLayout(bodyRange, rowTitleRange, columnTitleRange, cornerRange);
            return true;
        }

        private readonly struct PrintTitleLayout {
            internal PrintTitleLayout(string bodyRange, string? rowTitleRange, string? columnTitleRange, string? cornerRange) {
                BodyRange = bodyRange;
                RowTitleRange = rowTitleRange;
                ColumnTitleRange = columnTitleRange;
                CornerRange = cornerRange;
            }

            internal string BodyRange { get; }
            internal string? RowTitleRange { get; }
            internal string? ColumnTitleRange { get; }
            internal string? CornerRange { get; }
        }

        private sealed class PrintTitleComponent {
            internal PrintTitleComponent(
                string range,
                double x,
                double y,
                double width,
                double height,
                OfficeRasterImage? raster,
                string svgInner) {
                Range = range;
                X = x;
                Y = y;
                Width = width;
                Height = height;
                Raster = raster;
                SvgInner = svgInner;
            }

            internal string Range { get; }
            internal double X { get; }
            internal double Y { get; }
            internal double Width { get; }
            internal double Height { get; }
            internal OfficeRasterImage? Raster { get; }
            internal string SvgInner { get; }

            internal OfficeImageLayer ToLayer() =>
                Raster != null
                    ? OfficeImageLayer.FromRaster(Raster, X, Y, Width, Height)
                    : OfficeImageLayer.FromSvgInner(SvgInner, X, Y, Width, Height);
        }
    }
}
