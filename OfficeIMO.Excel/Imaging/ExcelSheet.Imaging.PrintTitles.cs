using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private OfficeImageExportResult RenderWorksheetImageResult(
            OfficeImageExportFormat format,
            WorksheetImageRangeResolution range,
            ExcelWorksheetImageExportOptions options,
            HeaderFooterSnapshot? headerFooterSnapshot,
            int pageNumber,
            int pageCount,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageExportFormat workingFormat = format == OfficeImageExportFormat.Svg
                ? OfficeImageExportFormat.Svg
                : OfficeImageExportFormat.Png;
            OfficeImageExportResult result;
            ExcelRasterRenderState rasterState;
            if (options.SplitByManualPageBreaks &&
                TryCreatePrintTitleLayout(range.Range, out PrintTitleLayout layout)) {
                result = RenderPrintTitleLayout(
                    workingFormat,
                    format,
                    range,
                    options,
                    layout,
                    out rasterState);
            } else {
                ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(this, range.Range, options, range.Diagnostics);
                result = ExcelRangeImageRenderer.Render(
                    snapshot,
                    workingFormat,
                    options,
                    format,
                    out rasterState,
                    cancellationToken);
            }

            if (options.SplitByManualPageBreaks) {
                result = ApplyPageSetupCanvas(
                    workingFormat,
                    format,
                    result,
                    options,
                    ref rasterState);
                result = ApplyHeaderFooterTextChrome(
                    workingFormat,
                    format,
                    result,
                    options,
                    headerFooterSnapshot,
                    pageNumber,
                    pageCount,
                    ref rasterState);
            }

            if (format == workingFormat) return result;
            cancellationToken.ThrowIfCancellationRequested();
            if (!OfficeRasterImageDecoder.TryDecode(result.Bytes, out OfficeRasterImage? image) || image == null) {
                throw new InvalidOperationException("The worksheet raster composition could not be decoded for final image encoding.");
            }

            return options.EnsureAccepted(new OfficeImageExportResult(
                format,
                result.Width,
                result.Height,
                OfficeRasterImageEncoder.Encode(image, format, rasterState.EncodingOptions),
                result.Name,
                result.Source,
                result.Diagnostics));
        }

        private OfficeImageExportResult RenderPrintTitleLayout(
            OfficeImageExportFormat format,
            OfficeImageExportFormat rasterPlanningFormat,
            WorksheetImageRangeResolution range,
            ExcelWorksheetImageExportOptions options,
            PrintTitleLayout layout,
            out ExcelRasterRenderState rasterState) {
            var diagnostics = new List<OfficeImageExportDiagnostic>(range.Diagnostics);
            List<PrintTitleVisualComponent> visuals = CreatePrintTitleVisualComponents(layout, options, diagnostics);
            double logicalWidth = visuals.Count == 0 ? 0D : visuals.Max(component => component.X + component.Snapshot.Width);
            double logicalHeight = visuals.Count == 0 ? 0D : visuals.Max(component => component.Y + component.Snapshot.Height);
            double renderScale = options.Scale;
            OfficeRasterExportPlan? rasterPlan = null;
            if (format == OfficeImageExportFormat.Svg) {
                rasterState = new ExcelRasterRenderState(renderScale, options.RasterEncoding);
            } else {
                rasterPlan = OfficeRasterExportPlanner.Resolve(
                    logicalWidth,
                    logicalHeight,
                    rasterPlanningFormat,
                    options,
                    Name + "!" + range.Range);
                rasterState = ExcelRasterRenderState.FromPlan(rasterPlan.Value);
                renderScale = rasterState.Scale;
                if (rasterPlan.Value.Diagnostic != null) {
                    diagnostics.Add(rasterPlan.Value.Diagnostic);
                }
            }

            List<PrintTitleComponent> components = RenderPrintTitleComponents(
                visuals,
                format,
                options,
                renderScale,
                diagnostics);
            int outputWidth = format == OfficeImageExportFormat.Svg
                ? Math.Max(1, (int)Math.Ceiling(components.Max(component => component.X + component.Width)))
                : rasterPlan!.Value.Limit.PixelWidth;
            int outputHeight = format == OfficeImageExportFormat.Svg
                ? Math.Max(1, (int)Math.Ceiling(components.Max(component => component.Y + component.Height)))
                : rasterPlan!.Value.Limit.PixelHeight;

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

            OfficeRasterImage image = OfficeImageComposer.ComposeRaster(
                outputWidth,
                outputHeight,
                options.BackgroundColor,
                components.Select(component => component.ToLayer()));
            return new OfficeImageExportResult(
                format,
                outputWidth,
                outputHeight,
                OfficeRasterImageEncoder.Encode(image, format, rasterState.EncodingOptions),
                Name,
                Name + "!" + range.Range,
                diagnostics.AsReadOnly());
        }

        private List<PrintTitleVisualComponent> CreatePrintTitleVisualComponents(
            PrintTitleLayout layout,
            ExcelWorksheetImageExportOptions options,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var components = new List<PrintTitleVisualComponent>();
            PrintTitleVisualComponent? corner = layout.CornerRange == null
                ? null
                : CreatePrintTitleVisualComponent(layout.CornerRange, 0D, 0D, options, diagnostics);
            PrintTitleVisualComponent? rowTitles = layout.RowTitleRange == null
                ? null
                : CreatePrintTitleVisualComponent(layout.RowTitleRange, corner?.Snapshot.Width ?? 0D, 0D, options, diagnostics);
            PrintTitleVisualComponent? columnTitles = layout.ColumnTitleRange == null
                ? null
                : CreatePrintTitleVisualComponent(layout.ColumnTitleRange, 0D, rowTitles?.Snapshot.Height ?? 0D, options, diagnostics);
            PrintTitleVisualComponent body = CreatePrintTitleVisualComponent(
                layout.BodyRange,
                columnTitles?.Snapshot.Width ?? 0D,
                rowTitles?.Snapshot.Height ?? 0D,
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

        private PrintTitleVisualComponent CreatePrintTitleVisualComponent(
            string range,
            double x,
            double y,
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

            return new PrintTitleVisualComponent(range, x, y, snapshot);
        }

        private static List<PrintTitleComponent> RenderPrintTitleComponents(
            IReadOnlyList<PrintTitleVisualComponent> visuals,
            OfficeImageExportFormat format,
            ExcelWorksheetImageExportOptions options,
            double renderScale,
            List<OfficeImageExportDiagnostic> diagnostics) {
            ExcelWorksheetImageExportOptions renderOptions = options.CloneWorksheet();
            renderOptions.TargetDpi = null;
            renderOptions.Scale = renderScale;
            var components = new List<PrintTitleComponent>(visuals.Count);
            foreach (PrintTitleVisualComponent visual in visuals) {
                components.Add(RenderPrintTitleComponent(
                    visual,
                    format,
                    renderOptions,
                    diagnostics));
            }

            return components;
        }

        private static PrintTitleComponent RenderPrintTitleComponent(
            PrintTitleVisualComponent visual,
            OfficeImageExportFormat format,
            ExcelWorksheetImageExportOptions renderOptions,
            List<OfficeImageExportDiagnostic> diagnostics) {
            var componentDiagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeRasterImage? raster = null;
            string svgInner = string.Empty;
            if (format == OfficeImageExportFormat.Svg) {
                string svg = ExcelRangeImageRenderer.RenderSvg(visual.Snapshot, renderOptions, componentDiagnostics);
                svgInner = OfficeSvgFormatting.ExtractSvgInner(svg);
            } else {
                raster = ExcelRangeImageRenderer.RenderRaster(visual.Snapshot, renderOptions, componentDiagnostics);
            }

            if (componentDiagnostics.Count > 0) {
                diagnostics.AddRange(componentDiagnostics);
            }

            double x = visual.X * renderOptions.Scale;
            double y = visual.Y * renderOptions.Scale;
            double width = visual.Snapshot.Width * renderOptions.Scale;
            double height = visual.Snapshot.Height * renderOptions.Scale;
            if (format == OfficeImageExportFormat.Svg) {
                x = Math.Ceiling(x);
                y = Math.Ceiling(y);
                width = Math.Ceiling(width);
                height = Math.Ceiling(height);
            }

            return new PrintTitleComponent(
                visual.Range,
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

        private sealed class PrintTitleVisualComponent {
            internal PrintTitleVisualComponent(
                string range,
                double x,
                double y,
                ExcelRangeVisualSnapshot snapshot) {
                Range = range;
                X = x;
                Y = y;
                Snapshot = snapshot;
            }

            internal string Range { get; }
            internal double X { get; }
            internal double Y { get; }
            internal ExcelRangeVisualSnapshot Snapshot { get; }
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
