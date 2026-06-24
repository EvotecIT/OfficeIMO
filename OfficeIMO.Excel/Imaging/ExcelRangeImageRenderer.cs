using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        internal static OfficeImageExportResult Render(ExcelRangeVisualSnapshot snapshot, OfficeImageExportFormat format, ExcelImageExportOptions options) {
            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            if (format == OfficeImageExportFormat.Svg) {
                string svg = RenderSvg(snapshot, options, diagnostics);
                return new OfficeImageExportResult(format, ScaledWidth(snapshot, options), ScaledHeight(snapshot, options), Encoding.UTF8.GetBytes(svg), snapshot.SheetName, snapshot.SheetName + "!" + snapshot.Range, diagnostics.AsReadOnly());
            }

            OfficeRasterImage image = RenderRaster(snapshot, options, diagnostics);
            return new OfficeImageExportResult(format, image.Width, image.Height, OfficePngWriter.Encode(image), snapshot.SheetName, snapshot.SheetName + "!" + snapshot.Range, diagnostics.AsReadOnly());
        }

        internal static OfficeRasterImage RenderRaster(ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics = null) {
            int width = ScaledWidth(snapshot, options);
            int height = ScaledHeight(snapshot, options);
            OfficeRasterImage image = new OfficeRasterImage(width, height, options.BackgroundColor);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            double scale = options.Scale;
            Dictionary<string, ExcelVisualConditionalDataBar> dataBars = BuildDataBarMap(snapshot.ConditionalDataBars);
            Dictionary<string, ExcelVisualCell> cellsByAddress = BuildCellMap(snapshot.Cells);

            foreach (ExcelVisualCell cell in snapshot.Cells) {
                if (cell.CoveredByMerge) {
                    continue;
                }

                double x = cell.X * scale;
                double y = cell.Y * scale;
                double w = cell.Width * scale;
                double h = cell.Height * scale;
                DrawRasterCellFill(canvas, cell, snapshot, options, scale, diagnostics);
                if (dataBars.TryGetValue(Key(cell.Row, cell.Column), out ExcelVisualConditionalDataBar? dataBar)) {
                    DrawDataBar(canvas, dataBar, scale);
                }

                if (options.ShowGridlines) {
                    canvas.DrawRectangle(x, y, w, h, options.GridlineColor, Math.Max(1D, scale));
                }

                DrawBorders(canvas, cell, scale);
            }

            foreach (ExcelVisualCell cell in snapshot.Cells) {
                if (cell.CoveredByMerge) {
                    continue;
                }

                DrawRasterCellText(canvas, cell, snapshot, options, scale, cellsByAddress, diagnostics);
            }

            RenderRasterSparklines(canvas, snapshot, options);
            RenderRasterCommentIndicators(canvas, snapshot, options);
            RenderRasterDrawingLayers(canvas, snapshot, options, diagnostics);

            return image;
        }

        internal static string RenderSvg(ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics = null) {
            int width = ScaledWidth(snapshot, options);
            int height = ScaledHeight(snapshot, options);
            double scale = options.Scale;
            StringBuilder builder = new StringBuilder();
            builder.Append("<svg xmlns=\"http://www.w3.org/2000/svg\"");
            builder.AppendNumberAttribute("width", width)
                .AppendNumberAttribute("height", height)
                .AppendAttribute("viewBox", "0 0 " + Number(width) + " " + Number(height))
                .Append('>');
            var backgroundAttributes = new StringBuilder();
            backgroundAttributes.AppendPaintAttribute("fill", options.BackgroundColor);
            builder.AppendRectElement(0D, 0D, width, height, backgroundAttributes.ToString());
            OfficeRasterCanvas textMeasureCanvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1, OfficeColor.Transparent));
            Dictionary<string, ExcelVisualConditionalDataBar> dataBars = BuildDataBarMap(snapshot.ConditionalDataBars);
            Dictionary<string, ExcelVisualCell> cellsByAddress = BuildCellMap(snapshot.Cells);

            foreach (ExcelVisualCell cell in snapshot.Cells) {
                if (cell.CoveredByMerge) {
                    continue;
                }

                double x = cell.X * scale;
                double y = cell.Y * scale;
                double w = cell.Width * scale;
                double h = cell.Height * scale;
                AppendSvgCellFill(builder, cell, snapshot, options, scale, diagnostics);
                if (dataBars.TryGetValue(Key(cell.Row, cell.Column), out ExcelVisualConditionalDataBar? dataBar)) {
                    AppendSvgDataBar(builder, dataBar, scale);
                }

                if (options.ShowGridlines) {
                    var gridlineAttributes = new StringBuilder();
                    gridlineAttributes
                        .AppendAttribute("fill", "none")
                        .AppendPaintAttribute("stroke", options.GridlineColor)
                        .AppendNumberAttribute("stroke-width", Math.Max(1D, scale));
                    builder.AppendRectElement(x, y, w, h, gridlineAttributes.ToString());
                }

                AppendSvgBorders(builder, cell, scale);
            }

            foreach (ExcelVisualCell cell in snapshot.Cells) {
                if (cell.CoveredByMerge) {
                    continue;
                }

                AppendSvgCellText(builder, cell, snapshot, options, textMeasureCanvas, cellsByAddress, diagnostics);
            }

            AppendSvgSparklines(builder, snapshot, options);
            AppendSvgCommentIndicators(builder, snapshot, options);
            AppendSvgDrawingLayers(builder, snapshot, options, diagnostics, textMeasureCanvas);

            builder.Append("</svg>");
            return builder.ToString();
        }

        private static void DrawDataBar(OfficeRasterCanvas canvas, ExcelVisualConditionalDataBar dataBar, double scale) {
            OfficeColor color = ResolveArgb(dataBar.ColorArgb) ?? OfficeColor.FromRgb(91, 155, 213);
            OfficeDataBarRenderer.DrawRaster(
                canvas,
                dataBar.X * scale,
                dataBar.Y * scale,
                dataBar.Width * scale,
                dataBar.Height * scale,
                dataBar.StartRatio,
                dataBar.Ratio,
                color,
                2D * scale);
        }

        private static void AppendSvgDataBar(StringBuilder builder, ExcelVisualConditionalDataBar dataBar, double scale) {
            OfficeColor color = ResolveArgb(dataBar.ColorArgb) ?? OfficeColor.FromRgb(91, 155, 213);
            OfficeDataBarRenderer.AppendSvg(
                builder,
                dataBar.X * scale,
                dataBar.Y * scale,
                dataBar.Width * scale,
                dataBar.Height * scale,
                dataBar.StartRatio,
                dataBar.Ratio,
                color,
                2D * scale);
        }

        private static Dictionary<string, ExcelVisualConditionalDataBar> BuildDataBarMap(IReadOnlyList<ExcelVisualConditionalDataBar> dataBars) {
            var resolved = new Dictionary<string, ExcelVisualConditionalDataBar>(StringComparer.Ordinal);
            foreach (ExcelVisualConditionalDataBar dataBar in dataBars) {
                resolved[Key(dataBar.Row, dataBar.Column)] = dataBar;
            }

            return resolved;
        }

        private static Dictionary<string, ExcelVisualCell> BuildCellMap(IReadOnlyList<ExcelVisualCell> cells) {
            var resolved = new Dictionary<string, ExcelVisualCell>(cells.Count, StringComparer.Ordinal);
            foreach (ExcelVisualCell cell in cells) {
                resolved[Key(cell.Row, cell.Column)] = cell;
            }

            return resolved;
        }

        private static void RenderRasterCharts(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            foreach (ExcelVisualChart chart in snapshot.Charts) {
                RenderRasterChart(canvas, snapshot, chart, options, diagnostics);
            }
        }

        private static void AppendSvgCharts(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            foreach (ExcelVisualChart chart in snapshot.Charts) {
                AppendSvgChart(builder, snapshot, chart, options, diagnostics);
            }
        }

        private static void RenderRasterChart(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelVisualChart chart, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            double scale = options.Scale;
            if (!TryCreateOfficeChartSnapshot(chart.Snapshot, chart.Width, chart.Height, diagnostics, snapshot.SheetName, out OfficeChartSnapshot? officeSnapshot) || officeSnapshot == null) {
                return;
            }

            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(officeSnapshot);
            OfficeRasterImage chartImage = OfficeDrawingRasterRenderer.Render(drawing, scale, OfficeColor.White);
            canvas.DrawImage(chartImage, chart.X * scale, chart.Y * scale, chart.Width * scale, chart.Height * scale);
        }

        private static void AppendSvgChart(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelVisualChart chart, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            double scale = options.Scale;
            if (!TryCreateOfficeChartSnapshot(chart.Snapshot, chart.Width * scale, chart.Height * scale, diagnostics, snapshot.SheetName, out OfficeChartSnapshot? officeSnapshot) || officeSnapshot == null) {
                return;
            }

            string chartSvg = OfficeDrawingSvgExporter.ToSvg(OfficeChartDrawingRenderer.Render(officeSnapshot));
            builder.AppendNestedSvg(
                chart.X * scale,
                chart.Y * scale,
                chart.Width * scale,
                chart.Height * scale,
                OfficeSvgFormatting.ExtractSvgInner(chartSvg));
        }

        private static bool TryCreateOfficeChartSnapshot(ExcelChartSnapshot snapshot, double width, double height, List<OfficeImageExportDiagnostic>? diagnostics, string sheetName, out OfficeChartSnapshot? officeSnapshot) {
            officeSnapshot = null;
            if (!TryMapChartKind(snapshot.ChartType, out OfficeChartKind kind, out string? approximation)) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartKindUnsupported,
                    "Worksheet chart type is not supported by the shared image renderer yet.",
                    sheetName + "!" + snapshot.Name));
                return false;
            }

            if (approximation != null) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartKindApproximated,
                    approximation,
                    sheetName + "!" + snapshot.Name));
            }

            OfficeChartData data = new OfficeChartData(
                snapshot.Data.Categories,
                snapshot.Data.Series.Select(series => new OfficeChartSeries(
                    series.Name,
                    series.Values,
                    xValues: null,
                    ResolveArgb(series.SeriesColorArgb),
                    ResolvePointColors(series.PointColorArgb),
                    series.ShowMarkers,
                    markerSize: series.MarkerSize,
                    markerShape: series.MarkerShape,
                    markerOutlineColor: ResolveArgb(series.MarkerOutlineColorArgb),
                    markerOutlineWidth: series.MarkerOutlineWidth,
                    strokeWidth: series.SeriesLineWidth,
                    strokeDashStyle: series.SeriesLineDashStyle)));
            officeSnapshot = new OfficeChartSnapshot(snapshot.Name, snapshot.Title, kind, data, Math.Max(1D, width), Math.Max(1D, height), snapshot.Style, snapshot.Layout);
            return true;
        }

        private static bool TryMapChartKind(ExcelChartType type, out OfficeChartKind kind, out string? approximation) {
            approximation = null;
            switch (type) {
                case ExcelChartType.ColumnClustered:
                    kind = OfficeChartKind.ColumnClustered;
                    return true;
                case ExcelChartType.ColumnStacked:
                    kind = OfficeChartKind.ColumnStacked;
                    return true;
                case ExcelChartType.ColumnStacked100:
                    kind = OfficeChartKind.ColumnStacked100;
                    return true;
                case ExcelChartType.BarClustered:
                    kind = OfficeChartKind.BarClustered;
                    return true;
                case ExcelChartType.BarStacked:
                    kind = OfficeChartKind.BarStacked;
                    return true;
                case ExcelChartType.BarStacked100:
                    kind = OfficeChartKind.BarStacked100;
                    return true;
                case ExcelChartType.Line:
                    kind = OfficeChartKind.Line;
                    return true;
                case ExcelChartType.LineStacked:
                    kind = OfficeChartKind.LineStacked;
                    return true;
                case ExcelChartType.LineStacked100:
                    kind = OfficeChartKind.LineStacked100;
                    return true;
                case ExcelChartType.Area:
                    kind = OfficeChartKind.Area;
                    return true;
                case ExcelChartType.AreaStacked:
                    kind = OfficeChartKind.AreaStacked;
                    return true;
                case ExcelChartType.AreaStacked100:
                    kind = OfficeChartKind.AreaStacked100;
                    return true;
                case ExcelChartType.Pie:
                case ExcelChartType.Pie3D:
                case ExcelChartType.PieOfPie:
                case ExcelChartType.BarOfPie:
                    kind = OfficeChartKind.Pie;
                    approximation = type == ExcelChartType.Pie ? null : "Excel pie variant is rendered as a standard 2D pie chart.";
                    return true;
                case ExcelChartType.Doughnut:
                    kind = OfficeChartKind.Doughnut;
                    return true;
                case ExcelChartType.Scatter:
                case ExcelChartType.Bubble:
                    kind = OfficeChartKind.Scatter;
                    approximation = type == ExcelChartType.Bubble ? "Excel bubble chart is rendered as a scatter chart without bubble-size encoding." : null;
                    return true;
                case ExcelChartType.Radar:
                    kind = OfficeChartKind.Radar;
                    return true;
                case ExcelChartType.Column3DClustered:
                case ExcelChartType.Column3DStacked:
                case ExcelChartType.Column3DStacked100:
                    kind = type == ExcelChartType.Column3DStacked ? OfficeChartKind.ColumnStacked : type == ExcelChartType.Column3DStacked100 ? OfficeChartKind.ColumnStacked100 : OfficeChartKind.ColumnClustered;
                    approximation = "Excel 3D column chart is rendered as a 2D column chart.";
                    return true;
                case ExcelChartType.Bar3DClustered:
                case ExcelChartType.Bar3DStacked:
                case ExcelChartType.Bar3DStacked100:
                    kind = type == ExcelChartType.Bar3DStacked ? OfficeChartKind.BarStacked : type == ExcelChartType.Bar3DStacked100 ? OfficeChartKind.BarStacked100 : OfficeChartKind.BarClustered;
                    approximation = "Excel 3D bar chart is rendered as a 2D bar chart.";
                    return true;
                case ExcelChartType.Line3D:
                    kind = OfficeChartKind.Line;
                    approximation = "Excel 3D line chart is rendered as a 2D line chart.";
                    return true;
                case ExcelChartType.Area3D:
                case ExcelChartType.Area3DStacked:
                case ExcelChartType.Area3DStacked100:
                    kind = type == ExcelChartType.Area3DStacked ? OfficeChartKind.AreaStacked : type == ExcelChartType.Area3DStacked100 ? OfficeChartKind.AreaStacked100 : OfficeChartKind.Area;
                    approximation = "Excel 3D area chart is rendered as a 2D area chart.";
                    return true;
                case ExcelChartType.Stock:
                case ExcelChartType.Surface:
                case ExcelChartType.SurfaceWireframe:
                case ExcelChartType.SurfaceContour:
                case ExcelChartType.SurfaceContourWireframe:
                    kind = OfficeChartKind.Line;
                    approximation = "Excel stock/surface chart is rendered as a line chart approximation.";
                    return true;
                default:
                    kind = OfficeChartKind.ColumnClustered;
                    return false;
            }
        }

        private static void DrawBorders(OfficeRasterCanvas canvas, ExcelVisualCell cell, double scale) {
            ExcelCellBorderSnapshot? border = cell.Style.Border;
            if (border == null) {
                return;
            }

            double x = cell.X * scale;
            double y = cell.Y * scale;
            double w = cell.Width * scale;
            double h = cell.Height * scale;
            DrawBorder(canvas, x, y, x + w, y, border.Top, scale);
            DrawBorder(canvas, x + w, y, x + w, y + h, border.Right, scale);
            DrawBorder(canvas, x, y + h, x + w, y + h, border.Bottom, scale);
            DrawBorder(canvas, x, y, x, y + h, border.Left, scale);
            if (border.DiagonalDown) {
                DrawBorder(canvas, x, y, x + w, y + h, border.Diagonal, scale);
            }

            if (border.DiagonalUp) {
                DrawBorder(canvas, x, y + h, x + w, y, border.Diagonal, scale);
            }
        }

        private static void DrawBorder(OfficeRasterCanvas canvas, double x1, double y1, double x2, double y2, ExcelBorderSideSnapshot? border, double scale) {
            if (border == null || !TryResolveBorderStroke(border.Style, scale, out ExcelBorderStroke stroke)) {
                return;
            }

            OfficeColor color = ResolveArgb(border.ColorArgb) ?? OfficeColor.Black;
            if (stroke.DoubleLine) {
                canvas.DrawParallelStyledLine(x1, y1, x2, y2, color, stroke.Width, stroke.DoubleLineOffset);
                return;
            }

            canvas.DrawStyledLine(x1, y1, x2, y2, color, stroke.Width, stroke.DashStyle);
        }

        private static void AppendSvgBorders(StringBuilder builder, ExcelVisualCell cell, double scale) {
            ExcelCellBorderSnapshot? border = cell.Style.Border;
            if (border == null) {
                return;
            }

            double x = cell.X * scale;
            double y = cell.Y * scale;
            double w = cell.Width * scale;
            double h = cell.Height * scale;
            AppendSvgLine(builder, x, y, x + w, y, border.Top, scale);
            AppendSvgLine(builder, x + w, y, x + w, y + h, border.Right, scale);
            AppendSvgLine(builder, x, y + h, x + w, y + h, border.Bottom, scale);
            AppendSvgLine(builder, x, y, x, y + h, border.Left, scale);
            if (border.DiagonalDown) {
                AppendSvgLine(builder, x, y, x + w, y + h, border.Diagonal, scale);
            }

            if (border.DiagonalUp) {
                AppendSvgLine(builder, x, y + h, x + w, y, border.Diagonal, scale);
            }
        }

        private static void AppendSvgLine(StringBuilder builder, double x1, double y1, double x2, double y2, ExcelBorderSideSnapshot? border, double scale) {
            if (border == null || !TryResolveBorderStroke(border.Style, scale, out ExcelBorderStroke stroke)) {
                return;
            }

            OfficeColor color = ResolveArgb(border.ColorArgb) ?? OfficeColor.Black;
            if (stroke.DoubleLine) {
                builder.AppendParallelLineElements(x1, y1, x2, y2, color, stroke.Width, stroke.DoubleLineOffset);
                return;
            }

            AppendSvgStyledLine(builder, x1, y1, x2, y2, color, stroke.Width, stroke.DashStyle);
        }

        private static void AppendSvgStyledLine(StringBuilder builder, double x1, double y1, double x2, double y2, OfficeColor color, double width, OfficeStrokeDashStyle dashStyle) {
            string? dashArray = dashStyle.GetSvgDashArray(width);
            builder.AppendLineElement(
                x1,
                y1,
                x2,
                y2,
                color,
                width,
                dashStyle,
                dashArray != null && dashStyle == OfficeStrokeDashStyle.Dot ? OfficeStrokeLineCap.Round : null);
        }

        private static bool TryResolveBorderStroke(string? style, double scale, out ExcelBorderStroke stroke) {
            double thin = Math.Max(1D, scale);
            string normalized = string.IsNullOrWhiteSpace(style) ? "thin" : style!.Trim().ToLowerInvariant();
            switch (normalized) {
                case "none":
                    stroke = default;
                    return false;
                case "hair":
                    stroke = new ExcelBorderStroke(Math.Max(0.5D, scale * 0.5D), OfficeStrokeDashStyle.Solid, false, 0D);
                    return true;
                case "medium":
                    stroke = new ExcelBorderStroke(Math.Max(2D, scale * 2D), OfficeStrokeDashStyle.Solid, false, 0D);
                    return true;
                case "thick":
                    stroke = new ExcelBorderStroke(Math.Max(3D, scale * 3D), OfficeStrokeDashStyle.Solid, false, 0D);
                    return true;
                case "dotted":
                    stroke = new ExcelBorderStroke(thin, OfficeStrokeDashStyle.Dot, false, 0D);
                    return true;
                case "dashed":
                    stroke = new ExcelBorderStroke(thin, OfficeStrokeDashStyle.Dash, false, 0D);
                    return true;
                case "dashdot":
                    stroke = new ExcelBorderStroke(thin, OfficeStrokeDashStyle.DashDot, false, 0D);
                    return true;
                case "dashdotdot":
                    stroke = new ExcelBorderStroke(thin, OfficeStrokeDashStyle.DashDotDot, false, 0D);
                    return true;
                case "mediumdashed":
                    stroke = new ExcelBorderStroke(Math.Max(2D, scale * 2D), OfficeStrokeDashStyle.Dash, false, 0D);
                    return true;
                case "mediumdashdot":
                case "slantdashdot":
                    stroke = new ExcelBorderStroke(Math.Max(2D, scale * 2D), OfficeStrokeDashStyle.DashDot, false, 0D);
                    return true;
                case "mediumdashdotdot":
                    stroke = new ExcelBorderStroke(Math.Max(2D, scale * 2D), OfficeStrokeDashStyle.DashDotDot, false, 0D);
                    return true;
                case "double":
                    stroke = new ExcelBorderStroke(thin, OfficeStrokeDashStyle.Solid, true, Math.Max(3D, scale * 3D));
                    return true;
                default:
                    stroke = new ExcelBorderStroke(thin, OfficeStrokeDashStyle.Solid, false, 0D);
                    return true;
            }
        }

        private static OfficeTextAlignment ResolveTextAlignment(string? alignment) {
            if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Center;
            }

            if (string.Equals(alignment, "right", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Right;
            }

            return OfficeTextAlignment.Left;
        }

        private static OfficeTextVerticalAlignment ResolveTextVerticalAlignment(string? alignment) {
            if (string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextVerticalAlignment.Center;
            }

            if (string.Equals(alignment, "top", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextVerticalAlignment.Top;
            }

            return OfficeTextVerticalAlignment.Bottom;
        }

        private static OfficeFontStyle ResolveFontStyle(ExcelCellStyleSnapshot style) {
            OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
            if (style.Bold) {
                fontStyle |= OfficeFontStyle.Bold;
            }

            if (style.Italic) {
                fontStyle |= OfficeFontStyle.Italic;
            }

            if (style.Underline) {
                fontStyle |= OfficeFontStyle.Underline;
            }

            return fontStyle;
        }

        private static OfficeColor? ResolveArgb(string? argb) {
            if (string.IsNullOrWhiteSpace(argb)) {
                return null;
            }

            string value = argb!.Trim().TrimStart('#');
            if (value.Length == 8) {
                string rgba = value.Substring(2) + value.Substring(0, 2);
                return OfficeColor.TryParseHex(rgba, out OfficeColor color) ? color : null;
            }

            return OfficeColor.TryParseHex(value, out OfficeColor rgbColor) ? rgbColor : null;
        }

        private static IReadOnlyList<OfficeColor?>? ResolvePointColors(IReadOnlyList<string?>? colors) {
            if (colors == null) {
                return null;
            }

            var resolved = new OfficeColor?[colors.Count];
            bool any = false;
            for (int i = 0; i < colors.Count; i++) {
                resolved[i] = ResolveArgb(colors[i]);
                any |= resolved[i].HasValue;
            }

            return any ? resolved : null;
        }

        private static int ScaledWidth(ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) => Math.Max(1, (int)Math.Ceiling(snapshot.Width * options.Scale));

        private static int ScaledHeight(ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) => Math.Max(1, (int)Math.Ceiling(snapshot.Height * options.Scale));

        private static string Number(double value) => OfficeSvgFormatting.FormatNumber(value);

        private static string Key(int row, int column) => row.ToString(CultureInfo.InvariantCulture) + ":" + column.ToString(CultureInfo.InvariantCulture);

        private static string EscapeXml(string text) => OfficeSvgFormatting.Escape(text);

        private readonly struct ExcelBorderStroke {
            internal ExcelBorderStroke(double width, OfficeStrokeDashStyle dashStyle, bool doubleLine, double doubleLineOffset) {
                Width = width;
                DashStyle = dashStyle;
                DoubleLine = doubleLine;
                DoubleLineOffset = doubleLineOffset;
            }

            internal double Width { get; }

            internal OfficeStrokeDashStyle DashStyle { get; }

            internal bool DoubleLine { get; }

            internal double DoubleLineOffset { get; }
        }
    }
}
