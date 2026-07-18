using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        internal static OfficeImageExportResult Render(
            ExcelRangeVisualSnapshot snapshot,
            OfficeImageExportFormat format,
            ExcelImageExportOptions options,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            if (format == OfficeImageExportFormat.Svg) {
                string svg = RenderSvg(snapshot, options, diagnostics);
                cancellationToken.ThrowIfCancellationRequested();
                return options.EnsureAccepted(new OfficeImageExportResult(format, ScaledWidth(snapshot, options), ScaledHeight(snapshot, options), Encoding.UTF8.GetBytes(svg), snapshot.SheetName, snapshot.SheetName + "!" + snapshot.Range, diagnostics.AsReadOnly()));
            }

            string source = snapshot.SheetName + "!" + snapshot.Range;
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                snapshot.Width,
                snapshot.Height,
                format,
                options,
                source);
            if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
            ExcelImageExportOptions renderOptions = options.Clone();
            renderOptions.Scale = plan.Limit.Scale;
            OfficeRasterImage image = RenderRaster(snapshot, renderOptions, diagnostics, cancellationToken);
            byte[] bytes = OfficeRasterImageEncoder.Encode(image, format, options.RasterEncoding);
            cancellationToken.ThrowIfCancellationRequested();
            return options.EnsureAccepted(new OfficeImageExportResult(format, image.Width, image.Height, bytes, snapshot.SheetName, source, diagnostics.AsReadOnly()));
        }

        internal static OfficeRasterImage RenderRaster(
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            List<OfficeImageExportDiagnostic>? diagnostics = null,
            CancellationToken cancellationToken = default) {
            int width = ScaledWidth(snapshot, options);
            int height = ScaledHeight(snapshot, options);
            OfficeRasterImage image = new OfficeRasterImage(width, height, options.BackgroundColor);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image, fonts: options.Fonts);
            double scale = options.Scale;
            Dictionary<string, ExcelVisualConditionalDataBar> dataBars = BuildDataBarMap(snapshot.ConditionalDataBars);
            Dictionary<string, ExcelVisualCell> cellsByAddress = BuildCellMap(snapshot.Cells);

            foreach (ExcelVisualCell cell in snapshot.Cells) {
                cancellationToken.ThrowIfCancellationRequested();
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
                cancellationToken.ThrowIfCancellationRequested();
                if (cell.CoveredByMerge) {
                    continue;
                }

                DrawRasterCellText(canvas, cell, snapshot, options, scale, cellsByAddress, diagnostics);
            }

            RenderRasterConditionalIcons(canvas, snapshot, options);
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
            OfficeTextMeasurer textMeasurer = OfficeTextMeasurer.Create();
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

                AppendSvgCellText(builder, cell, snapshot, options, textMeasurer, cellsByAddress, diagnostics);
            }

            AppendSvgConditionalIcons(builder, snapshot, options);
            AppendSvgSparklines(builder, snapshot, options);
            AppendSvgCommentIndicators(builder, snapshot, options);
            AppendSvgDrawingLayers(builder, snapshot, options, diagnostics, textMeasurer);

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
                string key = Key(dataBar.Row, dataBar.Column);
                if (!resolved.ContainsKey(key)) {
                    resolved[key] = dataBar;
                }
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

        private static void RenderRasterChart(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelVisualChart chart, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            double scale = options.Scale;
            if (!TryCreateOfficeChartSnapshot(chart.Snapshot, chart.Width, chart.Height, diagnostics, snapshot.SheetName, out OfficeChartSnapshot? officeSnapshot) || officeSnapshot == null) {
                return;
            }

            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(officeSnapshot);
            drawing.Fonts.AddRange(options.Fonts);
            drawing.AppendFontDiagnostics(
                diagnostics ?? new List<OfficeImageExportDiagnostic>(),
                snapshot.SheetName + "!" + chart.Snapshot.Name);
            OfficeColor chartBackground = officeSnapshot.Style?.ShowBackground == false ? OfficeColor.Transparent : OfficeColor.White;
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                drawing.Width,
                drawing.Height,
                OfficeImageExportFormat.Png,
                options,
                snapshot.SheetName + "!" + chart.Snapshot.Name);
            if (plan.Diagnostic != null) diagnostics?.Add(plan.Diagnostic);
            OfficeRasterImage chartImage = OfficeDrawingRasterRenderer.Render(drawing, plan.Limit.Scale, chartBackground);
            canvas.DrawImage(chartImage, chart.X * scale, chart.Y * scale, chart.Width * scale, chart.Height * scale);
        }

        private static void AppendSvgChart(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelVisualChart chart, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            double scale = options.Scale;
            if (!TryCreateOfficeChartSnapshot(chart.Snapshot, chart.Width, chart.Height, diagnostics, snapshot.SheetName, out OfficeChartSnapshot? officeSnapshot) || officeSnapshot == null) {
                return;
            }

            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(officeSnapshot);
            drawing.Fonts.AddRange(options.Fonts);
            drawing.AppendFontDiagnostics(
                diagnostics ?? new List<OfficeImageExportDiagnostic>(),
                snapshot.SheetName + "!" + chart.Snapshot.Name);
            string chartSvg = OfficeDrawingSvgExporter.ToSvg(drawing);
            builder.AppendNestedSvg(
                chart.X * scale,
                chart.Y * scale,
                chart.Width * scale,
                chart.Height * scale,
                drawing.Width,
                drawing.Height,
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

            if (snapshot.Data.Series.Any(series => series.ChartType.HasValue &&
                series.ChartType.Value != snapshot.ChartType &&
                (!TryMapSeriesRenderKind(series.ChartType.Value, out _, out string? seriesApproximation) || seriesApproximation != null))) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartKindApproximated,
                    "Excel combo chart includes a series type that is rendered through a fallback chart kind.",
                    sheetName + "!" + snapshot.Name));
            }

            OfficeChartData data = new OfficeChartData(
                snapshot.Data.Categories,
                snapshot.Data.Series.Select(series => new OfficeChartSeries(
                    series.Name,
                    series.Values,
                    series.XValues,
                    ResolveArgb(series.SeriesColorArgb),
                    ResolvePointColors(series.PointColorArgb),
                    series.ShowMarkers,
                    connectLine: series.ConnectLine,
                    markerSize: series.MarkerSize,
                    markerShape: series.MarkerShape,
                    markerOutlineColor: ResolveArgb(series.MarkerOutlineColorArgb),
                    markerOutlineWidth: series.MarkerOutlineWidth,
                    strokeWidth: series.SeriesLineWidth,
                    strokeDashStyle: series.SeriesLineDashStyle,
                    renderKind: TryMapSeriesRenderKind(series.ChartType ?? snapshot.ChartType, out OfficeChartKind seriesKind, out _) ? seriesKind : null,
                    axisGroup: series.AxisGroup == ExcelChartAxisGroup.Secondary
                        ? OfficeChartAxisGroup.Secondary
                        : OfficeChartAxisGroup.Primary)));
            officeSnapshot = new OfficeChartSnapshot(snapshot.Name, snapshot.Title, kind, data, Math.Max(1D, width), Math.Max(1D, height), snapshot.Style, snapshot.Layout);
            return true;
        }

        private static bool TryMapSeriesRenderKind(ExcelChartType type, out OfficeChartKind kind, out string? approximation) {
            if (!TryMapChartKind(type, out kind, out approximation)) {
                return false;
            }

            return kind == OfficeChartKind.ColumnClustered ||
                kind == OfficeChartKind.ColumnStacked ||
                kind == OfficeChartKind.ColumnStacked100 ||
                kind == OfficeChartKind.BarClustered ||
                kind == OfficeChartKind.BarStacked ||
                kind == OfficeChartKind.BarStacked100 ||
                kind == OfficeChartKind.Line ||
                kind == OfficeChartKind.LineStacked ||
                kind == OfficeChartKind.LineStacked100 ||
                kind == OfficeChartKind.Area ||
                kind == OfficeChartKind.AreaStacked ||
                kind == OfficeChartKind.AreaStacked100 ||
                kind == OfficeChartKind.Scatter;
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
            if (TryCreateBorderBox(border, scale, out OfficeBorderBox borderBox)) {
                OfficeBorderBoxRenderer.DrawRaster(canvas, x, y, w, h, borderBox);
                return;
            }

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
            if (TryCreateBorderBox(border, scale, out OfficeBorderBox borderBox)) {
                OfficeBorderBoxRenderer.AppendSvg(builder, x, y, w, h, borderBox);
                return;
            }

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

        private static bool TryCreateBorderBox(ExcelCellBorderSnapshot border, double scale, out OfficeBorderBox borderBox) {
            borderBox = default;
            if (!TryCreateBorderSide(border.Left, scale, out OfficeBorderSide? left) ||
                !TryCreateBorderSide(border.Top, scale, out OfficeBorderSide? top) ||
                !TryCreateBorderSide(border.Right, scale, out OfficeBorderSide? right) ||
                !TryCreateBorderSide(border.Bottom, scale, out OfficeBorderSide? bottom)) {
                return false;
            }

            OfficeBorderSide? diagonalDown = null;
            OfficeBorderSide? diagonalUp = null;
            if (border.DiagonalDown || border.DiagonalUp) {
                if (!TryCreateBorderSide(border.Diagonal, scale, out OfficeBorderSide? diagonal)) {
                    return false;
                }

                diagonalDown = border.DiagonalDown ? diagonal : null;
                diagonalUp = border.DiagonalUp ? diagonal : null;
            }

            borderBox = new OfficeBorderBox(left, top, right, bottom, diagonalDown, diagonalUp);
            return borderBox.HasVisibleSide;
        }

        private static bool TryCreateBorderSide(ExcelBorderSideSnapshot? border, double scale, out OfficeBorderSide? side) {
            side = null;
            if (border == null || !TryResolveBorderStroke(border.Style, scale, out ExcelBorderStroke stroke)) {
                return true;
            }

            OfficeColor color = ResolveArgb(border.ColorArgb) ?? OfficeColor.Black;
            side = new OfficeBorderSide(
                color,
                stroke.Width,
                stroke.DashStyle,
                stroke.DoubleLine ? OfficeBorderLineKind.Double : OfficeBorderLineKind.Single,
                stroke.DoubleLineOffset);
            return true;
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
