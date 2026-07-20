using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void DrawRasterCellFill(
            OfficeRasterCanvas canvas,
            ExcelVisualCell cell,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            double scale,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            string pattern = NormalizeFillPattern(cell.Style.FillPatternType);
            OfficeColor background = ResolveFillBackground(cell.Style, options);
            double x = cell.X * scale;
            double y = cell.Y * scale;
            double w = cell.Width * scale;
            double h = cell.Height * scale;
            OfficeLinearGradient? gradient = CreateLinearGradient(cell.Style);

            if (gradient != null) {
                canvas.FillLinearGradientRectangle(x, y, w, h, gradient);
            } else {
                canvas.FillRectangle(x, y, w, h, background);
                AddGradientDiagnosticIfNeeded(snapshot, cell, diagnostics);
            }

            if (!IsPatternFill(pattern)) {
                return;
            }

            AddPatternApproximationDiagnostic(snapshot, cell, pattern, diagnostics);
            OfficeColor foreground = ResolveFillForeground(cell.Style, background);
            using (canvas.PushClipRectangle(x, y, w, h)) {
                DrawRasterPattern(canvas, x, y, w, h, pattern, foreground, scale);
            }
        }

        private static void AppendSvgCellFill(
            StringBuilder builder,
            ExcelVisualCell cell,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            double scale,
            List<OfficeImageExportDiagnostic>? diagnostics) {
            string pattern = NormalizeFillPattern(cell.Style.FillPatternType);
            OfficeColor background = ResolveFillBackground(cell.Style, options);
            double x = cell.X * scale;
            double y = cell.Y * scale;
            double w = cell.Width * scale;
            double h = cell.Height * scale;
            OfficeLinearGradient? gradient = CreateLinearGradient(cell.Style);

            string? gradientId = null;
            if (gradient != null) {
                gradientId = "xl-gradient-" + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + cell.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);
                builder.AppendLinearGradientDefinition(gradientId, gradient);
            }

            var fillAttributes = new StringBuilder();
            if (gradientId != null) {
                fillAttributes.AppendAttribute("fill", "url(#" + gradientId + ")");
            } else {
                fillAttributes.AppendPaintAttribute("fill", background);
                AddGradientDiagnosticIfNeeded(snapshot, cell, diagnostics);
            }

            builder.AppendRectElement(x, y, w, h, fillAttributes.ToString());
            if (!IsPatternFill(pattern)) {
                return;
            }

            AddPatternApproximationDiagnostic(snapshot, cell, pattern, diagnostics);
            OfficeColor foreground = ResolveFillForeground(cell.Style, background);
            string clipId = "xl-fill-" + cell.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + cell.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);
            builder.AppendRectClipPathDefinition(clipId, x, y, w, h);
            builder.Append("<g").AppendClipPathReference(clipId).Append(">");
            AppendSvgPattern(builder, x, y, w, h, pattern, foreground, scale);
            builder.Append("</g>");
        }

        private static OfficeColor ResolveFillBackground(ExcelCellStyleSnapshot style, ExcelImageExportOptions options) {
            string pattern = NormalizeFillPattern(style.FillPatternType);
            if (pattern == "solid") {
                return ResolveArgb(style.FillColorArgb)
                    ?? ResolveArgb(style.FillPatternForegroundColorArgb)
                    ?? ResolveArgb(style.FillPatternBackgroundColorArgb)
                    ?? options.BackgroundColor;
            }

            return ResolveArgb(style.FillPatternBackgroundColorArgb)
                ?? ResolveArgb(style.FillColorArgb)
                ?? options.BackgroundColor;
        }

        private static OfficeColor ResolveFillForeground(ExcelCellStyleSnapshot style, OfficeColor background) =>
            ResolveArgb(style.FillPatternForegroundColorArgb)
            ?? ResolveArgb(style.FillColorArgb)
            ?? OfficeColor.FromRgba(background.R, background.G, background.B, 180);

        private static OfficeLinearGradient? CreateLinearGradient(ExcelCellStyleSnapshot style) {
            if (style.FillGradientStops.Count >= 2) {
                var stops = new List<OfficeGradientStop>(style.FillGradientStops.Count);
                for (int i = 0; i < style.FillGradientStops.Count; i++) {
                    ExcelGradientFillStopSnapshot stop = style.FillGradientStops[i];
                    OfficeColor? color = ResolveArgb(stop.ColorArgb);
                    if (!color.HasValue) {
                        return null;
                    }

                    stops.Add(new OfficeGradientStop(stop.Offset, color.Value));
                }

                return OfficeLinearGradient.FromAngle(stops, style.FillGradientDegree ?? 0D);
            }

            if (style.FillGradientStartColorArgb == null || style.FillGradientEndColorArgb == null) {
                return null;
            }

            OfficeColor? start = ResolveArgb(style.FillGradientStartColorArgb);
            OfficeColor? end = ResolveArgb(style.FillGradientEndColorArgb);
            if (!start.HasValue || !end.HasValue) {
                return null;
            }

            return OfficeLinearGradient.FromAngle(start.Value, end.Value, style.FillGradientDegree ?? 0D);
        }

        private static string NormalizeFillPattern(string? pattern) {
            if (string.IsNullOrWhiteSpace(pattern)) {
                return string.Empty;
            }

            return pattern!.Trim().Replace("_", string.Empty).Replace("-", string.Empty).ToLowerInvariant();
        }

        private static bool IsPatternFill(string pattern) =>
            pattern.Length > 0 && pattern != "none" && pattern != "solid" && pattern != "gradient";

        private static void DrawRasterPattern(OfficeRasterCanvas canvas, double x, double y, double w, double h, string pattern, OfficeColor foreground, double scale) {
            PatternMetrics metrics = GetPatternMetrics(pattern, scale);
            canvas.DrawHatchPatternRectangle(x, y, w, h, foreground, metrics.Step, metrics.Width, MapExcelPatternToHatchPattern(pattern));
        }

        private static void AppendSvgPattern(StringBuilder builder, double x, double y, double w, double h, string pattern, OfficeColor foreground, double scale) {
            PatternMetrics metrics = GetPatternMetrics(pattern, scale);
            builder.AppendHatchPatternRectangle(x, y, w, h, foreground, metrics.Step, metrics.Width, MapExcelPatternToHatchPattern(pattern));
        }

        private static OfficeHatchPatternKind MapExcelPatternToHatchPattern(string pattern) {
            switch (pattern) {
                case "gray0625":
                    return OfficeHatchPatternKind.Percent6_25;
                case "gray125":
                    return OfficeHatchPatternKind.Percent12_5;
                case "lightgray":
                    return OfficeHatchPatternKind.Percent25;
                case "mediumgray":
                    return OfficeHatchPatternKind.Percent50;
                case "darkgray":
                    return OfficeHatchPatternKind.Percent75;
                case "lighthorizontal":
                case "darkhorizontal":
                    return OfficeHatchPatternKind.Horizontal;
                case "lightvertical":
                case "darkvertical":
                    return OfficeHatchPatternKind.Vertical;
                case "lightdown":
                case "darkdown":
                    return OfficeHatchPatternKind.DiagonalDown;
                case "lightup":
                case "darkup":
                    return OfficeHatchPatternKind.DiagonalUp;
                case "lightgrid":
                case "darkgrid":
                    return OfficeHatchPatternKind.Grid;
                case "lighttrellis":
                case "darktrellis":
                    return OfficeHatchPatternKind.Trellis;
                default:
                    return OfficeHatchPatternKind.Dotted;
            }
        }

        private static PatternMetrics GetPatternMetrics(string pattern, double scale) {
            if (IsGrayPattern(pattern)) {
                return new PatternMetrics(Math.Max(4D, 4D * scale), Math.Max(1D, scale));
            }

            bool dark = pattern.StartsWith("dark", StringComparison.Ordinal) || pattern == "darkgray";
            bool medium = pattern == "mediumgray" || pattern == "gray125";
            double step = Math.Max(3D, (dark ? 5D : medium ? 6D : 8D) * scale);
            double width = Math.Max(1D, (dark ? 1.5D : 1D) * scale);
            return new PatternMetrics(step, width);
        }

        private static bool IsGrayPattern(string pattern) =>
            pattern == "gray0625" ||
            pattern == "gray125" ||
            pattern == "lightgray" ||
            pattern == "mediumgray" ||
            pattern == "darkgray";

        private static void AddGradientDiagnosticIfNeeded(ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (!cell.Style.FillGradientUnsupported) {
                return;
            }

            diagnostics?.Add(ExcelImageExportDiagnosticClassifier.Create(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.FillGradientUnsupported,
                "Excel gradient cell fill is not rendered by the dependency-free image exporter yet.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private static void AddPatternApproximationDiagnostic(ExcelRangeVisualSnapshot snapshot, ExcelVisualCell cell, string pattern, List<OfficeImageExportDiagnostic>? diagnostics) {
            diagnostics?.Add(ExcelImageExportDiagnosticClassifier.Create(
                OfficeImageExportDiagnosticSeverity.Info,
                ExcelImageExportDiagnosticCodes.FillPatternApproximation,
                "Excel pattern cell fill '" + pattern + "' is rendered as a deterministic hatch approximation.",
                GetCellDiagnosticSource(snapshot, cell)));
        }

        private readonly struct PatternMetrics {
            internal PatternMetrics(double step, double width) {
                Step = step;
                Width = width;
            }

            internal double Step { get; }

            internal double Width { get; }
        }
    }
}
