using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterSparklines(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            double scale = options.Scale;
            foreach (ExcelVisualSparkline sparkline in snapshot.Sparklines) {
                using (canvas.PushClipRectangle(sparkline.X * scale, sparkline.Y * scale, sparkline.Width * scale, sparkline.Height * scale)) {
                    RenderRasterSparkline(canvas, sparkline, scale);
                }
            }
        }

        private static void RenderRasterSparkline(OfficeRasterCanvas canvas, ExcelVisualSparkline sparkline, double scale) {
            OfficeSparklineRenderer.DrawRaster(
                canvas,
                sparkline.X * scale,
                sparkline.Y * scale,
                sparkline.Width * scale,
                sparkline.Height * scale,
                sparkline.Values,
                ResolveSparklineKind(sparkline.Kind),
                CreateSparklineStyle(sparkline, scale));
        }

        private static void AppendSvgSparklines(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            double scale = options.Scale;
            int index = 0;
            foreach (ExcelVisualSparkline sparkline in snapshot.Sparklines) {
                string clipId = "officeimo-sparkline-clip-" + index.ToString(CultureInfo.InvariantCulture);
                builder.AppendRectClipPathDefinition(clipId, sparkline.X * scale, sparkline.Y * scale, sparkline.Width * scale, sparkline.Height * scale, wrapInDefs: true);
                builder.Append("<g").AppendClipPathReference(clipId).Append(">");
                AppendSvgSparkline(builder, sparkline, scale);
                builder.Append("</g>");
                index++;
            }
        }

        private static void AppendSvgSparkline(StringBuilder builder, ExcelVisualSparkline sparkline, double scale) {
            OfficeSparklineRenderer.AppendSvg(
                builder,
                sparkline.X * scale,
                sparkline.Y * scale,
                sparkline.Width * scale,
                sparkline.Height * scale,
                sparkline.Values,
                ResolveSparklineKind(sparkline.Kind),
                CreateSparklineStyle(sparkline, scale));
        }

        private static bool ShouldDrawSparklineMarker(ExcelVisualSparkline sparkline, int index, double min, double max) {
            double value = sparkline.Values[index];
            return sparkline.DisplayMarkers ||
                (sparkline.DisplayFirst && index == 0) ||
                (sparkline.DisplayLast && index == sparkline.Values.Count - 1) ||
                (sparkline.DisplayHigh && Math.Abs(value - max) < 0.000001D) ||
                (sparkline.DisplayLow && Math.Abs(value - min) < 0.000001D) ||
                (sparkline.DisplayNegative && value < 0D);
        }

        private static OfficeSparklineStyle CreateSparklineStyle(ExcelVisualSparkline sparkline, double scale) {
            double min = sparkline.Values.Count == 0 ? 0D : sparkline.Values.Min();
            double max = sparkline.Values.Count == 0 ? 0D : sparkline.Values.Max();
            var pointStyles = new List<OfficeSparklinePointStyle>(sparkline.Values.Count);
            for (int i = 0; i < sparkline.Values.Count; i++) {
                pointStyles.Add(new OfficeSparklinePointStyle(
                    ResolveSparklinePointColor(sparkline, i, min, max),
                    ShouldDrawSparklineMarker(sparkline, i, min, max)));
            }

            return new OfficeSparklineStyle {
                SeriesColor = ResolveArgb(sparkline.SeriesColorArgb) ?? OfficeColor.FromRgb(37, 99, 235),
                AxisColor = ResolveArgb(sparkline.AxisColorArgb) ?? OfficeColor.FromRgb(128, 128, 128),
                DisplayAxis = sparkline.DisplayAxis,
                MinimumValue = sparkline.ScaleMinimum,
                MaximumValue = sparkline.ScaleMaximum,
                Padding = 3D * scale,
                LineStrokeWidth = Math.Max(1D, scale * 1.35D),
                AxisStrokeWidth = Math.Max(1D, scale),
                AxisInset = 2D * scale,
                MarkerDiameter = Math.Max(3D, scale * 3D),
                PointStyles = pointStyles
            };
        }

        private static OfficeColor ResolveSparklinePointColor(ExcelVisualSparkline sparkline, int index, double min, double max) {
            double value = sparkline.Values[index];
            string? color = null;
            if (sparkline.DisplayFirst && index == 0) {
                color = sparkline.FirstColorArgb;
            } else if (sparkline.DisplayLast && index == sparkline.Values.Count - 1) {
                color = sparkline.LastColorArgb;
            } else if (sparkline.DisplayHigh && Math.Abs(value - max) < 0.000001D) {
                color = sparkline.HighColorArgb;
            } else if (sparkline.DisplayLow && Math.Abs(value - min) < 0.000001D) {
                color = sparkline.LowColorArgb;
            } else if (sparkline.DisplayNegative && value < 0D) {
                color = sparkline.NegativeColorArgb;
            } else if (sparkline.DisplayMarkers) {
                color = sparkline.MarkersColorArgb;
            }

            return ResolveArgb(color) ?? ResolveArgb(sparkline.SeriesColorArgb) ?? OfficeColor.FromRgb(37, 99, 235);
        }

        private static OfficeSparklineKind ResolveSparklineKind(string kind) {
            string normalized = NormalizeSparklineKind(kind);
            if (normalized == "column") {
                return OfficeSparklineKind.Column;
            }

            return normalized == "stacked" || normalized == "winloss"
                ? OfficeSparklineKind.WinLoss
                : OfficeSparklineKind.Line;
        }

        private static string NormalizeSparklineKind(string kind) =>
            (kind ?? string.Empty).Replace("-", string.Empty).Trim().ToLowerInvariant();
    }
}
