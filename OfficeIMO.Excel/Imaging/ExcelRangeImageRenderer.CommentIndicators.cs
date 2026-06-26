using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static readonly OfficeColor LegacyCommentIndicatorColor = OfficeColor.FromRgb(192, 0, 0);
        private static readonly OfficeColor ThreadedCommentIndicatorColor = OfficeColor.FromRgb(124, 58, 237);

        private static void RenderRasterCommentIndicators(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            double scale = options.Scale;
            foreach (ExcelVisualCommentIndicator indicator in snapshot.CommentIndicators) {
                double size = ResolveCommentIndicatorSize(indicator, scale);
                double right = (indicator.X + indicator.Width) * scale;
                double top = indicator.Y * scale;
                OfficeColor color = indicator.Threaded ? ThreadedCommentIndicatorColor : LegacyCommentIndicatorColor;
                canvas.FillPolygon(
                    new[] {
                        new OfficePoint(right - size, top),
                        new OfficePoint(right, top),
                        new OfficePoint(right, top + size)
                    },
                    color);
            }
        }

        private static void RenderRasterCommentBody(OfficeRasterCanvas canvas, ExcelVisualCommentBody body, ExcelImageExportOptions options) {
            OfficeCalloutRenderer.DrawRaster(
                canvas,
                CreateCommentBodyCallout(body),
                CreateCommentBodyStyle(body.Threaded),
                options.Scale);
        }

        private static void AppendSvgCommentIndicators(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            double scale = options.Scale;
            foreach (ExcelVisualCommentIndicator indicator in snapshot.CommentIndicators) {
                double size = ResolveCommentIndicatorSize(indicator, scale);
                double right = (indicator.X + indicator.Width) * scale;
                double top = indicator.Y * scale;
                OfficeColor color = indicator.Threaded ? ThreadedCommentIndicatorColor : LegacyCommentIndicatorColor;
                var points = new[] {
                    new OfficePoint(right - size, top),
                    new OfficePoint(right, top),
                    new OfficePoint(right, top + size)
                };
                builder.AppendPolygonElement(points, color);
            }
        }

        private static void AppendSvgCommentBody(StringBuilder builder, ExcelVisualCommentBody body, ExcelImageExportOptions options, OfficeTextMeasurer textMeasurer) {
            OfficeCalloutRenderer.AppendSvg(
                builder,
                CreateCommentBodyCallout(body),
                CreateCommentBodyStyle(body.Threaded),
                (text, size, family) => MeasureSvgText(textMeasurer, text, size, family),
                options.Scale,
                "xl-comment-" + body.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + body.Column.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        private static double ResolveCommentIndicatorSize(ExcelVisualCommentIndicator indicator, double scale) {
            double cellLimit = Math.Min(indicator.Width, indicator.Height) * scale * 0.42D;
            return Math.Max(4D * scale, Math.Min(8D * scale, cellLimit));
        }

        private static OfficeCallout CreateCommentBodyCallout(ExcelVisualCommentBody body) =>
            new OfficeCallout(
                body.X,
                body.Y,
                body.Width,
                body.Height,
                body.AnchorX,
                body.AnchorY,
                body.Title,
                body.Text);

        private static OfficeCalloutStyle CreateCommentBodyStyle(bool threaded) =>
            new OfficeCalloutStyle {
                AccentColor = threaded ? ThreadedCommentIndicatorColor : LegacyCommentIndicatorColor
            };
    }
}
