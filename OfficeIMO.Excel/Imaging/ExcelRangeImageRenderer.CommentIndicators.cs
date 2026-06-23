using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static readonly OfficeColor LegacyCommentIndicatorColor = OfficeColor.FromRgb(192, 0, 0);
        private static readonly OfficeColor ThreadedCommentIndicatorColor = OfficeColor.FromRgb(124, 58, 237);
        private static readonly OfficeColor CommentBodyFillColor = OfficeColor.FromRgb(255, 251, 230);
        private static readonly OfficeColor CommentBodyHeaderFillColor = OfficeColor.FromRgb(255, 242, 204);
        private static readonly OfficeColor CommentBodyStrokeColor = OfficeColor.FromRgb(214, 168, 67);
        private static readonly OfficeColor CommentBodyShadowColor = OfficeColor.FromRgba(15, 23, 42, 46);
        private static readonly OfficeColor CommentBodyTitleColor = OfficeColor.FromRgb(92, 64, 14);
        private static readonly OfficeColor CommentBodyTextColor = OfficeColor.FromRgb(31, 41, 55);
        private const double CommentBodyPadding = 7D;
        private const double CommentBodyHeaderHeight = 20D;
        private const double CommentBodyTitleFontSize = 9.5D;
        private const double CommentBodyTextFontSize = 9D;
        private const double CommentBodyLineHeightFactor = 1.18D;

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
            double scale = options.Scale;
            double x = body.X * scale;
            double y = body.Y * scale;
            double width = body.Width * scale;
            double height = body.Height * scale;
            OfficeDrawing drawing = CreateCommentBodyDrawing(body, scale);
            OfficeRasterImage drawingImage = OfficeDrawingRasterRenderer.Render(drawing);
            canvas.FillRectangle(x + (2D * scale), y + (2D * scale), width, height, CommentBodyShadowColor);
            DrawRasterCommentBodyPointer(canvas, body, scale, shadow: true);
            DrawRasterCommentBodyPointer(canvas, body, scale, shadow: false);
            canvas.DrawImage(drawingImage, x, y, width, height);
            DrawRasterCommentBodyText(canvas, body, scale, x, y, width, height);
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

        private static void AppendSvgCommentBody(StringBuilder builder, ExcelVisualCommentBody body, ExcelImageExportOptions options, OfficeRasterCanvas textMeasureCanvas) {
            double scale = options.Scale;
            double x = body.X * scale;
            double y = body.Y * scale;
            double width = body.Width * scale;
            double height = body.Height * scale;
            OfficeDrawing drawing = CreateCommentBodyDrawing(body, scale);
            string inner = OfficeSvgFormatting.ExtractSvgInner(OfficeDrawingSvgExporter.ToSvg(drawing));
            AppendSvgCommentBodyPointer(builder, body, scale, shadow: true);
            AppendSvgCommentBodyPointer(builder, body, scale, shadow: false);
            builder.AppendNestedSvgStart(x, y, width, height);
            builder.Append(inner);
            AppendSvgCommentBodyText(builder, body, scale, width, height, textMeasureCanvas);
            builder.AppendNestedSvgEnd();
        }

        private static double ResolveCommentIndicatorSize(ExcelVisualCommentIndicator indicator, double scale) {
            double cellLimit = Math.Min(indicator.Width, indicator.Height) * scale * 0.42D;
            return Math.Max(4D * scale, Math.Min(8D * scale, cellLimit));
        }

        private static OfficeDrawing CreateCommentBodyDrawing(ExcelVisualCommentBody body, double scale) {
            double width = Math.Max(1D, body.Width * scale);
            double height = Math.Max(1D, body.Height * scale);
            double radius = Math.Min(7D * scale, Math.Min(width, height) / 7D);
            OfficeColor accent = body.Threaded ? ThreadedCommentIndicatorColor : LegacyCommentIndicatorColor;
            var drawing = new OfficeDrawing(width, height);

            OfficeShape background = OfficeShape.RoundedRectangle(width, height, radius);
            background.FillColor = CommentBodyFillColor;
            background.StrokeColor = CommentBodyStrokeColor;
            background.StrokeWidth = Math.Max(1D, scale);
            drawing.AddShape(background, 0D, 0D);

            OfficeShape header = OfficeShape.Rectangle(width, Math.Min(height, CommentBodyHeaderHeight * scale));
            header.FillColor = CommentBodyHeaderFillColor;
            header.StrokeColor = null;
            header.ClipPath = OfficeClipPath.RoundedRectangle(width, height, radius);
            drawing.AddShape(header, 0D, 0D);

            OfficeShape accentLine = OfficeShape.Rectangle(Math.Max(2D, 3D * scale), height);
            accentLine.FillColor = accent;
            accentLine.StrokeColor = null;
            accentLine.ClipPath = OfficeClipPath.RoundedRectangle(width, height, radius);
            drawing.AddShape(accentLine, 0D, 0D);

            return drawing;
        }

        private static void DrawRasterCommentBodyPointer(OfficeRasterCanvas canvas, ExcelVisualCommentBody body, double scale, bool shadow) {
            IReadOnlyList<OfficePoint> points = GetCommentBodyPointerPoints(body, scale, shadow);
            if (points.Count < 3) {
                return;
            }

            if (shadow) {
                canvas.FillPolygon(points, CommentBodyShadowColor);
                return;
            }

            canvas.FillPolygon(points, CommentBodyFillColor);
            canvas.DrawPolygon(points, CommentBodyStrokeColor, Math.Max(1D, scale));
        }

        private static void AppendSvgCommentBodyPointer(StringBuilder builder, ExcelVisualCommentBody body, double scale, bool shadow) {
            IReadOnlyList<OfficePoint> points = GetCommentBodyPointerPoints(body, scale, shadow);
            if (points.Count < 3) {
                return;
            }

            builder.AppendPolygonElement(
                points,
                shadow ? CommentBodyShadowColor : CommentBodyFillColor,
                shadow ? null : CommentBodyStrokeColor,
                shadow ? 0D : Math.Max(1D, scale));
        }

        private static IReadOnlyList<OfficePoint> GetCommentBodyPointerPoints(ExcelVisualCommentBody body, double scale, bool shadow) {
            double x = body.X * scale;
            double y = body.Y * scale;
            double width = body.Width * scale;
            double height = body.Height * scale;
            double anchorX = body.AnchorX * scale;
            double anchorY = body.AnchorY * scale;
            double offset = shadow ? 2D * scale : 0D;
            double inset = Math.Min(height - (8D * scale), Math.Max(10D * scale, CommentBodyHeaderHeight * scale * 0.65D));
            double half = Math.Max(5D * scale, Math.Min(10D * scale, height / 7D));

            if (anchorX <= x) {
                double sideY = Math.Min(y + height - half - (4D * scale), Math.Max(y + half + (4D * scale), y + inset));
                return new[] {
                    new OfficePoint(anchorX + offset, anchorY + offset),
                    new OfficePoint(x + offset, sideY - half + offset),
                    new OfficePoint(x + offset, sideY + half + offset)
                };
            }

            if (anchorX >= x + width) {
                double sideY = Math.Min(y + height - half - (4D * scale), Math.Max(y + half + (4D * scale), y + inset));
                return new[] {
                    new OfficePoint(anchorX + offset, anchorY + offset),
                    new OfficePoint(x + width + offset, sideY - half + offset),
                    new OfficePoint(x + width + offset, sideY + half + offset)
                };
            }

            if (anchorY <= y) {
                double sideX = Math.Min(x + width - half - (4D * scale), Math.Max(x + half + (4D * scale), anchorX));
                return new[] {
                    new OfficePoint(anchorX + offset, anchorY + offset),
                    new OfficePoint(sideX - half + offset, y + offset),
                    new OfficePoint(sideX + half + offset, y + offset)
                };
            }

            double bottomSideX = Math.Min(x + width - half - (4D * scale), Math.Max(x + half + (4D * scale), anchorX));
            return new[] {
                new OfficePoint(anchorX + offset, anchorY + offset),
                new OfficePoint(bottomSideX - half + offset, y + height + offset),
                new OfficePoint(bottomSideX + half + offset, y + height + offset)
            };
        }

        private static void DrawRasterCommentBodyText(OfficeRasterCanvas canvas, ExcelVisualCommentBody body, double scale, double x, double y, double width, double height) {
            double padding = CommentBodyPadding * scale;
            double titleHeight = CommentBodyHeaderHeight * scale;
            double titleFontSize = CommentBodyTitleFontSize * scale;
            double bodyFontSize = CommentBodyTextFontSize * scale;
            double textX = x + padding + (2D * scale);
            double textWidth = Math.Max(1D, width - (padding * 2D) - (2D * scale));

            using (canvas.PushClipRectangle(x, y, width, height)) {
                canvas.DrawTextLine(body.Title, textX, y + Math.Max(2D * scale, (titleHeight - titleFontSize) / 2D), titleFontSize, CommentBodyTitleColor, bold: true, alignment: OfficeTextAlignment.Left);

                double bodyTop = y + titleHeight + (4D * scale);
                double bodyHeight = Math.Max(1D, height - titleHeight - (padding * 1.4D));
                OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                    body.Text,
                    bodyFontSize,
                    textWidth,
                    bodyHeight,
                    CommentBodyLineHeightFactor,
                    Math.Max(6D * scale, scale),
                    canvas.MeasureText,
                    wrap: true,
                    forceSingleLine: false,
                    shrinkToFit: false);
                OfficeTextBlockRenderer.DrawRasterTextBlock(
                    canvas,
                    layout,
                    textX,
                    bodyTop,
                    textWidth,
                    bodyHeight,
                    CommentBodyTextColor,
                    OfficeTextAlignment.Left,
                    OfficeTextVerticalAlignment.Top,
                    centerLineInLineHeight: false);
            }
        }

        private static void AppendSvgCommentBodyText(StringBuilder builder, ExcelVisualCommentBody body, double scale, double width, double height, OfficeRasterCanvas textMeasureCanvas) {
            double padding = CommentBodyPadding * scale;
            double titleHeight = CommentBodyHeaderHeight * scale;
            double titleFontSize = CommentBodyTitleFontSize * scale;
            double bodyFontSize = CommentBodyTextFontSize * scale;
            double textX = padding + (2D * scale);
            double textWidth = Math.Max(1D, width - (padding * 2D) - (2D * scale));
            string clipId = "xl-comment-" + body.Row.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" + body.Column.ToString(System.Globalization.CultureInfo.InvariantCulture);

            builder.AppendRectClipPathDefinition(clipId, 0D, 0D, width, height);
            builder.Append("<g").AppendClipPathReference(clipId).Append(">");
            builder.AppendSvgTextElement(
                body.Title,
                textX,
                Math.Max(titleFontSize + (3D * scale), titleHeight - (5D * scale)),
                titleFontSize,
                CommentBodyTitleColor,
                "Calibri, Arial, sans-serif",
                titleFontSize,
                OfficeTextAlignment.Left,
                bold: true);

            double bodyTop = titleHeight + (4D * scale);
            double bodyHeight = Math.Max(1D, height - titleHeight - (padding * 1.4D));
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                body.Text,
                bodyFontSize,
                textWidth,
                bodyHeight,
                CommentBodyLineHeightFactor,
                Math.Max(6D * scale, scale),
                textMeasureCanvas.MeasureText,
                wrap: true,
                forceSingleLine: false,
                shrinkToFit: false);
            builder.AppendSvgTextBlock(
                layout,
                textX,
                bodyTop,
                textWidth,
                bodyHeight,
                CommentBodyTextColor,
                "Calibri, Arial, sans-serif",
                OfficeTextAlignment.Left,
                OfficeTextVerticalAlignment.Top,
                centerLineInLineHeight: false);

            builder.Append("</g>");
        }
    }
}
