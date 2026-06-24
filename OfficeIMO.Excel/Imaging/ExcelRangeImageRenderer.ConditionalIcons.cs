using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterConditionalIcons(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            double scale = options.Scale;
            foreach (ExcelVisualConditionalIcon icon in snapshot.ConditionalIcons) {
                DrawConditionalIcon(canvas, icon, scale);
            }
        }

        private static void AppendSvgConditionalIcons(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            double scale = options.Scale;
            foreach (ExcelVisualConditionalIcon icon in snapshot.ConditionalIcons) {
                AppendSvgConditionalIcon(builder, icon, scale);
            }
        }

        private static void DrawConditionalIcon(OfficeRasterCanvas canvas, ExcelVisualConditionalIcon icon, double scale) {
            IconBounds bounds = GetConditionalIconBounds(icon, scale);
            OfficeColor fill = GetConditionalIconColor(icon.Kind);
            OfficeColor stroke = GetConditionalIconStroke(icon.Kind);
            switch (icon.Kind) {
                case ExcelConditionalIconKind.GreenCheck:
                    canvas.DrawLine(bounds.X + bounds.Size * 0.22D, bounds.Y + bounds.Size * 0.54D, bounds.X + bounds.Size * 0.42D, bounds.Y + bounds.Size * 0.74D, fill, Math.Max(2D, bounds.Size * 0.14D));
                    canvas.DrawLine(bounds.X + bounds.Size * 0.42D, bounds.Y + bounds.Size * 0.74D, bounds.X + bounds.Size * 0.80D, bounds.Y + bounds.Size * 0.28D, fill, Math.Max(2D, bounds.Size * 0.14D));
                    break;
                case ExcelConditionalIconKind.YellowExclamation:
                    canvas.FillEllipse(bounds.X, bounds.Y, bounds.Size, bounds.Size, fill);
                    canvas.DrawEllipse(bounds.X, bounds.Y, bounds.Size, bounds.Size, stroke, Math.Max(1D, scale));
                    canvas.DrawLine(bounds.X + bounds.Size * 0.5D, bounds.Y + bounds.Size * 0.25D, bounds.X + bounds.Size * 0.5D, bounds.Y + bounds.Size * 0.60D, OfficeColor.White, Math.Max(2D, bounds.Size * 0.12D));
                    canvas.FillEllipse(bounds.X + bounds.Size * 0.44D, bounds.Y + bounds.Size * 0.72D, bounds.Size * 0.12D, bounds.Size * 0.12D, OfficeColor.White);
                    break;
                case ExcelConditionalIconKind.RedCross:
                    canvas.DrawLine(bounds.X + bounds.Size * 0.25D, bounds.Y + bounds.Size * 0.25D, bounds.X + bounds.Size * 0.75D, bounds.Y + bounds.Size * 0.75D, fill, Math.Max(2D, bounds.Size * 0.14D));
                    canvas.DrawLine(bounds.X + bounds.Size * 0.75D, bounds.Y + bounds.Size * 0.25D, bounds.X + bounds.Size * 0.25D, bounds.Y + bounds.Size * 0.75D, fill, Math.Max(2D, bounds.Size * 0.14D));
                    break;
                case ExcelConditionalIconKind.GreenCircle:
                case ExcelConditionalIconKind.YellowCircle:
                case ExcelConditionalIconKind.RedCircle:
                    canvas.FillEllipse(bounds.X, bounds.Y, bounds.Size, bounds.Size, fill);
                    canvas.DrawEllipse(bounds.X, bounds.Y, bounds.Size, bounds.Size, stroke, Math.Max(1D, scale));
                    break;
                case ExcelConditionalIconKind.GreenUpArrow:
                case ExcelConditionalIconKind.YellowSideArrow:
                case ExcelConditionalIconKind.RedDownArrow:
                    IReadOnlyList<OfficePoint> points = CreateArrowPoints(bounds, icon.Kind);
                    canvas.FillPolygon(points, fill);
                    canvas.DrawPolygon(points, stroke, Math.Max(1D, scale));
                    break;
            }
        }

        private static void AppendSvgConditionalIcon(StringBuilder builder, ExcelVisualConditionalIcon icon, double scale) {
            IconBounds bounds = GetConditionalIconBounds(icon, scale);
            OfficeColor fill = GetConditionalIconColor(icon.Kind);
            OfficeColor stroke = GetConditionalIconStroke(icon.Kind);
            switch (icon.Kind) {
                case ExcelConditionalIconKind.GreenCheck:
                    AppendSvgIconLine(builder, bounds.X + bounds.Size * 0.22D, bounds.Y + bounds.Size * 0.54D, bounds.X + bounds.Size * 0.42D, bounds.Y + bounds.Size * 0.74D, fill, bounds.Size * 0.14D);
                    AppendSvgIconLine(builder, bounds.X + bounds.Size * 0.42D, bounds.Y + bounds.Size * 0.74D, bounds.X + bounds.Size * 0.80D, bounds.Y + bounds.Size * 0.28D, fill, bounds.Size * 0.14D);
                    break;
                case ExcelConditionalIconKind.YellowExclamation:
                    AppendSvgIconCircle(builder, bounds, fill, stroke);
                    AppendSvgIconLine(builder, bounds.X + bounds.Size * 0.5D, bounds.Y + bounds.Size * 0.25D, bounds.X + bounds.Size * 0.5D, bounds.Y + bounds.Size * 0.60D, OfficeColor.White, bounds.Size * 0.12D);
                    builder.AppendCircleElement(bounds.X + bounds.Size * 0.5D, bounds.Y + bounds.Size * 0.78D, bounds.Size * 0.06D, OfficeColor.White);
                    break;
                case ExcelConditionalIconKind.RedCross:
                    AppendSvgIconLine(builder, bounds.X + bounds.Size * 0.25D, bounds.Y + bounds.Size * 0.25D, bounds.X + bounds.Size * 0.75D, bounds.Y + bounds.Size * 0.75D, fill, bounds.Size * 0.14D);
                    AppendSvgIconLine(builder, bounds.X + bounds.Size * 0.75D, bounds.Y + bounds.Size * 0.25D, bounds.X + bounds.Size * 0.25D, bounds.Y + bounds.Size * 0.75D, fill, bounds.Size * 0.14D);
                    break;
                case ExcelConditionalIconKind.GreenCircle:
                case ExcelConditionalIconKind.YellowCircle:
                case ExcelConditionalIconKind.RedCircle:
                    AppendSvgIconCircle(builder, bounds, fill, stroke);
                    break;
                case ExcelConditionalIconKind.GreenUpArrow:
                case ExcelConditionalIconKind.YellowSideArrow:
                case ExcelConditionalIconKind.RedDownArrow:
                    IReadOnlyList<OfficePoint> points = CreateArrowPoints(bounds, icon.Kind);
                    var attributes = new StringBuilder()
                        .AppendPaintAttribute("fill", fill)
                        .AppendPaintAttribute("stroke", stroke)
                        .AppendNumberAttribute("stroke-width", Math.Max(1D, scale))
                        .ToString();
                    builder.AppendPathElement(OfficeSvgFormatting.FormatMoveLinePathData(points, closePath: true), attributes);
                    break;
            }
        }

        private static void AppendSvgIconCircle(StringBuilder builder, IconBounds bounds, OfficeColor fill, OfficeColor stroke) {
            var attributes = new StringBuilder()
                .AppendPaintAttribute("fill", fill)
                .AppendPaintAttribute("stroke", stroke)
                .AppendNumberAttribute("stroke-width", Math.Max(1D, bounds.Size / 14D))
                .ToString();
            builder.AppendCircleElement(bounds.X + bounds.Size / 2D, bounds.Y + bounds.Size / 2D, bounds.Size / 2D, attributes);
        }

        private static void AppendSvgIconLine(StringBuilder builder, double x1, double y1, double x2, double y2, OfficeColor color, double width) {
            builder.AppendLineElement(x1, y1, x2, y2, color, Math.Max(1D, width), OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap.Round);
        }

        private static IconBounds GetConditionalIconBounds(ExcelVisualConditionalIcon icon, double scale) {
            double cellX = icon.X * scale;
            double cellY = icon.Y * scale;
            double cellWidth = icon.Width * scale;
            double cellHeight = icon.Height * scale;
            double size = Math.Max(8D * scale, Math.Min(cellHeight * 0.62D, Math.Min(cellWidth * 0.38D, 16D * scale)));
            double x = cellX + Math.Max(3D * scale, cellWidth * 0.08D);
            double y = cellY + (cellHeight - size) / 2D;
            return new IconBounds(x, y, size);
        }

        private static IReadOnlyList<OfficePoint> CreateArrowPoints(IconBounds bounds, ExcelConditionalIconKind kind) {
            double x = bounds.X;
            double y = bounds.Y;
            double s = bounds.Size;
            if (kind == ExcelConditionalIconKind.RedDownArrow) {
                return new[] {
                    new OfficePoint(x + s * 0.36D, y + s * 0.08D),
                    new OfficePoint(x + s * 0.64D, y + s * 0.08D),
                    new OfficePoint(x + s * 0.64D, y + s * 0.54D),
                    new OfficePoint(x + s * 0.86D, y + s * 0.54D),
                    new OfficePoint(x + s * 0.50D, y + s * 0.92D),
                    new OfficePoint(x + s * 0.14D, y + s * 0.54D),
                    new OfficePoint(x + s * 0.36D, y + s * 0.54D)
                };
            }

            if (kind == ExcelConditionalIconKind.YellowSideArrow) {
                return new[] {
                    new OfficePoint(x + s * 0.10D, y + s * 0.36D),
                    new OfficePoint(x + s * 0.56D, y + s * 0.36D),
                    new OfficePoint(x + s * 0.56D, y + s * 0.14D),
                    new OfficePoint(x + s * 0.92D, y + s * 0.50D),
                    new OfficePoint(x + s * 0.56D, y + s * 0.86D),
                    new OfficePoint(x + s * 0.56D, y + s * 0.64D),
                    new OfficePoint(x + s * 0.10D, y + s * 0.64D)
                };
            }

            return new[] {
                new OfficePoint(x + s * 0.36D, y + s * 0.92D),
                new OfficePoint(x + s * 0.64D, y + s * 0.92D),
                new OfficePoint(x + s * 0.64D, y + s * 0.46D),
                new OfficePoint(x + s * 0.86D, y + s * 0.46D),
                new OfficePoint(x + s * 0.50D, y + s * 0.08D),
                new OfficePoint(x + s * 0.14D, y + s * 0.46D),
                new OfficePoint(x + s * 0.36D, y + s * 0.46D)
            };
        }

        private static OfficeColor GetConditionalIconColor(ExcelConditionalIconKind kind) =>
            kind switch {
                ExcelConditionalIconKind.GreenUpArrow or ExcelConditionalIconKind.GreenCheck or ExcelConditionalIconKind.GreenCircle => OfficeColor.FromRgb(22, 163, 74),
                ExcelConditionalIconKind.YellowSideArrow or ExcelConditionalIconKind.YellowExclamation or ExcelConditionalIconKind.YellowCircle => OfficeColor.FromRgb(245, 158, 11),
                _ => OfficeColor.FromRgb(220, 38, 38)
            };

        private static OfficeColor GetConditionalIconStroke(ExcelConditionalIconKind kind) =>
            kind switch {
                ExcelConditionalIconKind.GreenUpArrow or ExcelConditionalIconKind.GreenCheck or ExcelConditionalIconKind.GreenCircle => OfficeColor.FromRgb(21, 128, 61),
                ExcelConditionalIconKind.YellowSideArrow or ExcelConditionalIconKind.YellowExclamation or ExcelConditionalIconKind.YellowCircle => OfficeColor.FromRgb(180, 83, 9),
                _ => OfficeColor.FromRgb(185, 28, 28)
            };

        private readonly struct IconBounds {
            internal IconBounds(double x, double y, double size) {
                X = x;
                Y = y;
                Size = size;
            }

            internal double X { get; }

            internal double Y { get; }

            internal double Size { get; }
        }
    }
}
