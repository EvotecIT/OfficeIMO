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

        private static void AppendSvgConditionalIcons(
            StringBuilder builder,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            System.Threading.CancellationToken cancellationToken) {
            double scale = options.Scale;
            foreach (ExcelVisualConditionalIcon icon in snapshot.ConditionalIcons) {
                cancellationToken.ThrowIfCancellationRequested();
                AppendSvgConditionalIcon(builder, icon, scale);
            }
        }

        private static void DrawConditionalIcon(OfficeRasterCanvas canvas, ExcelVisualConditionalIcon icon, double scale) {
            IconBounds bounds = GetConditionalIconBounds(icon, scale);
            OfficeConditionalIconRenderer.DrawRaster(canvas, bounds.X, bounds.Y, bounds.Size, MapConditionalIconKind(icon.Kind), scale);
        }

        private static void AppendSvgConditionalIcon(StringBuilder builder, ExcelVisualConditionalIcon icon, double scale) {
            IconBounds bounds = GetConditionalIconBounds(icon, scale);
            OfficeConditionalIconRenderer.AppendSvg(builder, bounds.X, bounds.Y, bounds.Size, MapConditionalIconKind(icon.Kind), scale);
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

        private static OfficeConditionalIconKind MapConditionalIconKind(OfficeConditionalIconKind kind) =>
            kind switch {
                OfficeConditionalIconKind.GreenUpArrow => OfficeConditionalIconKind.GreenUpArrow,
                OfficeConditionalIconKind.YellowUpArrow => OfficeConditionalIconKind.YellowUpArrow,
                OfficeConditionalIconKind.YellowSideArrow => OfficeConditionalIconKind.YellowSideArrow,
                OfficeConditionalIconKind.YellowDownArrow => OfficeConditionalIconKind.YellowDownArrow,
                OfficeConditionalIconKind.RedDownArrow => OfficeConditionalIconKind.RedDownArrow,
                OfficeConditionalIconKind.GreenCheck => OfficeConditionalIconKind.GreenCheck,
                OfficeConditionalIconKind.YellowExclamation => OfficeConditionalIconKind.YellowExclamation,
                OfficeConditionalIconKind.RedCross => OfficeConditionalIconKind.RedCross,
                OfficeConditionalIconKind.GreenCircle => OfficeConditionalIconKind.GreenCircle,
                OfficeConditionalIconKind.LightGreenCircle => OfficeConditionalIconKind.LightGreenCircle,
                OfficeConditionalIconKind.YellowCircle => OfficeConditionalIconKind.YellowCircle,
                OfficeConditionalIconKind.OrangeCircle => OfficeConditionalIconKind.OrangeCircle,
                OfficeConditionalIconKind.RedCircle => OfficeConditionalIconKind.RedCircle,
                OfficeConditionalIconKind.RatingOne => OfficeConditionalIconKind.RatingOne,
                OfficeConditionalIconKind.RatingTwo => OfficeConditionalIconKind.RatingTwo,
                OfficeConditionalIconKind.RatingThree => OfficeConditionalIconKind.RatingThree,
                OfficeConditionalIconKind.RatingFour => OfficeConditionalIconKind.RatingFour,
                OfficeConditionalIconKind.RatingFive => OfficeConditionalIconKind.RatingFive,
                OfficeConditionalIconKind.QuarterEmpty => OfficeConditionalIconKind.QuarterEmpty,
                OfficeConditionalIconKind.QuarterOne => OfficeConditionalIconKind.QuarterOne,
                OfficeConditionalIconKind.QuarterTwo => OfficeConditionalIconKind.QuarterTwo,
                OfficeConditionalIconKind.QuarterThree => OfficeConditionalIconKind.QuarterThree,
                OfficeConditionalIconKind.QuarterFull => OfficeConditionalIconKind.QuarterFull,
                OfficeConditionalIconKind.GreenFlag => OfficeConditionalIconKind.GreenFlag,
                OfficeConditionalIconKind.YellowFlag => OfficeConditionalIconKind.YellowFlag,
                OfficeConditionalIconKind.RedFlag => OfficeConditionalIconKind.RedFlag,
                _ => OfficeConditionalIconKind.RedCross
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
