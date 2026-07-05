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

        private static OfficeConditionalIconKind MapConditionalIconKind(ExcelConditionalIconKind kind) =>
            kind switch {
                ExcelConditionalIconKind.GreenUpArrow => OfficeConditionalIconKind.GreenUpArrow,
                ExcelConditionalIconKind.YellowUpArrow => OfficeConditionalIconKind.YellowUpArrow,
                ExcelConditionalIconKind.YellowSideArrow => OfficeConditionalIconKind.YellowSideArrow,
                ExcelConditionalIconKind.YellowDownArrow => OfficeConditionalIconKind.YellowDownArrow,
                ExcelConditionalIconKind.RedDownArrow => OfficeConditionalIconKind.RedDownArrow,
                ExcelConditionalIconKind.GreenCheck => OfficeConditionalIconKind.GreenCheck,
                ExcelConditionalIconKind.YellowExclamation => OfficeConditionalIconKind.YellowExclamation,
                ExcelConditionalIconKind.RedCross => OfficeConditionalIconKind.RedCross,
                ExcelConditionalIconKind.GreenCircle => OfficeConditionalIconKind.GreenCircle,
                ExcelConditionalIconKind.LightGreenCircle => OfficeConditionalIconKind.LightGreenCircle,
                ExcelConditionalIconKind.YellowCircle => OfficeConditionalIconKind.YellowCircle,
                ExcelConditionalIconKind.OrangeCircle => OfficeConditionalIconKind.OrangeCircle,
                ExcelConditionalIconKind.RedCircle => OfficeConditionalIconKind.RedCircle,
                ExcelConditionalIconKind.RatingOne => OfficeConditionalIconKind.RatingOne,
                ExcelConditionalIconKind.RatingTwo => OfficeConditionalIconKind.RatingTwo,
                ExcelConditionalIconKind.RatingThree => OfficeConditionalIconKind.RatingThree,
                ExcelConditionalIconKind.RatingFour => OfficeConditionalIconKind.RatingFour,
                ExcelConditionalIconKind.RatingFive => OfficeConditionalIconKind.RatingFive,
                ExcelConditionalIconKind.QuarterEmpty => OfficeConditionalIconKind.QuarterEmpty,
                ExcelConditionalIconKind.QuarterOne => OfficeConditionalIconKind.QuarterOne,
                ExcelConditionalIconKind.QuarterTwo => OfficeConditionalIconKind.QuarterTwo,
                ExcelConditionalIconKind.QuarterThree => OfficeConditionalIconKind.QuarterThree,
                ExcelConditionalIconKind.QuarterFull => OfficeConditionalIconKind.QuarterFull,
                ExcelConditionalIconKind.GreenFlag => OfficeConditionalIconKind.GreenFlag,
                ExcelConditionalIconKind.YellowFlag => OfficeConditionalIconKind.YellowFlag,
                ExcelConditionalIconKind.RedFlag => OfficeConditionalIconKind.RedFlag,
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
