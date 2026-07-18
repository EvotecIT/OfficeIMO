using System;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioPngRenderer {

        private static void DrawText(
            RasterCanvas canvas,
            string text,
            double centerX,
            double centerY,
            VisioTextStyle? style,
            double defaultSize,
            double maxWidth,
            double maxHeight,
            double rotateRadians,
            bool drawLabelBackground) {
            double pointSize = style?.Size ?? defaultSize;
            double pixelHeight = Math.Max(canvas.Supersampling * 7D, pointSize * canvas.Scale / 72D);
            Color color = style?.Color ?? Color.FromRgb(17, 24, 39);
            string fontFamily = string.IsNullOrWhiteSpace(style?.FontFamily)
                ? "Aptos, Calibri, Arial, sans-serif"
                : style!.FontFamily!;
            OfficeFontStyle fontStyle =
                (style?.Bold == true ? OfficeFontStyle.Bold : OfficeFontStyle.Regular) |
                (style?.Italic == true ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
            OfficeTextAlignment alignment = VisioDrawingTextAlignment.ToOfficeTextAlignment(style?.HorizontalAlignment);
            OfficeTextVerticalAlignment verticalAlignment = VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment);
            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFittedFromCenter(
                text,
                pixelHeight,
                centerX,
                centerY,
                maxWidth,
                maxHeight,
                (value, size) => canvas.MeasureText(value, size, fontFamily, fontStyle),
                alignment,
                verticalAlignment,
                lineHeightFactor: 1.25D,
                minimumFontSize: canvas.Supersampling * 5D);
            pixelHeight = plan.Layout.FontSize;

            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            double padX = Math.Max(canvas.Supersampling * 3D, pixelHeight * 0.22D);
            double padY = Math.Max(canvas.Supersampling * 2D, pixelHeight * 0.16D);

            canvas.DrawTextBox(
                plan,
                color,
                style?.Bold == true,
                style?.Italic == true,
                style?.Underline == true,
                fontFamily,
                rotateRadians,
                centerX,
                centerY,
                backgroundColor,
                padX,
                padY);
        }

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }
    }
}
