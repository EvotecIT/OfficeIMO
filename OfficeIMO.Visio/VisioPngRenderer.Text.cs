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
            text = ResolveRasterDisplayText(text, style);
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
            OfficeTextBaseline baseline = style?.Baseline ?? OfficeTextBaseline.Normal;
            double baselineScale = baseline == OfficeTextBaseline.Normal ? 1D : 0.65D;
            OfficeTextBlockRenderPlan plan = OfficeTextBlockRenderPlan.CreateFittedFromCenter(
                text,
                pixelHeight,
                centerX,
                centerY,
                maxWidth,
                maxHeight,
                (value, size) => canvas.MeasureText(value, size * baselineScale, fontFamily, fontStyle),
                alignment,
                verticalAlignment,
                lineHeightFactor: 1.25D * baselineScale,
                minimumFontSize: canvas.Supersampling * 5D);
            pixelHeight = plan.Layout.FontSize;
            double renderedPixelHeight = pixelHeight * baselineScale;

            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            double padX = Math.Max(canvas.Supersampling * 3D, renderedPixelHeight * 0.22D);
            double padY = Math.Max(canvas.Supersampling * 2D, renderedPixelHeight * 0.16D);

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
                padY,
                style?.UnderlineStyle ?? OfficeTextDecorationStyle.None,
                style?.StrikethroughStyle ?? OfficeTextDecorationStyle.None,
                baseline);
        }

        private static string ResolveRasterDisplayText(string text, VisioTextStyle? style) {
            if (style?.SmallCaps == true || style?.Capitalization == VisioTextCapitalization.AllCaps) {
                return OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Uppercase, System.Globalization.CultureInfo.InvariantCulture);
            }

            return style?.Capitalization == VisioTextCapitalization.InitialCaps
                ? OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Capitalize, System.Globalization.CultureInfo.InvariantCulture)
                : text;
        }

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }
    }
}
