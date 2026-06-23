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
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.FitWrappedText(
                text,
                pixelHeight,
                maxWidth,
                maxHeight,
                lineHeightFactor: 1.25D,
                minimumFontSize: canvas.Supersampling * 5D,
                canvas.MeasureText);
            pixelHeight = layout.FontSize;

            OfficeTextAlignment alignment = VisioDrawingTextAlignment.ToOfficeTextAlignment(style?.HorizontalAlignment);
            OfficeTextVerticalAlignment verticalAlignment = VisioDrawingTextAlignment.ToOfficeTextVerticalAlignment(style?.VerticalAlignment);
            double top = OfficeTextPlacement.ResolveTopFromCenter(centerY, maxHeight, layout.Height, verticalAlignment);
            double anchorX = OfficeTextPlacement.ResolveAnchorXFromCenter(centerX, maxWidth, alignment);
            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            if (backgroundColor.HasValue) {
                double padX = Math.Max(canvas.Supersampling * 3D, pixelHeight * 0.22D);
                double padY = Math.Max(canvas.Supersampling * 2D, pixelHeight * 0.16D);
                double backgroundLeft = OfficeTextPlacement.ResolveLeftFromAnchor(anchorX, layout.Width, alignment) - padX;
                double backgroundTop = top - padY;
                double backgroundWidth = layout.Width + (padX * 2D);
                double backgroundHeight = layout.Height + (padY * 2D);
                if (Math.Abs(rotateRadians) < TextRotationEpsilon) {
                    canvas.FillRectangle(backgroundLeft, backgroundTop, backgroundWidth, backgroundHeight, backgroundColor.Value);
                } else {
                    canvas.FillPolygon(new[] {
                        OfficeGeometry.RotatePoint((backgroundLeft, backgroundTop), centerX, centerY, -rotateRadians),
                        OfficeGeometry.RotatePoint((backgroundLeft + backgroundWidth, backgroundTop), centerX, centerY, -rotateRadians),
                        OfficeGeometry.RotatePoint((backgroundLeft + backgroundWidth, backgroundTop + backgroundHeight), centerX, centerY, -rotateRadians),
                        OfficeGeometry.RotatePoint((backgroundLeft, backgroundTop + backgroundHeight), centerX, centerY, -rotateRadians)
                    }, backgroundColor.Value);
                }
            }

            canvas.DrawTextBlock(
                layout,
                centerX - (maxWidth / 2D),
                centerY - (maxHeight / 2D),
                maxWidth,
                maxHeight,
                color,
                style?.Bold == true,
                style?.Italic == true,
                style?.Underline == true,
                alignment,
                verticalAlignment,
                rotateRadians,
                centerX,
                centerY);
        }

        private const double TextRotationEpsilon = 1e-9;

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }
    }
}
