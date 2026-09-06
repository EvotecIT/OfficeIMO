using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static void AddTextSpan(OfficeDrawing drawing, double pageHeight, PdfTextSpan span) {
        if (string.IsNullOrEmpty(span.Text) || !span.IsVisible) {
            return;
        }

        double height = Math.Max(1D, span.FontSize * 1.25D);
        double width = Math.Max(span.Advance, span.Text.Length * span.FontSize * 0.55D);
        double rawX = span.X;
        double rawY = pageHeight - span.Y - span.FontSize;
        if (!HasVisibleOverlap(rawX, rawY, width, height, drawing.Width, drawing.Height)) {
            return;
        }

        double x = rawX;
        double y = rawY;
        double clippedRight = Math.Min(rawX + width, drawing.Width);
        double clippedBottom = Math.Min(rawY + height, drawing.Height);
        double baselineY = pageHeight - span.Y;
        if (!span.ClipPath.HasValue &&
            (rawX < 0D || rawY < 0D || rawX + width > drawing.Width || rawY + height > drawing.Height)) {
            PdfPageClipPath pageClip = PdfPageClipPath.Rectangle(0D, 0D, drawing.Width, drawing.Height);
            if (TryAddClippedTextSpan(drawing, span, x, y, width, height, baselineY, pageClip)) {
                return;
            }
        }

        x = Clamp(rawX, 0D, drawing.Width);
        y = Clamp(rawY, 0D, drawing.Height);
        baselineY = Clamp(baselineY, 0D, drawing.Height);
        width = Math.Max(1D, clippedRight - x);
        height = Math.Max(1D, clippedBottom - y);
        if (TryAddClippedTextSpan(drawing, span, x, y, width, height, baselineY)) {
            return;
        }

        if (TryGetSafePositionedAdvance(span, out double textAdvance)) {
            drawing.AddPositionedText(
                span.Text,
                x,
                y,
                width,
                height,
                new OfficeImageFrameTransform(-span.RotationDegrees, x, baselineY),
                ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
                span.Color ?? OfficeColor.Black,
                textAdvanceWidth: textAdvance);
        } else {
            drawing.AddText(
                span.Text,
                x,
                y,
                width,
                height,
                ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
                span.Color ?? OfficeColor.Black,
                rotationDegrees: -span.RotationDegrees,
                rotationCenterX: x,
                rotationCenterY: baselineY,
                wrapText: false);
        }
    }

    private static bool TryAddClippedTextSpan(OfficeDrawing drawing, PdfTextSpan span, double x, double y, double width, double height, double baselineY, PdfPageClipPath? overrideClipPath = null) {
        PdfPageClipPath? activeClipPath = overrideClipPath ?? span.ClipPath;
        if (!activeClipPath.HasValue) {
            return false;
        }

        PdfPageClipPath clip = activeClipPath.Value;
        if (clip.Width <= 0D || clip.Height <= 0D) {
            return true;
        }

        OfficeClipPath? officeClipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (officeClipPath == null) {
            return false;
        }

        double clipRight = clip.X + clip.Width;
        double clipBottom = clip.Y + clip.Height;
        if (clip.IsRectangle && x >= clip.X && y >= clip.Y && x + width <= clipRight && y + height <= clipBottom) {
            return false;
        }

        if (x + width <= clip.X || y + height <= clip.Y || x >= clipRight || y >= clipBottom) {
            return true;
        }

        double localX = x - clip.X;
        double localY = y - clip.Y;
        if (clip.X < 0D ||
            clip.Y < 0D ||
            clipRight > drawing.Width ||
            clipBottom > drawing.Height) {
            if (!TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out PdfPageClipPath drawingClip)) {
                return true;
            }

            clip = drawingClip;
            officeClipPath = clip.ToOfficeClipPath(clip.X, clip.Y);
            if (officeClipPath == null) {
                return false;
            }

            localX = x - clip.X;
            localY = y - clip.Y;
        }

        double textWidth = Math.Max(1D, width);
        double textHeight = Math.Max(1D, height);
        if (TryGetSafePositionedAdvance(span, out double textAdvance)) {
            drawing.AddClippedPositionedText(
                span.Text,
                x,
                y,
                textWidth,
                textHeight,
                clip.X,
                clip.Y,
                officeClipPath,
                new OfficeImageFrameTransform(-span.RotationDegrees, x, baselineY),
                ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
                span.Color ?? OfficeColor.Black,
                textAdvanceWidth: textAdvance);
        } else {
            drawing.AddClippedText(
                span.Text,
                x,
                y,
                textWidth,
                textHeight,
                clip.X,
                clip.Y,
                officeClipPath,
                ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
                span.Color ?? OfficeColor.Black,
                rotationDegrees: -span.RotationDegrees,
                rotationCenterX: x,
                rotationCenterY: baselineY,
                wrapText: false);
        }
        return true;
    }

    private static bool TryGetSafePositionedAdvance(PdfTextSpan span, out double advance) {
        advance = span.Advance;
        return span.CanScaleAggregateAdvance && advance > 0D && !double.IsNaN(advance) && !double.IsInfinity(advance);
    }

}
