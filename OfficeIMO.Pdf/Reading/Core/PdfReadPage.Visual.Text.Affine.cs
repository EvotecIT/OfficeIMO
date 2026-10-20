using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static bool TryGetAffineTextLinearTransform(PdfTextSpan span, out OfficeTransform transform) {
        transform = default;
        if (span.TextToPageTransform is not Matrix2D matrix) return false;
        double xLength = Math.Sqrt(matrix.A * matrix.A + matrix.B * matrix.B);
        double yLength = Math.Sqrt(matrix.C * matrix.C + matrix.D * matrix.D);
        if (!(xLength > 0D) || !(yLength > 0D) || double.IsInfinity(xLength) || double.IsInfinity(yLength)) return false;
        double a = matrix.A / xLength, b = matrix.B / xLength;
        double c = matrix.C / yLength, d = matrix.D / yLength;
        // Orthogonal axes already have an allocation-free rotation/reflection
        // projection. Only nonorthogonal axes need the complete affine group.
        if (Math.Abs(a * c + b * d) <= .000000001D || Math.Abs(a * d - b * c) <= .000000000001D) return false;

        // The parser has already applied the horizontal axis length to glyph
        // advances and the vertical axis length to the displayed font size.
        // Retain the directions, not those lengths a second time. PDF's upward
        // vertical axis becomes a downward axis in the drawing coordinate space.
        double direction = span.CharacterAdvanceDirection < 0D ? -1D : 1D;
        transform = new OfficeTransform(a * direction, -b * direction, -c, d, 0D, 0D);
        return true;
    }

    private static void AddAffineTextSpan(OfficeDrawing drawing, double pageHeight, PdfTextSpan span,
        OfficeTransform linear, (double Left, double Top, double Right, double Bottom)? measuredPaint,
        PageContentBudget budget) {
        var cancellationToken = budget.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        var frame = GetTextFrame(span.FontSize, 0D, 0D, span.FontSize, span.Text.Length, span.Advance);
        // Local text has its baseline at FontSize. Its actual source baseline is
        // already in page space, including text rise and Form/page transforms.
        OfficeTransform localToPage = OfficeTransform.Translate(0D, -span.FontSize).Then(linear)
            .Then(OfficeTransform.Translate(span.X, pageHeight - span.Y));
        (double Left, double Top, double Right, double Bottom) localPaint;
        if (measuredPaint.HasValue) {
            // Spaced glyphs have already paid for and cached their outline
            // measurement. An inverse bounding box conservatively retains that
            // paint without measuring every repeated glyph a second time.
            var paint = measuredPaint.Value;
            localPaint = localToPage.Invert().TransformRectangleBounds(paint.Left, paint.Top,
                paint.Right - paint.Left, paint.Bottom - paint.Top);
        } else {
            budget.ChargePositionedTextWorkCharacters(span.Text.Length);
            localPaint = MeasureTextPaintBounds(OfficeDrawingTextLayout.CreateMetrics(drawing, cancellationToken),
                span.Text, span, frame, span.FontSize, span.Advance, applyFrameTransform: false);
        }
        var pagePaint = measuredPaint ?? localToPage.TransformRectangleBounds(localPaint.Left, localPaint.Top,
            localPaint.Right - localPaint.Left, localPaint.Bottom - localPaint.Top);
        double strokePadding = Math.Max(0D, span.VisibleStrokePadding);
        if (strokePadding > 0D) {
            double xPadding = strokePadding * (Math.Abs(linear.M11) + Math.Abs(linear.M21));
            double yPadding = strokePadding * (Math.Abs(linear.M12) + Math.Abs(linear.M22));
            pagePaint = (pagePaint.Left - xPadding, pagePaint.Top - yPadding,
                pagePaint.Right + xPadding, pagePaint.Bottom + yPadding);
        }
        PdfPageClipPath clip = span.ClipPath ?? PdfPageClipPath.Rectangle(0D, 0D, drawing.Width, drawing.Height);
        OfficeClipPath? officeClip = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (officeClip == null) {
            // Match the ordinary text projection's fallback for an unsupported
            // source clip while still retaining the page's physical boundary.
            clip = PdfPageClipPath.Rectangle(0D, 0D, drawing.Width, drawing.Height);
        }
        if (!TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out clip) ||
            clip.Width <= 0D || clip.Height <= 0D || !HasPaintOverlap(drawing, pagePaint, clip)) return;

        double left = Math.Floor(Math.Min(0D, localPaint.Left) - strokePadding);
        double top = Math.Floor(Math.Min(0D, localPaint.Top) - strokePadding);
        double right = Math.Ceiling(Math.Max(frame.Width, localPaint.Right) + strokePadding);
        double bottom = Math.Ceiling(Math.Max(frame.Height, localPaint.Bottom) + strokePadding);
        budget.ChargeClippedTextFontCopyWork(drawing.Fonts.Faces.Count);
        var child = new OfficeDrawing(right - left, bottom - top) {
            TextShapingProvider = drawing.TextShapingProvider,
            TextShapingLanguage = drawing.TextShapingLanguage
        };
        child.Fonts.AddRange(drawing.Fonts);
        child.AddPositionedTextWithResolvedDirection(span.Text, -left, -top, frame.Width, frame.Height,
            ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
            span.Color ?? OfficeColor.Black, lineHeight: null,
            textAdvanceWidth: TryGetSafePositionedAdvance(span, out double advance) ? advance : null,
            textDirection: OfficeTextDirection.Auto);
        // These helpers annotate the text element itself. An affine wrapper is
        // intentionally added afterward so logical/painted glyphs and stroked
        // text do not disappear behind an extra level of drawing hierarchy.
        MarkPaintedGlyphs(child, span);
        OfficeTransform childToPage = OfficeTransform.Translate(left, top).Then(localToPage);
        cancellationToken.ThrowIfCancellationRequested();
        if (clip.IsRectangle && pagePaint.Left >= clip.X && pagePaint.Top >= clip.Y &&
            pagePaint.Right <= clip.X + clip.Width && pagePaint.Bottom <= clip.Y + clip.Height) {
            drawing.AddEffectDrawing(child, childToPage);
            return;
        }

        officeClip = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (officeClip == null) return;
        budget.ChargeClippedTextFontCopyWork(drawing.Fonts.Faces.Count);
        var clipped = new OfficeDrawing(clip.Width, clip.Height) {
            TextShapingProvider = drawing.TextShapingProvider,
            TextShapingLanguage = drawing.TextShapingLanguage
        };
        clipped.AddEffectDrawing(child, childToPage.Then(OfficeTransform.Translate(-clip.X, -clip.Y)));
        // PDF clipping is expressed in destination coordinates. Put it outside
        // the affine paint, rather than shearing the clip together with the text.
        drawing.AddClippedDrawing(clipped, clip.X, clip.Y, officeClip);
    }
}
