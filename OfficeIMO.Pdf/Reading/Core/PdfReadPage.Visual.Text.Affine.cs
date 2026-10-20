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
        var frame = GetTextFrame(pageHeight, span.X, span.Y, span.FontSize, span.Text.Length, span.Advance);
        double baseline = pageHeight - span.Y;
        OfficeTransform glyphTransform = OfficeTransform.Translate(-span.X, -baseline).Then(linear)
            .Then(OfficeTransform.Translate(span.X, baseline));
        // Spaced glyphs already paid for their cached outline measurement.
        var pagePaint = measuredPaint ?? GetTextPaintBounds(drawing, span, frame, baseline, budget);
        double strokePadding = Math.Max(0D, span.VisibleStrokePadding);
        if (strokePadding > 0D) {
            // This envelope is already in CTM/page space. Applying the text
            // matrix again can shrink a visible pen at the destination clip.
            pagePaint = (pagePaint.Left - strokePadding, pagePaint.Top - strokePadding,
                pagePaint.Right + strokePadding, pagePaint.Bottom + strokePadding);
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

        cancellationToken.ThrowIfCancellationRequested();
        if (clip.IsRectangle && pagePaint.Left >= clip.X && pagePaint.Top >= clip.Y &&
            pagePaint.Right <= clip.X + clip.Width && pagePaint.Bottom <= clip.Y + clip.Height) {
            AddAffinePositionedText(drawing, span, frame.X, frame.Y, frame.Width, frame.Height, glyphTransform);
            return;
        }

        officeClip = clip.ToOfficeClipPath(clip.X, clip.Y);
        if (officeClip == null) return;
        budget.ChargeClippedTextFontCopyWork(drawing.Fonts.Faces.Count);
        var clipped = new OfficeDrawing(clip.Width, clip.Height) {
            TextShapingProvider = drawing.TextShapingProvider,
            TextShapingLanguage = drawing.TextShapingLanguage
        };
        OfficeTransform localGlyph = OfficeTransform.Translate(clip.X, clip.Y).Then(glyphTransform)
            .Then(OfficeTransform.Translate(-clip.X, -clip.Y));
        AddAffinePositionedText(clipped, span, frame.X - clip.X, frame.Y - clip.Y, frame.Width, frame.Height, localGlyph);
        // The glyph transform precedes stroking in Core. The independent PDF
        // destination clip still encloses both the filled glyphs and their pen.
        drawing.AddClippedDrawing(clipped, clip.X, clip.Y, officeClip);
    }

    private static void AddAffinePositionedText(OfficeDrawing drawing, PdfTextSpan span,
        double x, double y, double width, double height, OfficeTransform transform) {
        drawing.AddPositionedTextWithResolvedDirection(span.Text, x, y, width, height,
            ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
            span.Color ?? OfficeColor.Black, lineHeight: null,
            textAdvanceWidth: TryGetSafePositionedAdvance(span, out double advance) ? advance : null,
            textDirection: OfficeTextDirection.Auto);
        MarkPaintedGlyphs(drawing, span);
        drawing.SetLastTextGlyphTransform(transform);
    }
}
