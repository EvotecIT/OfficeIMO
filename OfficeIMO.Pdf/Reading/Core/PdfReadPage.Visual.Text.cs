using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static void AddTextSpan(OfficeDrawing drawing, double pageHeight, PdfTextSpan span, PageContentBudget pageContentBudget,
        System.Threading.CancellationToken cancellationToken = default) =>
        AddTextSpanCore(drawing, pageHeight, span, pageContentBudget, null, cancellationToken);

    private static void AddTextSpanCore(OfficeDrawing drawing, double pageHeight, PdfTextSpan span, PageContentBudget pageContentBudget,
        (double Left, double Top, double Right, double Bottom)? measuredPaint, System.Threading.CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(span.Text) || !span.IsVisible) {
            return;
        }

        if (!span.CanScaleAggregateAdvance && TryAddSpacedText(drawing, pageHeight, span, pageContentBudget, cancellationToken)) {
            return;
        }

        var frame = GetTextFrame(pageHeight, span.X, span.Y, span.FontSize, span.Text.Length, span.Advance);
        double height = frame.Height;
        double width = frame.Width;
        double rawX = frame.X;
        double rawY = frame.Y;
        var paint = measuredPaint ?? GetTextPaintBounds(drawing, span, frame, pageHeight - span.Y, pageContentBudget.CancellationToken);
        if (!HasVisibleOverlap(paint.Left, paint.Top, paint.Right - paint.Left, paint.Bottom - paint.Top, drawing.Width, drawing.Height)) {
            return;
        }

        if (TryAddClippedTextSpan(drawing, span, rawX, rawY, width, height, pageHeight - span.Y,
            span.ClipPath ?? PdfPageClipPath.Rectangle(0D, 0D, drawing.Width, drawing.Height), paint)) return;

        double x = rawX;
        double y = rawY;
        double baselineY = pageHeight - span.Y;
        // An unsupported source clip still needs page clipping. Preserve the raw
        // origin and rotation center; clamping either moves visible glyph ink.
        if (span.ClipPath.HasValue && TryAddClippedTextSpan(drawing, span, x, y, width, height, baselineY,
            PdfPageClipPath.Rectangle(0D, 0D, drawing.Width, drawing.Height), paint)) return;

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

    private static bool TryAddClippedTextSpan(OfficeDrawing drawing, PdfTextSpan span, double x, double y, double width, double height, double baselineY, PdfPageClipPath? overrideClipPath = null,
        (double Left, double Top, double Right, double Bottom)? paintBounds = null) {
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
        var paint = paintBounds ?? (x, y, x + width, y + height);
        if (clip.IsRectangle && paint.Item1 >= clip.X && paint.Item2 >= clip.Y && paint.Item3 <= clipRight && paint.Item4 <= clipBottom &&
            x >= 0D && y >= 0D && x + width <= drawing.Width && y + height <= drawing.Height) {
            return false;
        }

        if (paint.Item3 <= clip.X || paint.Item4 <= clip.Y || paint.Item1 >= clipRight || paint.Item2 >= clipBottom) {
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
            if (clip.Width <= 0D || clip.Height <= 0D) return true;
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

    // Character and word spacing move the next glyph without stretching the painted glyph.
    // Project these runs individually instead of fitting the whole run as a drawing label.
    private static bool TryAddSpacedText(OfficeDrawing drawing, double pageHeight, PdfTextSpan span, PageContentBudget pageContentBudget,
        System.Threading.CancellationToken cancellationToken) {
        // Nested forms, patterns and transparency groups share the render invocation token.
        cancellationToken = pageContentBudget.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        PdfPageClipPath? activeClip = span.ClipPath;
        OfficeClipPath? sharedClip = null;
        PdfPageClipPath sharedClipBounds = default;
        bool canCullClip = false;
        if (activeClip.HasValue) {
            PdfPageClipPath clip = activeClip.Value;
            if (clip.Width <= 0D || clip.Height <= 0D) return true;
            canCullClip = clip.IsRectangle || clip.ToOfficeClipPath(clip.X, clip.Y) != null;
            if (canCullClip && !HasVisibleOverlap(clip.X, clip.Y, clip.Width, clip.Height, drawing.Width, drawing.Height)) return true;
            if (canCullClip && !clip.IsRectangle && TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out sharedClipBounds)) {
                if (sharedClipBounds.Width <= 0D || sharedClipBounds.Height <= 0D) return true;
                sharedClip = sharedClipBounds.ToOfficeClipPath(sharedClipBounds.X, sharedClipBounds.Y);
            }
        }
        if (!CanExpandSpacedText(span, cancellationToken, out double direction)) return false;
        IReadOnlyList<int> characterLengths = span.GlyphCharacterLengths!;
        IReadOnlyList<double> paintedAdvances = span.GlyphPaintedAdvances!;
        IReadOnlyList<double> characterAdvances = span.CharacterAdvances!;
        var measuredGlyphs = new System.Collections.Generic.Dictionary<(string Text, double Advance), (double Left, double Top, double Right, double Bottom)>();
        var metrics = OfficeDrawingTextLayout.CreateMetrics(drawing, cancellationToken);
        // A path can contain thousands of commands. Retain it once for the run,
        // rather than cloning and rasterizing it separately for every glyph.
        OfficeDrawing glyphDrawing = sharedClip == null ? drawing : new OfficeDrawing(drawing.Width, drawing.Height);
        PdfPageClipPath? glyphClip = sharedClip == null ? span.ClipPath :
            PdfPageClipPath.Rectangle(sharedClipBounds.X, sharedClipBounds.Y, sharedClipBounds.Width, sharedClipBounds.Height);

        double radians = span.RotationDegrees * Math.PI / 180D;
        double alongX = Math.Cos(radians);
        double alongY = Math.Sin(radians);
        int characterOffset = 0;
        double offset = 0D;
        for (int index = 0; index < characterLengths.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = characterLengths[index];
            double x = span.X + alongX * offset;
            double y = span.Y + alongY * offset;
            var frame = GetTextFrame(pageHeight, x, y, span.FontSize, length, paintedAdvances[index]);
            string glyphText = span.Text.Substring(characterOffset, length);
            var key = (glyphText, paintedAdvances[index]);
            if (!measuredGlyphs.TryGetValue(key, out var localPaint)) {
                var localFrame = GetTextFrame(span.FontSize, 0, 0, span.FontSize, length, paintedAdvances[index]);
                localPaint = MeasureTextPaintBounds(metrics, glyphText, span, localFrame, span.FontSize, paintedAdvances[index]);
                if (measuredGlyphs.Count < 256) measuredGlyphs[key] = localPaint;
            }
            var paint = (Left: localPaint.Left + x, Top: localPaint.Top + frame.Y,
                Right: localPaint.Right + x, Bottom: localPaint.Bottom + frame.Y);
            if (HasPaintOverlap(drawing, paint, canCullClip ? activeClip : null)) {
                // Charge only scene expansion, before creating any glyph string or object.
                pageContentBudget.ChargePositionedTextCharacters(length);
                var glyph = new PdfTextSpan(glyphText, span.FontResource, span.FontSize,
                    x, y, paintedAdvances[index], span.Color, span.IsVisible, span.RotationDegrees, span.BaseFont, glyphClip,
                    drawingFontFamily: span.DrawingFontFamily, fontWeight: span.FontWeight,
                    fontDescriptorFlags: span.FontDescriptorFlags);
                AddTextSpanCore(glyphDrawing, pageHeight, glyph, pageContentBudget, paint, cancellationToken);
            }
            // Stream origins instead of allocating an array for a potentially huge
            // off-page run. Nonpainting prefixes still move later glyphs correctly.
            int end = characterOffset + length;
            for (; characterOffset < end; characterOffset++) {
                cancellationToken.ThrowIfCancellationRequested();
                offset += characterAdvances[characterOffset] * direction;
            }
        }
        if (sharedClip != null && glyphDrawing.Elements.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            drawing.AddClippedDrawing(glyphDrawing, sharedClipBounds.X, sharedClipBounds.Y, sharedClip,
                -sharedClipBounds.X, -sharedClipBounds.Y);
        }
        return true;
    }

    private static (double X, double Y, double Width, double Height) GetTextFrame(
        double pageHeight, double x, double y, double fontSize, int textLength, double advance) =>
        (x, pageHeight - y - fontSize, Math.Max(advance, textLength * fontSize * .55D), Math.Max(1D, fontSize * 1.25D));

    private static (double Left, double Top, double Right, double Bottom) GetTextPaintBounds(OfficeDrawing drawing, PdfTextSpan span,
        (double X, double Y, double Width, double Height) frame, double baseline, System.Threading.CancellationToken cancellationToken) =>
        MeasureTextPaintBounds(OfficeDrawingTextLayout.CreateMetrics(drawing, cancellationToken), span.Text, span, frame, baseline, span.Advance);

    private static (double Left, double Top, double Right, double Bottom) MeasureTextPaintBounds(OfficeRasterCanvas metrics, string text, PdfTextSpan span,
        (double X, double Y, double Width, double Height) frame, double baseline, double advance) {
        var bounds = metrics.MeasurePositionedTextBounds(text, frame.X, frame.Y, frame.Width, frame.Height, Math.Max(1D, span.FontSize),
            ToOfficeFontInfo(span.BaseFont, span.FontSize, span.DrawingFontFamily, span.IsBold, span.IsItalic),
            advance > 0D ? advance : frame.Width, OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "", Math.Max(1D, span.FontSize),
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None);
        return new OfficeImageFrameTransform(-span.RotationDegrees, frame.X, baseline).CreateDestinationTransform()
            .TransformRectangleBounds(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top);
    }

    private static bool HasPaintOverlap(OfficeDrawing drawing,
        (double Left, double Top, double Right, double Bottom) paint, PdfPageClipPath? clip) =>
        HasVisibleOverlap(paint.Left, paint.Top, paint.Right - paint.Left, paint.Bottom - paint.Top, drawing.Width, drawing.Height) &&
        (!clip.HasValue || paint.Right > clip.Value.X && paint.Bottom > clip.Value.Y &&
            paint.Left < clip.Value.X + clip.Value.Width && paint.Top < clip.Value.Y + clip.Value.Height);

    private static bool CanExpandSpacedText(PdfTextSpan span, System.Threading.CancellationToken cancellationToken,
        out double direction) {
        direction = 0D;
        cancellationToken.ThrowIfCancellationRequested();
        // PDF character codes are not shaping clusters. Expand only basic Latin
        // display runs; retain whole-run shaping for other scripts, marks and emoji.
        // Supporting those at individual origins requires a shaped-glyph run contract.
        if (span.HasActualText) return false;
        for (int index = 0; index < span.Text.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (span.Text[index] < ' ' || span.Text[index] > '~') return false;
        }
        IReadOnlyList<int>? lengths = span.GlyphCharacterLengths;
        IReadOnlyList<double>? widths = span.GlyphPaintedAdvances;
        if (lengths == null || widths == null || lengths.Count == 0 || lengths.Count != widths.Count) return false;
        long count = 0;
        for (int index = 0; index < lengths.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (lengths[index] <= 0 || widths[index] <= 0D || double.IsNaN(widths[index]) || double.IsInfinity(widths[index])) return false;
            count += lengths[index];
        }
        if (count != span.Text.Length) return false;
        return PdfTextAdvanceProjection.TryGetResolvedDirection(span, cancellationToken, out direction);
    }

}
