namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string FitInlineCheckBoxAppearance(string content, InlineFormAppearance appearance,
        double textX, PositionedTextBoxMeasurement measured, LayoutResult.Page page, PdfCreateOptions options) {
        // Editing retains a fixed PDF rectangle. Advances locate the next text
        // origin, while selected glyph outlines own any overhang or bearing.
        // Both contribute to fitting; neither changes source-state paragraph flow.
        double left = textX, right = textX + measured.Width;
        double bottom = appearance.Baseline - measured.Descent;
        double top = appearance.Baseline + measured.Ascent;
        foreach (InlineFormTextFragment fragment in appearance.Fragments) {
            IncludeInlineFormPaintBounds(page, fragment.Segment, fragment.X, fragment.Baseline,
                options, ref left, ref bottom, ref right, ref top);
        }

        double scaleX = right > left ? Math.Min(1D, appearance.Right / (right - left)) : 1D;
        double translateX = textX * (1D - scaleX);
        translateX = Math.Max(-left * scaleX, Math.Min(translateX, appearance.Right - right * scaleX));

        double scaleY = 1D;
        double above = top - appearance.Baseline, below = appearance.Baseline - bottom;
        if (above > 0D) scaleY = Math.Min(scaleY, (appearance.Top - appearance.Baseline) / above);
        if (below > 0D) scaleY = Math.Min(scaleY, appearance.Baseline / below);
        double translateY;
        if (scaleY > 0D) {
            translateY = appearance.Baseline * (1D - scaleY);
        } else {
            // A baseline at the rectangle's edge cannot retain an alternate
            // glyph's opposite-side ink at any positive baseline-anchored scale.
            // Fit that state inside the rectangle instead of making it vanish.
            scaleY = top > bottom ? Math.Min(1D, appearance.Top / (top - bottom)) : 1D;
            translateY = Math.Max(-bottom * scaleY,
                Math.Min(appearance.Baseline * (1D - scaleY), appearance.Top - top * scaleY));
        }
        if (scaleX == 1D && scaleY == 1D && translateX == 0D && translateY == 0D) return content;

        var fitted = new StringBuilder();
        new ContentStreamBuilder(fitted).SaveState().TransformMatrix(scaleX, 0D, 0D, scaleY,
            translateX, translateY, preciseCoordinates: true);
        fitted.Append(content);
        new ContentStreamBuilder(fitted).RestoreState();
        return fitted.ToString();
    }
}
