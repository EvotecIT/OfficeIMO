namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static RichSeg CreateInlineAppearanceSegment(PdfTextRun run, PdfCreateOptions options) {
        RichSeg segment = CreatePositionedTextSegment(run, options.DefaultFontSize, options);
        PrepareInlineAppearanceCommand(segment, options);
        return segment;
    }

    private static void PrepareInlineAppearanceCommand(RichSeg segment, PdfCreateOptions options) {
        // Register glyphs during layout, before font programs are subset and
        // serialized. The appearance later paints this exact cached command.
        segment.KerningCommand ??= EncodeTextShowCommand(segment.Text, segment.Font, segment.NamedFont,
            options, segment.FeatureSettings, segment.TextDirection, segment.FontMetricScale,
            preservesPaintedGlyphs: segment.PreservesPaintedGlyphs);
    }

    private static void PrepareInlineCheckBoxGlyphUsage(PdfInlineFormField field, PdfCreateOptions options) {
        foreach (PdfTextRun? source in new[] { field.CheckedAppearanceRun, field.UncheckedAppearanceRun }) {
            if (source == null) continue;
            foreach (PdfTextRun run in NormalizeFallbackRuns(new[] { source }, ChooseNormal(options.DefaultFont), options))
                _ = CreateInlineAppearanceSegment(run, options);
        }
    }

    private static void IncludeInlineFormPaintBounds(LayoutResult.Page page, RichSeg segment,
        double x, double baseline, PdfCreateOptions options,
        ref double left, ref double bottom, ref double right, ref double top) {
        page.OperationCancellationToken.ThrowIfCancellationRequested();
        PrepareInlineAppearanceCommand(segment, options);
        double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
        double paintedBaseline = baseline + TextRiseForBaseline(segment.FontSize, segment.Baseline) + segment.LineAlignmentRise;
        try {
            if (GetGlyphPaintFont(segment, options, page.InlineFormGlyphFonts) == null) {
                // Core 14 resources supply nominal metrics, not an outline program.
                // Reuse the font's conservative paint envelope: a viewer can resolve
                // a compatible face whose ink differs from the individual AFM box.
                var horizontal = GetStandardFontHorizontalPaintBounds(segment.Font, size, 0D);
                double scale = segment.HorizontalTextScaling / 100D;
                left = Math.Min(left, x + horizontal.Left * scale);
                right = Math.Max(right, x + GetRichSegmentWidth(segment) + horizontal.Right * scale);
            }
            if (GetGlyphPaintBounds(segment, segment.Text, x, paintedBaseline, options,
                page.InlineFormGlyphFonts, paintCommand: segment.KerningCommand,
                cancellationToken: page.OperationCancellationToken) is { } ink) {
                left = Math.Min(left, ink.Left); bottom = Math.Min(bottom, ink.Bottom);
                right = Math.Max(right, ink.Right); top = Math.Max(top, ink.Top);
            }
        } catch (NotSupportedException error) {
            // Preserve renderability where exact outline inspection is unavailable.
            // The declared font box is conservative, and the report retains that gap.
            int[]? fontBox = null;
            if (segment.NamedFont is { } named) {
                if (options.TryGetNamedFontProgram(named, out var trueType) && trueType != null) fontBox = trueType.FontBBox;
                else if (options.TryGetNamedOpenTypeCffFontProgram(named, out var cff) && cff != null) fontBox = cff.FontBBox;
            } else {
                if (options.TryGetEmbeddedStandardFontProgram(segment.Font, out var trueType) && trueType != null) fontBox = trueType.FontBBox;
                else if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out var cff) && cff != null) fontBox = cff.FontBBox;
            }
            if (fontBox == null) {
                var box = GetStandardFontBoundingBox(segment.Font);
                fontBox = new[] { box.Left, box.Bottom, box.Right, box.Top };
            }
            double scale = size / 1000D;
            double horizontal = segment.HorizontalTextScaling / 100D;
            left = Math.Min(left, x + fontBox[0] * scale * horizontal);
            right = Math.Max(right, x + GetRichSegmentWidth(segment) + fontBox[2] * scale * horizontal);
            bottom = Math.Min(bottom, paintedBaseline + fontBox[1] * scale);
            top = Math.Max(top, paintedBaseline + fontBox[3] * scale);
            options.AddUnassessedLayoutWarning("inline-form-paint-bounds-unmeasured", "inline form", error.Message);
        }
        if (segment.Underline)
            IncludeDecorationVerticalBounds(segment.UnderlineStyle, paintedBaseline - size * .15D, .5D, ref bottom, ref top);
        if (segment.Strike)
            IncludeDecorationVerticalBounds(segment.StrikeStyle, paintedBaseline + size * .32D, .5D, ref bottom, ref top);
    }
}
