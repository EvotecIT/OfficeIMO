namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // A wrapped line owns its placement as well as its advance. Keeping that
    // placement on the line preserves it when tables and columns slice lines,
    // including paragraphs whose fallback size differs from the cell size.
    private sealed class RichLine : List<RichSeg> {
        internal double BaselineOffset { get; set; }
        internal double? FrameFitHeight { get; set; }
        internal double MinimumContentHeight { get; set; }
    }

    private static double ResolveRichLineBaselineOffset(
        IReadOnlyList<RichSeg> segments, PdfStandardFont baseFont, double fontSize,
        PdfOptions? options, PdfLineSpacing? spacing) {
        double fallbackAscent = GetAscenderForOptions(baseFont, fontSize, options);
        if (spacing?.IsExact == true) return spacing.FixedLineBoxBaselineOffset ?? fallbackAscent;

        double ascent = 0D;
        double bodyAscent = 0D;
        double inlineAscent = 0D;
        double descent = 0D;
        foreach (RichSeg segment in segments) {
            // Preserved spaces carry their source font for painting and later
            // remeasurement. They have no visible glyph to move the line baseline.
            if (segment.Text.Length == 0 && segment.LeadingSpace && segment.InlineElement == null) continue;
            if (segment.InlineElement is PdfInlineFormWidget { SourceRun: not null } widget) {
                GetInlineFormWidgetLineMetrics(widget, baseFont, fontSize, options, spacing,
                    out double runAscent, out double runDescent);
                ascent = Math.Max(ascent, runAscent);
                if (segment.ContributesLineBoxDescent) {
                    bodyAscent = Math.Max(bodyAscent, runAscent);
                    descent = Math.Max(descent, runDescent);
                }
            } else if (segment.InlineElement is { } inline && !UsesInlineRunTextMetrics(inline)) {
                ascent = Math.Max(ascent, inline.BaselineOffset + inline.Height);
                inlineAscent = Math.Max(inlineAscent, inline.BaselineOffset + inline.Height);
                descent = Math.Max(descent, -inline.BaselineOffset);
            } else {
                GetRichRunLineMetrics(segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                    options, spacing, out double runAscent, out double runDescent);
                ascent = Math.Max(ascent, runAscent);
                if (segment.ContributesLineBoxDescent) {
                    bodyAscent = Math.Max(bodyAscent, runAscent);
                    descent = Math.Max(descent, runDescent);
                }
            }
        }
        if (spacing?.LineGridPitch != null)
            ascent += spacing.GetGridBaselineExpansion(ascent + descent);
        else if (spacing?.FontLineBoxMultiplier is double natural && spacing.Rule == PdfLineSpacingRule.Multiple)
            ascent = Math.Max(inlineAscent, ascent - bodyAscent * (1D - Math.Min(1D, spacing.Value / natural)));
        else if (spacing?.FontLineBoxMultiplier != null && spacing.Rule == PdfLineSpacingRule.AtLeast)
            ascent = Math.Max(ascent, spacing.Value - descent);
        return ascent == 0D ? fallbackAscent : ascent;
    }

    private static void GetRichRunLineMetrics(PdfStandardFont font, PdfNamedFontFace? namedFont,
        double fontSize, PdfTextBaseline baseline, PdfOptions? options, PdfLineSpacing? spacing,
        out double ascent, out double descent) {
        double size = EffectiveRichFontSize(fontSize, baseline);
        double rise = TextRiseForBaseline(fontSize, baseline);
        if (spacing?.IsExact == false)
            rise += ResolveFontLineAlignmentRise(font, namedFont, size, options, spacing);
        if (spacing?.FontLineBoxMultiplier is double natural) {
            OfficeIMO.Drawing.OfficeOpenTypeLineMetrics? metrics = ResolveFontLineMetrics(font, namedFont, options);
            double fontDescent = metrics.HasValue ? size * metrics.Value.DescentRatio
                : GetDescenderForOptions(font, namedFont, size, options);
            ascent = rise + size * (metrics?.HorizontalAdvanceRatio ?? natural) - fontDescent;
            descent = fontDescent - rise;
        } else {
            ascent = rise + GetAscenderForOptions(font, namedFont, size, options);
            descent = GetDescenderForOptions(font, namedFont, size, options) - rise;
        }
    }

    // A source glyph keeps its text line-box policy even when its ink is owned
    // by an editable widget. Its annotation rectangle is a separate paint box.
    private static void GetInlineFormWidgetLineMetrics(PdfInlineFormWidget widget, PdfStandardFont baseFont,
        double fontSize, PdfOptions? options, PdfLineSpacing? spacing, out double ascent, out double descent) {
        ascent = 0D;
        descent = 0D;
        foreach (PdfTextRun run in NormalizeFallbackRuns(new[] { widget.SourceRun! }, baseFont, options)) {
            PdfStandardFont font = ResolveFontForRun(run, baseFont);
            PdfNamedFontFace? namedFont = options != null &&
                options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolved)
                    ? resolved : null;
            GetRichRunLineMetrics(font, namedFont, run.FontSize ?? fontSize, run.Baseline,
                options, spacing, out double runAscent, out double runDescent);
            ascent = Math.Max(ascent, runAscent);
            if (run.ContributesLineBoxDescent) descent = Math.Max(descent, runDescent);
        }
    }

    private static OfficeIMO.Drawing.OfficeOpenTypeLineMetrics? ResolveFontLineMetrics(
        PdfStandardFont font, PdfNamedFontFace? namedFont, PdfOptions? options) {
        if (options == null) return null;
        if (namedFont.HasValue) {
            if (options.TryGetNamedFontProgram(namedFont.Value, out PdfTrueTypeFontProgram? namedTrueType) && namedTrueType != null)
                return namedTrueType.LineMetrics;
            if (options.TryGetNamedOpenTypeCffFontProgram(namedFont.Value, out PdfOpenTypeCffFontProgram? namedCff) && namedCff != null)
                return namedCff.LineMetrics;
        }
        if (options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? trueType) && trueType != null)
            return trueType.LineMetrics;
        if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cff) && cff != null)
            return cff.LineMetrics;
        return null;
    }

    private static double AdjustRichLineBaseline(double baseline, IReadOnlyList<RichSeg> segments,
        PdfOptions options, double fontSize, PdfStandardFont? baselineFont = null) {
        double baseAscent = GetAscenderForOptions(baselineFont ?? ChooseNormal(options.DefaultFont), fontSize, options);
        if (segments is RichLine line) return baseline + baseAscent - line.BaselineOffset;

        // Positioned text already carries an authored baseline. Retain that
        // geometry; these lines do not participate in paragraph line layout.
        double requiredAscent = baseAscent;
        foreach (RichSeg segment in segments) {
            if (segment.InlineElement is { } inline)
                requiredAscent = Math.Max(requiredAscent, inline.BaselineOffset + inline.Height);
        }
        return baseline - Math.Max(0D, requiredAscent - baseAscent);
    }

    private static void GetRichLineInkMetrics(IReadOnlyList<RichSeg> line, PdfOptions options,
        out double ascender, out double descender) {
        ascender = 0D; descender = 0D;
        foreach (RichSeg segment in line) {
            if (segment.IsLineBoxStrut) continue;
            if (segment.InlineElement is { } inline) {
                if (inline is PdfInlineBox { SourceRule: { } rule } box) {
                    if (box.Width > 0D) {
                        foreach (PdfInlineRule.InkBand band in rule.Bands) {
                            ascender = Math.Max(ascender, segment.LineAlignmentRise + band.BaselineOffset + band.Height);
                            descender = Math.Max(descender, -band.BaselineOffset - segment.LineAlignmentRise);
                        }
                    }
                    continue;
                }
                ascender = Math.Max(ascender, inline.BaselineOffset + inline.Height);
                descender = Math.Max(descender, -inline.BaselineOffset);
            } else {
                double rise = TextRiseForBaseline(segment.FontSize, segment.Baseline) + segment.LineAlignmentRise;
                double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
                ascender = Math.Max(ascender, rise + GetAscenderForOptions(segment.Font, segment.NamedFont, size, options));
                descender = Math.Max(descender, GetDescenderForOptions(segment.Font, segment.NamedFont, size, options) - rise);
            }
        }
    }
}
