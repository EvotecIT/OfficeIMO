namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // A wrapped line owns its placement as well as its advance. Keeping that
    // placement on the line preserves it when tables and columns slice lines,
    // including paragraphs whose fallback size differs from the cell size.
    private sealed class RichLine : List<RichSeg> {
        // Fragment plans survive ordinary row slicing, but never a segment-copy transformation.
        internal IReadOnlyList<RichLineFragment>? Fragments { get; set; }
        internal double FragmentTopGap { get; set; }
        internal bool IsFragmentContext { get; set; }
        internal double BaselineOffset { get; set; }
        internal double? FrameFitHeight { get; set; }
        internal double MinimumContentHeight { get; set; }
        internal PdfTabStop[]? BarTabStops { get; set; }
        internal double BarOriginOffset { get; set; }
        internal double BarLineHeight { get; set; }
        internal double BarSpacingBefore { get; set; }
        internal double BarSpacingAfter { get; set; }
        internal bool IsFirstParagraphLine { get; set; }
        internal bool IsLastParagraphLine { get; set; }

        /// <summary>Creates an empty line retaining placement and paragraph decorations independently of its text.</summary>
        internal RichLine CopyPlacement() => new RichLine {
            IsFragmentContext = IsFragmentContext, BaselineOffset = BaselineOffset, FrameFitHeight = FrameFitHeight, MinimumContentHeight = MinimumContentHeight,
            BarTabStops = BarTabStops, BarOriginOffset = BarOriginOffset, BarLineHeight = BarLineHeight,
            BarSpacingBefore = BarSpacingBefore, BarSpacingAfter = BarSpacingAfter,
            IsFirstParagraphLine = IsFirstParagraphLine, IsLastParagraphLine = IsLastParagraphLine
        };
    }

    private static double ResolveRichContentHeight(double contentHeight, PdfLineSpacing? spacing,
        double textAscent, double textDescent, double bodyAscent,
        double inlineAscent, double inlineDescent, bool hasInline) {
        if (spacing?.IsExact == true) return contentHeight;
        double combined = Math.Max(textAscent, inlineAscent) + Math.Max(textDescent, inlineDescent);
        if (spacing?.FontLineBoxMultiplier is not double natural)
            return hasInline ? Math.Max(contentHeight, combined) : contentHeight;
        if (spacing.LineGridPitch.HasValue)
            return Math.Max(contentHeight, spacing.GetGridAdvance(combined));
        double textAdvance = textAscent + textDescent;
        if (spacing.Rule == PdfLineSpacingRule.Multiple)
            textAdvance = bodyAscent == 0D && textDescent == 0D ? 0D
                : (bodyAscent + textDescent) * spacing.MultipleValue!.Value / natural
                    + Math.Max(0D, textAscent - bodyAscent);
        contentHeight = Math.Max(contentHeight, textAdvance);
        return hasInline && (inlineAscent > textAscent || inlineDescent > textDescent)
            ? Math.Max(contentHeight, combined) : contentHeight;
    }

    // A line's retained tab gaps can change after exclusion/page placement.
    // Rebuild its owned metrics from the segments that will actually paint,
    // while keeping the caller's line break and exact-spacing policy intact.
    private static double RefreshRichLineMetrics(RichLine line, PdfStandardFont baseFont,
        double fontSize, double leading, PdfCreateOptions options, PdfLineSpacing? spacing) {
        double advance = leading;
        if (spacing?.IsExact != true) {
            double textAscent = 0D, textDescent = 0D, bodyAscent = 0D;
            double inlineAscent = 0D, inlineDescent = 0D;
            bool hasInline = false;
            double ratio = fontSize > 0D ? leading / fontSize : 1.2D;
            foreach (RichSeg segment in line) {
                bool emptySeparator = segment.Text.Length == 0 && (segment.LeadingSpace || segment.LeadingIsTab) && segment.InlineElement == null;
                if (emptySeparator && segment.LeadingIsTab && !HasPaintedRichTabLeader(segment, options)) continue;
                if (segment.InlineElement is { } inline && !UsesInlineRunTextMetrics(inline)) {
                    double ascent = Math.Max(0D, inline.BaselineOffset.ToPoints() + inline.Height.ToPoints());
                    double descent = Math.Max(0D, -inline.BaselineOffset.ToPoints());
                    inlineAscent = Math.Max(inlineAscent, ascent);
                    inlineDescent = Math.Max(inlineDescent, descent);
                    hasInline = true;
                    double inlineAdvance = spacing?.FontLineBoxMultiplier != null
                        ? Math.Max(leading, ascent + descent)
                        : Math.Max(GetAscenderForOptions(baseFont, fontSize, options), ascent)
                            + Math.Max(GetDescenderForOptions(baseFont, fontSize, options), descent);
                    advance = Math.Max(advance, inlineAdvance);
                    continue;
                }
                if (segment.ContributesLineBoxDescent)
                    advance = Math.Max(advance, spacing?.GetAdvance(segment.FontSize) ?? segment.FontSize * ratio);
                // Literal source spaces retain their font advance. As with the
                // wrapper's baseline policy, an empty separator has text metrics
                // only when its resolved tab leader actually paints glyphs.
                if (emptySeparator && !HasPaintedRichTabLeader(segment, options)) continue;
                double runAscent, runDescent;
                if (segment.InlineElement is PdfInlineFormWidget { SourceRun: not null } widget)
                    GetInlineFormWidgetLineMetrics(widget, baseFont, fontSize, options, spacing, out runAscent, out runDescent);
                else GetRichRunLineMetrics(segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                    options, spacing, out runAscent, out runDescent);
                textAscent = Math.Max(textAscent, runAscent);
                if (segment.ContributesLineBoxDescent) {
                    textDescent = Math.Max(textDescent, runDescent);
                    bodyAscent = Math.Max(bodyAscent, runAscent);
                }
            }
            advance = ResolveRichContentHeight(advance, spacing, textAscent, textDescent,
                bodyAscent, inlineAscent, inlineDescent, hasInline);
        }
        RefreshRichLinePlacementMetrics(line, baseFont, fontSize, options, spacing);
        if (line.BarTabStops != null) line.BarLineHeight = advance;
        return advance;
    }

    private static void RefreshRichLinePlacementMetrics(RichLine line, PdfStandardFont baseFont,
        double fontSize, PdfCreateOptions? options, PdfLineSpacing? spacing) {
        foreach (RichSeg segment in line) segment.LineAlignmentRise = ResolveRichRunLineAlignmentRise(segment, options, spacing);
        line.BaselineOffset = ResolveRichLineBaselineOffset(line, baseFont, fontSize, options, spacing);
        line.MinimumContentHeight = 0D;
        line.FrameFitHeight = null;
        if (spacing == null) {
            GetRichLineInkMetrics(line, options ?? new PdfCreateOptions(), out _, out double descent);
            line.MinimumContentHeight = line.BaselineOffset + descent;
        }
        if (spacing?.FontLineBoxMultiplier != null && spacing.Rule == PdfLineSpacingRule.Multiple)
            line.FrameFitHeight = ResolveImportedLineFrameFitHeight(line, baseFont, fontSize, options, spacing);
    }

    private static double ResolveRichLineBaselineOffset(
        IReadOnlyList<RichSeg> segments, PdfStandardFont baseFont, double fontSize,
        PdfCreateOptions? options, PdfLineSpacing? spacing) {
        double fallbackAscent = GetAscenderForOptions(baseFont, fontSize, options);
        if (spacing?.IsExact == true) return spacing.FixedLineBoxBaselineOffset ?? fallbackAscent;

        double ascent = 0D;
        double bodyAscent = 0D;
        double inlineAscent = 0D;
        double descent = 0D;
        foreach (RichSeg segment in segments) {
            // Empty separators move the baseline only when they paint a tab leader.
            if (segment.Text.Length == 0 && (segment.LeadingSpace || segment.LeadingIsTab) && segment.InlineElement == null &&
                !HasPaintedRichTabLeader(segment, options)) continue;
            if (segment.InlineElement is PdfInlineFormWidget { SourceRun: not null } widget) {
                GetInlineFormWidgetLineMetrics(widget, baseFont, fontSize, options, spacing,
                    out double runAscent, out double runDescent);
                ascent = Math.Max(ascent, runAscent);
                if (segment.ContributesLineBoxDescent) {
                    bodyAscent = Math.Max(bodyAscent, runAscent);
                    descent = Math.Max(descent, runDescent);
                }
            } else if (segment.InlineElement is { } inline && !UsesInlineRunTextMetrics(inline)) {
                ascent = Math.Max(ascent, inline.BaselineOffset.ToPoints() + inline.Height.ToPoints());
                inlineAscent = Math.Max(inlineAscent, inline.BaselineOffset.ToPoints() + inline.Height.ToPoints());
                descent = Math.Max(descent, -inline.BaselineOffset.ToPoints());
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
            ascent = Math.Max(inlineAscent, ascent - bodyAscent * (1D - Math.Min(1D, spacing.MultipleValue!.Value / natural)));
        else if (spacing?.FontLineBoxMultiplier != null && spacing.Rule == PdfLineSpacingRule.AtLeast)
            ascent = Math.Max(ascent, spacing.Height!.Value.ToPoints() - descent);
        return ascent == 0D ? fallbackAscent : ascent;
    }

    private static void GetRichRunLineMetrics(PdfStandardFont font, PdfNamedFontFace? namedFont,
        double fontSize, PdfTextBaseline baseline, PdfCreateOptions? options, PdfLineSpacing? spacing,
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
        double fontSize, PdfCreateOptions? options, PdfLineSpacing? spacing, out double ascent, out double descent) {
        ascent = 0D;
        descent = 0D;
        foreach (PdfTextRun run in NormalizeFallbackRuns(new[] { widget.SourceRun! }, baseFont, options)) {
            PdfStandardFont font = ResolveFontForRun(run, baseFont);
            PdfNamedFontFace? namedFont = options != null &&
                options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolved)
                    ? resolved : null;
            GetRichRunLineMetrics(font, namedFont, (run.FontSize?.ToPoints()) ?? fontSize, run.Baseline,
                options, spacing, out double runAscent, out double runDescent);
            ascent = Math.Max(ascent, runAscent);
            if (run.ContributesLineBoxDescent) descent = Math.Max(descent, runDescent);
        }
    }

    private static OfficeIMO.Drawing.OfficeOpenTypeLineMetrics? ResolveFontLineMetrics(
        PdfStandardFont font, PdfNamedFontFace? namedFont, PdfCreateOptions? options) {
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
        PdfCreateOptions options, double fontSize, PdfStandardFont? baselineFont = null) {
        double baseAscent = GetAscenderForOptions(baselineFont ?? ChooseNormal(options.DefaultFont), fontSize, options);
        if (segments is RichLine line) return baseline + baseAscent - line.BaselineOffset;

        // Positioned text already carries an authored baseline. Retain that
        // geometry; these lines do not participate in paragraph line layout.
        double requiredAscent = baseAscent;
        foreach (RichSeg segment in segments) {
            if (segment.InlineElement is { } inline)
                requiredAscent = Math.Max(requiredAscent, inline.BaselineOffset.ToPoints() + inline.Height.ToPoints());
        }
        return baseline - Math.Max(0D, requiredAscent - baseAscent);
    }

    private static void GetRichLineInkMetrics(IReadOnlyList<RichSeg> line, PdfCreateOptions options,
        out double ascender, out double descender) {
        ascender = 0D; descender = 0D;
        foreach (RichSeg segment in line) {
            if (segment.IsLineBoxStrut) continue;
            if (segment.Text.Length == 0 && segment.LeadingIsTab && segment.InlineElement == null &&
                !HasPaintedRichTabLeader(segment, options)) continue;
            if (segment.InlineElement is { } inline) {
                if (inline is PdfInlineBox { SourceRule: { } rule } box) {
                    if (box.Width.ToPoints() > 0D) {
                        foreach (PdfInlineRule.InkBand band in rule.Bands) {
                            ascender = Math.Max(ascender, segment.LineAlignmentRise + band.BaselineOffset + band.Height);
                            descender = Math.Max(descender, -band.BaselineOffset - segment.LineAlignmentRise);
                        }
                    }
                    continue;
                }
                ascender = Math.Max(ascender, inline.BaselineOffset.ToPoints() + inline.Height.ToPoints());
                descender = Math.Max(descender, -inline.BaselineOffset.ToPoints());
            } else {
                double rise = TextRiseForBaseline(segment.FontSize, segment.Baseline) + segment.LineAlignmentRise;
                double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
                ascender = Math.Max(ascender, rise + GetAscenderForOptions(segment.Font, segment.NamedFont, size, options));
                descender = Math.Max(descender, GetDescenderForOptions(segment.Font, segment.NamedFont, size, options) - rise);
            }
        }
    }
}
