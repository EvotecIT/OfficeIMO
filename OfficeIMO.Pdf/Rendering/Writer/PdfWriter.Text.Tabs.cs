namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double CalculateDefaultTabAdvance(double lineWidth, double spaceWidth, double tabStopWidth = DefaultParagraphTabStopWidth, double? explicitTabStopPosition = null, double? maxWidth = null) {
        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            (!explicitTabStopPosition.HasValue && (tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)))) {
            return spaceWidth;
        }

        double nextStop = GetTabTargetPosition(lineWidth, tabStopWidth, explicitTabStopPosition);
        if (TryGetBoundedMaxWidth(maxWidth, out double boundedMaxWidth)) {
            nextStop = Math.Min(nextStop, boundedMaxWidth);
        }

        return Math.Max(spaceWidth, nextStop - lineWidth);
    }

    private static double GetTabTargetPosition(double lineWidth, double tabStopWidth, double? explicitTabStopPosition) {
        if (explicitTabStopPosition.HasValue &&
            explicitTabStopPosition.Value >= 0 &&
            !double.IsNaN(explicitTabStopPosition.Value) &&
            !double.IsInfinity(explicitTabStopPosition.Value) &&
            explicitTabStopPosition.Value > lineWidth) {
            return explicitTabStopPosition.Value;
        }

        return (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
    }

    private static bool TryGetBoundedMaxWidth(double? maxWidth, out double boundedMaxWidth) {
        if (maxWidth.HasValue &&
            maxWidth.Value > 0 &&
            !double.IsNaN(maxWidth.Value) &&
            !double.IsInfinity(maxWidth.Value)) {
            boundedMaxWidth = maxWidth.Value;
            return true;
        }

        boundedMaxWidth = 0D;
        return false;
    }

    private static PdfTabStop? GetNextExplicitTabStop(double lineWidth, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops) {
        if (tabStops == null || tabStops.Count == 0 || lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth)) {
            return null;
        }

        PdfTabStop? next = null;
        for (int i = 0; i < tabStops.Count; i++) {
            PdfTabStop tabStop = tabStops[i];
            if (tabStop.Position <= lineWidth + 0.001D) {
                continue;
            }

            if (next == null || tabStop.Position < next.Position) {
                next = tabStop;
            }
        }

        return next;
    }

    private static PdfTabAlignment ResolveTabAlignment(PdfTabAlignment runAlignment, PdfTabStop? tabStop) =>
        tabStop != null && runAlignment == PdfTabAlignment.Left ? tabStop.Alignment : runAlignment;

    private static PdfTabLeaderStyle ResolveTabLeader(PdfTabLeaderStyle runLeader, PdfTabStop? tabStop) =>
        tabStop != null && runLeader == PdfTabLeaderStyle.None ? tabStop.Leader : runLeader;

    private static double MeasureDecimalAnchorWidth(string text, PdfStandardFont font, double fontSize, PdfTextBaseline baseline, PdfOptions? options = null, PdfNamedFontFace? namedFont = null, OfficeIMO.Drawing.OfficeTextFeatureSettings? featureSettings = null, double horizontalTextScaling = 100D, double characterSpacing = 0D, char? decimalSeparator = null) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        int decimalIndex = decimalSeparator.HasValue ? text.IndexOf(decimalSeparator.Value) : text.IndexOfAny(DecimalTabAnchorChars);
        if (decimalIndex < 0) {
            return MeasureRichText(text, font, namedFont, fontSize, baseline, options, featureSettings, horizontalTextScaling, characterSpacing);
        }

        // Shape the whole token: a pair spanning the separator can adjust both
        // the preceding advance and the separator's own painted position.
        PdfTextShowCommand command = EncodeTextShowCommand(text, font, namedFont, options, featureSettings,
            layoutHorizontalScaling: horizontalTextScaling);
        if (command.VisualGlyphs is { Count: > 0 } glyphs) {
            double effectiveSize = EffectiveRichFontSize(fontSize, baseline);
            double scale = horizontalTextScaling / 100D;
            double advance = 0D;
            double tracking = GetIntrinsicGlyphTrackingAdvance(command, effectiveSize);
            for (int index = 0; index < glyphs.Count; index++) {
                PdfGlyphInfo glyph = glyphs[index];
                if (glyph.TextIndex == decimalIndex) {
                    return advance + glyph.OffsetX1000 * effectiveSize / 1000D * scale;
                }
                double glyphAdvance = glyph.AdvanceWidth1000 * effectiveSize / 1000D;
                if (command.TrackingBoundaries?[index] == true) glyphAdvance += tracking;
                advance += glyphAdvance * scale + characterSpacing;
            }
        }

        return MeasureRichText(text.Substring(0, decimalIndex), font, namedFont, fontSize, baseline, options, featureSettings, horizontalTextScaling, characterSpacing);
    }

    private static double CalculateTabAdvance(double lineWidth, double followingTextWidth, double spaceWidth, PdfTabAlignment alignment, double tabStopWidth = DefaultParagraphTabStopWidth, string followingText = "", PdfStandardFont followingFont = PdfStandardFont.Helvetica, double fontSize = 12D, PdfTextBaseline baseline = PdfTextBaseline.Normal, PdfOptions? options = null, double? maxWidth = null) =>
        CalculateTabAdvanceToStop(lineWidth, followingTextWidth, spaceWidth, alignment, tabStopWidth, followingText, followingFont, fontSize, baseline, options, maxWidth, explicitTabStopPosition: null);

    private static double CalculateTabAdvanceToStop(double lineWidth, double followingTextWidth, double spaceWidth, PdfTabAlignment alignment, double tabStopWidth = DefaultParagraphTabStopWidth, string followingText = "", PdfStandardFont followingFont = PdfStandardFont.Helvetica, double fontSize = 12D, PdfTextBaseline baseline = PdfTextBaseline.Normal, PdfOptions? options = null, double? maxWidth = null, double? explicitTabStopPosition = null) {
        if (alignment == PdfTabAlignment.Left) {
            return CalculateDefaultTabAdvance(lineWidth, spaceWidth, tabStopWidth, explicitTabStopPosition, maxWidth);
        }

        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            followingTextWidth < 0 || double.IsNaN(followingTextWidth) || double.IsInfinity(followingTextWidth) ||
            (!explicitTabStopPosition.HasValue && (tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)))) {
            return spaceWidth;
        }

        double? boundedMaxWidth = TryGetBoundedMaxWidth(maxWidth, out double bounded) ? bounded : (double?)null;
        double anchorWidth = alignment switch {
            PdfTabAlignment.Center => followingTextWidth / 2D,
            PdfTabAlignment.Right => followingTextWidth,
            PdfTabAlignment.DecimalSeparator => MeasureDecimalAnchorWidth(followingText, followingFont, fontSize, baseline, options),
            _ => followingTextWidth
        };
        double nextStop = GetTabTargetPosition(lineWidth, tabStopWidth, explicitTabStopPosition);
        if (boundedMaxWidth.HasValue) {
            nextStop = Math.Min(nextStop, boundedMaxWidth.Value);
        }

        double advance = nextStop - anchorWidth - lineWidth;
        if (advance < spaceWidth) {
            if (boundedMaxWidth.HasValue && nextStop >= boundedMaxWidth.Value) {
                return Math.Max(0D, advance);
            }

            double stopsToAdd = Math.Ceiling((spaceWidth - advance) / tabStopWidth);
            nextStop += Math.Max(1D, stopsToAdd) * tabStopWidth;
            if (boundedMaxWidth.HasValue) {
                nextStop = Math.Min(nextStop, boundedMaxWidth.Value);
            }

            advance = nextStop - anchorWidth - lineWidth;
            if (boundedMaxWidth.HasValue && nextStop >= boundedMaxWidth.Value) {
                return Math.Max(0D, advance);
            }
        }

        return Math.Max(spaceWidth, advance);
    }

    // A short effective gap can select a leader without painting any glyphs.
    // Keep line metrics on exactly the same decision as the painter.
    private static bool HasPaintedRichTabLeader(RichSeg segment, PdfOptions? options) =>
        segment.LeadingSpace && segment.LeadingIsTab &&
        segment.LeadingTabLeader != PdfTabLeaderStyle.None &&
        BuildTabLeaderText(segment.LeadingAdvance, segment, options ?? new PdfOptions()).Length > 0;
    private const int MaxTabLeaderGlyphCount = 4_096;

    private static string BuildTabLeaderText(double gap, RichSeg segment, PdfOptions options) {
        string leaderGlyph = segment.LeadingTabLeader switch {
            PdfTabLeaderStyle.Dots => ".",
            PdfTabLeaderStyle.Hyphens => "-",
            PdfTabLeaderStyle.Underscores => "_",
            _ => string.Empty
        };

        if (leaderGlyph.Length == 0) {
            return string.Empty;
        }

        return OfficeIMO.Drawing.OfficeTextTabLeaderLayout.Create(leaderGlyph, gap, MaxTabLeaderGlyphCount,
            value => MeasureRichText(value, segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                options, segment.FeatureSettings, segment.HorizontalTextScaling, segment.CharacterSpacing,
                segment.TextDirection, segment.FontMetricScale), System.Threading.CancellationToken.None,
            minimumGapGlyphs: 3).Text ?? string.Empty;
    }
}
