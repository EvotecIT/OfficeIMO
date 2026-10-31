namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool AlignsCompleteTabField(PdfTabAlignment alignment) =>
        alignment == PdfTabAlignment.Center || alignment == PdfTabAlignment.Right || alignment == PdfTabAlignment.DecimalSeparator;

    private static RichTabFieldMeasurement MeasureRichTabField(IReadOnlyList<PdfTextRun> runs, int runIndex, int textOffset,
        PdfStandardFont baseFont, double fontSize, PdfCreateOptions? options, PdfTextWhitespaceMode whitespaceMode) {
        var field = new List<PdfTextRun>();
        for (int index = runIndex; index < runs.Count; index++) {
            PdfTextRun run = runs[index];
            if (run.InlineElement != null) {
                field.Add(run);
                continue;
            }

            string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            int start = index == runIndex ? textOffset : 0;
            int end = start;
            while (end < text.Length && text[end] != '\t' && text[end] != '\n') end++;
            if (end > start) field.Add(CreateStyledTextRun(text.Substring(start, end - start), run, run.Font));
            if (end < text.Length) break;
        }

        // The slice contains no tabs or hard breaks, so this cannot recurse into
        // another field measurement. Reuse the actual token/space/kerning layout
        // rather than measuring a differently normalized copy of its text.
        var wrapped = WrapRichRunsCoreWithFirstLineOrigin(field, double.MaxValue / 4D, fontSize, baseFont,
            fontSize * 1.2D, null, null, DefaultParagraphTabStopWidth, options, textWhitespaceMode: whitespaceMode,
            purpose: RichWrapPurpose.IntrinsicMeasurement);
        return new RichTabFieldMeasurement(wrapped.Lines.Count == 0 ? Array.Empty<RichSeg>() : wrapped.Lines[0], 0);
    }

    private static RichTabFieldMeasurement MeasureWrappedTabField(IReadOnlyList<RichSeg> line, int start) =>
        new(line, start);

    private static double? ResolveRichDecimalTabAnchor(RichTabFieldMeasurement? field, PdfTabAlignment alignment,
        PdfTabStop? stop, PdfCreateOptions? options, double leadingAdvance = 0D) =>
        field != null && (stop?.Alignment ?? alignment) == PdfTabAlignment.DecimalSeparator
            ? leadingAdvance + field.GetDecimalAnchorWidth(stop?.DecimalSeparator, options) : null;

    /// <summary>Retains the actual styled field layout for its selected tab anchor and later frame refits.</summary>
    private sealed class RichTabFieldMeasurement {
        private readonly IReadOnlyList<RichSeg> _segments;
        private readonly int _start;
        private readonly int _end;

        internal RichTabFieldMeasurement(IReadOnlyList<RichSeg> segments, int start) {
            _segments = segments;
            _start = start;
            int end = start;
            for (; end < segments.Count; end++) {
                RichSeg segment = segments[end];
                if (end > start && segment.LeadingIsTab) break;
                if (segment.LeadingSpace && (end > start || !segment.LeadingIsTab)) Width += segment.LeadingAdvance;
                Width += GetRichSegmentWidth(segment);
            }
            _end = end;
        }

        internal double Width { get; }

        internal double GetDecimalAnchorWidth(char? separator, PdfCreateOptions? options) {
            double advance = 0D;
            for (int index = _start; index < _end; index++) {
                RichSeg segment = _segments[index];
                if (segment.LeadingSpace && (index > _start || !segment.LeadingIsTab)) advance += segment.LeadingAdvance;
                string paintedText = WithoutOptionalBreaks(segment.Text);
                int decimalIndex = segment.InlineElement != null ? -1 : separator.HasValue
                    ? paintedText.IndexOf(separator.Value) : paintedText.IndexOfAny(DecimalTabAnchorChars);
                if (decimalIndex >= 0) return advance + MeasureRichSegmentDecimalAnchor(segment, paintedText, decimalIndex, options);
                advance += GetRichSegmentWidth(segment);
            }
            // A field without the selected separator is right-aligned in full.
            return Width;
        }
    }

    /// <summary>Uses the prepared token's encoded glyph origins, including cross-run kerning and authored tracking.</summary>
    private static double MeasureRichSegmentDecimalAnchor(RichSeg segment, string paintedText, int decimalIndex, PdfCreateOptions? options) {
        PdfTextShowCommand command = segment.KerningCommand ?? EncodeTextShowCommand(segment.Text, segment.Font,
            segment.NamedFont, options, segment.FeatureSettings, segment.TextDirection, segment.FontMetricScale,
            segment.HorizontalTextScaling, segment.PreservesPaintedGlyphs, requireGlyphClusters: true);
        double advance = 0D;
        for (int index = 0; index < command.GlyphCount; index++) {
            int textIndex = command.VisualGlyphs?[index].TextIndex ?? index;
            if (textIndex == decimalIndex) {
                double offset = command.VisualGlyphs?[index].OffsetX1000 ?? 0D;
                return advance + offset * EffectiveRichFontSize(segment.FontSize, segment.Baseline) / 1000D
                    * segment.HorizontalTextScaling / 100D;
            }
            advance += GetRichGlyphAdvance(segment, paintedText, command, index, options);
        }
        // Preserve the existing per-token behavior if an inseparable shaped
        // cluster supplies no distinct source glyph for the selected character.
        return MeasureDecimalAnchorWidth(paintedText, segment.Font, segment.FontSize, segment.Baseline,
            options, segment.NamedFont, segment.FeatureSettings, segment.HorizontalTextScaling,
            segment.CharacterSpacing, paintedText[decimalIndex]);
    }
}
