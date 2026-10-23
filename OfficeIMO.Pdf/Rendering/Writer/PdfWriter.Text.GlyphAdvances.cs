namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Returns the actual encoded glyph advance, including authored and line-distribution tracking.</summary>
    private static double GetRichGlyphAdvance(RichSeg segment, string text, PdfTextShowCommand command,
        int index, PdfOptions? options) {
        double advance;
        if (command.VisualGlyphs is { } glyphs) advance = glyphs[index].AdvanceWidth1000;
        else if (command.GlyphHex.Length == command.GlyphCount * 2) {
            advance = segment.Font is PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique
                ? 600D : PdfStandardFontWidths.TryGetWidth1000(segment.Font, text[index], out int width) ? width : 0D;
            advance += command.SimpleFontKerning?[index] ?? 0;
        } else {
            int glyphId = GetGlyphPaintId(command.GlyphHex, index * 4);
            if (options != null && segment.NamedFont.HasValue && options.TryGetNamedFontProgram(segment.NamedFont.Value, out var named) && named != null)
                advance = named.GetGlyphWidth1000(glyphId);
            else if (options != null && segment.NamedFont.HasValue && options.TryGetNamedOpenTypeCffFontProgram(segment.NamedFont.Value, out var namedCff) && namedCff != null)
                advance = namedCff.GetGlyphWidth1000(glyphId);
            else if (options != null && options.TryGetEmbeddedStandardFontProgram(segment.Font, out var standard) && standard != null)
                advance = standard.GetGlyphWidth1000(glyphId);
            else if (options != null && options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out var standardCff) && standardCff != null)
                advance = standardCff.GetGlyphWidth1000(glyphId);
            else throw new InvalidOperationException("The selected embedded glyph has no PDF advance metrics.");
        }
        double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
        double tracking = command.TrackingBoundaries?[index] == true ? GetIntrinsicGlyphTrackingAdvance(command, size) : 0D;
        return (advance * size / 1000D + tracking) * segment.HorizontalTextScaling / 100D
            + segment.CharacterSpacing + command.GetJustificationAdvance(index);
    }

    /// <summary>Uses the paint command's word boundaries and advances for words-only underline geometry.</summary>
    private static void VisitRichWordDecorationAdvances(RichSeg segment, PdfTextShowCommand command,
        PdfOptions options, Action<double, double> visit) {
        string paintedText = WithoutOptionalBreaks(segment.Text);
        if (command.VisualGlyphs is not { Count: > 0 } && !command.HasJustificationAdvances) {
            VisitWordDecorationAdvances(paintedText,
                span => MeasurePositionedTextWidth(span, segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                    options, segment.FeatureSettings, segment.TextDirection, segment.HorizontalTextScaling,
                    segment.CharacterSpacing, segment.FontMetricScale, segment.PreservesPaintedGlyphs), visit);
            return;
        }
        double advance = 0D;
        double? wordStart = null;
        for (int index = 0; index < command.GlyphCount; index++) {
            int sourceIndex = command.VisualGlyphs?[index].TextIndex ?? index;
            bool whitespace = sourceIndex >= 0 && sourceIndex < paintedText.Length && char.IsWhiteSpace(paintedText[sourceIndex]);
            if (whitespace) {
                if (wordStart.HasValue) {
                    visit(wordStart.Value, advance);
                    wordStart = null;
                }
            } else if (!wordStart.HasValue) wordStart = advance;
            advance += GetRichGlyphAdvance(segment, paintedText, command, index, options);
        }
        if (wordStart.HasValue) visit(wordStart.Value, advance);
    }
}
