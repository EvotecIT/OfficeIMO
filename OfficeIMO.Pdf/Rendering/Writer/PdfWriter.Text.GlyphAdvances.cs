namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Returns the actual encoded glyph advance, including authored and line-distribution tracking.</summary>
    private static double GetRichGlyphAdvance(RichSeg segment, string text, PdfTextShowCommand command,
        int index, PdfCreateOptions? options) {
        double advance;
        if (command.VisualGlyphs is { } glyphs) advance = glyphs[index].AdvanceWidth1000;
        else if (command.GlyphHex.Length == command.GlyphCount * 2) advance = GetSimpleGlyphAdvance1000(segment, text, command, index);
        else {
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
        return ScaleRichGlyphAdvance(segment, command, index, advance);
    }

    private static double GetSimpleGlyphAdvance1000(RichSeg segment, string text, PdfTextShowCommand command, int index) {
        double advance = segment.Font is PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique
            ? 600D : PdfStandardFontWidths.TryGetWidth1000(segment.Font, text[index], out int width) ? width : 0D;
        return advance + (command.SimpleFontKerning?[index] ?? 0);
    }

    private static double ScaleRichGlyphAdvance(RichSeg segment, PdfTextShowCommand command, int index, double advance) {
        double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
        double tracking = command.TrackingBoundaries?[index] == true ? GetIntrinsicGlyphTrackingAdvance(command, size) : 0D;
        return (advance * size / 1000D + tracking) * segment.HorizontalTextScaling / 100D
            + segment.CharacterSpacing + command.GetJustificationAdvance(index);
    }

    /// <summary>
    /// Reads the advances of one segment's glyphs in order. Each embedded glyph would otherwise look its
    /// font program up again; the program belongs to the segment, so it is resolved once, when first needed.
    /// </summary>
    private sealed class RichGlyphAdvanceReader {
        private readonly RichSeg _segment;
        private readonly string _text;
        private readonly PdfTextShowCommand _command;
        private readonly PdfCreateOptions? _options;
        private PdfTrueTypeFontProgram? _trueType;
        private PdfOpenTypeCffFontProgram? _cff;

        internal RichGlyphAdvanceReader(RichSeg segment, string text, PdfTextShowCommand command, PdfCreateOptions? options) {
            _segment = segment;
            _text = text;
            _command = command;
            _options = options;
        }

        internal double Get(int index) {
            if (_command.VisualGlyphs != null || _command.GlyphHex.Length == _command.GlyphCount * 2)
                return GetRichGlyphAdvance(_segment, _text, _command, index, _options);
            if (_trueType == null && _cff == null) ResolveProgram();
            int glyphId = GetGlyphPaintId(_command.GlyphHex, index * 4);
            double advance = _trueType != null ? _trueType.GetGlyphWidth1000(glyphId) : _cff!.GetGlyphWidth1000(glyphId);
            return ScaleRichGlyphAdvance(_segment, _command, index, advance);
        }

        private void ResolveProgram() {
            PdfCreateOptions? options = _options;
            RichSeg segment = _segment;
            if (options != null && segment.NamedFont.HasValue && options.TryGetNamedFontProgram(segment.NamedFont.Value, out var named) && named != null)
                _trueType = named;
            else if (options != null && segment.NamedFont.HasValue && options.TryGetNamedOpenTypeCffFontProgram(segment.NamedFont.Value, out var namedCff) && namedCff != null)
                _cff = namedCff;
            else if (options != null && options.TryGetEmbeddedStandardFontProgram(segment.Font, out var standard) && standard != null)
                _trueType = standard;
            else if (options != null && options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out var standardCff) && standardCff != null)
                _cff = standardCff;
            else throw new InvalidOperationException("The selected embedded glyph has no PDF advance metrics.");
        }
    }

    /// <summary>Uses the paint command's word boundaries and advances for words-only underline geometry.</summary>
    private static void VisitRichWordDecorationAdvances(RichSeg segment, PdfTextShowCommand command,
        PdfCreateOptions options, Action<double, double> visit) {
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
