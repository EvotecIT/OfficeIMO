using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    private readonly StringBuilder _sb;

    public ContentStreamBuilder(StringBuilder sb) {
        Guard.NotNull(sb, nameof(sb));
        _sb = sb;
    }

    public ContentStreamBuilder SaveState() {
        _textStates.Push((_textScale, _textLeading, _textWordSpacing, _textCharacterSpacing, _syntheticOblique));
        InvalidateTextStateCache();
        _sb.Append("q\n");
        return this;
    }

    public ContentStreamBuilder RestoreState() {
        if (_textStates.Count != 0) {
            var state = _textStates.Pop();
            _textScale = state.Scale; _textLeading = state.Leading; _textWordSpacing = state.WordSpacing;
            _textCharacterSpacing = state.CharacterSpacing;
            _syntheticOblique = state.SyntheticOblique;
        }
        InvalidateTextStateCache();
        _sb.Append("Q\n");
        return this;
    }

    public ContentStreamBuilder GraphicsState(string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (resourceName[0] == '/') {
            _sb.Append(resourceName);
        } else {
            _sb.Append('/').Append(resourceName);
        }

        // An ExtGState can select a font.
        InvalidateTextStateCache();
        _sb.Append(" gs\n");
        return this;
    }

    public ContentStreamBuilder FillColor(PdfColor color) {
        if (SyncTextState() && _fillColor.HasValue && _fillColor.Value.Equals(color)) return this;
        _sb.Append(F(color.R)).Append(' ').Append(F(color.G)).Append(' ').Append(F(color.B)).Append(" rg\n");
        if (_inText) {
            _fillColor = color;
            OwnWrite();
        }
        return this;
    }

    public ContentStreamBuilder FillGray(double gray) {
        _fillColor = null;
        _sb.Append(F(gray)).Append(" g\n");
        return this;
    }

    public ContentStreamBuilder StrokeColor(PdfColor color) {
        _sb.Append(F(color.R)).Append(' ').Append(F(color.G)).Append(' ').Append(F(color.B)).Append(" RG\n");
        return this;
    }

    public ContentStreamBuilder LineWidth(double width, bool preciseCoordinates = false) {
        if (preciseCoordinates) {
            _sb.Append(PdfNumberFormatter.Precise(width)).Append(" w\n");
            return this;
        }
        _sb.Append(F(width)).Append(" w\n");
        return this;
    }

    public ContentStreamBuilder LineCap(int lineCap) {
        _sb.Append(lineCap.ToString(CultureInfo.InvariantCulture)).Append(" J\n");
        return this;
    }

    public ContentStreamBuilder LineJoin(int lineJoin) {
        _sb.Append(lineJoin.ToString(CultureInfo.InvariantCulture)).Append(" j\n");
        return this;
    }

    public ContentStreamBuilder MiterLimit(double miterLimit) {
        _sb.Append(F(miterLimit)).Append(" M\n");
        return this;
    }

    public ContentStreamBuilder StrokeDash(params double[] pattern) {
        return StrokeDash(pattern, 0D);
    }

    public ContentStreamBuilder StrokeDash(IReadOnlyList<double> pattern, double phase) {
        Guard.NotNull(pattern, nameof(pattern));
        _sb.Append('[');
        for (int i = 0; i < pattern.Count; i++) {
            if (i > 0) {
                _sb.Append(' ');
            }

            _sb.Append(F(pattern[i]));
        }

        _sb.Append("] ").Append(F(phase)).Append(" d\n");
        return this;
    }

    public ContentStreamBuilder Rectangle(double x, double y, double width, double height, bool preciseCoordinates = false) {
        if (preciseCoordinates) {
            _sb.Append(PdfNumberFormatter.Precise(x)).Append(' ').Append(PdfNumberFormatter.Precise(y)).Append(' ')
                .Append(PdfNumberFormatter.Precise(width)).Append(' ').Append(PdfNumberFormatter.Precise(height)).Append(" re");
            return this;
        }
        _sb.Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(width)).Append(' ').Append(F(height)).Append(" re");
        return this;
    }

    public ContentStreamBuilder FillPath(OfficeFillRule fillRule = OfficeFillRule.NonZero) {
        _sb.Append(fillRule == OfficeFillRule.EvenOdd ? " f*\n" : " f\n");
        return this;
    }

    public ContentStreamBuilder FillStrokePath(OfficeFillRule fillRule = OfficeFillRule.NonZero) {
        _sb.Append(fillRule == OfficeFillRule.EvenOdd ? " B*\n" : " B\n");
        return this;
    }

    public ContentStreamBuilder StrokePath() {
        _sb.Append(" S\n");
        return this;
    }

    public ContentStreamBuilder MoveTo(double x, double y, bool preciseCoordinates = false) {
        if (preciseCoordinates) {
            _sb.Append(PdfNumberFormatter.Precise(x)).Append(' ').Append(PdfNumberFormatter.Precise(y)).Append(" m");
            return this;
        }
        _sb.Append(F(x)).Append(' ').Append(F(y)).Append(" m");
        return this;
    }

    public ContentStreamBuilder PathSeparator() {
        _sb.Append('\n');
        return this;
    }

    public ContentStreamBuilder LineTo(double x, double y, bool preciseCoordinates = false) {
        if (preciseCoordinates) {
            _sb.Append(' ').Append(PdfNumberFormatter.Precise(x)).Append(' ').Append(PdfNumberFormatter.Precise(y)).Append(" l");
            return this;
        }
        _sb.Append(' ').Append(F(x)).Append(' ').Append(F(y)).Append(" l");
        return this;
    }

    public ContentStreamBuilder CubicTo(double x1, double y1, double x2, double y2, double x3, double y3) {
        _sb.Append(' ')
            .Append(F(x1)).Append(' ').Append(F(y1)).Append(' ')
            .Append(F(x2)).Append(' ').Append(F(y2)).Append(' ')
            .Append(F(x3)).Append(' ').Append(F(y3)).Append(" c");
        return this;
    }

    public ContentStreamBuilder ClosePath() {
        _sb.Append(" h");
        return this;
    }

    public ContentStreamBuilder EndPath() {
        _sb.Append(" n\n");
        return this;
    }

    public ContentStreamBuilder ClipPath() {
        _sb.Append(" W");
        return this;
    }

    public ContentStreamBuilder ClipPath(OfficeFillRule fillRule) {
        _sb.Append(fillRule == OfficeFillRule.EvenOdd ? " W*" : " W");
        return this;
    }

    public ContentStreamBuilder TransformMatrix(double a, double b, double c, double d, double e, double f, bool preciseCoordinates = false) {
        if (preciseCoordinates) {
            _sb.Append(PdfNumberFormatter.Precise(a)).Append(' ').Append(PdfNumberFormatter.Precise(b)).Append(' ')
                .Append(PdfNumberFormatter.Precise(c)).Append(' ').Append(PdfNumberFormatter.Precise(d)).Append(' ')
                .Append(PdfNumberFormatter.Precise(e)).Append(' ').Append(PdfNumberFormatter.Precise(f)).Append(" cm\n");
            return this;
        }
        _sb.Append(MatrixNumber(a)).Append(' ')
            .Append(MatrixNumber(b)).Append(' ')
            .Append(MatrixNumber(c)).Append(' ')
            .Append(MatrixNumber(d)).Append(' ')
            .Append(MatrixNumber(e)).Append(' ')
            .Append(MatrixNumber(f)).Append(" cm\n");
        return this;
    }

    public ContentStreamBuilder TransformMatrix(OfficeTransform transform) =>
        TransformMatrix(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX, transform.OffsetY);

    public ContentStreamBuilder XObject(string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (resourceName[0] == '/') {
            _sb.Append(resourceName);
        } else {
            _sb.Append('/').Append(resourceName);
        }

        _sb.Append(" Do\n");
        return this;
    }

    public ContentStreamBuilder Shading(string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (resourceName[0] == '/') {
            _sb.Append(resourceName);
        } else {
            _sb.Append('/').Append(resourceName);
        }

        _sb.Append(" sh\n");
        return this;
    }

    public ContentStreamBuilder BeginText() {
        ResetTrackedTextMatrix();
        InvalidateTextStateCache();
        _showEnd = -1;
        _sb.Append("BT\n");
        _inText = true;
        _ownedLength = _sb.Length;
        return this;
    }

    public ContentStreamBuilder EndText() {
        _inText = false;
        _showEnd = -1;
        InvalidateTextStateCache();
        _sb.Append("ET\n");
        return this;
    }

    public ContentStreamBuilder Font(string resourceName, double size, bool syntheticOblique = false, bool preserveLogicalPrecision = false) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (SyncTextState() && _syntheticOblique == syntheticOblique && _fontSize == size &&
            _fontPrecise == preserveLogicalPrecision && string.Equals(_fontResource, resourceName, StringComparison.Ordinal)) return this;
        _sb.Append('/').Append(resourceName).Append(' ');
        if (preserveLogicalPrecision) _sb.Append(MatrixNumber(size));
        else _sb.Append(F(size));
        _sb.Append(" Tf\n");
        if (_inText) {
            _fontResource = resourceName;
            _fontSize = size;
            _fontPrecise = preserveLogicalPrecision;
            OwnWrite();
        }
        if (_syntheticOblique != syntheticOblique) {
            double previousShear = _syntheticOblique ? SyntheticObliqueShear : 0D;
            _syntheticOblique = syntheticOblique;
            if (_hasTextMatrix) {
                double nextShear = syntheticOblique ? SyntheticObliqueShear : 0D;
                TextMatrixApplied(_textA, _textB,
                    _textC + (nextShear - previousShear) * _textA,
                    _textD + (nextShear - previousShear) * _textB,
                    _textE, _textF);
            }
        }
        return this;
    }

    public ContentStreamBuilder TextLeading(double leading) {
        _textLeading = leading;
        SyncTextState();
        _sb.Append(F(leading)).Append(" TL\n");
        OwnWrite();
        return this;
    }

    public ContentStreamBuilder TextMatrix(double x, double y) {
        return TextMatrix(1, 0, 0, 1, x, y);
    }

    public ContentStreamBuilder TextMatrix(double a, double b, double c, double d, double e, double f) {
        double shear = _syntheticOblique ? SyntheticObliqueShear : 0D;
        return TextMatrixApplied(a, b, c + shear * a, d + shear * b, e, f);
    }

    private ContentStreamBuilder TextMatrixApplied(double a, double b, double c, double d, double e, double f) {
        if (SyncTextState() && TryFoldTextMatrix(a, b, c, d, e, f)) return this;
        // Track the serialized matrix: later TJ moves are relative to what viewers read.
        TrackTextMatrix(Serialized(a), Serialized(b), Serialized(c), Serialized(d), Serialized(e), Serialized(f));
        _sb.Append(F(a)).Append(' ')
            .Append(F(b)).Append(' ')
            .Append(F(c)).Append(' ')
            .Append(F(d)).Append(' ')
            .Append(F(e)).Append(' ')
            .Append(F(f)).Append(" Tm\n");
        OwnWrite();
        return this;
    }

    public ContentStreamBuilder MoveText(double x, double y) {
        double shear = _syntheticOblique ? SyntheticObliqueShear : 0D;
        double xAdvance = _textA * x + (_textC - shear * _textA) * y;
        double yAdvance = _textB * x + (_textD - shear * _textB) * y;
        if (_isolatedText || _syntheticOblique) return TextMatrixApplied(_textA, _textB, _textC, _textD, _lineE + xAdvance, _lineF + yAdvance);
        _lineE += xAdvance; _lineF += yAdvance;
        _textE = _lineE; _textF = _lineF;
        SyncTextState();
        _sb.Append(F(x)).Append(' ').Append(F(y)).Append(" Td\n");
        OwnWrite();
        return this;
    }

    public ContentStreamBuilder NextTextLine() {
        if (_isolatedText || _syntheticOblique) return MoveText(0, -_textLeading);
        _lineE -= _textC * _textLeading; _lineF -= _textD * _textLeading;
        _textE = _lineE; _textF = _lineF;
        SyncTextState();
        _sb.Append("T*\n");
        OwnWrite();
        return this;
    }

    public ContentStreamBuilder WordSpacing(double spacing) {
        if (SyncTextState() && _wordSpacingKnown && _textWordSpacing == spacing) return this;
        _textWordSpacing = spacing;
        _sb.Append(F(spacing)).Append(" Tw\n");
        if (_inText) {
            _wordSpacingKnown = true;
            OwnWrite();
        }
        return this;
    }

    public ContentStreamBuilder CharacterSpacing(double spacing) {
        if (double.IsNaN(spacing) || double.IsInfinity(spacing)) {
            throw new ArgumentOutOfRangeException(nameof(spacing), "PDF character spacing must be finite.");
        }
        if (SyncTextState() && _characterSpacingKnown && _textCharacterSpacing == spacing) return this;
        _textCharacterSpacing = spacing;
        _sb.Append(F(spacing)).Append(" Tc\n");
        if (_inText) {
            _characterSpacingKnown = true;
            OwnWrite();
        }
        return this;
    }

    public ContentStreamBuilder HorizontalTextScaling(double percentage) {
        if (percentage <= 0D || double.IsNaN(percentage) || double.IsInfinity(percentage)) {
            throw new ArgumentOutOfRangeException(nameof(percentage), "PDF horizontal text scaling must be positive and finite.");
        }

        _textScale = percentage / 100D;
        SyncTextState();
        _sb.Append(F(percentage)).Append(" Tz\n");
        OwnWrite();
        return this;
    }

    public ContentStreamBuilder TextRise(double rise) {
        SyncTextState();
        _sb.Append(F(rise)).Append(" Ts\n");
        OwnWrite();
        return this;
    }

    public ContentStreamBuilder TextRenderingMode(int mode) {
        if (mode < 0 || mode > 7) {
            throw new ArgumentOutOfRangeException(nameof(mode), "PDF text rendering mode must be between 0 and 7.");
        }

        SyncTextState();
        _sb.Append(mode.ToString(CultureInfo.InvariantCulture)).Append(" Tr\n");
        OwnWrite();
        return this;
    }

    /// <summary>Ends text-show coalescing so the next text matrix is written explicitly (tab stops, leaders).</summary>
    public ContentStreamBuilder BreakTextRun() {
        _showEnd = -1;
        return this;
    }

    public ContentStreamBuilder ShowHexText(string hexText) {
        Guard.NotNull(hexText, nameof(hexText));
        // The caller owns positioning for raw show strings.
        _showEnd = -1;
        _positionKnown = false;
        _sb.Append('<').Append(hexText).Append("> Tj\n");
        return this;
    }

    public ContentStreamBuilder ShowText(PdfTextShowCommand command, double fontSize, double currentTextRise = 0D, bool suppressActualText = false) {
        Guard.NotNull(command, nameof(command));
        if (fontSize <= 0 || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
            throw new ArgumentOutOfRangeException(nameof(fontSize), "PDF text font size must be positive and finite.");
        }

        SyncTextState();
        if (!suppressActualText && command.LogicalGlyphs is { } logicalGlyphs) {
            if (command.IsolatedRightToLeft) _showEnd = -1;
            WriteLogicalGlyphs(logicalGlyphs, command, fontSize, currentTextRise);
            if (command.IsolatedRightToLeft) _showEnd = -1;
            return this;
        }

        bool marked = !suppressActualText && command.ActualText != null;
        if (marked) {
            _sb.Append("/Span << /ActualText ")
                .Append(PdfSyntaxEscaper.TextString(command.ActualText!))
                .Append(" >> BDC\n");
        }

        if (!command.HasPositioning) {
            OpenShow(fontSize);
            if (command.SimpleFontKerning is { } kerning) {
                string hex = command.GlyphHex;
                for (int index = 0; index < kerning.Length && index * 2 + 1 < hex.Length; index++) {
                    ShowHexElement(hex.Substring(index * 2, 2));
                    if (kerning[index] != 0) ShowNumberElement(-kerning[index]);
                }
            } else {
                ShowHexElement(command.GlyphHex);
            }
            CloseShow();
        } else {
            double tracking1000 = (command.Tracking?.GetAdjustment(fontSize / command.FontMetricScale) ?? 0D) * 1000D / command.UnitsPerEm;
            AppendPositionedGlyphs(command.PositionedGlyphs!, fontSize, currentTextRise,
                command.NegativeTracking ? -tracking1000 : tracking1000, command.TrackingBoundaries);
        }

        if (marked) {
            _sb.Append("EMC\n");
            OwnWrite();
        }

        if (command.AdvanceWidth1000.HasValue) {
            double trackingAdvance = (command.Tracking?.GetAdjustment(fontSize / command.FontMetricScale) ?? 0D)
                * fontSize / command.UnitsPerEm * (command.TrackingBoundaries?.Count(boundary => boundary) ?? 0);
            AdvanceTrackedText(command.AdvanceWidth1000.Value * fontSize / 1000D
                + (command.NegativeTracking ? -trackingAdvance : trackingAdvance) + command.WordSpaceCount * _textWordSpacing + command.GlyphCount * _textCharacterSpacing);
        } else {
            _positionKnown = false;
        }
        return this;
    }

    private void AppendPositionedGlyphs(IReadOnlyList<PdfGlyphInfo> glyphs, double fontSize, double baseTextRise,
        double tracking1000 = 0D, bool[]? trackingBoundaries = null) {
        int currentOffsetY1000 = 0;
        bool open = false;
        for (int index = 0; index < glyphs.Count; index++) {
            PdfGlyphInfo glyph = glyphs[index];
            if (glyph.OffsetY1000 != currentOffsetY1000) {
                if (open) {
                    CloseShow();
                    open = false;
                }
                _sb.Append(F(baseTextRise + glyph.OffsetY1000 * fontSize / 1000D)).Append(" Ts\n");
                currentOffsetY1000 = glyph.OffsetY1000;
            }
            if (!open) {
                OpenShow(fontSize);
                open = true;
            }

            if (glyph.OffsetX1000 != 0) ShowNumberElement(-glyph.OffsetX1000);
            ShowGlyphElement(glyph.GlyphId);
            ShowNumberElement(glyph.OffsetX1000 + glyph.NominalWidth1000 - glyph.AdvanceWidth1000
                - (trackingBoundaries?[index] == true ? tracking1000 : 0D));
        }

        if (open) CloseShow();
        if (currentOffsetY1000 != 0) {
            _sb.Append(F(baseTextRise)).Append(" Ts\n");
            OwnWrite();
        }
    }

    private static double Serialized(double value) {
        if (Math.Abs(value) < 0.0005D) return 0D;
        double rounded = Math.Round(value, 3, MidpointRounding.AwayFromZero);
        return double.IsNaN(rounded) || double.IsInfinity(rounded) ? value : rounded;
    }

#if NET6_0_OR_GREATER
    [ThreadStatic] private static char[]? _numberBuffer;

    // Matrix coefficients multiply page coordinates; three decimals can move glyphs by tenths of a point.
    private static ReadOnlySpan<char> MatrixNumber(double value) {
        value = Math.Abs(value) < 0.0000005D ? 0D : value;
        return FormatNumber(value, "0.######");
    }

    private static ReadOnlySpan<char> F(double value) {
        if (Math.Abs(value) < 0.0005D) {
            value = 0D;
        }

        // Content streams repeat the same few numbers (colour components, font sizes, line widths, column
        // edges), and the custom "0.###" format is slow, so each thread memoises recent results in a small
        // direct-mapped table. A hit returns the exact characters the format produced for that value. The
        // clamp above is what keeps -0 (which == 0 but formats as "-0") out of the table; NaN never matches.
        double[] keys = _fKeys ??= new double[FCacheSlots];
        byte[] lengths = _fLengths ??= new byte[FCacheSlots];
        char[] chars = _fChars ??= new char[FCacheSlots * FSlotChars];
        long bits = BitConverter.DoubleToInt64Bits(value);
        int slot = (int)((ulong)(bits * unchecked((long)0x9E3779B97F4A7C15UL)) >> (64 - FCacheBits));
        int at = slot * FSlotChars;
        if (lengths[slot] != 0 && keys[slot] == value) {
            return chars.AsSpan(at, lengths[slot]);
        }

        if (!value.TryFormat(chars.AsSpan(at, FSlotChars), out int written, "0.###", CultureInfo.InvariantCulture)) {
            lengths[slot] = 0;
            return value.ToString("0.###", CultureInfo.InvariantCulture).AsSpan();
        }

        keys[slot] = value;
        lengths[slot] = (byte)written;
        return chars.AsSpan(at, written);
    }

    private const int FCacheBits = 10, FCacheSlots = 1 << FCacheBits, FSlotChars = 24;
    [ThreadStatic] private static double[]? _fKeys;
    [ThreadStatic] private static byte[]? _fLengths;
    [ThreadStatic] private static char[]? _fChars;

    // Formats into a reused per-thread buffer and returns a span. Every call site is _sb.Append(F(x)),
    // which copies the span immediately, so the buffer can be overwritten by the next call; this avoids
    // a throwaway string per formatted number. Span formatting / StringBuilder.Append(ReadOnlySpan)
    // are not available on netstandard2.0/net472, which fall back to ToString below.
    private static ReadOnlySpan<char> FormatNumber(double value, string format) {
        char[] buffer = _numberBuffer ??= new char[32];
        return value.TryFormat(buffer, out int written, format, CultureInfo.InvariantCulture)
            ? buffer.AsSpan(0, written)
            : value.ToString(format, CultureInfo.InvariantCulture).AsSpan();
    }
#else
    // Matrix coefficients multiply page coordinates; three decimals can move glyphs by tenths of a point.
    private static string MatrixNumber(double value) =>
        (Math.Abs(value) < 0.0000005D ? 0D : value).ToString("0.######", CultureInfo.InvariantCulture);

    private static string F(double value) {
        if (Math.Abs(value) < 0.0005D) {
            value = 0D;
        }

        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
#endif
}
