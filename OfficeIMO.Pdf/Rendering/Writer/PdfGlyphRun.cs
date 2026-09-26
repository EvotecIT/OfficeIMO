using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfGlyphRun {
    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs)
        : this(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), actualText: null, OfficeTextDirection.Auto) {
    }

    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs, IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics, string? actualText = null, OfficeTextDirection direction = OfficeTextDirection.Auto, bool hasCompleteVerticalAdvances = false, OfficeTextShapingResult? sourceShapingResult = null, bool preserveGlyphUnicode = false) {
        Glyphs = glyphs ?? throw new ArgumentNullException(nameof(glyphs));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        ActualText = string.IsNullOrEmpty(actualText) ? null : actualText;
        Direction = direction;
        HasCompleteVerticalAdvances = hasCompleteVerticalAdvances;
        SourceShapingResult = sourceShapingResult;
        PreserveGlyphUnicode = preserveGlyphUnicode;
    }

    internal bool PreserveGlyphUnicode { get; }
    public IReadOnlyList<PdfGlyphInfo> Glyphs { get; }
    public IReadOnlyList<PdfTextEncodingDiagnostic> Diagnostics { get; }
    public string? ActualText { get; }
    public OfficeTextDirection Direction { get; }
    public bool HasCompleteVerticalAdvances { get; }
    internal OfficeTextShapingResult? SourceShapingResult { get; }
    public bool HasMissingGlyphs => Diagnostics.Count > 0;
    public int TotalAdvanceWidth1000 {
        get {
            int total = 0;
            for (int i = 0; i < Glyphs.Count; i++) {
                total = checked(total + Glyphs[i].AdvanceWidth1000);
            }
            return total;
        }
    }
    public bool HasPositioning {
        get {
            for (int i = 0; i < Glyphs.Count; i++) if (Glyphs[i].HasPositioning) return true;
            return false;
        }
    }

    private const string HexChars = "0123456789ABCDEF";

    // Glyph-hex show-strings are built once per drawn run; a fresh StringBuilder (and its char[] backing)
    // per run is a top allocator. Reuse one per thread. Detached while in use so nested/reentrant callers
    // fall back to a fresh instance, and oversized buffers are dropped rather than retained.
    [ThreadStatic] private static StringBuilder? _hexBuilder;

    internal static StringBuilder RentHexBuilder(int capacityHint) {
        StringBuilder sb = _hexBuilder ?? new StringBuilder(256);
        _hexBuilder = null;
        sb.Clear();
        if (sb.Capacity < capacityHint && capacityHint <= 8192) sb.EnsureCapacity(capacityHint);
        return sb;
    }

    internal static string ReturnHexBuilder(StringBuilder sb) {
        string result = sb.ToString();
        if (sb.Capacity <= 8192) {
            sb.Clear();
            _hexBuilder = sb;
        }

        return result;
    }

    // GlyphId is a 16-bit TrueType/CFF index, so four nibbles match ToString("X4") without the
    // per-glyph string allocation.
    internal static void AppendGlyphHex(StringBuilder sb, int glyphId) {
        sb.Append(HexChars[(glyphId >> 12) & 0xF]);
        sb.Append(HexChars[(glyphId >> 8) & 0xF]);
        sb.Append(HexChars[(glyphId >> 4) & 0xF]);
        sb.Append(HexChars[glyphId & 0xF]);
    }

    public string ToGlyphHex() {
        var sb = RentHexBuilder(Glyphs.Count * 4);
        for (int i = 0; i < Glyphs.Count; i++) {
            AppendGlyphHex(sb, Glyphs[i].GlyphId);
        }

        return ReturnHexBuilder(sb);
    }

    public PdfTextShowCommand ToTextShowCommand() {
        if (Direction == OfficeTextDirection.TopToBottom) {
            throw new InvalidOperationException("PDF horizontal text operators cannot publish a top-to-bottom shaped glyph run. Use the diagnosed vertical drawing route.");
        }
        return new PdfTextShowCommand(ToGlyphHex(), HasPositioning ? Glyphs : null, ActualText,
            PreserveGlyphUnicode ? Glyphs : null, TotalAdvanceWidth1000);
    }
}

internal sealed class PdfTextShowCommand {
    internal PdfTextShowCommand(string glyphHex, IReadOnlyList<PdfGlyphInfo>? positionedGlyphs = null, string? actualText = null, IReadOnlyList<PdfGlyphInfo>? logicalGlyphs = null, double? advanceWidth1000 = null, int wordSpaceCount = 0) {
        LogicalGlyphs = logicalGlyphs; AdvanceWidth1000 = advanceWidth1000; WordSpaceCount = wordSpaceCount;
        GlyphHex = glyphHex ?? throw new ArgumentNullException(nameof(glyphHex));
        PositionedGlyphs = positionedGlyphs;
        ActualText = string.IsNullOrEmpty(actualText) ? null : actualText;
    }

    internal IReadOnlyList<PdfGlyphInfo>? LogicalGlyphs { get; }
    internal double? AdvanceWidth1000 { get; }
    internal int WordSpaceCount { get; }
    internal string GlyphHex { get; }
    internal IReadOnlyList<PdfGlyphInfo>? PositionedGlyphs { get; }
    internal string? ActualText { get; }
    internal bool HasPositioning => PositionedGlyphs != null && PositionedGlyphs.Count > 0;
}

internal readonly struct PdfGlyphInfo {
    public PdfGlyphInfo(int glyphId, int unicodeScalar, int textIndex, int advanceWidth1000)
        : this(glyphId, char.ConvertFromUtf32(unicodeScalar), unicodeScalar, textIndex, advanceWidth1000, advanceWidth1000, 0, 0, 0) {
    }

    public PdfGlyphInfo(int glyphId, string unicodeText, int textIndex, int advanceWidth1000)
        : this(glyphId, unicodeText, unicodeText != null && unicodeText.Length > 0 ? char.ConvertToUtf32(unicodeText, 0) : 0, textIndex, advanceWidth1000, advanceWidth1000, 0, 0, 0) {
    }

    public PdfGlyphInfo(int glyphId, string unicodeText, int textIndex, int nominalWidth1000, int advanceWidth1000, int offsetX1000, int offsetY1000)
        : this(glyphId, unicodeText, unicodeText != null && unicodeText.Length > 0 ? char.ConvertToUtf32(unicodeText, 0) : 0, textIndex, nominalWidth1000, advanceWidth1000, 0, offsetX1000, offsetY1000) {
    }

    public PdfGlyphInfo(int glyphId, string unicodeText, int textIndex, int nominalWidth1000, int advanceWidth1000, int advanceHeight1000, int offsetX1000, int offsetY1000)
        : this(glyphId, unicodeText, unicodeText != null && unicodeText.Length > 0 ? char.ConvertToUtf32(unicodeText, 0) : 0, textIndex, nominalWidth1000, advanceWidth1000, advanceHeight1000, offsetX1000, offsetY1000) { }

    private PdfGlyphInfo(int glyphId, string unicodeText, int unicodeScalar, int textIndex, int nominalWidth1000, int advanceWidth1000, int advanceHeight1000, int offsetX1000, int offsetY1000) {
        GlyphId = glyphId;
        UnicodeText = unicodeText ?? string.Empty;
        UnicodeScalar = unicodeScalar;
        TextIndex = textIndex;
        NominalWidth1000 = nominalWidth1000;
        AdvanceWidth1000 = advanceWidth1000;
        AdvanceHeight1000 = advanceHeight1000;
        OffsetX1000 = offsetX1000;
        OffsetY1000 = offsetY1000;
    }

    public int GlyphId { get; }
    public string UnicodeText { get; }
    public int UnicodeScalar { get; }
    public int TextIndex { get; }
    public int NominalWidth1000 { get; }
    public int AdvanceWidth1000 { get; }
    public int AdvanceHeight1000 { get; }
    public int OffsetX1000 { get; }
    public int OffsetY1000 { get; }
    public bool HasPositioning => AdvanceWidth1000 != NominalWidth1000 || AdvanceHeight1000 != 0 || OffsetX1000 != 0 || OffsetY1000 != 0;
}

internal readonly struct PdfTextShapingOptions {
    public PdfTextShapingOptions(bool recordGlyphUsage, bool throwOnMissingGlyph, bool skipLayoutControls, bool reportControlCharacters, string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, Action<string, string, bool>? providerShapedTextRecorder = null, string? language = null, OfficeTextFeatureSettings? featureSettings = null, OfficeTextDirection direction = OfficeTextDirection.Auto) {
        RecordGlyphUsage = recordGlyphUsage;
        ThrowOnMissingGlyph = throwOnMissingGlyph;
        SkipLayoutControls = skipLayoutControls;
        ReportControlCharacters = reportControlCharacters;
        Source = source ?? string.Empty;
        FontName = fontName ?? string.Empty;
        ShapingMode = shapingMode;
        ShapingProvider = shapingProvider;
        ProviderShapedTextRecorder = providerShapedTextRecorder;
        Language = string.IsNullOrWhiteSpace(language) ? null : language;
        FeatureSettings = featureSettings ?? OfficeTextFeatureSettings.Default;
        Direction = direction;
    }

    public bool RecordGlyphUsage { get; }
    public bool ThrowOnMissingGlyph { get; }
    public bool SkipLayoutControls { get; }
    public bool ReportControlCharacters { get; }
    public string Source { get; }
    public string FontName { get; }
    public PdfTextShapingMode ShapingMode { get; }
    public IOfficeTextShapingProvider? ShapingProvider { get; }
    public Action<string, string, bool>? ProviderShapedTextRecorder { get; }
    public string? Language { get; }
    public OfficeTextFeatureSettings FeatureSettings { get; }
    public OfficeTextDirection Direction { get; }

    public static PdfTextShapingOptions ForRendering(string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, Action<string, string, bool>? providerShapedTextRecorder = null, string? language = null, OfficeTextFeatureSettings? featureSettings = null, OfficeTextDirection direction = OfficeTextDirection.Auto) =>
        new PdfTextShapingOptions(recordGlyphUsage: true, throwOnMissingGlyph: true, skipLayoutControls: false, reportControlCharacters: false, source: string.Empty, fontName: fontName, shapingMode: shapingMode, shapingProvider: shapingProvider, providerShapedTextRecorder: providerShapedTextRecorder, language: language, featureSettings: featureSettings, direction: direction);

    public static PdfTextShapingOptions ForDiagnostics(string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null) =>
        new PdfTextShapingOptions(recordGlyphUsage: false, throwOnMissingGlyph: false, skipLayoutControls: true, reportControlCharacters: true, source: source, fontName: fontName, shapingMode: shapingMode, shapingProvider: shapingProvider);
}

internal interface IPdfTextShaper {
    PdfGlyphRun ShapeText(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options);
}

internal sealed class PdfUnicodeScalarTextShaper : IPdfTextShaper {
    public static PdfUnicodeScalarTextShaper Instance { get; } = new PdfUnicodeScalarTextShaper();

    private PdfUnicodeScalarTextShaper() {
    }

    public PdfGlyphRun ShapeText(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        var glyphs = new List<PdfGlyphInfo>();
        var diagnostics = new List<PdfTextEncodingDiagnostic>();
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            if (options.ShapingMode == PdfTextShapingMode.LatinLigatures &&
                OfficeTextLigatures.TryGetLatinPresentationForm(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                font.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                if (options.RecordGlyphUsage) {
                    font.RecordGlyphUsage(ligatureGlyphId, text.Substring(scalarStart, ligatureLength));
                }

                glyphs.Add(new PdfGlyphInfo(ligatureGlyphId, ligatureScalar, scalarStart, font.GetGlyphWidth1000(ligatureGlyphId)));
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (options.SkipLayoutControls && (scalar == '\n' || scalar == '\r' || scalar == '\t')) {
                continue;
            }

            if (options.ReportControlCharacters && (scalar < ' ' || scalar == '\u007F')) {
                diagnostics.Add(PdfTextDiagnostics.CreateControlCharacterDiagnostic(scalarStart, scalar, options.Source));
                continue;
            }

            if (!font.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                diagnostics.Add(PdfTextDiagnostics.CreateEmbeddedFontDiagnostic(scalarStart, scalar, options.Source, ResolveFontName(font, options)));
                if (options.ThrowOnMissingGlyph) {
                    throw PdfTrueTypeFontProgram.CreateUnsupportedGlyphException(text, scalarStart, scalar);
                }

                continue;
            }

            if (options.RecordGlyphUsage) {
                font.RecordGlyphUsage(glyphId, scalar);
            }

            glyphs.Add(new PdfGlyphInfo(glyphId, scalar, scalarStart, font.GetGlyphWidth1000(glyphId)));
        }

        // The scalar fallback does not perform bidirectional layout. Preserve the
        // caller's logical string as ActualText so readers do not have to guess
        // whether the emitted glyph order is logical or visual.
        string? actualText = OfficeTextElements.ResolveBaseDirection(text) == OfficeTextDirection.RightToLeft
            ? text
            : null;
        return new PdfGlyphRun(glyphs, diagnostics, actualText, preserveGlyphUnicode:
            options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures && options.ShapingProvider == null && !OfficeManagedTextShaper.RequiresComplexLayout(text));
    }

    // Total advance width (1000-em units) without materializing a glyph run. Mirrors ShapeText's loop
    // (same ligature/scalar handling, same usage recording, same missing-glyph throw) but skips the
    // glyph/diagnostic list allocation, so widths and font subsetting are unchanged. Used by the
    // line-break measurement path, which only needs the width. ForRendering never reports control
    // characters, so that (list-producing) branch is not part of the measurement contract.
    public static int MeasureAdvanceWidth1000(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        int totalWidth = 0;
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            if (options.ShapingMode == PdfTextShapingMode.LatinLigatures &&
                OfficeTextLigatures.TryGetLatinPresentationForm(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                font.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                if (options.RecordGlyphUsage) {
                    font.RecordGlyphUsage(ligatureGlyphId, text.Substring(scalarStart, ligatureLength));
                }

                totalWidth = checked(totalWidth + font.GetGlyphWidth1000(ligatureGlyphId));
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (options.SkipLayoutControls && (scalar == '\n' || scalar == '\r' || scalar == '\t')) {
                continue;
            }

            if (!font.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                if (options.ThrowOnMissingGlyph) {
                    throw PdfTrueTypeFontProgram.CreateUnsupportedGlyphException(text, scalarStart, scalar);
                }

                continue;
            }

            if (options.RecordGlyphUsage) {
                font.RecordGlyphUsage(glyphId, scalar);
            }

            totalWidth = checked(totalWidth + font.GetGlyphWidth1000(glyphId));
        }

        return totalWidth;
    }

    // Emits the glyph hex show-string directly, skipping the per-run List<PdfGlyphInfo>. Scalar shaping
    // never positions glyphs, so the emission path needs only this hex plus ActualText. Mirrors
    // MeasureAdvanceWidth1000's loop (same usage recording, same missing-glyph throw); ForRendering never
    // reports control characters, so that branch is not part of the contract and the hex equals
    // ToGlyphHex over the same shaped glyphs.
    public static string EncodeGlyphHex(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, out string? actualText) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        var sb = PdfGlyphRun.RentHexBuilder(text.Length * 4);
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            if (options.ShapingMode == PdfTextShapingMode.LatinLigatures &&
                OfficeTextLigatures.TryGetLatinPresentationForm(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                font.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                if (options.RecordGlyphUsage) {
                    font.RecordGlyphUsage(ligatureGlyphId, text.Substring(scalarStart, ligatureLength));
                }

                PdfGlyphRun.AppendGlyphHex(sb, ligatureGlyphId);
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (options.SkipLayoutControls && (scalar == '\n' || scalar == '\r' || scalar == '\t')) {
                continue;
            }

            if (!font.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                if (options.ThrowOnMissingGlyph) {
                    throw PdfTrueTypeFontProgram.CreateUnsupportedGlyphException(text, scalarStart, scalar);
                }

                continue;
            }

            if (options.RecordGlyphUsage) {
                font.RecordGlyphUsage(glyphId, scalar);
            }

            PdfGlyphRun.AppendGlyphHex(sb, glyphId);
        }

        actualText = OfficeTextElements.ResolveBaseDirection(text) == OfficeTextDirection.RightToLeft ? text : null;
        return PdfGlyphRun.ReturnHexBuilder(sb);
    }


    private static string ResolveFontName(PdfTrueTypeFontProgram font, PdfTextShapingOptions options) =>
        string.IsNullOrWhiteSpace(options.FontName) ? font.FontName : options.FontName;

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch)) {
            if (index < text.Length && char.IsLowSurrogate(text[index])) {
                return char.ConvertToUtf32(ch, text[index++]);
            }

            throw new ArgumentException("Text contains an unmatched high surrogate at index " + (index - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(text));
        }

        if (char.IsLowSurrogate(ch)) {
            throw new ArgumentException("Text contains an unmatched low surrogate at index " + (index - 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(text));
        }

        return ch;
    }
}
