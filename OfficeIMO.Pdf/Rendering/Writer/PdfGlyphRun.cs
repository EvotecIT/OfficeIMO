using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfGlyphRun {
    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs)
        : this(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), actualText: null) {
    }

    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs, IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics, string? actualText = null) {
        Glyphs = glyphs ?? throw new ArgumentNullException(nameof(glyphs));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        ActualText = string.IsNullOrEmpty(actualText) ? null : actualText;
    }

    public IReadOnlyList<PdfGlyphInfo> Glyphs { get; }
    public IReadOnlyList<PdfTextEncodingDiagnostic> Diagnostics { get; }
    public string? ActualText { get; }
    public bool HasMissingGlyphs => Diagnostics.Count > 0;
    public int TotalAdvanceWidth1000 => Glyphs.Sum(glyph => glyph.AdvanceWidth1000);
    public bool HasPositioning => Glyphs.Any(glyph => glyph.HasPositioning);

    public string ToGlyphHex() {
        var sb = new StringBuilder(Glyphs.Count * 4);
        foreach (PdfGlyphInfo glyph in Glyphs) {
            sb.Append(glyph.GlyphId.ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    public PdfTextShowCommand ToTextShowCommand() =>
        new(ToGlyphHex(), HasPositioning ? Glyphs : null, ActualText);
}

internal sealed class PdfTextShowCommand {
    internal PdfTextShowCommand(string glyphHex, IReadOnlyList<PdfGlyphInfo>? positionedGlyphs = null, string? actualText = null) {
        GlyphHex = glyphHex ?? throw new ArgumentNullException(nameof(glyphHex));
        PositionedGlyphs = positionedGlyphs;
        ActualText = string.IsNullOrEmpty(actualText) ? null : actualText;
    }

    internal string GlyphHex { get; }
    internal IReadOnlyList<PdfGlyphInfo>? PositionedGlyphs { get; }
    internal string? ActualText { get; }
    internal bool HasPositioning => PositionedGlyphs != null && PositionedGlyphs.Count > 0;
}

internal readonly struct PdfGlyphInfo {
    public PdfGlyphInfo(int glyphId, int unicodeScalar, int textIndex, int advanceWidth1000)
        : this(glyphId, char.ConvertFromUtf32(unicodeScalar), unicodeScalar, textIndex, advanceWidth1000, advanceWidth1000, 0, 0) {
    }

    public PdfGlyphInfo(int glyphId, string unicodeText, int textIndex, int advanceWidth1000)
        : this(glyphId, unicodeText, unicodeText != null && unicodeText.Length > 0 ? char.ConvertToUtf32(unicodeText, 0) : 0, textIndex, advanceWidth1000, advanceWidth1000, 0, 0) {
    }

    public PdfGlyphInfo(int glyphId, string unicodeText, int textIndex, int nominalWidth1000, int advanceWidth1000, int offsetX1000, int offsetY1000)
        : this(glyphId, unicodeText, unicodeText != null && unicodeText.Length > 0 ? char.ConvertToUtf32(unicodeText, 0) : 0, textIndex, nominalWidth1000, advanceWidth1000, offsetX1000, offsetY1000) {
    }

    private PdfGlyphInfo(int glyphId, string unicodeText, int unicodeScalar, int textIndex, int nominalWidth1000, int advanceWidth1000, int offsetX1000, int offsetY1000) {
        GlyphId = glyphId;
        UnicodeText = unicodeText ?? string.Empty;
        UnicodeScalar = unicodeScalar;
        TextIndex = textIndex;
        NominalWidth1000 = nominalWidth1000;
        AdvanceWidth1000 = advanceWidth1000;
        OffsetX1000 = offsetX1000;
        OffsetY1000 = offsetY1000;
    }

    public int GlyphId { get; }
    public string UnicodeText { get; }
    public int UnicodeScalar { get; }
    public int TextIndex { get; }
    public int NominalWidth1000 { get; }
    public int AdvanceWidth1000 { get; }
    public int OffsetX1000 { get; }
    public int OffsetY1000 { get; }
    public bool HasPositioning => AdvanceWidth1000 != NominalWidth1000 || OffsetX1000 != 0 || OffsetY1000 != 0;
}

internal readonly struct PdfTextShapingOptions {
    public PdfTextShapingOptions(bool recordGlyphUsage, bool throwOnMissingGlyph, bool skipLayoutControls, bool reportControlCharacters, string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, Action<string, string, bool>? providerShapedTextRecorder = null, string? language = null) {
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

    public static PdfTextShapingOptions ForRendering(string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, Action<string, string, bool>? providerShapedTextRecorder = null, string? language = null) =>
        new PdfTextShapingOptions(recordGlyphUsage: true, throwOnMissingGlyph: true, skipLayoutControls: false, reportControlCharacters: false, source: string.Empty, fontName: fontName, shapingMode: shapingMode, shapingProvider: shapingProvider, providerShapedTextRecorder: providerShapedTextRecorder, language: language);

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

        return new PdfGlyphRun(glyphs, diagnostics);
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
