namespace OfficeIMO.Pdf;

internal sealed class PdfGlyphRun {
    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs)
        : this(glyphs, Array.Empty<PdfTextEncodingDiagnostic>()) {
    }

    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs, IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) {
        Glyphs = glyphs ?? throw new ArgumentNullException(nameof(glyphs));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    public IReadOnlyList<PdfGlyphInfo> Glyphs { get; }
    public IReadOnlyList<PdfTextEncodingDiagnostic> Diagnostics { get; }
    public bool HasMissingGlyphs => Diagnostics.Count > 0;
    public int TotalAdvanceWidth1000 => Glyphs.Sum(glyph => glyph.AdvanceWidth1000);

    public string ToGlyphHex() {
        var sb = new StringBuilder(Glyphs.Count * 4);
        foreach (PdfGlyphInfo glyph in Glyphs) {
            sb.Append(glyph.GlyphId.ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }
}

internal readonly struct PdfGlyphInfo {
    public PdfGlyphInfo(int glyphId, int unicodeScalar, int textIndex, int advanceWidth1000) {
        GlyphId = glyphId;
        UnicodeScalar = unicodeScalar;
        TextIndex = textIndex;
        AdvanceWidth1000 = advanceWidth1000;
    }

    public int GlyphId { get; }
    public int UnicodeScalar { get; }
    public int TextIndex { get; }
    public int AdvanceWidth1000 { get; }
}

internal readonly struct PdfTextShapingOptions {
    public PdfTextShapingOptions(bool recordGlyphUsage, bool throwOnMissingGlyph, bool skipLayoutControls, bool reportControlCharacters, string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        RecordGlyphUsage = recordGlyphUsage;
        ThrowOnMissingGlyph = throwOnMissingGlyph;
        SkipLayoutControls = skipLayoutControls;
        ReportControlCharacters = reportControlCharacters;
        Source = source ?? string.Empty;
        FontName = fontName ?? string.Empty;
        ShapingMode = shapingMode;
    }

    public bool RecordGlyphUsage { get; }
    public bool ThrowOnMissingGlyph { get; }
    public bool SkipLayoutControls { get; }
    public bool ReportControlCharacters { get; }
    public string Source { get; }
    public string FontName { get; }
    public PdfTextShapingMode ShapingMode { get; }

    public static PdfTextShapingOptions ForRendering(string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) =>
        new PdfTextShapingOptions(recordGlyphUsage: true, throwOnMissingGlyph: true, skipLayoutControls: false, reportControlCharacters: false, source: string.Empty, fontName: fontName, shapingMode: shapingMode);

    public static PdfTextShapingOptions ForDiagnostics(string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) =>
        new PdfTextShapingOptions(recordGlyphUsage: false, throwOnMissingGlyph: false, skipLayoutControls: true, reportControlCharacters: true, source: source, fontName: fontName, shapingMode: shapingMode);
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
                PdfLatinLigatureSubstitution.TryGetPresentationLigature(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
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

internal static class PdfLatinLigatureSubstitution {
    internal static bool TryGetPresentationLigature(string text, int index, out int ligatureScalar, out int utf16Length) {
        if (index < 0 || index >= text.Length) {
            ligatureScalar = 0;
            utf16Length = 0;
            return false;
        }

        if (StartsWithOrdinal(text, index, "ffi")) {
            ligatureScalar = 0xFB03;
            utf16Length = 3;
            return true;
        }

        if (StartsWithOrdinal(text, index, "ffl")) {
            ligatureScalar = 0xFB04;
            utf16Length = 3;
            return true;
        }

        if (StartsWithOrdinal(text, index, "ff")) {
            ligatureScalar = 0xFB00;
            utf16Length = 2;
            return true;
        }

        if (StartsWithOrdinal(text, index, "fi")) {
            ligatureScalar = 0xFB01;
            utf16Length = 2;
            return true;
        }

        if (StartsWithOrdinal(text, index, "fl")) {
            ligatureScalar = 0xFB02;
            utf16Length = 2;
            return true;
        }

        ligatureScalar = 0;
        utf16Length = 0;
        return false;
    }

    private static bool StartsWithOrdinal(string text, int index, string value) =>
        index <= text.Length - value.Length &&
        string.Compare(text, index, value, 0, value.Length, StringComparison.Ordinal) == 0;
}
