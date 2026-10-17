using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfGlyphRun {
    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs)
        : this(glyphs, Array.Empty<PdfTextEncodingDiagnostic>(), actualText: null, OfficeTextDirection.Auto) {
    }

    public PdfGlyphRun(IReadOnlyList<PdfGlyphInfo> glyphs, IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics, string? actualText = null, OfficeTextDirection direction = OfficeTextDirection.Auto, bool hasCompleteVerticalAdvances = false, OfficeTextShapingResult? sourceShapingResult = null, bool preserveGlyphUnicode = false, bool isAutomaticallyShaped = false, bool isScalarBidi = false) {
        Glyphs = glyphs ?? throw new ArgumentNullException(nameof(glyphs));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        ActualText = string.IsNullOrEmpty(actualText) ? null : actualText;
        Direction = direction;
        HasCompleteVerticalAdvances = hasCompleteVerticalAdvances;
        SourceShapingResult = sourceShapingResult;
        PreserveGlyphUnicode = preserveGlyphUnicode;
        IsAutomaticallyShaped = isAutomaticallyShaped;
        IsScalarBidi = isScalarBidi;
    }

    internal bool PreserveGlyphUnicode { get; }
    internal bool IsAutomaticallyShaped { get; }
    internal bool IsScalarBidi { get; }
    /// <summary>Default pair kerning in font design units, aligned to <see cref="Glyphs"/>, when no provider result exists.</summary>
    internal int[]? DesignAdvanceAdjustments { get; set; }
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

    public string ToGlyphHex() => ToGlyphHex(Glyphs);

    internal static string ToGlyphHex(IReadOnlyList<PdfGlyphInfo> glyphs) {
        var sb = RentHexBuilder(glyphs.Count * 4);
        for (int i = 0; i < glyphs.Count; i++) {
            AppendGlyphHex(sb, glyphs[i].GlyphId);
        }

        return ReturnHexBuilder(sb);
    }

    public PdfTextShowCommand ToTextShowCommand() {
        if (Direction == OfficeTextDirection.TopToBottom) {
            throw new InvalidOperationException("PDF horizontal text operators cannot publish a top-to-bottom shaped glyph run. Use the diagnosed vertical drawing route.");
        }
        return new PdfTextShowCommand(this);
    }
}

internal sealed class PdfTextShowCommand {
    private string? _glyphHex;

    /// <summary>Imported layout advance, including the requested horizontal scale, before PDF rounding.</summary>
    internal double? FontUnitLayoutAdvanceWidth1000 { get; private set; }

    /// <summary>
    /// Right-to-left glyphs in visual order, mapped through ToUnicode only. They get their own show operator so
    /// readers see a pure right-to-left run they can reorder, rather than a mixed-direction TJ array.
    /// </summary>
    internal bool IsolatedRightToLeft { get; }

    internal PdfTextShowCommand WithFontUnitLayoutAdvance(double advanceWidth1000) {
        var copy = (PdfTextShowCommand)MemberwiseClone();
        copy.FontUnitLayoutAdvanceWidth1000 = advanceWidth1000;
        return copy;
    }

    internal PdfTextShowCommand(PdfGlyphRun run) {
        PositionedGlyphs = run.HasPositioning ? run.Glyphs : null;
        ActualText = run.ActualText;
        IsScalarBidi = run.IsScalarBidi;
        IsolatedRightToLeft = run.Direction == OfficeTextDirection.RightToLeft && run.ActualText == null;
        LogicalGlyphs = run.PreserveGlyphUnicode ? run.Glyphs : null;
        AdvanceWidth1000 = run.TotalAdvanceWidth1000;
        VisualGlyphs = run.Glyphs;
        GlyphCount = run.Glyphs.Count;
        UnitsPerEm = 1000;
        FontMetricScale = 1D;
        // Logical glyphs are emitted directly. Materialize hex only if a caller
        // requests the ordinary show-string (including suppressed ActualText).
        _glyphHex = run.PreserveGlyphUnicode ? null : run.ToGlyphHex();
    }

    internal PdfTextShowCommand(string glyphHex, IReadOnlyList<PdfGlyphInfo>? positionedGlyphs = null, string? actualText = null, IReadOnlyList<PdfGlyphInfo>? logicalGlyphs = null, double? advanceWidth1000 = null, int wordSpaceCount = 0, IReadOnlyList<PdfGlyphInfo>? visualGlyphs = null, OfficeOpenTypeTracking? tracking = null, int unitsPerEm = 1000, bool[]? trackingBoundaries = null, bool negativeTracking = false, double fontMetricScale = 1D, int? glyphCount = null, int[]? simpleFontKerning = null, bool isScalarBidi = false) {
        if (fontMetricScale <= 0D || double.IsNaN(fontMetricScale) || double.IsInfinity(fontMetricScale))
            throw new ArgumentOutOfRangeException(nameof(fontMetricScale));
        SimpleFontKerning = simpleFontKerning;
        IsScalarBidi = isScalarBidi;
        LogicalGlyphs = logicalGlyphs; AdvanceWidth1000 = advanceWidth1000; WordSpaceCount = wordSpaceCount;
        VisualGlyphs = visualGlyphs;
        _glyphHex = glyphHex ?? throw new ArgumentNullException(nameof(glyphHex));
        GlyphCount = glyphCount ?? visualGlyphs?.Count ?? positionedGlyphs?.Count ?? logicalGlyphs?.Count ?? glyphHex.Length / 2;
        PositionedGlyphs = positionedGlyphs;
        ActualText = string.IsNullOrEmpty(actualText) ? null : actualText;
        Tracking = tracking;
        UnitsPerEm = unitsPerEm;
        TrackingBoundaries = trackingBoundaries;
        NegativeTracking = negativeTracking;
        FontMetricScale = fontMetricScale;
    }

    /// <summary>Scalar bidi output uses isolated directional shows to avoid reader double-reordering.</summary>
    internal bool IsScalarBidi { get; }
    internal IReadOnlyList<PdfGlyphInfo>? LogicalGlyphs { get; }
    /// <summary>Single-byte font pair kerning after each code, in 1/1000 em (negative tightens).</summary>
    internal int[]? SimpleFontKerning { get; }
    /// <summary>Decides which logical glyphs can rely on ToUnicode instead of ActualText.</summary>
    internal IPdfGlyphTextPinner? Pinner { get; private set; }
    internal IReadOnlyList<PdfGlyphInfo>? VisualGlyphs { get; }

    internal PdfTextShowCommand WithPinner(IPdfGlyphTextPinner? pinner) {
        Pinner = pinner;
        return this;
    }
    internal double? AdvanceWidth1000 { get; }
    internal int WordSpaceCount { get; }
    internal int GlyphCount { get; }
    internal string GlyphHex => _glyphHex ??= PdfGlyphRun.ToGlyphHex(VisualGlyphs!);
    internal IReadOnlyList<PdfGlyphInfo>? PositionedGlyphs { get; }
    internal string? ActualText { get; }
    internal OfficeOpenTypeTracking? Tracking { get; }
    internal int UnitsPerEm { get; }
    internal bool[]? TrackingBoundaries { get; }
    internal bool NegativeTracking { get; }
    internal double FontMetricScale { get; }
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

    public PdfGlyphInfo(int glyphId, string unicodeText, int textIndex, int nominalWidth1000, int advanceWidth1000, int advanceHeight1000, int offsetX1000, int offsetY1000, int? logicalClusterStart = null)
        : this(glyphId, unicodeText, unicodeText != null && unicodeText.Length > 0 ? char.ConvertToUtf32(unicodeText, 0) : 0, textIndex, nominalWidth1000, advanceWidth1000, advanceHeight1000, offsetX1000, offsetY1000, logicalClusterStart) { }

    private PdfGlyphInfo(int glyphId, string unicodeText, int unicodeScalar, int textIndex, int nominalWidth1000, int advanceWidth1000, int advanceHeight1000, int offsetX1000, int offsetY1000, int? logicalClusterStart = null) {
        GlyphId = glyphId;
        UnicodeText = unicodeText ?? string.Empty;
        UnicodeScalar = unicodeScalar;
        TextIndex = textIndex;
        LogicalClusterStart = logicalClusterStart ?? textIndex;
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
    internal int LogicalClusterStart { get; }
    public int NominalWidth1000 { get; }
    public int AdvanceWidth1000 { get; }
    public int AdvanceHeight1000 { get; }
    public int OffsetX1000 { get; }
    public int OffsetY1000 { get; }
    public bool HasPositioning => AdvanceWidth1000 != NominalWidth1000 || AdvanceHeight1000 != 0 || OffsetX1000 != 0 || OffsetY1000 != 0;
}

internal readonly struct PdfTextShapingOptions {
    public PdfTextShapingOptions(bool recordGlyphUsage, bool throwOnMissingGlyph, bool skipLayoutControls, bool reportControlCharacters, string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, Action<string, string, bool, bool>? providerShapedTextRecorder = null, string? language = null, OfficeTextFeatureSettings? featureSettings = null, OfficeTextDirection direction = OfficeTextDirection.Auto, bool preservesPaintedGlyphs = false) {
        RecordGlyphUsage = recordGlyphUsage;
        ThrowOnMissingGlyph = throwOnMissingGlyph;
        SkipLayoutControls = skipLayoutControls;
        ReportControlCharacters = reportControlCharacters;
        Source = source ?? string.Empty;
        FontName = fontName ?? string.Empty;
        PreservesPaintedGlyphs = preservesPaintedGlyphs;
        ShapingMode = preservesPaintedGlyphs ? PdfTextShapingMode.UnicodeScalar : shapingMode;
        ShapingProvider = preservesPaintedGlyphs ? null : shapingProvider;
        ProviderShapedTextRecorder = preservesPaintedGlyphs ? null : providerShapedTextRecorder;
        Language = string.IsNullOrWhiteSpace(language) ? null : language;
        FeatureSettings = preservesPaintedGlyphs ? OfficeTextFeatureSettings.Default : featureSettings ?? OfficeTextFeatureSettings.Default;
        Direction = preservesPaintedGlyphs ? OfficeTextDirection.LeftToRight : direction;
    }

    public bool RecordGlyphUsage { get; }
    public bool ThrowOnMissingGlyph { get; }
    public bool SkipLayoutControls { get; }
    public bool ReportControlCharacters { get; }
    public string Source { get; }
    public string FontName { get; }
    public PdfTextShapingMode ShapingMode { get; }
    public IOfficeTextShapingProvider? ShapingProvider { get; }
    public Action<string, string, bool, bool>? ProviderShapedTextRecorder { get; }
    public string? Language { get; }
    public OfficeTextFeatureSettings FeatureSettings { get; }
    public OfficeTextDirection Direction { get; }
    public bool PreservesPaintedGlyphs { get; }

    public static PdfTextShapingOptions ForRendering(string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar, IOfficeTextShapingProvider? shapingProvider = null, Action<string, string, bool, bool>? providerShapedTextRecorder = null, string? language = null, OfficeTextFeatureSettings? featureSettings = null, OfficeTextDirection direction = OfficeTextDirection.Auto, bool preservesPaintedGlyphs = false) =>
        new PdfTextShapingOptions(recordGlyphUsage: true, throwOnMissingGlyph: true, skipLayoutControls: false, reportControlCharacters: false, source: string.Empty, fontName: fontName, shapingMode: shapingMode, shapingProvider: shapingProvider, providerShapedTextRecorder: providerShapedTextRecorder, language: language, featureSettings: featureSettings, direction: direction, preservesPaintedGlyphs: preservesPaintedGlyphs);

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

                glyphs.Add(new PdfGlyphInfo(ligatureGlyphId, text.Substring(scalarStart, ligatureLength), scalarStart, font.GetGlyphWidth1000(ligatureGlyphId)));
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

        // Without shaping, right-to-left text is still painted in UAX #9 visual order (joining forms
        // and mirrored glyphs need a shaping provider). A run that is an exact reversal maps through
        // ToUnicode, which readers reorder; other reordered runs keep the logical string as ActualText.
        bool reordered = !options.PreservesPaintedGlyphs && OfficeUnicodeBidi.ContainsRightToLeft(text) && glyphs.Count > 1;
        if (reordered) {
            glyphs = ToVisualOrder(text, glyphs);
            if (PdfExternalTextShaper.IsReversedLogicalRun(text, glyphs)) {
                return new PdfGlyphRun(glyphs, diagnostics, actualText: null, OfficeTextDirection.RightToLeft, preserveGlyphUnicode: true);
            }
        }
        string? actualText = !options.PreservesPaintedGlyphs && (reordered || OfficeTextElements.ResolveBaseDirection(text) == OfficeTextDirection.RightToLeft)
            ? text
            : null;
        return new PdfGlyphRun(glyphs, diagnostics, actualText, preserveGlyphUnicode:
            options.PreservesPaintedGlyphs || reordered || options.ShapingMode == PdfTextShapingMode.OpenTypeLigatures && options.ShapingProvider == null && !OfficeManagedTextShaper.RequiresComplexLayout(text),
            isScalarBidi: reordered);
    }

    internal static List<PdfGlyphInfo> ToVisualOrder(string text, List<PdfGlyphInfo> glyphs) {
        int[] glyphLevels = ResolveGlyphLevels(text, glyphs);
        int[] order = OfficeUnicodeBidi.ReorderItems(glyphLevels);
        var visual = new List<PdfGlyphInfo>(glyphs.Count);
        foreach (int index in order) visual.Add(glyphs[index]);
        return visual;
    }

    internal static int[] ResolveGlyphLevels(string text, IReadOnlyList<PdfGlyphInfo> glyphs) {
        int[] scalars = OfficeUnicodeBidi.ToScalars(text, out int[] starts);
        byte[] levels = OfficeUnicodeBidi.Resolve(scalars, paragraphLevel: null).GetLineLevels(0, scalars.Length);
        var glyphLevels = new int[glyphs.Count];
        for (int index = 0; index < glyphs.Count; index++) {
            int scalarIndex = Array.BinarySearch(starts, glyphs[index].TextIndex);
            glyphLevels[index] = levels[scalarIndex >= 0 ? scalarIndex : Math.Max(0, ~scalarIndex - 1)];
        }
        return glyphLevels;
    }

    // Total advance width (1000-em units) without materializing a glyph run. Mirrors ShapeText's loop
    // (same ligature/scalar handling, same usage recording, same missing-glyph throw) but skips the
    // glyph/diagnostic list allocation, so widths and font subsetting are unchanged. Used by the
    // line-break measurement path, which only needs the width. ForRendering never reports control
    // characters, so that (list-producing) branch is not part of the measurement contract.
    public static int MeasureAdvanceWidth1000(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, Action<int, string>? observeUsage = null) {
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
                observeUsage?.Invoke(ligatureGlyphId, text.Substring(scalarStart, ligatureLength));
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
            observeUsage?.Invoke(glyphId, char.ConvertFromUtf32(scalar));
        }

        return totalWidth;
    }

    // Emits the glyph hex show-string directly, skipping the per-run List<PdfGlyphInfo>. Scalar shaping
    // never positions glyphs, so the emission path needs only this hex plus ActualText. Mirrors
    // MeasureAdvanceWidth1000's loop (same usage recording, same missing-glyph throw); ForRendering never
    // reports control characters, so that branch is not part of the contract and the hex equals
    // ToGlyphHex over the same shaped glyphs.
    public static string EncodeGlyphHex(string text, PdfTrueTypeFontProgram font, PdfTextShapingOptions options, out string? actualText, out int advanceWidth1000) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));

        var sb = PdfGlyphRun.RentHexBuilder(text.Length * 4);
        advanceWidth1000 = 0;
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
                advanceWidth1000 = checked(advanceWidth1000 + font.GetGlyphWidth1000(ligatureGlyphId));
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
            advanceWidth1000 = checked(advanceWidth1000 + font.GetGlyphWidth1000(glyphId));
        }

        actualText = !options.PreservesPaintedGlyphs && OfficeTextElements.ResolveBaseDirection(text) == OfficeTextDirection.RightToLeft ? text : null;
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
