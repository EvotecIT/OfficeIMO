using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable text preflight helpers for generated PDF output.
/// </summary>
internal static class PdfTextDiagnostics {
    private const string WinAnsiEncodingDescription = "PDF WinAnsiEncoding";
    private const string WinAnsiGlyphRemediation = "Embedded Unicode fonts are required for this text.";
    private const string ControlCharacterEncodingDescription = "PDF text output";
    private const string ControlCharacterRemediation = "Use paragraphs, line breaks, tables, or spacing primitives for layout instead of literal control characters.";
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<PdfEmbeddedFontFallbackCandidate, EmbeddedFontFallbackProgramBox> FallbackProgramCache = new();
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<PdfTrueTypeFontProgram, OpenTypeFontInfoBox> TrueTypeLayoutInfoCache = new();
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<PdfOpenTypeCffFontProgram, OpenTypeFontInfoBox> CffLayoutInfoCache = new();

    /// <summary>
    /// Finds text that cannot be written through the current generated standard-font WinAnsi path.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <returns>Encoding diagnostics in source order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeWinAnsiText(string text, string source = "", string location = "") {
        return AnalyzeWinAnsiTextCore(text, source, location, null);
    }

    private static List<PdfTextEncodingDiagnostic> AnalyzeWinAnsiTextCore(string text, string source, string location, int? runIndex) {
        Guard.NotNull(text, nameof(text));
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        int index = 0;
        while (index < text.Length) {
            char ch = text[index];
            if (ch == '\n' || ch == '\r' || ch == '\t') {
                index++;
                continue;
            }

            if (!PdfWinAnsiEncoding.CanEncode(ch.ToString(), out _)) {
                diagnostics.Add(CreateDiagnostic(text, index, source, location, runIndex));
                if (char.IsHighSurrogate(ch) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                    index += 2;
                    continue;
                }
            }

            index++;
        }

        return diagnostics;
    }

    /// <summary>
    /// Finds text runs that cannot be written through the current generated standard-font WinAnsi path.
    /// Explicit PDF line-break and tab runs are treated as layout controls rather than literal text.
    /// </summary>
    /// <param name="runs">Text runs to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <returns>Encoding diagnostics in run order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeWinAnsiTextRuns(IEnumerable<TextRun> runs, string source = "", string location = "") {
        Guard.NotNull(runs, nameof(runs));
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        int runIndex = 0;
        foreach (TextRun run in runs) {
            if (run == null || IsLayoutControlRun(run)) {
                runIndex++;
                continue;
            }

            diagnostics.AddRange(AnalyzeWinAnsiTextCore(run.Text, source, AppendRunLocation(location, runIndex), runIndex));
            runIndex++;
        }

        return diagnostics;
    }

    /// <summary>
    /// Finds text that cannot be written with the generated PDF text path selected by the supplied options and font.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="options">PDF options that may provide embedded font coverage for the selected generated font.</param>
    /// <param name="font">Generated PDF font slot to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <returns>Encoding diagnostics in source order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeGeneratedText(string text, PdfOptions options, PdfStandardFont font, string source = "", string location = "") {
        return AnalyzeGeneratedTextCore(text, options, font, source, location, null);
    }

    private static List<PdfTextEncodingDiagnostic> AnalyzeGeneratedTextCore(string text, PdfOptions options, PdfStandardFont font, string source, string location, int? runIndex) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(options, nameof(options));
        Guard.StandardFont(font, nameof(font), "Generated PDF text diagnostics require a supported PDF font.");
        PdfTextShapingMode shapingMode = options.TextShapingModeSnapshot;

        PdfEmbeddedFontFallbackSet? fallbackSet = options.EmbeddedFontFallbacksSnapshot;

        if (options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            if (fallbackSet != null) {
                return AnalyzeGeneratedTextWithFallback(
                    text,
                    fallbackSet,
                    source,
                    location,
                    runIndex,
                    shapingMode,
                    (string value, int index, out int length) => TryGetCoveredTextLength(value, index, fontProgram, shapingMode, out length));
            }

            return AnalyzeEmbeddedFontText(text, fontProgram, source, location, runIndex, shapingMode);
        }

        if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            if (fallbackSet != null) {
                return AnalyzeGeneratedTextWithFallback(
                    text,
                    fallbackSet,
                    source,
                    location,
                    runIndex,
                    shapingMode,
                    (string value, int index, out int length) => TryGetCoveredTextLength(value, index, cffFontProgram, shapingMode, out length));
            }

            return AnalyzeEmbeddedFontText(text, cffFontProgram, source, location, runIndex, shapingMode);
        }

        if (fallbackSet != null) {
            return AnalyzeGeneratedTextWithFallback(
                text,
                fallbackSet,
                source,
                location,
                runIndex,
                shapingMode,
                TryGetWinAnsiCoveredTextLength);
        }

        return AnalyzeWinAnsiTextCore(text, source, location, runIndex);
    }

    /// <summary>
    /// Finds text runs that cannot be written with the generated PDF text path selected by the supplied options and font.
    /// Explicit PDF line-break and tab runs are treated as layout controls rather than literal text.
    /// </summary>
    /// <param name="runs">Text runs to inspect.</param>
    /// <param name="options">PDF options that may provide embedded font coverage for the selected generated font.</param>
    /// <param name="font">Generated PDF font slot to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="location">Optional generated document location such as a block, table cell, or canvas item path.</param>
    /// <returns>Encoding diagnostics in run order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeGeneratedTextRuns(IEnumerable<TextRun> runs, PdfOptions options, PdfStandardFont font, string source = "", string location = "") {
        Guard.NotNull(runs, nameof(runs));
        Guard.NotNull(options, nameof(options));
        Guard.StandardFont(font, nameof(font), "Generated PDF text diagnostics require a supported PDF font.");
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        int runIndex = 0;
        foreach (TextRun run in runs) {
            if (run == null || IsLayoutControlRun(run)) {
                runIndex++;
                continue;
            }

            PdfStandardFont runFont = ResolveRunFont(font, run);
            diagnostics.AddRange(AnalyzeGeneratedTextCore(run.Text, options, runFont, source, AppendRunLocation(location, runIndex), runIndex));
            runIndex++;
        }

        return diagnostics;
    }

    /// <summary>
    /// Finds text scalars that are not covered by a caller-supplied embedded TrueType or OpenType/CFF font program.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="trueTypeFont">TrueType or OpenType/CFF font bytes that will be embedded for generated PDF text.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="fontName">Optional display name used in diagnostic messages.</param>
    /// <returns>Missing-glyph diagnostics in source order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontText(string text, byte[] trueTypeFont, string source = "", string? fontName = null) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(trueTypeFont, nameof(trueTypeFont));
        if (IsOpenTypeCffFontData(trueTypeFont)) {
            PdfOpenTypeCffFontProgram cffFont = PdfOpenTypeCffFontProgram.Parse(trueTypeFont, fontName);
            return AnalyzeEmbeddedFontTextCore(text, cffFont, source, string.IsNullOrWhiteSpace(fontName) ? cffFont.FontName : fontName!);
        }

        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(trueTypeFont, fontName);
        return AnalyzeEmbeddedFontTextCore(text, font, source, string.IsNullOrWhiteSpace(fontName) ? font.FontName : fontName!);
    }

    internal static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontText(string text, PdfTrueTypeFontProgram font, string source = "", string? fontName = null) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));
        return AnalyzeEmbeddedFontTextCore(text, font, source, string.IsNullOrWhiteSpace(fontName) ? font.FontName : fontName!);
    }

    internal static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontText(string text, PdfOpenTypeCffFontProgram font, string source = "", string? fontName = null) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));
        return AnalyzeEmbeddedFontTextCore(text, font, source, string.IsNullOrWhiteSpace(fontName) ? font.FontName : fontName!);
    }

    /// <summary>
    /// Finds text-run scalars that are not covered by a caller-supplied embedded TrueType or OpenType/CFF font program.
    /// Explicit PDF line-break and tab runs are treated as layout controls rather than literal text.
    /// </summary>
    /// <param name="runs">Text runs to inspect.</param>
    /// <param name="trueTypeFont">TrueType or OpenType/CFF font bytes that will be embedded for generated PDF text.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="fontName">Optional display name used in diagnostic messages.</param>
    /// <returns>Missing-glyph diagnostics in run order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontTextRuns(IEnumerable<TextRun> runs, byte[] trueTypeFont, string source = "", string? fontName = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.NotNull(trueTypeFont, nameof(trueTypeFont));
        PdfTrueTypeFontProgram? font = null;
        PdfOpenTypeCffFontProgram? cffFont = null;
        string resolvedFontName;
        if (IsOpenTypeCffFontData(trueTypeFont)) {
            cffFont = PdfOpenTypeCffFontProgram.Parse(trueTypeFont, fontName);
            resolvedFontName = string.IsNullOrWhiteSpace(fontName) ? cffFont.FontName : fontName!;
        } else {
            font = PdfTrueTypeFontProgram.Parse(trueTypeFont, fontName);
            resolvedFontName = string.IsNullOrWhiteSpace(fontName) ? font.FontName : fontName!;
        }

        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        foreach (TextRun run in runs) {
            if (run == null || IsLayoutControlRun(run)) {
                continue;
            }

            diagnostics.AddRange(font != null
                ? AnalyzeEmbeddedFontTextCore(run.Text, font, source, resolvedFontName)
                : AnalyzeEmbeddedFontTextCore(run.Text, cffFont!, source, resolvedFontName));
        }

        return diagnostics;
    }

    /// <summary>
    /// Finds text that currently requires shaping, bidirectional layout, mark positioning, or script-specific line breaking beyond OfficeIMO.Pdf's scalar text path.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <returns>Advanced text layout diagnostics in source order, with one diagnostic per detected limitation.</returns>
    public static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, string source = "") {
        Guard.NotNull(text, nameof(text));
        return AnalyzeAdvancedTextLayoutCore(text, source, indexOffset: 0);
    }

    private static List<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayoutCore(string text, string source, int indexOffset) {
        var diagnostics = new List<PdfTextShapingDiagnostic>();
        var reportedCodes = new HashSet<string>(StringComparer.Ordinal);

        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int sourceIndex = scalarStart + indexOffset;
            int scalar = ReadScalar(text, ref index);

            if (OfficeTextElements.IsRightToLeftScalar(scalar)) {
                AddDiagnostic(
                    diagnostics,
                    reportedCodes,
                    source,
                    sourceIndex,
                    scalar,
                    "right-to-left",
                    "unsupported-bidirectional-text-layout",
                    "Text contains right-to-left scalar " + FormatCodePoint(scalar) + " at index " + sourceIndex.ToString(CultureInfo.InvariantCulture) + ". OfficeIMO.Pdf writes generated text through a simplified left-to-right scalar path today; use a shaped/bidirectional source or review the visual output until advanced text layout support is added.");
            }

            if (TryGetComplexScriptName(scalar, out string complexScript)) {
                AddDiagnostic(
                    diagnostics,
                    reportedCodes,
                    source,
                    sourceIndex,
                    scalar,
                    complexScript,
                    "unsupported-complex-script-shaping",
                    "Text contains " + complexScript + " scalar " + FormatCodePoint(scalar) + " at index " + sourceIndex.ToString(CultureInfo.InvariantCulture) + ". OfficeIMO.Pdf does not apply OpenType shaping, contextual forms, or ligature substitution yet, so generated output may be visually simplified even when the font covers the glyph.");
            }

            if (IsCombiningMarkOrJoiner(scalar)) {
                AddDiagnostic(
                    diagnostics,
                    reportedCodes,
                    source,
                    sourceIndex,
                    scalar,
                    "combining-mark-or-joiner",
                    "unsupported-mark-positioning-or-joiner-shaping",
                    "Text contains combining mark or joiner " + FormatCodePoint(scalar) + " at index " + sourceIndex.ToString(CultureInfo.InvariantCulture) + ". OfficeIMO.Pdf does not apply mark positioning, glyph joining, or joiner-driven shaping yet, so generated output may be visually simplified.");
            }

            if (TryGetScriptLineBreakingName(scalar, out string lineBreakingScript)) {
                AddDiagnostic(
                    diagnostics,
                    reportedCodes,
                    source,
                    sourceIndex,
                    scalar,
                    lineBreakingScript,
                    "unsupported-script-specific-line-breaking",
                    "Text contains " + lineBreakingScript + " scalar " + FormatCodePoint(scalar) + " at index " + sourceIndex.ToString(CultureInfo.InvariantCulture) + ". OfficeIMO.Pdf has CJK-style and callback-based break support, but not dictionary/script-specific line breaking for this script yet.");
            }
        }

        return diagnostics;
    }

    /// <summary>
    /// Finds text that currently requires advanced layout support, including font-specific OpenType substitution and positioning features.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="fontData">OpenType or TrueType font bytes used for the generated text.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="fontName">Optional configured font name used in diagnostic messages.</param>
    /// <returns>Advanced text layout diagnostics in source order, with one diagnostic per detected limitation.</returns>
    public static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, byte[] fontData, string source = "", string? fontName = null) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(fontData, nameof(fontData));
        return AnalyzeAdvancedTextLayout(text, fontData, source, fontName, indexOffset: 0);
    }

    internal static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, byte[] fontData, string source, string? fontName, int indexOffset) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(fontData, nameof(fontData));
        PdfOpenTypeFontInspector.TryInspect(fontData, out PdfOpenTypeFontInfo? info, out _, fontName);
        return AnalyzeAdvancedTextLayout(text, info, source, indexOffset);
    }

    internal static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, PdfTrueTypeFontProgram font, string source = "", int indexOffset = 0) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));
        PdfOpenTypeFontInfo? info = TrueTypeLayoutInfoCache.GetValue(
            font,
            static value => new OpenTypeFontInfoBox(value.FontDataForInspection, value.FontName)).Info;
        return AnalyzeAdvancedTextLayout(text, info, source, indexOffset);
    }

    internal static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, PdfOpenTypeCffFontProgram font, string source = "", int indexOffset = 0) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(font, nameof(font));
        PdfOpenTypeFontInfo? info = CffLayoutInfoCache.GetValue(
            font,
            static value => new OpenTypeFontInfoBox(value.FontDataForInspection, value.FontName)).Info;
        return AnalyzeAdvancedTextLayout(text, info, source, indexOffset);
    }

    private static List<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayout(string text, PdfOpenTypeFontInfo? info, string source, int indexOffset) {
        var diagnostics = new List<PdfTextShapingDiagnostic>(AnalyzeAdvancedTextLayoutCore(text, source, indexOffset));
        if (info == null) return diagnostics;
        var reportedCodes = new HashSet<string>(diagnostics.Select(diagnostic => diagnostic.Code), StringComparer.Ordinal);
        AddFontLayoutDiagnostics(text, info, diagnostics, reportedCodes, source, indexOffset);
        return diagnostics;
    }

    /// <summary>
    /// Finds text runs that currently require shaping, bidirectional layout, mark positioning, or script-specific line breaking beyond OfficeIMO.Pdf's scalar text path.
    /// Explicit PDF line-break and tab runs are treated as layout controls rather than literal text.
    /// </summary>
    /// <param name="runs">Text runs to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <returns>Advanced text layout diagnostics in run order.</returns>
    public static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayoutRuns(IEnumerable<TextRun> runs, string source = "") {
        Guard.NotNull(runs, nameof(runs));
        var diagnostics = new List<PdfTextShapingDiagnostic>();

        foreach (TextRun run in runs) {
            if (run == null || IsLayoutControlRun(run)) {
                continue;
            }

            diagnostics.AddRange(AnalyzeAdvancedTextLayout(run.Text, source));
        }

        return diagnostics;
    }

    /// <summary>
    /// Finds text runs that currently require advanced layout support, including font-specific OpenType substitution and positioning features.
    /// Explicit PDF line-break and tab runs are treated as layout controls rather than literal text.
    /// </summary>
    /// <param name="runs">Text runs to inspect.</param>
    /// <param name="fontData">OpenType or TrueType font bytes used for the generated text.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="fontName">Optional configured font name used in diagnostic messages.</param>
    /// <returns>Advanced text layout diagnostics in run order.</returns>
    public static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedTextLayoutRuns(IEnumerable<TextRun> runs, byte[] fontData, string source = "", string? fontName = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.NotNull(fontData, nameof(fontData));
        var diagnostics = new List<PdfTextShapingDiagnostic>();
        var reportedCodes = new HashSet<string>(StringComparer.Ordinal);
        PdfOpenTypeFontInfo? info = null;
        if (!PdfOpenTypeFontInspector.TryInspect(fontData, out info, out _, fontName)) {
            info = null;
        }

        foreach (TextRun run in runs) {
            if (run == null || IsLayoutControlRun(run)) {
                continue;
            }

            foreach (PdfTextShapingDiagnostic diagnostic in AnalyzeAdvancedTextLayout(run.Text, source)) {
                if (reportedCodes.Add(diagnostic.Code)) {
                    diagnostics.Add(diagnostic);
                }
            }

            if (info != null) {
                AddFontLayoutDiagnostics(run.Text, info, diagnostics, reportedCodes, source, indexOffset: 0);
            }
        }

        return diagnostics;
    }

    /// <summary>
    /// Plans how generated text can be split across embedded font candidates before rendering.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="candidates">Candidate fonts in priority order.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <param name="shapingMode">Text shaping mode to use when checking fallback font coverage.</param>
    /// <returns>A fallback plan with covered text segments and missing-glyph diagnostics.</returns>
    public static PdfTextFallbackPlan PlanEmbeddedFontFallbackText(string text, IEnumerable<PdfEmbeddedFontFallbackCandidate> candidates, string source = "", PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(candidates, nameof(candidates));

        List<EmbeddedFontFallbackProgram> fonts = BuildFallbackPrograms(candidates);
        if (fonts.Count == 0) {
            throw new ArgumentException("At least one embedded font fallback candidate is required.", nameof(candidates));
        }

        var segments = new List<PdfTextFallbackSegment>();
        var diagnostics = new List<PdfTextEncodingDiagnostic>();
        int segmentStart = -1;
        int segmentFontIndex = -1;
        string segmentFontName = string.Empty;
        int index = 0;

        void FlushSegment(int endIndex) {
            if (segmentStart < 0 || endIndex <= segmentStart) {
                return;
            }

            segments.Add(new PdfTextFallbackSegment(
                text.Substring(segmentStart, endIndex - segmentStart),
                segmentStart,
                endIndex - segmentStart,
                segmentFontIndex,
                segmentFontName));
            segmentStart = -1;
            segmentFontIndex = -1;
            segmentFontName = string.Empty;
        }

        while (index < text.Length) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                FlushSegment(scalarStart);
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                FlushSegment(scalarStart);
                diagnostics.Add(CreateDiagnostic(text, scalarStart, source));
                continue;
            }

            int fontIndex = FindCoveringFont(fonts, text, scalarStart, shapingMode, out int coveredLength);
            if (fontIndex < 0) {
                FlushSegment(scalarStart);
                diagnostics.Add(CreateEmbeddedFallbackDiagnostic(scalarStart, scalar, source, fonts));
                continue;
            }

            if (segmentStart < 0) {
                segmentStart = scalarStart;
                segmentFontIndex = fontIndex;
                segmentFontName = fonts[fontIndex].FontName;
            } else if (segmentFontIndex != fontIndex) {
                FlushSegment(scalarStart);
                segmentStart = scalarStart;
                segmentFontIndex = fontIndex;
                segmentFontName = fonts[fontIndex].FontName;
            }

            index = scalarStart + coveredLength;
        }

        FlushSegment(text.Length);
        return new PdfTextFallbackPlan(text, segments, diagnostics);
    }

    private delegate bool TryGetSelectedTextLength(string text, int index, out int length);

    private static List<PdfTextEncodingDiagnostic> AnalyzeGeneratedTextWithFallback(
        string text,
        PdfEmbeddedFontFallbackSet fallbackSet,
        string source,
        string location,
        int? runIndex,
        PdfTextShapingMode shapingMode,
        TryGetSelectedTextLength tryGetSelectedTextLength) {
        List<EmbeddedFontFallbackProgram> fallbackFonts = BuildFallbackPrograms(fallbackSet.Candidates);
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                continue;
            }

            if (tryGetSelectedTextLength(text, scalarStart, out int selectedLength)) {
                index = scalarStart + selectedLength;
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                diagnostics.Add(CreateDiagnostic(text, scalarStart, source, location, runIndex));
                continue;
            }

            if (FindCoveringFont(fallbackFonts, text, scalarStart, shapingMode, out int fallbackCoveredLength) < 0) {
                diagnostics.Add(CreateEmbeddedFallbackDiagnostic(scalarStart, scalar, source, fallbackFonts, location, runIndex));
            } else {
                index = scalarStart + fallbackCoveredLength;
            }
        }

        return diagnostics;
    }

    private static bool TryGetWinAnsiCoveredTextLength(string text, int index, out int length) {
        int endIndex = index;
        _ = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return PdfWinAnsiEncoding.CanEncode(text.Substring(index, length), out _);
    }

    private static bool TryGetCoveredTextLength(string text, int index, PdfTrueTypeFontProgram fontProgram, PdfTextShapingMode shapingMode, out int length) {
        if (TrySkipCoveredLatinLigature(text, index, shapingMode, fontProgram, out length)) {
            return true;
        }

        int endIndex = index;
        int scalar = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return fontProgram.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private static bool TryGetCoveredTextLength(string text, int index, PdfOpenTypeCffFontProgram fontProgram, PdfTextShapingMode shapingMode, out int length) {
        if (TrySkipCoveredLatinLigature(text, index, shapingMode, fontProgram, out length)) {
            return true;
        }

        int endIndex = index;
        int scalar = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return fontProgram.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private static bool IsLayoutControlRun(TextRun run) =>
        string.Equals(run.Text, "\n", StringComparison.Ordinal) ||
        string.Equals(run.Text, "\t", StringComparison.Ordinal);

    private static List<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontTextCore(string text, PdfTrueTypeFontProgram font, string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        PdfGlyphRun glyphRun = font.ShapeText(text, PdfTextShapingOptions.ForDiagnostics(source, fontName, shapingMode));
        return glyphRun.Diagnostics.Count == 0
            ? new List<PdfTextEncodingDiagnostic>()
            : glyphRun.Diagnostics.ToList();
    }

    private static List<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontTextCore(string text, PdfOpenTypeCffFontProgram font, string source, string fontName, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        var diagnostics = new List<PdfTextEncodingDiagnostic>();
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            if (TrySkipCoveredLatinLigature(text, scalarStart, shapingMode, font, out int ligatureLength)) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                diagnostics.Add(CreateControlCharacterDiagnostic(scalarStart, scalar, source));
                continue;
            }

            if (!font.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                diagnostics.Add(CreateEmbeddedCffFontDiagnostic(scalarStart, scalar, source, fontName));
            }
        }

        return diagnostics;
    }

    private static string AppendRunLocation(string location, int runIndex) {
        if (string.IsNullOrWhiteSpace(location)) {
            return string.Empty;
        }

        return location + ".Run[" + runIndex.ToString(CultureInfo.InvariantCulture) + "]";
    }

    private static List<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontText(string text, PdfTrueTypeFontProgram fontProgram, string source, string location, int? runIndex, PdfTextShapingMode shapingMode) {
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        int index = 0;
        while (index < text.Length) {
            char ch = text[index];
            if (ch == '\n' || ch == '\r' || ch == '\t') {
                index++;
                continue;
            }

            int scalarStart = index;
            if (TrySkipCoveredLatinLigature(text, scalarStart, shapingMode, fontProgram, out int ligatureLength)) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (!fontProgram.TryGetGlyphId(scalar, out _)) {
                diagnostics.Add(CreateDiagnostic(
                    text,
                    scalarStart,
                    source,
                    location,
                    runIndex,
                    "embedded TrueType font '" + fontProgram.FontName + "'",
                    "Choose a font that contains this glyph or configure a fallback before rendering."));
            }
        }

        return diagnostics;
    }

    private static List<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontText(string text, PdfOpenTypeCffFontProgram fontProgram, string source, string location, int? runIndex, PdfTextShapingMode shapingMode) {
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        int index = 0;
        while (index < text.Length) {
            char ch = text[index];
            if (ch == '\n' || ch == '\r' || ch == '\t') {
                index++;
                continue;
            }

            int scalarStart = index;
            if (TrySkipCoveredLatinLigature(text, scalarStart, shapingMode, fontProgram, out int ligatureLength)) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (!fontProgram.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                diagnostics.Add(CreateDiagnostic(
                    text,
                    scalarStart,
                    source,
                    location,
                    runIndex,
                    "embedded OpenType/CFF font '" + fontProgram.FontName + "'",
                    "Choose a font that contains this glyph or configure a fallback before rendering."));
            }
        }

        return diagnostics;
    }

    private static bool TrySkipCoveredLatinLigature(string text, int index, PdfTextShapingMode shapingMode, PdfTrueTypeFontProgram fontProgram, out int ligatureLength) {
        ligatureLength = 0;
        if (shapingMode != PdfTextShapingMode.LatinLigatures ||
            !OfficeTextLigatures.TryGetLatinPresentationForm(text, index, out int ligatureScalar, out ligatureLength) ||
            !fontProgram.TryGetGlyphId(ligatureScalar, out int glyphId) ||
            glyphId <= 0) {
            ligatureLength = 0;
            return false;
        }

        return true;
    }

    private static bool TrySkipCoveredLatinLigature(string text, int index, PdfTextShapingMode shapingMode, PdfOpenTypeCffFontProgram fontProgram, out int ligatureLength) {
        ligatureLength = 0;
        if (shapingMode != PdfTextShapingMode.LatinLigatures ||
            !OfficeTextLigatures.TryGetLatinPresentationForm(text, index, out int ligatureScalar, out ligatureLength) ||
            !fontProgram.TryGetGlyphId(ligatureScalar, out int glyphId) ||
            glyphId <= 0) {
            ligatureLength = 0;
            return false;
        }

        return true;
    }

    private static List<EmbeddedFontFallbackProgram> BuildFallbackPrograms(IEnumerable<PdfEmbeddedFontFallbackCandidate> candidates) {
        var fonts = new List<EmbeddedFontFallbackProgram>();
        foreach (PdfEmbeddedFontFallbackCandidate candidate in candidates) {
            if (candidate == null) {
                throw new ArgumentException("Embedded font fallback candidates cannot contain null entries.", nameof(candidates));
            }

            fonts.Add(FallbackProgramCache.GetValue(candidate, CreateFallbackProgram).Program);
        }

        return fonts;
    }

    private static EmbeddedFontFallbackProgramBox CreateFallbackProgram(PdfEmbeddedFontFallbackCandidate candidate) {
        byte[] fontData = candidate.DataSnapshot;
        EmbeddedFontFallbackProgram program = IsOpenTypeCffFontData(fontData)
            ? new EmbeddedFontFallbackProgram(candidate.FontName, PdfOpenTypeCffFontProgram.Parse(fontData, candidate.FontName))
            : new EmbeddedFontFallbackProgram(candidate.FontName, PdfTrueTypeFontProgram.Parse(fontData, candidate.FontName));
        return new EmbeddedFontFallbackProgramBox(program);
    }

    private static int FindCoveringFont(IReadOnlyList<EmbeddedFontFallbackProgram> fonts, int scalar) {
        for (int index = 0; index < fonts.Count; index++) {
            if (fonts[index].TryGetGlyphId(scalar, out int glyphId) && glyphId > 0) {
                return index;
            }
        }

        return -1;
    }

    private static int FindCoveringFont(IReadOnlyList<EmbeddedFontFallbackProgram> fonts, string text, int textIndex, PdfTextShapingMode shapingMode, out int coveredLength) {
        coveredLength = 0;
        if (shapingMode == PdfTextShapingMode.LatinLigatures &&
            OfficeTextLigatures.TryGetLatinPresentationForm(text, textIndex, out int ligatureScalar, out int ligatureLength)) {
            int ligatureFontIndex = FindCoveringFont(fonts, ligatureScalar);
            if (ligatureFontIndex >= 0) {
                coveredLength = ligatureLength;
                return ligatureFontIndex;
            }
        }

        int endIndex = textIndex;
        int scalar = ReadScalar(text, ref endIndex);
        int fontIndex = FindCoveringFont(fonts, scalar);
        if (fontIndex >= 0) {
            coveredLength = endIndex - textIndex;
        }

        return fontIndex;
    }

    private static PdfTextEncodingDiagnostic CreateDiagnostic(string text, int index, string source) {
        return CreateDiagnostic(text, index, source, string.Empty, null);
    }

    private static PdfStandardFont ResolveRunFont(PdfStandardFont baseFont, TextRun run) {
        PdfStandardFont font = run.Font ?? baseFont;
        if (run.Bold && run.Italic) {
            return PdfStandardFontMapper.GetStyledFont(font, bold: true, italic: true);
        }

        if (run.Bold) {
            return PdfStandardFontMapper.GetStyledFont(font, bold: true, italic: false);
        }

        if (run.Italic) {
            return PdfStandardFontMapper.GetStyledFont(font, bold: false, italic: true);
        }

        return font;
    }

    private static PdfTextEncodingDiagnostic CreateDiagnostic(string text, int index, string source, string location, int? runIndex) {
        return CreateDiagnostic(text, index, source, location, runIndex, string.Empty, string.Empty);
    }

    private static PdfTextEncodingDiagnostic CreateDiagnostic(string text, int index, string source, string location, int? runIndex, string encoding, string remediation) {
        char ch = text[index];
        bool isSurrogatePair = char.IsHighSurrogate(ch) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]);
        int codePointValue = isSurrogatePair ? char.ConvertToUtf32(ch, text[index + 1]) : ch;
        string codePoint = "U+" + codePointValue.ToString(codePointValue <= 0xFFFF ? "X4" : "X", CultureInfo.InvariantCulture);
        bool isControlCharacter = ch < ' ' || ch == '\u007F';
        string display = isControlCharacter
            ? string.Empty
            : isSurrogatePair
                ? new string(new[] { ch, text[index + 1] })
                : ch.ToString();
        string diagnosticEncoding = string.IsNullOrWhiteSpace(encoding)
            ? isControlCharacter ? ControlCharacterEncodingDescription : WinAnsiEncodingDescription
            : encoding;
        string diagnosticRemediation = string.IsNullOrWhiteSpace(remediation)
            ? isControlCharacter ? ControlCharacterRemediation : WinAnsiGlyphRemediation
            : remediation;

        return new PdfTextEncodingDiagnostic(source, index, codePoint, display, isControlCharacter, diagnosticEncoding, diagnosticRemediation, location, runIndex);
    }

    internal static PdfTextEncodingDiagnostic CreateControlCharacterDiagnostic(int index, int scalar, string source) {
        string codePoint = FormatCodePoint(scalar);
        return new PdfTextEncodingDiagnostic(source, index, codePoint, string.Empty, isControlCharacter: true);
    }

    internal static PdfTextEncodingDiagnostic CreateEmbeddedFontDiagnostic(int index, int scalar, string source, string fontName) {
        string codePoint = FormatCodePoint(scalar);
        string display = GetDisplayText(scalar);
        string rendered = string.IsNullOrEmpty(display) ? string.Empty : " '" + display + "'";
        string message = "Text contains character " + codePoint + rendered + " at index " + index.ToString(CultureInfo.InvariantCulture) + " that is not covered by embedded TrueType font '" + fontName + "'. Configure a font that contains this glyph or split the run to a fallback font.";
        return new PdfTextEncodingDiagnostic(
            source,
            index,
            codePoint,
            display,
            isControlCharacter: false,
            code: "missing-embedded-font-glyph",
            message: message,
            customCode: true);
    }

    private static PdfTextEncodingDiagnostic CreateEmbeddedCffFontDiagnostic(int index, int scalar, string source, string fontName) {
        string codePoint = FormatCodePoint(scalar);
        string display = GetDisplayText(scalar);
        string rendered = string.IsNullOrEmpty(display) ? string.Empty : " '" + display + "'";
        string message = "Text contains character " + codePoint + rendered + " at index " + index.ToString(CultureInfo.InvariantCulture) + " that is not covered by embedded OpenType/CFF font '" + fontName + "'. Configure a font that contains this glyph or split the run to a fallback font.";
        return new PdfTextEncodingDiagnostic(
            source,
            index,
            codePoint,
            display,
            isControlCharacter: false,
            code: "missing-embedded-font-glyph",
            message: message,
            customCode: true);
    }

    private static void AddDiagnostic(List<PdfTextShapingDiagnostic> diagnostics, HashSet<string> reportedCodes, string source, int index, int scalar, string script, string code, string message, bool isCoveredByBuiltInShaping = false) {
        if (reportedCodes.Add(code)) {
            diagnostics.Add(new PdfTextShapingDiagnostic(source, index, scalar, script, code, message, isCoveredByBuiltInShaping));
        }
    }

    private static void AddFontLayoutDiagnostics(string text, PdfOpenTypeFontInfo info, List<PdfTextShapingDiagnostic> diagnostics, HashSet<string> reportedCodes, string source, int indexOffset) {
        if (HasAnyFeature(info.GlyphSubstitutionFeatureTags, "liga", "clig", "dlig", "rlig")) {
            int ligatureIndex = FindLatinLigatureSequenceIndex(text);
            if (ligatureIndex >= 0) {
                int sourceIndex = ligatureIndex + indexOffset;
                int scalar = char.ConvertToUtf32(text, ligatureIndex);
                bool isCoveredByBuiltInShaping =
                    OfficeTextLigatures.TryGetLatinPresentationForm(text, ligatureIndex, out int ligatureScalar, out _) &&
                    info.ContainsUnicodeScalar(ligatureScalar);
                AddDiagnostic(
                    diagnostics,
                    reportedCodes,
                    source,
                    sourceIndex,
                    scalar,
                    "OpenType GSUB ligature",
                    "unsupported-font-ligature-substitution",
                    "Text contains a Latin ligature sequence at index " + sourceIndex.ToString(CultureInfo.InvariantCulture) + ", and embedded font '" + info.FontName + "' advertises GSUB ligature features. OfficeIMO.Pdf currently writes scalar glyph ids without applying OpenType ligature substitution, so generated output may be visually simplified.",
                    isCoveredByBuiltInShaping);
            }
        }

        if (HasAnyFeature(info.GlyphPositioningFeatureTags, "mark", "mkmk")) {
            for (int index = 0; index < text.Length;) {
                int scalarStart = index;
                int sourceIndex = scalarStart + indexOffset;
                int scalar = ReadScalar(text, ref index);
                if (!IsCombiningMarkOrJoiner(scalar)) {
                    continue;
                }

                AddDiagnostic(
                    diagnostics,
                    reportedCodes,
                    source,
                    sourceIndex,
                    scalar,
                    "OpenType GPOS mark",
                    "unsupported-font-mark-positioning",
                    "Text contains a combining mark or joiner at index " + sourceIndex.ToString(CultureInfo.InvariantCulture) + ", and embedded font '" + info.FontName + "' advertises GPOS mark positioning features. OfficeIMO.Pdf currently writes scalar glyph ids without applying OpenType mark positioning, so generated output may be visually simplified.");
                break;
            }
        }
    }

    private static bool HasAnyFeature(IReadOnlyList<string> featureTags, params string[] expectedTags) {
        for (int index = 0; index < featureTags.Count; index++) {
            string tag = featureTags[index];
            for (int expectedIndex = 0; expectedIndex < expectedTags.Length; expectedIndex++) {
                if (string.Equals(tag, expectedTags[expectedIndex], StringComparison.Ordinal)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static int FindLatinLigatureSequenceIndex(string text) {
        string[] sequences = { "ffi", "ffl", "ff", "fi", "fl" };
        for (int index = 0; index < text.Length; index++) {
            for (int sequenceIndex = 0; sequenceIndex < sequences.Length; sequenceIndex++) {
                string sequence = sequences[sequenceIndex];
                if (index <= text.Length - sequence.Length &&
                    string.Compare(text, index, sequence, 0, sequence.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    return index;
                }
            }
        }

        return -1;
    }

    private static bool TryGetComplexScriptName(int scalar, out string script) {
        if (IsInRange(scalar, 0x0600, 0x06FF) ||
            IsInRange(scalar, 0x0750, 0x077F) ||
            IsInRange(scalar, 0x08A0, 0x08FF) ||
            IsInRange(scalar, 0xFB50, 0xFDFF) ||
            IsInRange(scalar, 0xFE70, 0xFEFF) ||
            IsInRange(scalar, 0x1EE00, 0x1EEFF)) {
            script = "Arabic";
            return true;
        }

        if (IsInRange(scalar, 0x0900, 0x0D7F) ||
            IsInRange(scalar, 0xA8E0, 0xA8FF) ||
            IsInRange(scalar, 0xA980, 0xA9DF)) {
            script = "Indic";
            return true;
        }

        if (IsInRange(scalar, 0x1780, 0x17FF) ||
            IsInRange(scalar, 0x1000, 0x109F)) {
            script = "Southeast Asian";
            return true;
        }

        if (IsInRange(scalar, 0x0700, 0x074F)) {
            script = "Syriac";
            return true;
        }

        script = string.Empty;
        return false;
    }

    private static bool TryGetScriptLineBreakingName(int scalar, out string script) {
        if (IsInRange(scalar, 0x0E00, 0x0E7F)) {
            script = "Thai";
            return true;
        }

        if (IsInRange(scalar, 0x0E80, 0x0EFF)) {
            script = "Lao";
            return true;
        }

        if (IsInRange(scalar, 0x1780, 0x17FF)) {
            script = "Khmer";
            return true;
        }

        if (IsInRange(scalar, 0x1000, 0x109F)) {
            script = "Myanmar";
            return true;
        }

        script = string.Empty;
        return false;
    }

    private static bool IsCombiningMarkOrJoiner(int scalar) =>
        scalar == 0x200C ||
        scalar == 0x200D ||
        IsInRange(scalar, 0x0300, 0x036F) ||
        IsInRange(scalar, 0x0591, 0x05BD) ||
        scalar == 0x05BF ||
        IsInRange(scalar, 0x05C1, 0x05C2) ||
        IsInRange(scalar, 0x05C4, 0x05C5) ||
        IsInRange(scalar, 0x0610, 0x061A) ||
        IsInRange(scalar, 0x064B, 0x065F) ||
        scalar == 0x0670 ||
        IsInRange(scalar, 0x06D6, 0x06ED) ||
        IsInRange(scalar, 0x0711, 0x074A) ||
        IsInRange(scalar, 0x07A6, 0x07B0) ||
        IsInRange(scalar, 0x0816, 0x082D) ||
        IsInRange(scalar, 0x0900, 0x0D7F) ||
        IsInRange(scalar, 0x1AB0, 0x1AFF) ||
        IsInRange(scalar, 0x1DC0, 0x1DFF) ||
        IsInRange(scalar, 0x20D0, 0x20FF) ||
        IsInRange(scalar, 0xFE20, 0xFE2F);

    private static bool IsInRange(int scalar, int first, int last) =>
        scalar >= first && scalar <= last;

    private static PdfTextEncodingDiagnostic CreateEmbeddedFallbackDiagnostic(int index, int scalar, string source, IReadOnlyList<EmbeddedFontFallbackProgram> fonts, string location = "", int? runIndex = null) {
        string codePoint = FormatCodePoint(scalar);
        string display = GetDisplayText(scalar);
        string rendered = string.IsNullOrEmpty(display) ? string.Empty : " '" + display + "'";
        string fontNames = string.Join(", ", fonts.Select(font => "'" + font.FontName + "'"));
        string message = "Text contains character " + codePoint + rendered + " at index " + index.ToString(CultureInfo.InvariantCulture) + " that is not covered by any embedded font fallback candidate: " + fontNames + ". Add a fallback font that contains this glyph or replace the character before rendering.";
        return new PdfTextEncodingDiagnostic(
            source,
            index,
            codePoint,
            display,
            isControlCharacter: false,
            code: "missing-embedded-font-fallback-glyph",
            message: message,
            customCode: true);
    }

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
    }

    private static string FormatCodePoint(int scalar) =>
        "U+" + scalar.ToString(scalar <= 0xFFFF ? "X4" : "X", CultureInfo.InvariantCulture);

    private static string GetDisplayText(int scalar) =>
        scalar < ' ' || scalar == '\u007F' || scalar > 0x10FFFF || (scalar >= 0xD800 && scalar <= 0xDFFF)
            ? string.Empty
            : char.ConvertFromUtf32(scalar);

    private static bool IsOpenTypeCffFontData(byte[] fontData) =>
        fontData.Length >= 4 &&
        fontData[0] == 0x4F &&
        fontData[1] == 0x54 &&
        fontData[2] == 0x54 &&
        fontData[3] == 0x4F;

    private readonly struct EmbeddedFontFallbackProgram {
        public EmbeddedFontFallbackProgram(string fontName, PdfTrueTypeFontProgram font) {
            FontName = fontName;
            _trueTypeFont = font;
            _cffFont = null;
        }

        public EmbeddedFontFallbackProgram(string fontName, PdfOpenTypeCffFontProgram font) {
            FontName = fontName;
            _trueTypeFont = null;
            _cffFont = font;
        }

        private readonly PdfTrueTypeFontProgram? _trueTypeFont;
        private readonly PdfOpenTypeCffFontProgram? _cffFont;

        public string FontName { get; }

        public bool TryGetGlyphId(int unicodeScalar, out int glyphId) {
            if (_trueTypeFont != null) {
                return _trueTypeFont.TryGetGlyphId(unicodeScalar, out glyphId);
            }

            if (_cffFont != null) {
                return _cffFont.TryGetGlyphId(unicodeScalar, out glyphId);
            }

            glyphId = 0;
            return false;
        }
    }

    private sealed class OpenTypeFontInfoBox {
        public OpenTypeFontInfoBox(byte[] fontData, string fontName) {
            PdfOpenTypeFontInspector.TryInspect(fontData, out PdfOpenTypeFontInfo? info, out _, fontName);
            Info = info;
        }

        public PdfOpenTypeFontInfo? Info { get; }
    }

    private sealed class EmbeddedFontFallbackProgramBox {
        public EmbeddedFontFallbackProgramBox(EmbeddedFontFallbackProgram program) {
            Program = program;
        }

        public EmbeddedFontFallbackProgram Program { get; }
    }
}
