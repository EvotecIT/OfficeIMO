using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable text preflight helpers for generated PDF output.
/// </summary>
public static class PdfTextDiagnostics {
    private const string WinAnsiEncodingDescription = "PDF WinAnsiEncoding";
    private const string WinAnsiGlyphRemediation = "Embedded Unicode fonts are required for this text.";
    private const string ControlCharacterEncodingDescription = "PDF text output";
    private const string ControlCharacterRemediation = "Use paragraphs, line breaks, tables, or spacing primitives for layout instead of literal control characters.";

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

        if (!options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) ||
            fontProgram == null) {
            return AnalyzeWinAnsiTextCore(text, source, location, runIndex);
        }

        return AnalyzeEmbeddedFontText(text, fontProgram, source, location, runIndex);
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

    private static bool IsLayoutControlRun(TextRun run) =>
        string.Equals(run.Text, "\n", StringComparison.Ordinal) ||
        string.Equals(run.Text, "\t", StringComparison.Ordinal);

    private static string AppendRunLocation(string location, int runIndex) {
        if (string.IsNullOrWhiteSpace(location)) {
            return string.Empty;
        }

        return location + ".Run[" + runIndex.ToString(CultureInfo.InvariantCulture) + "]";
    }

    private static List<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFontText(string text, PdfTrueTypeFontProgram fontProgram, string source, string location, int? runIndex) {
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        int index = 0;
        while (index < text.Length) {
            char ch = text[index];
            if (ch == '\n' || ch == '\r' || ch == '\t') {
                index++;
                continue;
            }

            int scalarStart = index;
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

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
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
}
