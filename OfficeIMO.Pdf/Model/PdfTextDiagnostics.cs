using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable text preflight helpers for generated PDF output.
/// </summary>
public static class PdfTextDiagnostics {
    /// <summary>
    /// Finds text that cannot be written through the current generated standard-font WinAnsi path.
    /// </summary>
    /// <param name="text">Text to inspect.</param>
    /// <param name="source">Optional caller-provided source label such as a block, field, sheet, slide, or converter area.</param>
    /// <returns>Encoding diagnostics in source order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeWinAnsiText(string text, string source = "") {
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
                diagnostics.Add(CreateDiagnostic(text, index, source));
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
    /// <returns>Encoding diagnostics in run order.</returns>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeWinAnsiTextRuns(IEnumerable<TextRun> runs, string source = "") {
        Guard.NotNull(runs, nameof(runs));
        var diagnostics = new List<PdfTextEncodingDiagnostic>();

        foreach (TextRun run in runs) {
            if (run == null || IsLayoutControlRun(run)) {
                continue;
            }

            diagnostics.AddRange(AnalyzeWinAnsiText(run.Text, source));
        }

        return diagnostics;
    }

    private static bool IsLayoutControlRun(TextRun run) =>
        string.Equals(run.Text, "\n", StringComparison.Ordinal) ||
        string.Equals(run.Text, "\t", StringComparison.Ordinal);

    private static PdfTextEncodingDiagnostic CreateDiagnostic(string text, int index, string source) {
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

        return new PdfTextEncodingDiagnostic(source, index, codePoint, display, isControlCharacter);
    }
}
