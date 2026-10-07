namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static (double Width, double FontSize)[] MeasureRichWordContinuations(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont,
        double fontSize, PdfOptions? options) {
        if (runs.Count < 2) return System.Array.Empty<(double Width, double FontSize)>();
        var continuations = new (double Width, double FontSize)[runs.Count];
        // Measure each run prefix once, working backwards through words that span
        // several differently formatted runs. Inline objects and whitespace end a word.
        for (int index = runs.Count - 1; index >= 1; index--) {
            PdfTextRun run = runs[index];
            if (run.InlineElement != null) continue;
            string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            int separator = text.IndexOfAny(TokenSplitChars);
            if (separator == 0) continue;
            string prefix = separator < 0 ? text : text.Substring(0, separator);
            PdfStandardFont font = ResolveFontForRun(run, baseFont);
            PdfNamedFontFace? namedFont = options != null &&
                options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace resolved)
                    ? resolved : null;
            double size = run.FontSize ?? fontSize;
            double width = MeasureRichText(prefix, font, namedFont, size, run.Baseline, options,
                run.FeatureSettings, run.HorizontalTextScaling, run.CharacterSpacing, run.TextDirection);
            double largestSize = prefix.Length > 0 ? size : 0D;
            if (separator < 0 && index + 1 < runs.Count) {
                width += continuations[index + 1].Width;
                largestSize = System.Math.Max(largestSize, continuations[index + 1].FontSize);
            }
            continuations[index] = (width, largestSize);
        }
        return continuations;
    }
}
