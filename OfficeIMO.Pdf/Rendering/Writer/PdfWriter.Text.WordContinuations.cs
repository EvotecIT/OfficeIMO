namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static (double Width, double FontSize, IReadOnlyList<RichSeg>? Parts, bool HasKerning)[] MeasureRichWordContinuations(
        System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, PdfStandardFont baseFont,
        double fontSize, PdfCreateOptions? options) {
        if (runs.Count < 2) return System.Array.Empty<(double Width, double FontSize, IReadOnlyList<RichSeg>? Parts, bool HasKerning)>();
        var continuations = new (double Width, double FontSize, IReadOnlyList<RichSeg>? Parts, bool HasKerning)[runs.Count];
        var prefixes = new List<RichSeg>();
        var starts = new int[runs.Count];
        var counts = new int[runs.Count];
        var extendsWord = new bool[runs.Count];
        // Every suffix views the same prefix array. Retaining a copy of each suffix
        // would require quadratic storage for a word split into many paint runs.
        for (int index = 1; index < runs.Count; index++) {
            starts[index] = prefixes.Count;
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
            double size = (run.FontSize?.ToPoints()) ?? fontSize;
            double width = MeasureRichText(prefix, font, namedFont, size, run.Baseline, options,
                run.FeatureSettings, run.HorizontalTextScaling, (run.CharacterSpacing.ToPoints()), run.TextDirection);
            double largestSize = prefix.Length > 0 ? size : 0D;
            if (prefix.Length > 0) {
                prefixes.Add(CreateRunKerningProbe(prefix, run, baseFont, fontSize, options));
                counts[index] = 1;
            }
            extendsWord[index] = separator < 0;
            continuations[index] = (width, largestSize, null,
                prefix.Length > 0 && run.FeatureSettings.TryGetValue("kern", out int kern) && kern > 0);
        }
        RichSeg[] parts = prefixes.ToArray();
        for (int index = runs.Count - 1; index >= 1; index--) {
            var current = continuations[index];
            if (extendsWord[index] && index + 1 < runs.Count) {
                var next = continuations[index + 1];
                current.Width += next.Width;
                current.FontSize = Math.Max(current.FontSize, next.FontSize);
                current.HasKerning |= next.HasKerning;
                counts[index] += counts[index + 1];
            }
            current.Parts = counts[index] == 0 ? (IReadOnlyList<RichSeg>?)null : new ArraySegment<RichSeg>(parts, starts[index], counts[index]);
            continuations[index] = current;
        }
        return continuations;
    }
}
