namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Measure each whitespace and word span once. Re-measuring every prefix makes
    // words-only decoration quadratic for a long run with many short words.
    private static void VisitWordDecorationAdvances(string text, System.Func<string, double> measure,
        System.Action<double, double> visit) {
        int index = 0;
        double advance = 0;
        while (index < text.Length) {
            int start = index;
            while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
            if (index > start) advance += measure(text.Substring(start, index - start));
            if (index == text.Length) break;

            start = index;
            while (index < text.Length && !char.IsWhiteSpace(text[index])) index++;
            double end = advance + measure(text.Substring(start, index - start));
            visit(advance, end);
            advance = end;
        }
    }
}
