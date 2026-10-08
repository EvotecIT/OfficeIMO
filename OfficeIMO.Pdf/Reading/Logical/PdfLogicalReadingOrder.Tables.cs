namespace OfficeIMO.Pdf;

public static partial class PdfLogicalReadingOrderAnalysis {
    /// <summary>Suppresses a paragraph only when every source line has retained text in a table at its location.</summary>
    private static bool IsParagraphRepresentedByTable(
        PdfLogicalPage page,
        IReadOnlyList<PdfLogicalTextBlock> lines,
        PdfVisualBounds?[] tableBounds,
        string[] tableTexts,
        Action<long>? consumeWork) {
        if (lines.Count == 0 || tableTexts.Length == 0) return false;
        foreach (PdfLogicalTextBlock line in lines) {
            consumeWork?.Invoke(1);
            if (!line.IsTableContent || !TryGetVisualBounds(page, line, out PdfVisualBounds lineBounds)) return false;
            string text = NormalizeTableProjectionText(line.Text, consumeWork);
            if (text.Length == 0) return false;
            bool represented = false;
            double centerX = (lineBounds.Left + lineBounds.Right) / 2D;
            double centerY = (lineBounds.Top + lineBounds.Bottom) / 2D;
            for (int index = 0; index < tableTexts.Length; index++) {
                consumeWork?.Invoke(Math.Max(1, tableTexts[index].Length));
                if (tableBounds[index] is PdfVisualBounds bounds &&
                    centerX >= bounds.Left - 1D && centerX <= bounds.Right + 1D &&
                    centerY >= bounds.Top - 1D && centerY <= bounds.Bottom + 1D &&
                    ContainsTableProjectionText(tableTexts[index], text)) {
                    represented = true;
                    break;
                }
            }
            if (!represented) return false;
        }
        return true;
    }

    private static bool ContainsTableProjectionText(string value, string text) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        return value.IndexOf(text, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(text, StringComparison.Ordinal);
#endif
    }

    private static string NormalizeTableProjectionText(string text, Action<long>? consumeWork) {
        consumeWork?.Invoke(Math.Max(1, text.Length));
        var result = new System.Text.StringBuilder(text.Length);
        foreach (char character in text) {
            if (!char.IsWhiteSpace(character)) result.Append(char.ToUpperInvariant(character));
        }
        return result.ToString();
    }
}
