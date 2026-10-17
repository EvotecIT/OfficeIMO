namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormDictionaryBuilder {
    /// <summary>Shares list-box row geometry, clipping and selection paint between authored and regenerated appearances.</summary>
    internal static string BuildChoiceListAppearanceContent(double width, double height, int optionCount,
        int topIndex, IReadOnlyList<int> selectedIndices, double fontSize, PdfFormFieldStyle style,
        Action<StringBuilder, int, double, PdfFormFieldStyle> appendRow) {
        var content = new StringBuilder("q\n");
        content.Append(BuildFormFieldBackgroundAppearanceContent(width, height, style));
        if (style.BorderColor.HasValue && style.BorderWidth > 0D)
            content.Append(BuildRectangularBorderAppearanceContent(width, height, style));
        content.Append("2 2 ").Append(Format(Math.Max(0D, width - 4D))).Append(' ')
            .Append(Format(Math.Max(0D, height - 4D))).Append(" re W n\n");
        double rowHeight = Math.Max(fontSize + 2D, fontSize * 1.2D);
        (int first, int count) = GetChoiceListVisibleRange(height, fontSize, optionCount, topIndex);
        var selected = new HashSet<int>(selectedIndices);
        for (int offset = 0; offset < count; offset++) {
            int index = first + offset;
            double bottom = height - 2D - (offset + 1D) * rowHeight;
            PdfFormFieldStyle textStyle = style.Clone();
            textStyle.BackgroundColor = null;
            textStyle.BorderColor = null;
            textStyle.IsMultiline = false;
            textStyle.IsComb = false;
            textStyle.IsPassword = false;
            if (selected.Contains(index)) {
                content.Append("0.153 0.392 0.8 rg 2 ").Append(Format(bottom)).Append(' ')
                    .Append(Format(Math.Max(0D, width - 4D))).Append(' ').Append(Format(rowHeight)).Append(" re f\n");
                textStyle.TextColor = PdfColor.White;
            }
            double baseline = bottom + Math.Max(2D, (rowHeight - fontSize) / 2D + fontSize * 0.2D);
            appendRow(content, index, baseline, textStyle);
        }
        return content.Append("Q\n").ToString();
    }
    internal static (int First, int Count) GetChoiceListVisibleRange(double height, double fontSize, int optionCount, int topIndex) {
        double rowHeight = Math.Max(fontSize + 2D, fontSize * 1.2D);
        int visibleRows = Math.Max(1, (int)Math.Ceiling(Math.Max(0D, height - 4D) / rowHeight));
        int first = Math.Max(0, Math.Min(topIndex, Math.Max(0, optionCount - 1)));
        return (first, Math.Min(optionCount - first, visibleRows));
    }

}
