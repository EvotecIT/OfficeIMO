using System.Collections.Generic;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static int GetNativeJoinedCellParagraphEnd(IReadOnlyList<WordElement> elements, int index,
        Func<WordParagraph, (int Level, string Marker)?>? getMarker,
        NativeDocumentDefaults defaults, NativeTableStyleDefaults tableStyle, NativeFontMap? fontMap, WordToPdfOptions? options) {
        int end = index;
        while (elements[end] is WordParagraph paragraph && !ShouldRenderNativeEmptyParagraphLineBox(paragraph) &&
               end + 1 < elements.Count && elements[end + 1] is WordParagraph) end++;
        if (end == index) return index;

        Func<WordParagraph, (int Level, string Marker)?> marker = getMarker ?? (_ => null);
        for (int current = index; current <= end; current++) {
            var paragraph = (WordParagraph)elements[current];
            if (!CanJoinNativeTextParagraph(paragraph, marker, defaults, fontMap) ||
                (current > index && HasNativePageBreakBefore(paragraph)) ||
                (tableStyle.UseConfiguredTypography && !ResolveNativeJoinedCellFontSize(paragraph, defaults, tableStyle).HasValue)) {
                if (options != null) AddNativeExportWarning(options, "NativeHiddenParagraphJoinUnsupported", "table cell",
                    "Hidden paragraph marks beside lists, headings, decorated paragraphs, flow breaks, objects or unresolved configured typography retain separate PDF paragraph boundaries.");
                return index;
            }
        }
        return end;
    }

    private static (double? FontSize, double? LineHeight, PdfCore.PdfLineSpacing? LineSpacing)
        ResolveNativeJoinedCellParagraphMetrics(IReadOnlyList<WordElement> elements,
            IReadOnlyList<List<PdfCore.PdfTextRun>?> preparedRuns, int firstIndex, int lastIndex,
            NativeDocumentDefaults defaults, NativeTableStyleDefaults tableStyle, NativeFontMap? fontMap) {
        var visible = new List<WordParagraph>();
        for (int index = firstIndex; index <= lastIndex; index++) {
            if (preparedRuns[index]!.Any(run => !string.IsNullOrEmpty(run.Text))) visible.Add((WordParagraph)elements[index]);
        }
        // An empty joined group renders only its final visible mark. Visible
        // source runs otherwise supply the same minimum used by rich body text.
        WordParagraph first = (WordParagraph)elements[visible.Count == 0 ? lastIndex : firstIndex];
        if (visible.Count == 0) visible.Add(first);
        double? fontSize = visible.Min(paragraph => ResolveNativeJoinedCellFontSize(paragraph, defaults, tableStyle));
        double naturalHeight = visible.Max(paragraph => ResolveNativeParagraphSingleLineHeight(paragraph, defaults,
            GetNativeParagraphStyleDefaults(paragraph), tableStyle.RunStyle, fontMap));
        NativeLineSpacing spacing = ResolveNativeParagraphLineSpacing(first, GetNativeParagraphStyleDefaults(first), defaults, tableStyle.LineSpacing);
        if (tableStyle.UseConfiguredTypography && !spacing.Value.HasValue) return (fontSize, null, null);
        double? lineHeight = spacing.Resolve(fontSize ?? defaults.FontSize, naturalHeight) ?? tableStyle.ParagraphLineHeight;
        return (fontSize, lineHeight, spacing.ToPdfLineSpacing(naturalHeight));
    }

    private static double? ResolveNativeJoinedCellFontSize(WordParagraph paragraph, NativeDocumentDefaults defaults,
        NativeTableStyleDefaults tableStyle) {
        NativeParagraphStyleDefaults style = GetNativeParagraphStyleDefaults(paragraph);
        if (!tableStyle.UseConfiguredTypography) return ResolveNativeParagraphLayoutFontSize(paragraph, defaults, style, tableStyle.RunStyle);
        string? markSize = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.FontSize>()?.Val?.Value;
        return double.TryParse(markSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) && halfPoints > 0D
            ? halfPoints / 2D : style.FontSize ?? tableStyle.RunStyle.FontSize;
    }
}
