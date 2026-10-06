using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Header/footer zones serialize a paragraph at a time. Combine source text
    // before zone normalization so a hidden mark does not trim an interior space.
    private static bool TryAddNativeHeaderFooterJoinedParagraphText(NativeHeaderFooterText parts,
        IReadOnlyList<WordElement> elements, ref int index,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
        NativeFontMap? fontMap, WordToPdfOptions? options, NativeHeaderFooterZone? forcedZone = null) {
        if (elements[index] is not WordParagraph first) return false;
        int end = index;
        while (elements[end] is WordParagraph paragraph && !ShouldRenderNativeEmptyParagraphLineBox(paragraph) &&
               end + 1 < elements.Count && elements[end + 1] is WordParagraph) end++;
        if (end == index) return false;

        NativeDocumentDefaults defaults = GetNativeDocumentDefaults(first._document, fontMap);
        NativeResolvedTextStyle? visibleStyle = null;
        Func<WordParagraph, (int Level, string Marker)?> marker = paragraph =>
            listMarkers.TryGetValue(paragraph, out var value) ? value : null;
        for (int current = index; current <= end; current++) {
            var paragraph = (WordParagraph)elements[current];
            W.Paragraph source = paragraph._paragraph;
            bool supported = CanJoinNativeTextParagraph(paragraph, marker, defaults, fontMap) &&
                (current == index || !HasNativePageBreakBefore(paragraph)) &&
                (!forcedZone.HasValue || ReferenceEquals(source.Parent, first._paragraph.Parent)) &&
                !source.Descendants<W.SimpleField>().Any() && !source.Descendants<W.FieldChar>().Any() &&
                !source.Descendants<W.FootnoteReference>().Any() && !source.Descendants<W.EndnoteReference>().Any();
            if (supported) {
                foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                    if (IsNativeHiddenTextRun(run, paragraph) || string.IsNullOrEmpty(run.Text)) continue;
                    NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraph, nativeDefaults: defaults, nativeFontMap: fontMap);
                    // An ordinary inherited size can remain unspecified in the
                    // run projection; compare its effective document value.
                    style = style with { FontSize = style.FontSize ?? defaults.FontSize };
                    if (visibleStyle.HasValue && visibleStyle.Value != style) { supported = false; break; }
                    visibleStyle = style;
                }
            }
            if (!supported) {
                if (options != null) AddNativeExportWarning(options, "NativeHiddenParagraphJoinUnsupported", "header/footer",
                    "Hidden paragraph marks beside lists, headings, decorated paragraphs, flow breaks, objects, fields, notes or mixed typography retain separate header/footer PDF paragraph boundaries.");
                return false;
            }
        }

        var text = new StringBuilder();
        for (int current = index; current <= end; current++)
            text.Append(GetNativeHeaderFooterParagraphText((WordParagraph)elements[current], listMarkers, out _));
        AddNativeHeaderFooterResolvedParagraphText(parts, first, text.ToString(), null, forcedZone, listMarkers, fontMap);
        index = end;
        return true;
    }
}
