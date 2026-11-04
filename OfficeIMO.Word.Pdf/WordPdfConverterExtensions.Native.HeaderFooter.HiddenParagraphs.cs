using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    // Header/footer zones serialize a paragraph at a time. Combine source text
    // before zone normalization so a hidden mark does not trim an interior space.
    private static bool TryAddNativeHeaderFooterJoinedParagraphText(NativeHeaderFooterText parts,
        IReadOnlyList<WordElement> elements, ref int index, ref int rejectedJoinUntil,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
        NativeFontMap? fontMap, WordPdfExportOptions? options, NativeHeaderFooterZone? forcedZone = null) {
        if (index <= rejectedJoinUntil || elements[index] is not WordParagraph first ||
            ShouldRenderNativeEmptyParagraphLineBox(first) || index + 1 >= elements.Count ||
            elements[index + 1] is not WordParagraph) return false;
        NativeDocumentDefaults defaults = GetNativeDocumentDefaults(first._document, fontMap);
        Func<WordParagraph, (int Level, string Marker)?> marker = paragraph =>
            listMarkers.TryGetValue(paragraph, out var value) ? value : null;
        if (!CanJoinIndependently(first, false)) {
            ReportUnsupportedJoin();
            return false;
        }
        int end = index;
        while (elements[end] is WordParagraph paragraph && !ShouldRenderNativeEmptyParagraphLineBox(paragraph) &&
               end + 1 < elements.Count && elements[end + 1] is WordParagraph) end++;
        if (end == index) return false;

        NativeResolvedTextStyle? visibleStyle = null;
        for (int current = index; current <= end; current++) {
            var paragraph = (WordParagraph)elements[current];
            if (!CanJoinIndependently(paragraph, current > index)) {
                rejectedJoinUntil = current - 1;
                ReportUnsupportedJoin();
                return false;
            }
            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (IsNativeHiddenTextRun(run, paragraph) || string.IsNullOrEmpty(run.Text)) continue;
                NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraph, nativeDefaults: defaults, nativeFontMap: fontMap);
                // An ordinary inherited size can remain unspecified in the
                // run projection; compare its effective document value.
                style = style with { FontSize = style.FontSize ?? defaults.FontSize };
                if (visibleStyle.HasValue && visibleStyle.Value != style) {
                    ReportUnsupportedJoin();
                    return false;
                }
                visibleStyle = style;
            }
        }

        bool CanJoinIndependently(WordParagraph paragraph, bool isContinuation) {
            W.Paragraph source = paragraph._paragraph;
            return CanJoinNativeTextParagraph(paragraph, marker, defaults, fontMap) &&
                (!isContinuation || !HasNativePageBreakBefore(paragraph)) &&
                (!forcedZone.HasValue || ReferenceEquals(source.Parent, first._paragraph.Parent)) &&
                !source.Descendants<W.SimpleField>().Any() && !source.Descendants<W.FieldChar>().Any() &&
                !source.Descendants<W.FootnoteReference>().Any() && !source.Descendants<W.EndnoteReference>().Any();
        }

        void ReportUnsupportedJoin() {
            if (options != null) AddNativeExportWarning(options, "native-hidden-paragraph-join-unsupported", "header/footer",
                "Hidden paragraph marks beside lists, headings, decorated paragraphs, flow breaks, objects, fields, notes or mixed typography retain separate header/footer PDF paragraph boundaries.");
        }

        var text = new StringBuilder();
        for (int current = index; current <= end; current++)
            text.Append(GetNativeHeaderFooterParagraphText((WordParagraph)elements[current], listMarkers, out _));
        string joinedText = text.ToString();
        IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements = null;
        if (visibleStyle?.SmallCaps == true) {
            var casedReplacements = new List<NativeHeaderFooterStyledReplacement>();
            for (int current = index; current <= end; current++) {
                var paragraph = (WordParagraph)elements[current];
                foreach (var run in GetNativeHeaderFooterVisibleTextRuns(paragraph))
                    casedReplacements.AddRange(CreateNativeHeaderFooterSpacingReplacements(paragraph, run.Run, run.Text, run.IsField, fontMap));
            }
            replacements = casedReplacements;
        } else if (visibleStyle.HasValue && (HasNativeTextSpacing(visibleStyle.Value.TextSpacing) ||
            HasNativeTextSpacing(ResolveNativeTextRunStyle(first, nativeFontMap: fontMap).TextSpacing))) {
            replacements = new[] { new NativeHeaderFooterStyledReplacement(joinedText,
                CreateNativeHeaderFooterStyledTextRun(joinedText, visibleStyle.Value, 0D)) };
        }
        AddNativeHeaderFooterResolvedParagraphText(parts, first, joinedText, null, forcedZone, listMarkers, fontMap, replacements);
        index = end;
        return true;
    }
}
