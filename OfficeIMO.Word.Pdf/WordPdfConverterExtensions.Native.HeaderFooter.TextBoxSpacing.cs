using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static IReadOnlyList<NativeHeaderFooterStyledReplacement> CreateNativeHeaderFooterTextBoxReplacements(
        WordParagraph root,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
        NativeFontMap? fontMap) {
        var replacements = new List<NativeHeaderFooterStyledReplacement>();
        AppendParagraph(root, 0);
        return replacements;

        void AppendParagraph(WordParagraph paragraph, int depth) {
            if (depth > 0) {
                EnsureNativeHeaderFooterTextBoxDepth(depth);
                WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
                if (info.HasValue && listMarkers.TryGetValue(paragraph, out var marker) && !string.IsNullOrEmpty(marker.Marker)) {
                    string prefix = NormalizeNativeDirectText(marker.Marker + ResolveNativeInlineListMarkerSuffix(info.Value.LevelSuffix));
                    NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph, nativeFontMap: fontMap);
                    var offsets = ResolveNativeHeaderFooterListOffsets(paragraph, info.Value);
                    replacements.Add(new NativeHeaderFooterStyledReplacement(prefix,
                        CreateNativeHeaderFooterListMarkerTextRun(marker.Marker, paragraph, info.Value, style, fontMap, offsets.MarkerOffset, offsets.TextOffset)));
                }
            }

            foreach (var item in GetNativeHeaderFooterVisibleTextRuns(paragraph)) {
                if (IsNativeHiddenTextRun(item.Run, paragraph)) continue;
                if (item.IsField || item.Run._run?.Descendants<W.TextBoxContent>().Any() != true) {
                    AppendRun(paragraph, item.Run, item.Text, item.IsField);
                    continue;
                }

                // Preserve document order, including direct prefixes/suffixes and
                // peer boxes. Matching replacements cannot then consume the text
                // belonging to an earlier identical but differently styled run.
                foreach (var child in item.Run.EnumerateEffectiveRunContent()) {
                    WordParagraph view = CreateNativeRunContentView(item.Run, new[] { child });
                    if (view.TextBox is { } box) {
                        foreach (WordParagraph inner in GetNativeTextBoxParagraphs(box)) AppendParagraph(inner, depth + 1);
                    } else {
                        AppendRun(paragraph, view, view.Text, false);
                    }
                }
            }
        }

        void AppendRun(WordParagraph paragraph, WordParagraph run, string text, bool fieldToken) {
            NativeHeaderFooterStyledReplacement? replacement = CreateNativeHeaderFooterSpacingReplacement(paragraph, run, text, fieldToken, fontMap);
            if (replacement != null) replacements.Add(replacement);
        }
    }
}
