using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasNativeTextSpacing(NativeTextSpacing spacing) =>
            (spacing.WidthPercentage ?? 100D) != 100D || (spacing.CharacterSpacing ?? 0D) != 0D ||
            spacing.KerningMinimumFontSize.HasValue || spacing.LigatureSetting != null;

        private static IReadOnlyList<NativeHeaderFooterStyledReplacement>? CreateNativeHeaderFooterSpacingReplacements(
            WordParagraph paragraph, NativeFontMap? fontMap) {
            var replacements = new List<NativeHeaderFooterStyledReplacement>();
            bool hasSpacing = HasNativeTextSpacing(
                ResolveNativeTextRunStyle(paragraph, nativeFontMap: fontMap).TextSpacing);
            bool hasLiteralPageTokens = false;
            foreach (var item in GetNativeHeaderFooterVisibleTextRuns(paragraph)) {
                hasSpacing |= ResolveNativeTextRunStyle(item.Run, paragraph, nativeFontMap: fontMap).SmallCaps;
                foreach (NativeHeaderFooterStyledReplacement replacement in CreateNativeHeaderFooterSpacingReplacements(
                    paragraph, item.Run, item.Text, item.IsField, fontMap)) {
                    hasSpacing |= replacement.StyledRun.HorizontalTextScaling != 100D || (replacement.StyledRun.CharacterSpacing.ToPoints()) != 0D;
                    hasSpacing |= replacement.StyledRun.FeatureSettings.TryGetValue("kern", out int kerning) && kerning != 0;
                    hasSpacing |= new[] { "liga", "clig", "dlig", "hlig" }.Any(feature =>
                        replacement.StyledRun.FeatureSettings.TryGetValue(feature, out int enabled) && enabled != 0);
                    string text = replacement.SerializedText;
                    hasLiteralPageTokens |= !item.IsField &&
                        (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         text.IndexOf("{documentpages}", StringComparison.OrdinalIgnoreCase) >= 0);
                    replacements.Add(replacement);
                }
            }
            // Keep authored braces distinct from fields even when another paragraph enables styled output.
            return hasSpacing || hasLiteralPageTokens ? replacements : null;
        }

        private static IEnumerable<(WordParagraph Run, string Text, bool IsField)> GetNativeHeaderFooterVisibleTextRuns(WordParagraph paragraph) {
            var serializedRuns = new List<(W.Run Run, string Text, bool IsField, DocumentFormat.OpenXml.OpenXmlElement? SourceChild, OfficeIMO.Pdf.PdfPageNumberStyle? FieldStyle)>();
            return TryBuildNativeHeaderFooterParagraphText(paragraph, out _, out _, serializedRuns)
                ? serializedRuns.Select(item => {
                    var run = new WordParagraph(paragraph._document,
                        item.Run.Ancestors<W.Paragraph>().FirstOrDefault() ?? paragraph._paragraph!, item.Run);
                    return (item.SourceChild == null ? run : CreateNativeRunContentView(run, new[] { item.SourceChild }), item.Text, item.IsField);
                })
                : GetNativeRuns(paragraph).Select(run => (run, run.Text, false));
        }

        private static IEnumerable<NativeHeaderFooterStyledReplacement> CreateNativeHeaderFooterSpacingReplacements(
            WordParagraph paragraph, WordParagraph run, string sourceText, bool fieldToken, NativeFontMap? fontMap) {
            if (IsNativeHiddenTextRun(run, paragraph) || string.IsNullOrEmpty(sourceText)) yield break;
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraph, nativeFontMap: fontMap);
            if (fieldToken) {
                yield return new NativeHeaderFooterStyledReplacement(sourceText, CreateNativeHeaderFooterStyledTextRun(sourceText, style, 0D), true);
                yield break;
            }
            foreach (NativeCasedTextSegment segment in GetNativeCasedTextSegments(sourceText, style))
                yield return new NativeHeaderFooterStyledReplacement(segment.Text, CreateNativeHeaderFooterStyledTextRun(segment.Text, segment.Style, 0D));
        }
    }
}
