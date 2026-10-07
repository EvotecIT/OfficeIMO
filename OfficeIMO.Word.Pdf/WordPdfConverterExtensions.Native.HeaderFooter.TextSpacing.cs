using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasNativeTextSpacing(NativeTextSpacing spacing) =>
            (spacing.WidthPercentage ?? 100D) != 100D || (spacing.CharacterSpacing ?? 0D) != 0D;

        private static IReadOnlyList<NativeHeaderFooterStyledReplacement>? CreateNativeHeaderFooterSpacingReplacements(
            WordParagraph paragraph, NativeFontMap? fontMap) {
            var replacements = new List<NativeHeaderFooterStyledReplacement>();
            bool hasSpacing = HasNativeTextSpacing(
                ResolveNativeTextRunStyle(paragraph, nativeFontMap: fontMap).TextSpacing);
            bool hasLiteralPageTokens = false;
            foreach (var item in GetNativeHeaderFooterVisibleTextRuns(paragraph)) {
                NativeHeaderFooterStyledReplacement? replacement = CreateNativeHeaderFooterSpacingReplacement(
                    paragraph, item.Run, item.Text, item.IsField, fontMap);
                if (replacement == null) continue;
                hasSpacing |= replacement.StyledRun.HorizontalTextScaling != 100D || replacement.StyledRun.CharacterSpacing != 0D;
                string text = replacement.SerializedText;
                hasLiteralPageTokens |= !item.IsField &&
                    (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     text.IndexOf("{documentpages}", StringComparison.OrdinalIgnoreCase) >= 0);
                replacements.Add(replacement);
            }
            // Keep authored braces distinct from fields even when another paragraph enables styled output.
            return hasSpacing || hasLiteralPageTokens ? replacements : null;
        }

        private static IEnumerable<(WordParagraph Run, string Text, bool IsField)> GetNativeHeaderFooterVisibleTextRuns(WordParagraph paragraph) {
            var serializedRuns = new List<(W.Run Run, string Text, bool IsField)>();
            return TryBuildNativeHeaderFooterParagraphText(paragraph, out _, out _, serializedRuns)
                ? serializedRuns.Select(item => (new WordParagraph(paragraph._document, paragraph._paragraph!, item.Run), item.Text, item.IsField))
                : GetNativeRuns(paragraph).Select(run => (run, run.Text, false));
        }

        private static NativeHeaderFooterStyledReplacement? CreateNativeHeaderFooterSpacingReplacement(
            WordParagraph paragraph, WordParagraph run, string sourceText, bool fieldToken, NativeFontMap? fontMap) {
            if (IsNativeHiddenTextRun(run, paragraph) || string.IsNullOrEmpty(sourceText)) return null;
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraph, nativeFontMap: fontMap);
            string text = fieldToken ? sourceText : ApplyNativeTextTransform(sourceText, run, paragraph, nativeFontMap: fontMap);
            return new NativeHeaderFooterStyledReplacement(text, CreateNativeHeaderFooterStyledTextRun(text, style, 0D), fieldToken);
        }
    }
}
