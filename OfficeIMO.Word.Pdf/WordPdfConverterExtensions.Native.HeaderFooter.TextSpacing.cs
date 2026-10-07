using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasNativeTextSpacing(NativeTextSpacing spacing) =>
            (spacing.WidthPercentage ?? 100D) != 100D || (spacing.CharacterSpacing ?? 0D) != 0D;

        private static IReadOnlyList<NativeHeaderFooterStyledReplacement>? CreateNativeHeaderFooterSpacingReplacements(
            WordParagraph paragraph, NativeFontMap? fontMap, bool preserveNaturalSpacing = false) {
            var replacements = new List<NativeHeaderFooterStyledReplacement>();
            bool hasSpacing = preserveNaturalSpacing || HasNativeTextSpacing(
                ResolveNativeTextRunStyle(paragraph, nativeFontMap: fontMap).TextSpacing);
            bool hasLiteralPageTokens = false;
            var serializedRuns = new List<(W.Run Run, string Text, bool IsField)>();
            bool hasFields = TryBuildNativeHeaderFooterParagraphText(paragraph, out _, out _, serializedRuns);
            IEnumerable<(WordParagraph Run, string Text, bool IsField)> visibleRuns = hasFields
                ? serializedRuns.Select(item => (new WordParagraph(paragraph._document, paragraph._paragraph!, item.Run), item.Text, item.IsField))
                : GetNativeRuns(paragraph).Select(run => (run, run.Text, false));
            foreach (var item in visibleRuns) {
                WordParagraph run = item.Run;
                string sourceText = item.Text;
                if (!item.IsField && run._run?.Descendants<W.TextBoxContent>().Any() == true) {
                    // Nested text boxes contribute their own runs below. Keep only
                    // the enclosing run's direct text in this replacement.
                    run = CreateNativeRunContentView(run, run._run.ChildElements
                        .Where(child => child is not W.TextBoxContent && !child.Descendants<W.TextBoxContent>().Any()).ToArray());
                    sourceText = run.Text;
                }
                if (IsNativeHiddenTextRun(run, paragraph) || string.IsNullOrEmpty(sourceText)) continue;
                NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraph, nativeFontMap: fontMap);
                hasSpacing |= HasNativeTextSpacing(style.TextSpacing);
                bool fieldToken = item.IsField;
                string text = fieldToken ? sourceText : ApplyNativeTextTransform(sourceText, run, paragraph, nativeFontMap: fontMap);
                hasLiteralPageTokens |= !fieldToken &&
                    (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     text.IndexOf("{documentpages}", StringComparison.OrdinalIgnoreCase) >= 0);
                replacements.Add(new NativeHeaderFooterStyledReplacement(text,
                    CreateNativeHeaderFooterStyledTextRun(text, style, 0D), fieldToken));
            }
            // Keep authored braces distinct from fields even when another paragraph enables styled output.
            return hasSpacing || hasLiteralPageTokens ? replacements : null;
        }
    }
}
