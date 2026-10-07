using System.Collections.Generic;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool HasNativeTextSpacing(NativeTextSpacing spacing) =>
            (spacing.WidthPercentage ?? 100D) != 100D || (spacing.CharacterSpacing ?? 0D) != 0D;

        private static IReadOnlyList<NativeHeaderFooterStyledReplacement>? CreateNativeHeaderFooterSpacingReplacements(
            WordParagraph paragraph, NativeFontMap? fontMap) {
            var replacements = new List<NativeHeaderFooterStyledReplacement>();
            bool hasSpacing = false;
            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (IsNativeHiddenTextRun(run, paragraph) || string.IsNullOrEmpty(run.Text)) continue;
                NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, paragraph, nativeFontMap: fontMap);
                hasSpacing |= HasNativeTextSpacing(style.TextSpacing);
                string text = ApplyNativeTextTransform(run.Text, run, paragraph, nativeFontMap: fontMap);
                replacements.Add(new NativeHeaderFooterStyledReplacement(text,
                    CreateNativeHeaderFooterStyledTextRun(text, style, 0D)));
            }
            return hasSpacing ? replacements : null;
        }
    }
}
