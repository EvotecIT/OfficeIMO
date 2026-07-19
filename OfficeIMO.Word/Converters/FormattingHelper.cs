using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Internal formatting projection shared by the public paragraph API and first-party adapters.</summary>
    internal static class FormattingHelper {
        /// <summary>
        /// Enumerates runs within the paragraph and returns their text and formatting flags.
        /// </summary>
        internal static IEnumerable<WordFormattedRun> GetFormattedRuns(WordParagraph paragraph) {
            if (paragraph == null) {
                yield break;
            }

            foreach (WordParagraph run in paragraph.GetRuns()) {
                if (run.IsImage && run.Image != null) {
                    yield return new WordFormattedRun(null, run.Image, false, false, false, false, false, false, false, null);
                    continue;
                }

                string? text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                string? hyperlink = run.IsHyperLink && run.Hyperlink != null ? run.Hyperlink.Uri?.ToString() : null;
                bool strike = run.Strike;
                bool superscript = run.VerticalTextAlignment == VerticalPositionValues.Superscript;
                bool subscript = run.VerticalTextAlignment == VerticalPositionValues.Subscript;
                string? monospace = FontResolver.Resolve("monospace");
                string? defaultFont = paragraph._document?.Settings.FontFamily;
                bool code = !string.IsNullOrEmpty(monospace) &&
                            string.Equals(run.FontFamily, monospace, StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(run.FontFamily, defaultFont, StringComparison.OrdinalIgnoreCase);
                yield return new WordFormattedRun(text, null, run.Bold, run.Italic, run.Underline != null, strike, superscript, subscript, code, hyperlink);
            }
        }
    }
}
