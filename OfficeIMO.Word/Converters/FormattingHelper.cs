using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Internal formatting projection shared by the public paragraph API and first-party adapters.</summary>
    internal static class FormattingHelper {
        // Character styles Word, Pandoc and the Markdown converters use for inline code.
        private static readonly HashSet<string> CodeCharacterStyles = new(StringComparer.OrdinalIgnoreCase) {
            "Code", "CodeChar", "InlineCode", "HTMLCode", "HtmlCode", "SourceCode", "VerbatimChar", "HTMLTypewriter"
        };

        // Monospace families by name, so the result does not depend on the fonts this machine has.
        private static readonly HashSet<string> MonospaceFamilies = new(StringComparer.OrdinalIgnoreCase) {
            "monospace", "Consolas", "Courier New", "Courier", "Cascadia Code", "Cascadia Mono", "Lucida Console",
            "Lucida Sans Typewriter", "Menlo", "Monaco", "SF Mono", "Source Code Pro", "DejaVu Sans Mono",
            "Liberation Mono", "Noto Sans Mono", "Roboto Mono", "Ubuntu Mono", "JetBrains Mono", "Fira Code",
            "Fira Mono", "IBM Plex Mono", "Inconsolata", "Andale Mono", "Cousine", "Nimbus Mono PS"
        };

        /// <summary>
        /// A run is code when it uses a code character style, or a monospace family that is not the
        /// document's own body font.
        /// </summary>
        internal static bool IsCodeRun(WordParagraph run, string? documentFont) {
            if (run.CharacterStyleId is string styleId && CodeCharacterStyles.Contains(styleId)) {
                return true;
            }

            string? family = run.FontFamily;
            if (string.IsNullOrWhiteSpace(family) || string.Equals(family, documentFont, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (MonospaceFamilies.Contains(family!.Trim())) {
                return true;
            }

            string? resolvedMonospace = WordFontResolver.Resolve("monospace");
            return !string.IsNullOrEmpty(resolvedMonospace) &&
                   string.Equals(family, resolvedMonospace, StringComparison.OrdinalIgnoreCase);
        }

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
                bool superscript = run.VerticalTextAlignment == WordVerticalTextPosition.Superscript;
                bool subscript = run.VerticalTextAlignment == WordVerticalTextPosition.Subscript;
                bool code = IsCodeRun(run, paragraph._document?.Settings.FontFamily);
                yield return new WordFormattedRun(text, null, run.Bold, run.Italic, run.Underline != null, strike, superscript, subscript, code, hyperlink);
            }
        }
    }
}
