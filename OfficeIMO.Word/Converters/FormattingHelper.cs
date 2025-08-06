using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helpers for working with runs and their formatting.
    /// </summary>
    public static class FormattingHelper {
        /// <summary>
        /// Represents a run of text or image with associated formatting flags.
        /// </summary>
        public readonly struct FormattedRun {
            public string? Text { get; }
            public WordImage? Image { get; }
            public bool Bold { get; }
            public bool Italic { get; }
            public bool Underline { get; }
            public string? Hyperlink { get; }

            public FormattedRun(string? text, WordImage? image, bool bold, bool italic, bool underline, string? hyperlink) {
                Text = text;
                Image = image;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Hyperlink = hyperlink;
            }
        }

        /// <summary>
        /// Enumerates runs within the paragraph and returns their text and formatting flags.
        /// </summary>
        public static IEnumerable<FormattedRun> GetFormattedRuns(WordParagraph paragraph) {
            if (paragraph == null) {
                yield break;
            }

            foreach (WordParagraph run in paragraph.GetRuns()) {
                if (run.IsImage && run.Image != null) {
                    yield return new FormattedRun(null, run.Image, false, false, false, null);
                    continue;
                }

                string? text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                string? hyperlink = run.IsHyperLink && run.Hyperlink != null ? run.Hyperlink.Uri?.ToString() : null;
                yield return new FormattedRun(text, null, run.Bold, run.Italic, run.Underline != null, hyperlink);
            }
        }
    }
}

