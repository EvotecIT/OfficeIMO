using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helpers for working with runs and their formatting.
    /// </summary>
    public static class FormattingHelper {
        /// <summary>
        /// Represents a run of text or image with associated formatting flags.
        /// </summary>
        public readonly struct FormattedRun {
            /// <summary>Text content of the run, if any.</summary>
            public string? Text { get; }
            /// <summary>Embedded image for the run, when present.</summary>
            public WordImage? Image { get; }
            /// <summary>Indicates whether bold formatting is applied.</summary>
            public bool Bold { get; }
            /// <summary>Indicates whether italic formatting is applied.</summary>
            public bool Italic { get; }
            /// <summary>Indicates whether underline formatting is applied.</summary>
            public bool Underline { get; }
            /// <summary>Indicates whether strike-through formatting is applied.</summary>
            public bool Strike { get; }
            /// <summary>Indicates whether superscript formatting is applied.</summary>
            public bool Superscript { get; }
            /// <summary>Indicates whether subscript formatting is applied.</summary>
            public bool Subscript { get; }
            /// <summary>Indicates whether the run should be rendered with monospace formatting.</summary>
            public bool Code { get; }
            /// <summary>Hyperlink target associated with the run.</summary>
            public string? Hyperlink { get; }

            /// <summary>
            /// Initializes a new instance of the <see cref="FormattedRun"/> struct.
            /// </summary>
            /// <param name="text">Text content of the run.</param>
            /// <param name="image">Embedded image for the run.</param>
            /// <param name="bold">Indicates whether bold formatting is applied.</param>
            /// <param name="italic">Indicates whether italic formatting is applied.</param>
            /// <param name="underline">Indicates whether underline formatting is applied.</param>
            /// <param name="strike">Indicates whether strike-through formatting is applied.</param>
            /// <param name="superscript">Indicates whether superscript formatting is applied.</param>
            /// <param name="subscript">Indicates whether subscript formatting is applied.</param>
            /// <param name="code">Indicates whether monospace formatting is applied.</param>
            /// <param name="hyperlink">Hyperlink target associated with the run.</param>
            public FormattedRun(string? text, WordImage? image, bool bold, bool italic, bool underline, bool strike, bool superscript, bool subscript, bool code, string? hyperlink) {
                Text = text;
                Image = image;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Strike = strike;
                Superscript = superscript;
                Subscript = subscript;
                Code = code;
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
                    yield return new FormattedRun(null, run.Image, false, false, false, false, false, false, false, null);
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
                bool code = !string.IsNullOrEmpty(monospace) && string.Equals(run.FontFamily, monospace, StringComparison.OrdinalIgnoreCase);
                yield return new FormattedRun(text, null, run.Bold, run.Italic, run.Underline != null, strike, superscript, subscript, code, hyperlink);
            }
        }
    }
}

