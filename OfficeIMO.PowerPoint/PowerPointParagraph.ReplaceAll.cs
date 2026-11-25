using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointParagraph {
        /// <summary>
        /// Replaces literal text in the paragraph's runs.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text. Pass <c>null</c> to remove each match.</param>
        /// <param name="options">Comparison settings. <c>null</c> matches ordinally, ignoring case, inside or across words.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>
        /// A match can span adjacent text runs and stops at line breaks and dynamic fields.
        /// Replacement text takes the formatting of the run where the match starts.
        /// </remarks>
        public int ReplaceAll(string text, string replacement, OfficeTextSearchOptions? options = null) {
            OfficeTextSearch.ValidateSearchText(text, nameof(text));
            StringComparison comparison = OfficeTextSearch.GetComparison(options);
            bool wholeWords = options?.WholeWords == true;
            int count = 0;

            OpenXmlElement? child = Paragraph.FirstChild;
            while (child != null) {
                if (child is not A.Run run) {
                    child = child.NextSibling();
                    continue;
                }

                List<A.Run>? runs = null;
                OpenXmlElement? next = run.NextSibling();
                while (next is A.Run adjacent) {
                    (runs ??= new List<A.Run> { run }).Add(adjacent);
                    next = adjacent.NextSibling();
                }

                // Field text affects word boundaries but remains outside the text that can be replaced.
                string? preceding = wholeWords ? (run.PreviousSibling() as A.Field)?.Text?.Text : null;
                string? following = wholeWords ? (next as A.Field)?.Text?.Text : null;
                runs ??= new List<A.Run> { run };
                string[] texts = runs.Select(item => item.Text?.Text ?? string.Empty).ToArray();
                string?[] updated = new string?[texts.Length];
                count += PowerPointRunReplace.Plan(texts, text, replacement, comparison, wholeWords, preceding, following, updated);
                // Assigning text may insert native line breaks.
                for (int index = 0; index < updated.Length; index++) {
                    if (updated[index] is { } value) new PowerPointTextRun(runs[index], _slidePart, _ownerPart).Text = value;
                }

                child = next;
            }

            return count;
        }
    }
}
