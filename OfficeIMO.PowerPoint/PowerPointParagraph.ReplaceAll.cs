using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

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

            OfficeXmlNode child = Element.FirstChild;
            while (!child.IsNone) {
                if (!child.Is(OfficeXmlNamespaces.DrawingML, "r")) {
                    child = child.NextSibling;
                    continue;
                }

                var runs = new List<OfficeXmlNode> { child };
                OfficeXmlNode next = child.NextSibling;
                while (next.Is(OfficeXmlNamespaces.DrawingML, "r")) {
                    runs.Add(next);
                    next = next.NextSibling;
                }

                // Field text affects word boundaries but remains outside the text that can be replaced.
                OfficeXmlNode previousSibling = child.PreviousSibling;
                string? preceding = wholeWords && previousSibling.Is(OfficeXmlNamespaces.DrawingML, "fld") ? previousSibling.ChildText(OfficeXmlNamespaces.DrawingML, "t") : null;
                string? following = wholeWords && next.Is(OfficeXmlNamespaces.DrawingML, "fld") ? next.ChildText(OfficeXmlNamespaces.DrawingML, "t") : null;
                string[] texts = runs.Select(item => item.ChildText(OfficeXmlNamespaces.DrawingML, "t") ?? string.Empty).ToArray();
                string?[] updated = new string?[texts.Length];
                count += PowerPointRunReplace.Plan(texts, text, replacement, comparison, wholeWords, preceding, following, updated);
                // Assigning text may insert native line breaks.
                for (int index = 0; index < updated.Length; index++) {
                    if (updated[index] is { } value) RunFor(runs[index]).Text = value;
                }

                child = next;
            }

            return count;
        }
    }
}
