using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                if (runs == null) {
                    int occurrences = 0;
                    string updated = OfficeTextSearch.Replace(run.Text?.Text ?? string.Empty, text,
                        replacement, comparison, wholeWords, ref occurrences, preceding, following);
                    if (occurrences > 0) {
                        new PowerPointTextRun(run, _slidePart, _ownerPart).Text = updated;
                        count += occurrences;
                    }
                } else {
                    count += ReplaceRunGroup(runs, text, replacement, comparison, wholeWords, preceding, following);
                }

                child = next;
            }

            return count;
        }

        private int ReplaceRunGroup(List<A.Run> runs, string text, string? replacement,
            StringComparison comparison, bool wholeWords, string? preceding, string? following) {
            string original = string.Concat(runs.Select(run => run.Text?.Text ?? string.Empty));
            List<int>? matches = null;
            int searchStart = 0;
            while (searchStart <= original.Length - text.Length) {
                int found = original.IndexOf(text, searchStart, comparison);
                if (found < 0) break;
                if (wholeWords && !OfficeTextSearch.HasWordBoundaries(original, found, text.Length, preceding, following)) {
                    searchStart = found + 1;
                    continue;
                }

                (matches ??= new List<int>()).Add(found);
                searchStart = found + text.Length;
            }

            if (matches == null) return 0;

            // Use original offsets and write each touched run once. Assigning text may insert native line breaks.
            int match = 0;
            int runStart = 0;
            foreach (A.Run run in runs) {
                if (match == matches.Count) break;

                string current = run.Text?.Text ?? string.Empty;
                int runEnd = runStart + current.Length;
                if (current.Length > 0 && matches[match] < runEnd) {
                    var updated = new StringBuilder(current.Length);
                    int copied = 0;
                    while (match < matches.Count && matches[match] < runEnd) {
                        int startInRun = matches[match] - runStart;
                        if (startInRun >= 0) {
                            updated.Append(current, copied, startInRun - copied);
                            updated.Append(replacement);
                        }

                        int endInRun = startInRun + text.Length;
                        copied = Math.Min(endInRun, current.Length);
                        if (endInRun > current.Length) break;
                        match++;
                    }
                    updated.Append(current, copied, current.Length - copied);
                    new PowerPointTextRun(run, _slidePart, _ownerPart).Text = updated.ToString();
                }
                runStart = runEnd;
            }

            return matches.Count;
        }
    }
}
